import mermaid from "mermaid";
import type { OfficeThemeSnapshot } from "./office";

let mermaidInitialized = false;
let renderCounter = 0;
const DRAWIO_EMBED_ORIGIN = "https://embed.diagrams.net";
const DRAWIO_EMBED_URL =
  `${DRAWIO_EMBED_ORIGIN}/?embed=1&proto=json&spin=0&modified=0&libraries=0&noSaveBtn=1&noExitBtn=1`;
let drawioFrame: HTMLIFrameElement | undefined;
let drawioFramePromise: Promise<HTMLIFrameElement> | undefined;
let drawioReadyPromise: Promise<void> | undefined;
let resolveDrawioReady: (() => void) | undefined;
let drawioListenerAttached = false;
let drawioActiveRequest:
  | {
      resolve: (svg: string) => void;
      reject: (error: Error) => void;
      timeoutId: number;
    }
  | undefined;
const drawioQueue: Array<{
  xml: string;
  resolve: (svg: string) => void;
  reject: (error: Error) => void;
}> = [];

function getMermaidTheme(theme: OfficeThemeSnapshot | undefined): {
  theme: "neutral" | "dark";
  themeVariables: Record<string, string>;
} {
  const isDark = theme?.isDarkTheme ?? false;
  const bg = theme?.bodyBackgroundColor ?? (isDark ? "#1e1e1e" : "#ffffff");
  const fg = theme?.bodyForegroundColor ?? (isDark ? "#d4d4d4" : "#201F1E");
  const ctrlBg = theme?.controlBackgroundColor ?? (isDark ? "#2d2d2d" : "#ffffff");

  return {
    theme: isDark ? "dark" : "neutral",
    themeVariables: {
      primaryColor: isDark ? "#3a3a5c" : "#e8edf3",
      primaryTextColor: fg,
      primaryBorderColor: isDark ? "#555577" : "#99aabb",
      lineColor: isDark ? "#888899" : "#667788",
      secondaryColor: isDark ? "#2d3a4a" : "#f0f4f8",
      tertiaryColor: isDark ? "#3a2d2d" : "#fdf6f0",
      background: bg,
      mainBkg: isDark ? "#2d3350" : "#e8edf3",
      nodeBorder: isDark ? "#6670aa" : "#8899bb",
      clusterBkg: isDark ? "#2a2a3a" : "#f5f7fa",
      titleColor: fg,
      edgeLabelBackground: ctrlBg,
      nodeTextColor: fg,
    },
  };
}

function initMermaid(theme: OfficeThemeSnapshot | undefined): void {
  const config = getMermaidTheme(theme);
  mermaid.initialize({
    startOnLoad: false,
    securityLevel: "strict",
    theme: config.theme,
    themeVariables: config.themeVariables,
    fontFamily: '"Aptos", "Segoe UI", sans-serif',
    flowchart: { useMaxWidth: true, htmlLabels: true, curve: "basis" },
    sequence: { useMaxWidth: true, actorMargin: 50, mirrorActors: false },
    gantt: { useMaxWidth: true },
  });
  mermaidInitialized = true;
}

export async function renderMermaidToSvg(
  code: string,
  theme: OfficeThemeSnapshot | undefined,
): Promise<string> {
  initMermaid(theme);
  renderCounter += 1;
  const id = `pi-mermaid-${renderCounter}-${Date.now()}`;
  const { svg } = await mermaid.render(id, code.trim());
  return svg;
}

function resetDrawioEmbed(): void {
  if (drawioFrame) {
    drawioFrame.remove();
    drawioFrame = undefined;
  }
  drawioFramePromise = undefined;
  drawioReadyPromise = undefined;
  resolveDrawioReady = undefined;
  drawioActiveRequest = undefined;
}

function decodeDrawioDataUri(data: string): string {
  const match = data.match(/^data:[^,]*?(;base64)?,([\s\S]*)$/);
  if (!match) return data;
  const payload = match[2] ?? "";
  if (match[1]) {
    const binary = atob(payload);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return new TextDecoder().decode(bytes);
  }
  return decodeURIComponent(payload);
}

function ensureDrawioListener(): void {
  if (drawioListenerAttached) return;
  window.addEventListener("message", (event) => {
    if (event.origin !== DRAWIO_EMBED_ORIGIN || typeof event.data !== "string") return;
    if (drawioFrame && event.source !== drawioFrame.contentWindow) return;

    let message: Record<string, unknown>;
    try {
      message = JSON.parse(event.data);
    } catch {
      return;
    }

    const eventName = String(message.event ?? "");
    if (eventName === "init") {
      resolveDrawioReady?.();
      resolveDrawioReady = undefined;
      return;
    }

    if (eventName === "export" && drawioActiveRequest) {
      const format = String(message.format ?? "svg");
      if (format !== "svg") return;
      const activeRequest = drawioActiveRequest;
      drawioActiveRequest = undefined;
      window.clearTimeout(activeRequest.timeoutId);
      try {
        activeRequest.resolve(decodeDrawioDataUri(String(message.data ?? "")));
      } catch (error) {
        activeRequest.reject(error instanceof Error ? error : new Error(String(error)));
      }
      void flushDrawioQueue();
      return;
    }

    if (eventName === "error" && drawioActiveRequest) {
      const activeRequest = drawioActiveRequest;
      drawioActiveRequest = undefined;
      window.clearTimeout(activeRequest.timeoutId);
      activeRequest.reject(new Error(String(message.message ?? "Draw.io export failed.")));
      resetDrawioEmbed();
      void flushDrawioQueue();
    }
  });
  drawioListenerAttached = true;
}

function ensureDrawioFrame(): Promise<HTMLIFrameElement> {
  if (drawioFramePromise) return drawioFramePromise;
  ensureDrawioListener();
  drawioReadyPromise = new Promise<void>((resolve) => {
    resolveDrawioReady = resolve;
  });
  drawioFramePromise = new Promise<HTMLIFrameElement>((resolve, reject) => {
    const iframe = document.createElement("iframe");
    iframe.src = DRAWIO_EMBED_URL;
    iframe.tabIndex = -1;
    iframe.setAttribute("aria-hidden", "true");
    iframe.style.cssText =
      "position:fixed;left:-99999px;top:0;width:1px;height:1px;border:0;opacity:0;pointer-events:none;";
    iframe.onload = () => {
      drawioFrame = iframe;
      resolve(iframe);
    };
    iframe.onerror = () => {
      resetDrawioEmbed();
      reject(new Error("Failed to load diagrams.net embed frame."));
    };
    document.body.appendChild(iframe);
  });
  return drawioFramePromise;
}

async function flushDrawioQueue(): Promise<void> {
  if (drawioActiveRequest || drawioQueue.length === 0) return;

  const nextRequest = drawioQueue.shift();
  if (!nextRequest) return;

  try {
    const iframe = await ensureDrawioFrame();
    await drawioReadyPromise;
    if (!iframe.contentWindow) {
      nextRequest.reject(new Error("Diagrams.net embed frame is unavailable."));
      resetDrawioEmbed();
      return;
    }

    const timeoutId = window.setTimeout(() => {
      if (!drawioActiveRequest) return;
      const activeRequest = drawioActiveRequest;
      drawioActiveRequest = undefined;
      resetDrawioEmbed();
      activeRequest.reject(new Error("Draw.io export timed out."));
      void flushDrawioQueue();
    }, 30000);

    drawioActiveRequest = {
      resolve: nextRequest.resolve,
      reject: nextRequest.reject,
      timeoutId,
    };

    // Load the XML first (handles all formats including <mxfile> with
    // compressed diagrams), then request SVG export of the loaded diagram.
    iframe.contentWindow.postMessage(
      JSON.stringify({ action: "load", xml: nextRequest.xml }),
      DRAWIO_EMBED_ORIGIN,
    );
    iframe.contentWindow.postMessage(
      JSON.stringify({ action: "export", format: "svg" }),
      DRAWIO_EMBED_ORIGIN,
    );
  } catch (error) {
    resetDrawioEmbed();
    nextRequest.reject(error instanceof Error ? error : new Error(String(error)));
    void flushDrawioQueue();
  }
}

export async function renderDrawioToSvg(xml: string): Promise<string> {
  return new Promise<string>((resolve, reject) => {
    drawioQueue.push({
      xml,
      resolve,
      reject,
    });
    void flushDrawioQueue();
  });
}

function parseSvgDimensions(svgEl: SVGSVGElement): { width: number; height: number } {
  const rawW = svgEl.getAttribute("width") ?? "0";
  const rawH = svgEl.getAttribute("height") ?? "0";

  let w = rawW.includes("%") ? 0 : parseFloat(rawW);
  let h = rawH.includes("%") ? 0 : parseFloat(rawH);

  if ((!w || !h) && svgEl.viewBox.baseVal) {
    const vb = svgEl.viewBox.baseVal;
    if (!w) w = vb.width || 0;
    if (!h) h = vb.height || 0;
  }

  if (!w || !h) {
    try {
      const tmpContainer = document.createElement("div");
      tmpContainer.style.cssText = "position:absolute;left:-99999px;top:0;width:2000px;visibility:hidden;";
      tmpContainer.appendChild(svgEl.cloneNode(true));
      document.body.appendChild(tmpContainer);
      const svgClone = tmpContainer.querySelector("svg")!;
      const bbox = svgClone.getBoundingClientRect();
      if (bbox.width > 1) w = bbox.width;
      if (bbox.height > 1) h = bbox.height;
      tmpContainer.remove();
    } catch { /* ignore */ }
  }

  return { width: w || 800, height: h || 600 };
}

export async function svgToPngBase64(
  svgString: string,
  options?: { maxWidth?: number; scale?: number },
): Promise<{ data: string; width: number; height: number }> {
  const parser = new DOMParser();
  const doc = parser.parseFromString(svgString, "image/svg+xml");
  const svgEl = doc.querySelector("svg");
  if (!svgEl) throw new Error("Invalid SVG string.");

  const { width: svgWidth, height: svgHeight } = parseSvgDimensions(svgEl);

  const maxWidth = options?.maxWidth ?? 1200;
  const scaleFactor = options?.scale ?? Math.min(3, window.devicePixelRatio || 2);
  const displayScale = Math.min(1, maxWidth / svgWidth);
  const renderWidth = Math.round(svgWidth * displayScale * scaleFactor);
  const renderHeight = Math.round(svgHeight * displayScale * scaleFactor);
  const outputWidth = Math.round(svgWidth * displayScale);
  const outputHeight = Math.round(svgHeight * displayScale);

  svgEl.setAttribute("width", String(svgWidth));
  svgEl.setAttribute("height", String(svgHeight));
  svgEl.removeAttribute("style");
  const serialized = new XMLSerializer().serializeToString(svgEl);
  const blob = new Blob([serialized], { type: "image/svg+xml;charset=utf-8" });
  const url = URL.createObjectURL(blob);

  try {
    const image = await new Promise<HTMLImageElement>((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = () => reject(new Error("Failed to load SVG as image."));
      img.src = url;
    });

    const canvas = document.createElement("canvas");
    canvas.width = renderWidth;
    canvas.height = renderHeight;
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Canvas 2D context unavailable.");

    ctx.fillStyle = "#FFFFFF";
    ctx.fillRect(0, 0, renderWidth, renderHeight);
    ctx.scale(scaleFactor * displayScale, scaleFactor * displayScale);
    ctx.drawImage(image, 0, 0, svgWidth, svgHeight);

    const dataUrl = canvas.toDataURL("image/png");
    const base64 = dataUrl.replace(/^data:image\/png;base64,/, "");
    return { data: base64, width: outputWidth, height: outputHeight };
  } finally {
    URL.revokeObjectURL(url);
  }
}

export async function renderDrawioToPngBase64(
  xml: string,
  options?: { maxWidth?: number; scale?: number },
): Promise<{ data: string; width: number; height: number }> {
  const scale = options?.scale ?? Math.min(3, window.devicePixelRatio || 2);

  // Wait for any queued SVG exports to finish before sending a PNG export,
  // so both don't overlap on the same embed frame.
  await new Promise<void>((resolve) => {
    const check = () => {
      if (!drawioActiveRequest && drawioQueue.length === 0) { resolve(); return; }
      setTimeout(check, 50);
    };
    check();
  });

  return new Promise<{ data: string; width: number; height: number }>((resolve, reject) => {
    let pngHandler: ((event: MessageEvent) => void) | undefined;

    const cleanup = () => {
      if (pngHandler) window.removeEventListener("message", pngHandler);
    };

    const timeoutId = window.setTimeout(() => {
      cleanup();
      reject(new Error("Draw.io PNG export timed out."));
    }, 30000);

    const doExport = async () => {
      const iframe = await ensureDrawioFrame();
      await drawioReadyPromise;
      if (!iframe.contentWindow) {
        cleanup();
        clearTimeout(timeoutId);
        reject(new Error("Diagrams.net embed frame is unavailable."));
        return;
      }

      pngHandler = (event: MessageEvent) => {
        if (event.origin !== DRAWIO_EMBED_ORIGIN || typeof event.data !== "string") return;
        if (drawioFrame && event.source !== drawioFrame.contentWindow) return;
        let message: Record<string, unknown>;
        try { message = JSON.parse(event.data); } catch { return; }

        if (String(message.event) !== "export") return;
        const format = String(message.format ?? "");
        if (format !== "png") return;

        cleanup();
        clearTimeout(timeoutId);

        try {
          const dataUri = String(message.data ?? "");
          const base64 = dataUri.replace(/^data:image\/png;base64,/, "");
          const binary = atob(base64);
          const bytes = new Uint8Array(binary.length);
          for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);

          const maxWidth = options?.maxWidth ?? 1200;
          const img = new Image();
          img.onload = () => {
            const w = img.naturalWidth;
            const h = img.naturalHeight;
            resolve({ data: base64, width: Math.round(w / scale), height: Math.round(h / scale) });
          };
          img.onerror = () => resolve({ data: base64, width: maxWidth, height: 800 });
          img.src = dataUri;
        } catch (error) {
          reject(error instanceof Error ? error : new Error(String(error)));
        }
      };

      window.addEventListener("message", pngHandler);

      // Load first (handles all formats including <mxfile>), then export.
      iframe.contentWindow.postMessage(
        JSON.stringify({ action: "load", xml }),
        DRAWIO_EMBED_ORIGIN,
      );
      iframe.contentWindow.postMessage(
        JSON.stringify({
          action: "export",
          format: "png",
          scale,
          background: "#FFFFFF",
        }),
        DRAWIO_EMBED_ORIGIN,
      );
    };

    if (!drawioListenerAttached) ensureDrawioListener();
    void doExport();
  });
}
