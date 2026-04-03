import { memo, useCallback, useEffect, useRef, useState } from "react";
import type { OfficeStateUpdate } from "@pi-office/pi-office-pack/protocol";
import type { OfficeThemeSnapshot } from "../../lib/office";
import { renderMermaidToSvg, renderDrawioToSvg, svgToPngBase64, renderDrawioToPngBase64 } from "../../lib/diagram-renderer";
import { applyHostAction } from "../../lib/office";
import { openExpandViewer } from "../../lib/diagram-dialog";
import { downloadPngBlob, base64ToBlob } from "../../lib/download-utils";
import { Toast, createToast, type ToastMessage } from "./Toast";

type DiagramType = "mermaid" | "drawio";
type RenderState = "idle" | "rendering" | "done" | "error";

const renderedDiagramCache = new Map<string, string>();

interface DiagramBlockProps {
  code: string;
  type: DiagramType;
  officeTheme: OfficeThemeSnapshot | undefined;
  officeState: OfficeStateUpdate | undefined;
}

function themeKey(theme: OfficeThemeSnapshot | undefined): string {
  if (!theme) return "";
  return `${theme.isDarkTheme}|${theme.bodyBackgroundColor}|${theme.bodyForegroundColor}|${theme.controlBackgroundColor}`;
}

export const DiagramBlock = memo(function DiagramBlock({
  code,
  type,
  officeTheme,
  officeState,
}: DiagramBlockProps) {
  const [editedCode, setEditedCode] = useState<string | undefined>(undefined);
  const activeCode = editedCode ?? code;
  const [renderState, setRenderState] = useState<RenderState>("idle");
  const [svgHtml, setSvgHtml] = useState<string>("");
  const [error, setError] = useState<string>("");
  const [showSource, setShowSource] = useState(false);
  const [inserting, setInserting] = useState(false);
  const [insertStatus, setInsertStatus] = useState<"" | "success" | "error">("");
  const [toast, setToast] = useState<ToastMessage | null>(null);
  const previewRef = useRef<HTMLDivElement>(null);
  const lastRenderedKey = useRef<string>("");

  const stableThemeKey = themeKey(officeTheme);

  useEffect(() => {
    const trimmedCode = activeCode.trim();
    if (!trimmedCode) {
      lastRenderedKey.current = "";
      setSvgHtml("");
      setError("");
      setRenderState("idle");
      return;
    }

    const renderKey = `${type}|${stableThemeKey}|${trimmedCode}`;
    if (renderKey === lastRenderedKey.current) return;

    const cachedSvg = renderedDiagramCache.get(renderKey);
    if (cachedSvg) {
      lastRenderedKey.current = renderKey;
      setSvgHtml(cachedSvg);
      setError("");
      setRenderState("done");
      return;
    }

    let cancelled = false;
    setRenderState("rendering");
    setError("");

    const render = async () => {
      if (cancelled) return;
      try {
        const svg =
          type === "mermaid"
            ? await renderMermaidToSvg(trimmedCode, officeTheme)
            : await renderDrawioToSvg(trimmedCode);
        if (!cancelled) {
          if (!svg || (type === "drawio" && !svg.includes("<svg"))) {
            throw new Error("Draw.io export returned invalid SVG.");
          }
          renderedDiagramCache.set(renderKey, svg);
          lastRenderedKey.current = renderKey;
          setSvgHtml(svg);
          setRenderState("done");
        }
      } catch (err) {
        if (!cancelled) {
          setError(err instanceof Error ? err.message : String(err));
          setRenderState("error");
        }
      }
    };

    void render();
    return () => {
      cancelled = true;
    };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeCode, type, stableThemeKey]);

  const getPngBase64 = useCallback(async (opts: { maxWidth: number; scale?: number }): Promise<string> => {
    if (type === "drawio") {
      const { data } = await renderDrawioToPngBase64(activeCode.trim(), opts);
      return data;
    }
    if (!svgHtml) throw new Error("No SVG available.");
    const { data } = await svgToPngBase64(svgHtml, opts);
    return data;
  }, [type, activeCode, svgHtml]);

  const handleInsert = useCallback(async () => {
    if (!officeState || (!svgHtml && type !== "drawio")) return;
    setInserting(true);
    setInsertStatus("");
    try {
      const data = await getPngBase64({ maxWidth: 1400, scale: 3 });
      await applyHostAction(officeState.host, {
        type: "insertInlinePicture",
        content: data,
        placement: "after",
        options: {
          altText: `${type === "mermaid" ? "Mermaid" : "Draw.io"} diagram`,
        },
      });
      setInsertStatus("success");
      setTimeout(() => setInsertStatus(""), 3000);
    } catch (err) {
      setInsertStatus("error");
      console.error("[diagram-insert]", err);
    } finally {
      setInserting(false);
    }
  }, [svgHtml, officeState, type, getPngBase64]);

  const handleCopy = useCallback(async () => {
    if (!svgHtml && type !== "drawio") return;
    try {
      const data = await getPngBase64({ maxWidth: 1600, scale: 3 });
      const byteString = atob(data);
      const bytes = new Uint8Array(byteString.length);
      for (let i = 0; i < byteString.length; i++) bytes[i] = byteString.charCodeAt(i);
      const blob = new Blob([bytes], { type: "image/png" });
      await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
    } catch (err) {
      console.error("[diagram-copy]", err);
      try { await navigator.clipboard.writeText(svgHtml || activeCode); } catch { /* ignore */ }
    }
  }, [svgHtml, type, activeCode, getPngBase64]);

  const handleCopyCode = useCallback(async () => {
    try {
      await navigator.clipboard.writeText(activeCode);
      setToast(createToast("Code copied to clipboard"));
    } catch {
      setToast(createToast("Failed to copy code", "error"));
    }
  }, [activeCode]);

  const handleDownload = useCallback(async () => {
    if (!svgHtml && type !== "drawio") return;
    try {
      const data = await getPngBase64({ maxWidth: 2400, scale: 3 });
      const blob = base64ToBlob(data, "image/png");
      const suggestedName = `${type === "mermaid" ? "mermaid" : "drawio"}-diagram-${Date.now()}.png`;
      const result = await downloadPngBlob(blob, suggestedName);
      if (result.ok) {
        setToast(createToast(`Saved${result.savedAs ? `: ${result.savedAs}` : ` as ${result.filename}`}`));
      }
    } catch (err) {
      console.error("[diagram-download]", err);
      setToast(createToast("Failed to save PNG", "error"));
    }
  }, [svgHtml, type, getPngBase64]);

  const handleExpand = useCallback(() => {
    const contentType = type === "drawio" ? "drawio" as const : "svg" as const;
    const data = type === "drawio" ? activeCode.trim() : svgHtml;
    if (!data) return;
    void openExpandViewer(contentType, data, {
      isDark: officeTheme?.isDarkTheme ?? false,
      onInsertRequest: officeState
        ? async (pngBase64: string) => {
            try {
              await applyHostAction(officeState.host, {
                type: "insertInlinePicture",
                content: pngBase64,
                placement: "after",
                options: { altText: `${type === "mermaid" ? "Mermaid" : "Draw.io"} diagram` },
              });
            } catch (err) {
              console.error("[diagram-dialog-insert]", err);
            }
          }
        : undefined,
      onDataUpdate: type === "drawio"
        ? (newData: string) => setEditedCode(newData)
        : undefined,
    });
  }, [activeCode, type, svgHtml, officeTheme?.isDarkTheme, officeState]);

  const label = type === "mermaid" ? "Mermaid Diagram" : "Draw.io Diagram";

  return (
    <div className="diagram-block">
      <div className="diagram-header">
        <span className="diagram-label">{label}</span>
        <button
          type="button"
          className="diagram-source-toggle"
          onClick={() => setShowSource((s) => !s)}
        >
          {showSource ? "Hide Source" : "Show Source"}
        </button>
      </div>

      {showSource && (
        <pre className="diagram-source">
          <code>{activeCode}</code>
        </pre>
      )}

      {!svgHtml && renderState === "rendering" && (
        <div className="diagram-preview diagram-loading">
          <span className="diagram-loading-text">Rendering diagram...</span>
        </div>
      )}

      {renderState === "error" && (
        <div className="diagram-preview diagram-error-state">
          <svg viewBox="0 0 20 20" width="14" height="14" aria-hidden="true">
            <path d="M10 0C4.48 0 0 4.48 0 10s4.48 10 10 10 10-4.48 10-10S15.52 0 10 0zm1 15H9v-2h2v2zm0-4H9V5h2v6z" fill="currentColor" />
          </svg>
          <span>{error || "Failed to render diagram."}</span>
        </div>
      )}

      {renderState !== "error" && svgHtml && (
        <div
          ref={previewRef}
          className="diagram-preview"
          dangerouslySetInnerHTML={{ __html: svgHtml }}
        />
      )}

      {renderState === "done" && svgHtml && (
        <div className="diagram-actions">
          <button
            type="button"
            className="button diagram-action-btn"
            disabled={inserting || !officeState}
            onClick={handleInsert}
          >
            {inserting ? "Inserting..." : insertStatus === "success" ? "Inserted!" : "Insert into Document"}
          </button>
          <button
            type="button"
            className="button diagram-action-btn"
            onClick={handleCopy}
          >
            Copy Image
          </button>
          <button
            type="button"
            className="button diagram-action-btn"
            onClick={handleCopyCode}
          >
            Copy Code
          </button>
          <button
            type="button"
            className="button diagram-action-btn"
            onClick={handleDownload}
          >
            Download PNG
          </button>
          <button
            type="button"
            className="button diagram-action-btn diagram-expand-btn"
            onClick={handleExpand}
            title="Open expanded viewer with zoom, pan, and drag"
          >
            <svg viewBox="0 0 16 16" width="12" height="12" aria-hidden="true" style={{ marginRight: 3, verticalAlign: -1 }}>
              <path d="M3.75 2h2.5a.75.75 0 000-1.5h-4a.75.75 0 00-.75.75v4a.75.75 0 001.5 0V2.56L6.22 5.78a.75.75 0 001.06-1.06L4.06 1.5zM9.75 14h-2.5a.75.75 0 000 1.5h4a.75.75 0 00.75-.75v-4a.75.75 0 00-1.5 0v2.69L7.28 10.22a.75.75 0 00-1.06 1.06L9.44 14.5z" fill="currentColor" />
            </svg>
            Expand
          </button>
        </div>
      )}

      {insertStatus === "error" && (
        <div className="diagram-insert-error">
          Insert failed. Make sure a document is open and try again.
        </div>
      )}

      <Toast toast={toast} onDismiss={() => setToast(null)} />
    </div>
  );
});
