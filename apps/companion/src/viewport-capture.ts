import { constants as fsConstants } from "node:fs";
import { access, mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import { spawn } from "node:child_process";
import { join } from "node:path";
import type { OfficeVisualSnapshot } from "@pi-office/pi-office-pack";
import type { CompanionConfig } from "./config.js";

const PYTHON_RUNTIME_MARKER = "pi-office-viewport-runtime.json";
const PYTHON_RUNTIME_PACKAGES = ["zbl==0.7.1", "pillow==12.2.0"] as const;

interface NativeRect {
  left: number;
  top: number;
  right: number;
  bottom: number;
  width: number;
  height: number;
}

interface NativeWindowDescriptor {
  handle: string;
  className: string;
  title?: string | undefined;
  rect: NativeRect;
}

interface NativeFrameDescriptor {
  width: number;
  height: number;
  scaleX: number;
  scaleY: number;
}

interface NativeViewportCaptureResponse {
  captureHandle: string;
  matchedBy?: string | undefined;
  window: NativeWindowDescriptor;
  viewport: NativeWindowDescriptor;
  frame: NativeFrameDescriptor;
  viewportImagePath: string;
  windowImagePath?: string | undefined;
}

export interface WordViewportCaptureResult {
  visuals: OfficeVisualSnapshot[];
  metadata: Record<string, unknown>;
}

interface CaptureWordViewportOptions {
  documentTitle?: string | undefined;
  windowCaption?: string | undefined;
  includeWindowFrame?: boolean | undefined;
}

let runtimeReadyPromise: Promise<string> | undefined;

function getRuntimeDir(config: CompanionConfig): string {
  return join(config.repoRoot, ".pi-office", "viewport-python");
}

function getRuntimePythonPath(config: CompanionConfig): string {
  return join(getRuntimeDir(config), "Scripts", "python.exe");
}

function getRuntimeMarkerPath(config: CompanionConfig): string {
  return join(getRuntimeDir(config), PYTHON_RUNTIME_MARKER);
}

function getHelperScriptPath(config: CompanionConfig): string {
  return join(config.repoRoot, "apps", "companion", "runtime", "word_viewport_capture.py");
}

async function readRuntimeMarker(config: CompanionConfig): Promise<boolean> {
  try {
    const markerPath = getRuntimeMarkerPath(config);
    const raw = await readFile(markerPath, "utf8");
    const parsed = JSON.parse(raw) as { packages?: string[] | undefined };
    return JSON.stringify(parsed.packages ?? []) === JSON.stringify(PYTHON_RUNTIME_PACKAGES);
  } catch {
    return false;
  }
}

async function runProcess(
  command: string,
  args: string[],
  options?: {
    cwd?: string | undefined;
    input?: string | undefined;
    timeoutMs?: number | undefined;
  },
): Promise<{ stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: options?.cwd,
      stdio: ["pipe", "pipe", "pipe"],
      windowsHide: true,
    });

    let stdout = "";
    let stderr = "";
    const timeout = setTimeout(() => {
      child.kill();
      reject(new Error(`Timed out running ${command}.`));
    }, options?.timeoutMs ?? 60_000);

    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stdout.on("data", (chunk) => {
      stdout += chunk;
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk;
    });
    child.on("error", (error) => {
      clearTimeout(timeout);
      reject(error);
    });
    child.on("close", (code) => {
      clearTimeout(timeout);
      if (code !== 0) {
        reject(new Error(stderr.trim() || stdout.trim() || `${command} exited with code ${code}.`));
        return;
      }
      resolve({ stdout, stderr });
    });

    if (options?.input) {
      child.stdin.write(options.input);
    }
    child.stdin.end();
  });
}

async function ensureViewportPythonRuntime(config: CompanionConfig): Promise<string> {
  if (runtimeReadyPromise) {
    return runtimeReadyPromise;
  }

  runtimeReadyPromise = (async () => {
    const runtimeDir = getRuntimeDir(config);
    const pythonPath = getRuntimePythonPath(config);
    const markerPath = getRuntimeMarkerPath(config);

    await mkdir(runtimeDir, { recursive: true });

    try {
      await access(pythonPath, fsConstants.F_OK);
    } catch {
      await runProcess("python", ["-m", "venv", runtimeDir], {
        cwd: config.repoRoot,
        timeoutMs: 120_000,
      });
    }

    const markerIsCurrent = await readRuntimeMarker(config);
    if (!markerIsCurrent) {
      await runProcess(
        pythonPath,
        ["-m", "pip", "install", "--disable-pip-version-check", ...PYTHON_RUNTIME_PACKAGES],
        {
          cwd: config.repoRoot,
          timeoutMs: 240_000,
        },
      );
      await writeFile(
        markerPath,
        JSON.stringify(
          {
            packages: [...PYTHON_RUNTIME_PACKAGES],
          },
          null,
          2,
        ),
        "utf8",
      );
    }

    return pythonPath;
  })().catch((error) => {
    runtimeReadyPromise = undefined;
    throw error;
  });

  return runtimeReadyPromise;
}

async function fileToVisual(
  imagePath: string,
  kind: OfficeVisualSnapshot["kind"],
  label: string,
  width: number,
  height: number,
): Promise<OfficeVisualSnapshot> {
  const bytes = await readFile(imagePath);
  return {
    kind,
    label,
    data: bytes.toString("base64"),
    mimeType: "image/png",
    width,
    height,
  };
}

export async function captureWordViewport(
  config: CompanionConfig,
  options: CaptureWordViewportOptions,
): Promise<WordViewportCaptureResult> {
  if (process.platform !== "win32") {
    throw new Error("office_capture_viewport is only supported on Windows desktop.");
  }

  const pythonPath = await ensureViewportPythonRuntime(config);
  const helperPath = getHelperScriptPath(config);
  const capturesDir = join(config.scratchDir, "viewport-captures");
  await mkdir(capturesDir, { recursive: true });
  const outputDir = await mkdtemp(join(capturesDir, "capture-"));

  const { stdout } = await runProcess(
    pythonPath,
    [helperPath],
    {
      cwd: config.repoRoot,
      input: JSON.stringify({
        documentTitle: options.documentTitle,
        windowCaption: options.windowCaption,
        includeWindowFrame: options.includeWindowFrame === true,
        outputDir,
      }),
      timeoutMs: 45_000,
    },
  );

  const parsed = JSON.parse(stdout.trim()) as NativeViewportCaptureResponse;
  const visuals: OfficeVisualSnapshot[] = [
    await fileToVisual(
      parsed.viewportImagePath,
      "viewport",
      "Current Word viewport",
      parsed.viewport.rect.width,
      parsed.viewport.rect.height,
    ),
  ];

  if (options.includeWindowFrame && parsed.windowImagePath) {
    visuals.push(
      await fileToVisual(
        parsed.windowImagePath,
        "window",
        "Word window capture",
        parsed.frame.width,
        parsed.frame.height,
      ),
    );
  }

  return {
    visuals,
    metadata: {
      matchedBy: parsed.matchedBy,
      captureHandle: parsed.captureHandle,
      window: parsed.window,
      viewport: parsed.viewport,
      frame: parsed.frame,
    },
  };
}
