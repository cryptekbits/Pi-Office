import { spawnSync } from "node:child_process";
import type {
  CompanionNativeCaptureCapability,
  CompanionNativeCaptureRequest,
  CompanionNativeCaptureResponse,
  OfficeHost,
  OfficeVisualSnapshot,
} from "@pi-office/pi-office-pack/protocol";

export const NATIVE_CAPTURE_CAPABILITY_VERSION = "native-capture-v1";

export interface NativeCaptureSession {
  host: OfficeHost;
  title: string;
}

export function createNativeCaptureCapability(platform = process.platform): CompanionNativeCaptureCapability {
  if (platform !== "win32") {
    return {
      state: "unavailable",
      available: false,
      version: NATIVE_CAPTURE_CAPABILITY_VERSION,
      platform,
      hosts: [],
      trueViewportScreenshot: false,
      includeWindowFrame: false,
      reason: "Native viewport/window capture is currently implemented for Windows only.",
    };
  }

  return {
    state: "available",
    available: true,
    version: NATIVE_CAPTURE_CAPABILITY_VERSION,
    platform,
    hosts: ["word", "excel"],
    trueViewportScreenshot: true,
    includeWindowFrame: true,
    reason: "Captures the current foreground desktop window through the Windows native screen API.",
  };
}

function parsePowerShellCapture(stdout: string): { data: string; width: number; height: number; left: number; top: number } {
  const trimmed = stdout.trim();
  const parsed = JSON.parse(trimmed) as Partial<{
    data: string;
    width: number;
    height: number;
    left: number;
    top: number;
  }>;
  if (!parsed.data || typeof parsed.data !== "string") {
    throw new Error("Native capture did not return image data.");
  }
  if (!Number.isFinite(parsed.width) || !Number.isFinite(parsed.height)) {
    throw new Error("Native capture did not return image dimensions.");
  }
  const width = Number(parsed.width);
  const height = Number(parsed.height);
  const left = Number(parsed.left);
  const top = Number(parsed.top);
  return {
    data: parsed.data,
    width: Math.trunc(width),
    height: Math.trunc(height),
    left: Number.isFinite(left) ? Math.trunc(left) : 0,
    top: Number.isFinite(top) ? Math.trunc(top) : 0,
  };
}

function windowsForegroundWindowCapture(): { data: string; width: number; height: number; left: number; top: number } {
  const script = `
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class PiOfficeNativeCapture {
  [DllImport("user32.dll")]
  public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")]
  public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
  [StructLayout(LayoutKind.Sequential)]
  public struct RECT {
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
  }
}
"@
$handle = [PiOfficeNativeCapture]::GetForegroundWindow()
$rect = New-Object PiOfficeNativeCapture+RECT
[PiOfficeNativeCapture]::GetWindowRect($handle, [ref]$rect) | Out-Null
$width = [Math]::Max(1, $rect.Right - $rect.Left)
$height = [Math]::Max(1, $rect.Bottom - $rect.Top)
$bitmap = New-Object System.Drawing.Bitmap($width, $height)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
try {
  $graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, $bitmap.Size)
  $stream = New-Object System.IO.MemoryStream
  try {
    $bitmap.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
    [Console]::Out.Write((ConvertTo-Json -Compress @{
      data = [Convert]::ToBase64String($stream.ToArray())
      width = $width
      height = $height
      left = $rect.Left
      top = $rect.Top
    }))
  } finally {
    $stream.Dispose()
  }
} finally {
  $graphics.Dispose()
  $bitmap.Dispose()
}
`;

  const result = spawnSync("powershell.exe", [
    "-NoProfile",
    "-NonInteractive",
    "-ExecutionPolicy",
    "Bypass",
    "-Command",
    script,
  ], {
    encoding: "utf8",
    maxBuffer: 24 * 1024 * 1024,
    timeout: 10_000,
    windowsHide: true,
  });

  if (result.error) {
    throw result.error;
  }
  if (result.status !== 0) {
    throw new Error(result.stderr.trim() || `PowerShell exited with status ${result.status}.`);
  }
  return parsePowerShellCapture(result.stdout);
}

export function captureNativeViewport(
  session: NativeCaptureSession,
  request: CompanionNativeCaptureRequest = {},
  platform = process.platform,
): CompanionNativeCaptureResponse {
  const host = request.host ?? session.host;
  const capability = createNativeCaptureCapability(platform);
  if (capability.state !== "available" || !capability.hosts.includes(host)) {
    return {
      ok: false,
      error: capability.reason ?? `Native capture is unavailable for ${host}.`,
      details: { capability },
    };
  }

  try {
    const captured = windowsForegroundWindowCapture();
    const visual: OfficeVisualSnapshot = {
      kind: request.includeWindowFrame === false ? "viewport" : "window",
      label: `${host} native ${request.includeWindowFrame === false ? "viewport" : "window"} capture`,
      data: captured.data,
      mimeType: "image/png",
      width: captured.width,
      height: captured.height,
    };
    return {
      ok: true,
      visual,
      details: {
        captureMode: "windows-foreground-window",
        host,
        title: session.title,
        includeWindowFrameRequested: request.includeWindowFrame !== false,
        includeWindowFrameCaptured: true,
        bounds: {
          left: captured.left,
          top: captured.top,
          width: captured.width,
          height: captured.height,
        },
        note: "Companion native capture uses the current foreground desktop window; bring the target Office window to the front before retrying if the frame is wrong.",
      },
    };
  } catch (error) {
    return {
      ok: false,
      error: error instanceof Error ? error.message : String(error),
      details: {
        captureMode: "windows-foreground-window",
        host,
      },
    };
  }
}
