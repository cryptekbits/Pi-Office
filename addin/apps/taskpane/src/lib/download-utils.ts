export interface DownloadResult {
  ok: boolean;
  filename: string;
  savedAs?: string;
}

export async function downloadBlob(
  blob: Blob,
  suggestedName: string,
  types?: Array<{ description: string; accept: Record<string, string[]> }>,
): Promise<DownloadResult> {
  if (typeof window.showSaveFilePicker === "function") {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName,
        ...(types ? { types } : {}),
      });
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
      return { ok: true, filename: handle.name, savedAs: handle.name };
    } catch (err: unknown) {
      if (err instanceof DOMException && err.name === "AbortError") {
        return { ok: false, filename: suggestedName };
      }
    }
  }

  try {
    const blobUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = blobUrl;
    link.download = suggestedName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(blobUrl), 500);
    return { ok: true, filename: suggestedName };
  } catch {
    return { ok: false, filename: suggestedName };
  }
}

export async function downloadPngBlob(
  blob: Blob,
  suggestedName: string,
): Promise<DownloadResult> {
  return downloadBlob(blob, suggestedName, [
    {
      description: "PNG Image",
      accept: { "image/png": [".png"] },
    },
  ]);
}

export async function downloadJsonBlob(
  value: unknown,
  suggestedName: string,
): Promise<DownloadResult> {
  const blob = new Blob([JSON.stringify(value, null, 2)], { type: "application/json" });
  return downloadBlob(blob, suggestedName, [
    {
      description: "JSON",
      accept: { "application/json": [".json"] },
    },
  ]);
}

export function base64ToBlob(base64: string, mimeType: string): Blob {
  const byteString = atob(base64);
  const bytes = new Uint8Array(byteString.length);
  for (let i = 0; i < byteString.length; i++) bytes[i] = byteString.charCodeAt(i);
  return new Blob([bytes], { type: mimeType });
}
