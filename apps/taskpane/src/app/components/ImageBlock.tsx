import { memo, useCallback, useState } from "react";
import type { OfficeStateUpdate } from "@pi-office/pi-office-pack/protocol";
import { applyHostAction } from "../../lib/office";
import { openExpandViewer } from "../../lib/diagram-dialog";
import { downloadPngBlob, base64ToBlob } from "../../lib/download-utils";
import { Toast, createToast, type ToastMessage } from "./Toast";

interface ImageBlockProps {
  base64: string;
  mimeType: string;
  width: number;
  height: number;
  modelName: string;
  prompt: string;
  inserted: boolean;
  officeState: OfficeStateUpdate | undefined;
}

export const ImageBlock = memo(function ImageBlock({
  base64,
  mimeType,
  width,
  height,
  modelName,
  prompt,
  inserted,
  officeState,
}: ImageBlockProps) {
  const [inserting, setInserting] = useState(false);
  const [insertStatus, setInsertStatus] = useState<"" | "success" | "error">("");
  const [toast, setToast] = useState<ToastMessage | null>(null);

  const handleInsert = useCallback(async () => {
    if (!officeState) return;
    setInserting(true);
    setInsertStatus("");
    try {
      await applyHostAction(officeState.host, {
        type: "insertInlinePicture",
        content: base64,
        placement: "after",
        options: { altText: prompt.slice(0, 120) },
      });
      setInsertStatus("success");
      setTimeout(() => setInsertStatus(""), 3000);
    } catch {
      setInsertStatus("error");
    } finally {
      setInserting(false);
    }
  }, [base64, prompt, officeState]);

  const handleCopy = useCallback(async () => {
    try {
      const byteString = atob(base64);
      const bytes = new Uint8Array(byteString.length);
      for (let i = 0; i < byteString.length; i++) bytes[i] = byteString.charCodeAt(i);
      const blob = new Blob([bytes], { type: mimeType });
      await navigator.clipboard.write([new ClipboardItem({ [mimeType]: blob })]);
    } catch {
      // silent fallback
    }
  }, [base64, mimeType]);

  const handleDownload = useCallback(async () => {
    try {
      const blob = base64ToBlob(base64, mimeType);
      const suggestedName = `generated-image-${Date.now()}.png`;
      const result = await downloadPngBlob(blob, suggestedName);
      if (result.ok) {
        setToast(createToast(`Saved${result.savedAs ? `: ${result.savedAs}` : ` as ${result.filename}`}`));
      }
    } catch {
      setToast(createToast("Failed to save image", "error"));
    }
  }, [base64, mimeType]);

  const handleExpand = useCallback(() => {
    if (!base64) return;
    void openExpandViewer("image", base64, {
      isDark: window.matchMedia?.("(prefers-color-scheme: dark)")?.matches ?? false,
      onInsertRequest: officeState
        ? async (pngBase64: string) => {
            try {
              await applyHostAction(officeState.host, {
                type: "insertInlinePicture",
                content: pngBase64,
                placement: "after",
                options: { altText: prompt.slice(0, 120) },
              });
            } catch {
              // silent
            }
          }
        : undefined,
    });
  }, [base64, prompt, officeState]);

  const dataUrl = `data:${mimeType};base64,${base64}`;

  return (
    <div className="diagram-block">
      <div className="diagram-header">
        <span className="diagram-label">Generated Image</span>
        <span className="diagram-source-toggle" style={{ cursor: "default" }}>
          {modelName} &middot; {width}&times;{height}
        </span>
      </div>

      <div className="diagram-preview">
        <img
          src={dataUrl}
          alt={prompt.slice(0, 120)}
          style={{ maxWidth: "100%", height: "auto", borderRadius: "4px" }}
        />
      </div>

      <div className="diagram-actions">
        {!inserted && (
          <button
            type="button"
            className="button diagram-action-btn"
            disabled={inserting || !officeState}
            onClick={handleInsert}
          >
            {inserting ? "Inserting..." : insertStatus === "success" ? "Inserted!" : "Insert into Document"}
          </button>
        )}
        {inserted && (
          <button
            type="button"
            className="button diagram-action-btn"
            disabled={inserting || !officeState}
            onClick={handleInsert}
          >
            {inserting ? "Inserting..." : insertStatus === "success" ? "Inserted!" : "Insert Again"}
          </button>
        )}
        <button type="button" className="button diagram-action-btn" onClick={handleCopy}>
          Copy Image
        </button>
        <button type="button" className="button diagram-action-btn" onClick={handleDownload}>
          Download PNG
        </button>
        <button
          type="button"
          className="button diagram-action-btn diagram-expand-btn"
          onClick={handleExpand}
          title="Open expanded view with zoom, pan, and drag"
        >
          <svg viewBox="0 0 16 16" width="12" height="12" aria-hidden="true" style={{ marginRight: 3, verticalAlign: -1 }}>
            <path d="M3.75 2h2.5a.75.75 0 000-1.5h-4a.75.75 0 00-.75.75v4a.75.75 0 001.5 0V2.56L6.22 5.78a.75.75 0 001.06-1.06L4.06 1.5zM9.75 14h-2.5a.75.75 0 000 1.5h4a.75.75 0 00.75-.75v-4a.75.75 0 00-1.5 0v2.69L7.28 10.22a.75.75 0 00-1.06 1.06L9.44 14.5z" fill="currentColor" />
          </svg>
          Expand
        </button>
      </div>

      {insertStatus === "error" && (
        <div className="diagram-insert-error">
          Insert failed. Make sure a document is open and try again.
        </div>
      )}

      <Toast toast={toast} onDismiss={() => setToast(null)} />
    </div>
  );
});
