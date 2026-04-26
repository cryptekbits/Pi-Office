import { useRef, useEffect } from "react";
import type { OfficeHost } from "@pi-office/pi-office-pack/protocol";

export type RewindMode = "conversation" | "conversation_and_document" | "cancel";

interface RewindDialogProps {
  messageText: string;
  host: OfficeHost | undefined;
  onSelect: (mode: RewindMode) => void;
  onClose: () => void;
}

function fidelityNote(host: OfficeHost | undefined): string | null {
  if (host === "excel") {
    return "Excel: cell values and formulas will be restored. Formatting, charts, and pivot tables may not be fully restored.";
  }
  if (host === "powerpoint") {
    return "PowerPoint: slides will be fully replaced from snapshot.";
  }
  if (host === "word") {
    return "Word: some list numbering or styles may shift slightly on restore.";
  }
  return null;
}

export function RewindDialog({ messageText, host, onSelect, onClose }: RewindDialogProps) {
  const ref = useRef<HTMLDivElement>(null);
  const closeFn = useRef(onClose);
  closeFn.current = onClose;

  useEffect(() => {
    let removed = false;
    function handleClick(e: MouseEvent) {
      if (ref.current && !ref.current.contains(e.target as Node)) closeFn.current();
    }
    function handleKey(e: KeyboardEvent) {
      if (e.key === "Escape") closeFn.current();
    }
    const handle = requestAnimationFrame(() => {
      if (removed) return;
      document.addEventListener("mousedown", handleClick);
      document.addEventListener("keydown", handleKey);
    });
    return () => {
      removed = true;
      cancelAnimationFrame(handle);
      document.removeEventListener("mousedown", handleClick);
      document.removeEventListener("keydown", handleKey);
    };
  }, []);

  const truncatedPrompt = messageText.length > 60 ? messageText.slice(0, 57) + "..." : messageText;
  const note = fidelityNote(host);

  return (
    <div ref={ref} className="rewind-dialog" role="dialog" aria-label="Rewind options">
      <div className="rewind-dialog-header">Rewind to before this message?</div>
      <div className="rewind-dialog-prompt">{truncatedPrompt}</div>
      <div className="rewind-dialog-options">
        <button
          type="button"
          className="rewind-dialog-option"
          onClick={() => onSelect("conversation")}
        >
          Rewind conversation only
          <span className="rewind-dialog-hint">Frees context, keeps document as-is</span>
        </button>
        <button
          type="button"
          className="rewind-dialog-option"
          onClick={() => onSelect("conversation_and_document")}
        >
          Rewind conversation + document
          <span className="rewind-dialog-hint">Restores document to snapshot state</span>
        </button>
      </div>
      {note && <div className="rewind-dialog-note">{note}</div>}
      <button type="button" className="rewind-dialog-cancel" onClick={onClose}>
        Cancel
      </button>
    </div>
  );
}
