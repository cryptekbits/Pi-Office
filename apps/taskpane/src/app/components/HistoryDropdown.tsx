import { useCallback, useEffect, useState } from "react";
import type { OfficeStateUpdate } from "@pi-office/pi-office-pack/protocol";
import type { ChatHistoryEntry, HistoryScope } from "../../hooks/useChatHistory";
import { TrashIcon } from "../../lib/icons";

const SKIP_DELETE_CONFIRM_KEY = "pi-office-skip-delete-confirm";

function loadSkipConfirm(): boolean {
  try {
    return localStorage.getItem(SKIP_DELETE_CONFIRM_KEY) === "true";
  } catch {
    return false;
  }
}

function relativeTime(iso: string): string {
  const diff = Date.now() - new Date(iso).getTime();
  const minutes = Math.floor(diff / 60_000);
  if (minutes < 1) return "now";
  if (minutes < 60) return `${minutes}m ago`;
  const hours = Math.floor(minutes / 60);
  if (hours < 24) return `${hours}h ago`;
  const days = Math.floor(hours / 24);
  if (days < 30) return `${days}d ago`;
  return new Date(iso).toLocaleDateString();
}

function entryLabel(entry: ChatHistoryEntry): string {
  if (entry.subject) return entry.subject;
  const firstUser = entry.messages.find((m) => m.role === "user");
  if (firstUser?.text) {
    return firstUser.text.length > 40 ? `${firstUser.text.slice(0, 40)}...` : firstUser.text;
  }
  return "Untitled chat";
}

interface HistoryDropdownProps {
  officeState: OfficeStateUpdate | undefined;
  listChats: (scope: HistoryScope, officeState: OfficeStateUpdate | undefined) => ChatHistoryEntry[];
  onDeleteChat: (chatId: string) => void;
  onSelectChat: (chatId: string) => void;
  onClose: () => void;
}

const SCOPE_LABELS: { scope: HistoryScope; label: string }[] = [
  { scope: "document", label: "Document" },
  { scope: "workspace", label: "Workspace" },
  { scope: "global", label: "Global" },
];

export function HistoryDropdown({
  officeState,
  listChats,
  onDeleteChat,
  onSelectChat,
  onClose,
}: HistoryDropdownProps) {
  const [scope, setScope] = useState<HistoryScope>("document");
  const [pendingDeleteId, setPendingDeleteId] = useState<string>();
  const [skipConfirm, setSkipConfirm] = useState(loadSkipConfirm);
  const [doNotAskChecked, setDoNotAskChecked] = useState(false);
  const [closing, setClosing] = useState(false);
  const chats = listChats(scope, officeState);

  const pendingEntry = pendingDeleteId ? chats.find((e) => e.chatId === pendingDeleteId) : undefined;

  const animateClose = useCallback(() => {
    setClosing(true);
  }, []);

  useEffect(() => {
    if (!closing) return;
    const timer = setTimeout(onClose, 180);
    return () => clearTimeout(timer);
  }, [closing, onClose]);

  const handleDeleteClick = (chatId: string) => {
    if (skipConfirm) {
      onDeleteChat(chatId);
      return;
    }
    setDoNotAskChecked(false);
    setPendingDeleteId(chatId);
  };

  const confirmDelete = () => {
    if (!pendingDeleteId) return;
    if (doNotAskChecked) {
      try { localStorage.setItem(SKIP_DELETE_CONFIRM_KEY, "true"); } catch {}
      setSkipConfirm(true);
    }
    onDeleteChat(pendingDeleteId);
    setPendingDeleteId(undefined);
  };

  const cancelDelete = () => {
    setPendingDeleteId(undefined);
  };

  return (
    <>
      <div className={`history-backdrop ${closing ? "history-backdrop-out" : ""}`} onClick={animateClose} />
      <div className={`history-dropdown ${closing ? "history-dropdown-out" : ""}`}>
        <div className="history-header">
          <span className="history-title">Recent chats</span>
          <div className="history-scope-toggle">
            {SCOPE_LABELS.map(({ scope: s, label }) => (
              <button
                key={s}
                type="button"
                className={`history-scope-btn ${scope === s ? "history-scope-btn-active" : ""}`}
                onClick={() => setScope(s)}
              >
                {label}
              </button>
            ))}
          </div>
        </div>

        <div className="history-list">
          {chats.length === 0 && (
            <div className="history-empty">No chats yet</div>
          )}
          {chats.map((entry) => (
            <button
              key={entry.chatId}
              type="button"
              className="history-item"
              onClick={() => {
                onSelectChat(entry.chatId);
                animateClose();
              }}
            >
              <span className="history-item-label">{entryLabel(entry)}</span>
              <span className="history-item-time">{relativeTime(entry.updatedAt)}</span>
              <button
                type="button"
                className="history-item-delete"
                aria-label="Delete chat"
                onClick={(e) => {
                  e.stopPropagation();
                  handleDeleteClick(entry.chatId);
                }}
              >
                <TrashIcon />
              </button>
            </button>
          ))}
        </div>

        {pendingEntry && (
          <div className="history-confirm-overlay" onClick={cancelDelete}>
            <div className="history-confirm-popup" onClick={(e) => e.stopPropagation()}>
              <div className="history-confirm-title">Delete this chat?</div>
              <div className="history-confirm-label">{entryLabel(pendingEntry)}</div>
              <label className="history-confirm-checkbox">
                <input
                  type="checkbox"
                  checked={doNotAskChecked}
                  onChange={(e) => setDoNotAskChecked(e.target.checked)}
                />
                Do not ask me again
              </label>
              <div className="history-confirm-actions">
                <button type="button" className="history-confirm-cancel" onClick={cancelDelete}>Cancel</button>
                <button type="button" className="history-confirm-delete" onClick={confirmDelete}>Delete</button>
              </div>
            </div>
          </div>
        )}
      </div>
    </>
  );
}
