import { ConversationLogIcon, GearIcon, HistoryIcon, NewChatIcon } from "../../lib/icons";
import { connectionLabel } from "../../lib/helpers";

interface HeaderProps {
  host: string | undefined;
  connectionState: string;
  chatSubject: string | undefined;
  historyOpen: boolean;
  debugLogCopyAvailable: boolean;
  debugLogCopying: boolean;
  onCopyConversationDebugLog: () => void;
  onSettingsClick: () => void;
  onNewChat: () => void;
  onHistoryToggle: () => void;
}

export function Header({
  host,
  connectionState,
  chatSubject,
  historyOpen,
  debugLogCopyAvailable,
  debugLogCopying,
  onCopyConversationDebugLog,
  onSettingsClick,
  onNewChat,
  onHistoryToggle,
}: HeaderProps) {
  return (
    <header className="topbar">
      <div className="topbar-left">
        <h1 className="topbar-title">Pi for Office</h1>
        <span
          className={[
            "status-badge",
            connectionState === "ready" ? "status-ready" : "",
            connectionState === "offline" ? "status-offline" : "",
          ]
            .filter(Boolean)
            .join(" ")}
        >
          {connectionLabel(connectionState)}
        </span>
      </div>

      {chatSubject && (
        <div className="topbar-center" title={chatSubject}>
          {chatSubject}
        </div>
      )}

      <div className="topbar-right">
        {debugLogCopyAvailable && (
          <button
            type="button"
            className="icon-button"
            aria-label="Copy conversation debug log"
            title={debugLogCopying ? "Copying conversation debug log..." : "Copy conversation debug log"}
            onClick={onCopyConversationDebugLog}
            disabled={debugLogCopying}
          >
            <ConversationLogIcon />
          </button>
        )}
        <button
          type="button"
          className="icon-button"
          aria-label="New chat"
          onClick={onNewChat}
        >
          <NewChatIcon />
        </button>
        <button
          type="button"
          className={`icon-button ${historyOpen ? "icon-button-active" : ""}`}
          aria-label="Chat history"
          onClick={onHistoryToggle}
        >
          <HistoryIcon />
        </button>
        <button
          type="button"
          className="icon-button"
          aria-label="Open settings"
          onClick={onSettingsClick}
        >
          <GearIcon />
        </button>
      </div>
    </header>
  );
}
