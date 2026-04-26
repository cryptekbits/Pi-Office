import { useCallback, useEffect, useMemo, useRef, useState, type KeyboardEvent, type MouseEvent } from "react";
import type {
  AutonomyLevel,
  CompanionState,
  ContextBreakdownEntry,
  OfficeDocumentState,
  OfficeStateUpdate,
  SessionStatsResponse,
  ThinkingLevel,
} from "@pi-office/pi-office-pack/protocol";
import { ArrowUpIcon, StopIcon } from "../../lib/icons";
import { clampNumber, formatTokenCount, getContextRingColor } from "../../lib/helpers";
import { useInputHistory } from "../../hooks/useInputHistory";
import { ModelSelector, type ConfiguredModelEntry } from "./ModelSelector";
import { AutonomyToggle } from "./AutonomyToggle";

interface ComposerProps {
  draft: string;
  setDraft: (value: string) => void;
  focusSignal: number;
  userMessages: string[];
  sessionId: string | undefined;
  isBusy: boolean;
  activeToolName: string | undefined;
  officeState: OfficeStateUpdate | undefined;
  documentState: OfficeDocumentState;
  companion: CompanionState;
  shouldAttachDraftVisuals: boolean;
  selectedLocalQueueId: string | undefined;
  configuredModels: ConfiguredModelEntry[];
  selectedModelKey: string;
  thinkingLevel: ThinkingLevel;
  availableThinkingLevels: ThinkingLevel[];
  sessionStats: SessionStatsResponse | undefined;
  autonomyLevel: AutonomyLevel;
  onSend: () => void;
  onAbort: () => void;
  onQueue: () => void;
  onSteer: () => void;
  onSelectModel: (key: string) => void;
  onSetThinkingLevel: (level: ThinkingLevel) => void;
  onSetAutonomyLevel: (level: AutonomyLevel) => void;
}

export function Composer({
  draft,
  setDraft,
  focusSignal,
  userMessages,
  sessionId,
  isBusy,
  activeToolName,
  officeState,
  documentState,
  companion,
  shouldAttachDraftVisuals,
  selectedLocalQueueId,
  configuredModels,
  selectedModelKey,
  thinkingLevel,
  availableThinkingLevels,
  sessionStats,
  onSend,
  onAbort,
  onQueue,
  onSteer,
  onSelectModel,
  onSetThinkingLevel,
  autonomyLevel,
  onSetAutonomyLevel,
}: ComposerProps) {
  const textareaRef = useRef<HTMLTextAreaElement | null>(null);
  const canSend = Boolean(draft.trim() && sessionId);

  const contextPercent = sessionStats?.contextUsage?.percent ?? null;
  const contextWindow = sessionStats?.contextUsage?.contextWindow;
  const contextTokens = sessionStats?.contextUsage?.tokens ?? null;
  const contextProgress = contextPercent == null ? 0 : clampNumber(contextPercent, 0, 100);
  const contextRingColor = useMemo(() => getContextRingColor(contextPercent), [contextPercent]);

  const [showContextPopover, setShowContextPopover] = useState(false);
  const contextBreakdown = sessionStats?.contextUsage?.breakdown;

  const contextTooltip = useMemo(() => {
    if (contextBreakdown?.length) return undefined;
    if (!contextWindow) return "Context usage will appear once Pi has active session stats.";
    if (contextTokens == null || contextPercent == null) {
      return `Context usage: ? / ${formatTokenCount(contextWindow)} tokens`;
    }
    return `Context usage: ${formatTokenCount(contextTokens)} / ${formatTokenCount(contextWindow)} tokens (${contextPercent.toFixed(1)}%)`;
  }, [contextBreakdown, contextPercent, contextTokens, contextWindow]);

  const busyHint = useMemo(() => {
    if (selectedLocalQueueId && draft.trim()) return "Press Enter to send the selected queued note as steer.";
    if (isBusy && draft.trim()) return "Press Enter to queue this as the next follow-up.";
    if (activeToolName) return `Running ${activeToolName}.`;
    if (shouldAttachDraftVisuals) return "Relevant visual context will be attached automatically.";
    if (documentState !== "saved") {
      return "Save the file to unlock document-folder context.";
    }
    if (companion.status === "connected" && companion.capabilities.fileRead) {
      return "Read-only local file access is available for this saved document.";
    }
    return "Document path context is available, but local files stay unavailable without the optional companion.";
  }, [activeToolName, companion.capabilities.fileRead, companion.status, documentState, draft, isBusy, selectedLocalQueueId, shouldAttachDraftVisuals]);

  const { handleHistoryNav, resetHistory } = useInputHistory(userMessages, setDraft);

  const handlePrimaryAction = useCallback(() => {
    if (isBusy) { onAbort(); return; }
    onSend();
  }, [isBusy, onAbort, onSend]);

  const focusTextarea = useCallback(() => {
    window.focus();
    textareaRef.current?.focus({ preventScroll: true });
  }, []);

  useEffect(() => {
    if (focusSignal <= 0) return;
    focusTextarea();
  }, [focusSignal, focusTextarea]);

  const handlePointerFocus = useCallback((event: MouseEvent<HTMLTextAreaElement>) => {
    event.stopPropagation();
    if (document.activeElement === event.currentTarget) return;
    queueMicrotask(focusTextarea);
  }, [focusTextarea]);

  const handleKeyDown = useCallback(
    (event: KeyboardEvent<HTMLTextAreaElement>) => {
      event.stopPropagation();
      if (handleHistoryNav(event)) return;

      if (event.key !== "Enter" || event.shiftKey) return;
      event.preventDefault();
      if (!draft.trim()) return;

      if (isBusy) {
        if (selectedLocalQueueId) { onSteer(); resetHistory(); return; }
        onQueue(); resetHistory(); return;
      }
      onSend(); resetHistory();
    },
    [draft, handleHistoryNav, isBusy, onQueue, onSend, onSteer, resetHistory, selectedLocalQueueId],
  );

  return (
    <section className="composer">
      <div className="composer-box">
        <textarea
          ref={textareaRef}
          value={draft}
          onChange={(e) => setDraft(e.target.value)}
          onMouseDown={handlePointerFocus}
          onClick={handlePointerFocus}
          onFocus={focusTextarea}
          onKeyUp={(event) => event.stopPropagation()}
          onKeyDown={handleKeyDown}
          placeholder="Ask Pi to work on the current document..."
          disabled={!sessionId}
        />

        <div className="composer-footer">
          <div className="composer-left">
            <div className="composer-controls">
              <ModelSelector
                configuredModels={configuredModels}
                selectedModelKey={selectedModelKey}
                thinkingLevel={thinkingLevel}
                availableThinkingLevels={availableThinkingLevels}
                onSelectModel={onSelectModel}
                onSetThinkingLevel={onSetThinkingLevel}
              />
              <AutonomyToggle level={autonomyLevel} onSetLevel={onSetAutonomyLevel} />
            </div>
            <p className="composer-hint">{busyHint}</p>
          </div>

          <div
            className="composer-right send-control"
            title={contextTooltip}
            onMouseEnter={() => contextBreakdown?.length && setShowContextPopover(true)}
            onMouseLeave={() => setShowContextPopover(false)}
          >
            {showContextPopover && contextBreakdown?.length && contextWindow ? (
              <ContextPopover
                breakdown={contextBreakdown}
                contextWindow={contextWindow}
                totalTokens={contextTokens}
                percent={contextPercent}
              />
            ) : null}
            <button
              type="button"
              className={["send-button", isBusy ? "send-button-busy" : ""].filter(Boolean).join(" ")}
              aria-label={isBusy ? "Abort current response" : "Send prompt"}
              disabled={isBusy ? false : !canSend}
              onClick={handlePrimaryAction}
            >
              <svg className="send-ring" viewBox="0 0 52 52" aria-hidden="true">
                <circle className="send-ring-track" cx="26" cy="26" r="21" />
                <circle
                  className="send-ring-progress"
                  cx="26"
                  cy="26"
                  r="21"
                  pathLength="100"
                  strokeDasharray={`${contextProgress} 100`}
                  style={{ stroke: contextRingColor }}
                />
              </svg>
              <span className="send-button-core">
                {isBusy ? <StopIcon /> : <ArrowUpIcon />}
              </span>
            </button>
          </div>
        </div>
      </div>
    </section>
  );
}

function ContextPopover({
  breakdown,
  contextWindow,
  totalTokens,
  percent,
}: {
  breakdown: ContextBreakdownEntry[];
  contextWindow: number;
  totalTokens: number | null;
  percent: number | null;
}) {
  const filled = breakdown.filter((e) => e.label !== "Available");
  const percentStr = percent != null ? `${percent.toFixed(1)}%` : "N/A";

  return (
    <div className="context-popover">
      <div className="context-popover-header">
        <span>Context Usage</span>
        <span>{percentStr}</span>
      </div>
      <div className="context-popover-bar">
        {filled
          .filter((e) => e.tokens > 0)
          .map((e) => (
            <div
              key={e.label}
              className="context-popover-bar-segment"
              style={{
                width: `${contextWindow > 0 ? (e.tokens / contextWindow) * 100 : 0}%`,
                backgroundColor: e.color,
              }}
            />
          ))}
      </div>
      <div className="context-popover-list">
        {breakdown.map((e) => {
          const pct = contextWindow > 0 ? ((e.tokens / contextWindow) * 100).toFixed(1) : "0.0";
          const isAvailable = e.label === "Available";
          return (
            <div key={e.label} className="context-popover-row">
              <div className="context-popover-label">
                <span
                  className="context-popover-dot"
                  style={{ backgroundColor: e.color, opacity: isAvailable ? 0.4 : 1 }}
                />
                <span className={isAvailable ? "context-popover-muted" : ""}>{e.label}</span>
              </div>
              <div className="context-popover-value">
                <span>{formatTokenCount(e.tokens)}</span>
                <span className="context-popover-muted">({pct}%)</span>
              </div>
            </div>
          );
        })}
      </div>
      <div className="context-popover-footer">
        <span>Total</span>
        <span>
          {formatTokenCount(totalTokens)} / {formatTokenCount(contextWindow)}
        </span>
      </div>
    </div>
  );
}
