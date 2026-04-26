import { useState } from "react";
import type { ToolCallEntry } from "../../lib/helpers";

interface ToolCallBlockProps {
  toolCalls: ToolCallEntry[];
  pending: boolean;
}

function humanizeToolName(name: string): string {
  return name
    .replace(/_/g, " ")
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .replace(/^./, (c) => c.toUpperCase());
}

function formatJson(value: unknown): string {
  if (value === undefined || value === null) return "";
  try {
    return JSON.stringify(value, null, 2);
  } catch {
    return String(value);
  }
}

function extractResultSummary(result: unknown): string | null {
  if (!result || typeof result !== "object") return null;
  const r = result as Record<string, unknown>;
  const content = r.content;
  if (Array.isArray(content)) {
    const textPart = content.find(
      (p: unknown) => p && typeof p === "object" && (p as { type?: string }).type === "text",
    ) as { text?: string } | undefined;
    if (textPart?.text) return textPart.text;
  }
  if (typeof r.text === "string") return r.text;
  return null;
}

function ToolCallDetail({ toolCall }: { toolCall: ToolCallEntry }) {
  const [showMore, setShowMore] = useState(false);
  const inputJson = formatJson(toolCall.input);
  const resultText = extractResultSummary(toolCall.result);
  const resultJson = formatJson(toolCall.result);
  const hasInput = Boolean(inputJson);
  const hasResult = Boolean(resultJson);

  return (
    <div className="tool-detail">
      {hasInput && (
        <div className="tool-detail-section">
          <div className="tool-detail-heading">Parameters:</div>
          <pre className="tool-detail-json">
            <code>{inputJson.length > 300 && !showMore ? `${inputJson.slice(0, 300)}...` : inputJson}</code>
          </pre>
          {inputJson.length > 300 && (
            <button type="button" className="tool-detail-toggle" onClick={() => setShowMore((c) => !c)}>
              {showMore ? "Show less" : "Show more"}
            </button>
          )}
        </div>
      )}
      {resultText && (
        <div className="tool-detail-section">
          <div className="tool-detail-result-label">Result</div>
          <pre className="tool-detail-result">{resultText}</pre>
        </div>
      )}
      {!resultText && hasResult && (
        <div className="tool-detail-section">
          <div className="tool-detail-result-label">Result</div>
          <pre className="tool-detail-json">
            <code>{resultJson.length > 400 ? `${resultJson.slice(0, 400)}...` : resultJson}</code>
          </pre>
        </div>
      )}
      {toolCall.status === "error" && toolCall.errorMessage && (
        <div className="tool-detail-error">{toolCall.errorMessage}</div>
      )}
      {!hasInput && !hasResult && !toolCall.errorMessage && (
        <div className="tool-detail-empty">No details available.</div>
      )}
    </div>
  );
}

function ToolCallRow({ toolCall, defaultOpen }: { toolCall: ToolCallEntry; defaultOpen?: boolean }) {
  const name = humanizeToolName(toolCall.toolName);
  const isRunning = toolCall.status === "running";
  const isError = toolCall.status === "error";
  const dotClass = isRunning ? "tool-line-dot-running" : isError ? "tool-line-dot-error" : "tool-line-dot-done";

  if (isRunning) {
    return (
      <div className="tool-line tool-line-running">
        <span className={`tool-line-dot ${dotClass}`} aria-hidden="true" />
        <span className="tool-line-name">{name}</span>
      </div>
    );
  }

  return (
    <details className="tool-line-details" open={defaultOpen}>
      <summary className={`tool-line-summary ${isError ? "tool-line-error" : ""}`}>
        <span className={`tool-line-dot ${dotClass}`} aria-hidden="true" />
        <span className="tool-line-name">{name}</span>
        <span className="tool-line-chevron" aria-hidden="true" />
      </summary>
      <ToolCallDetail toolCall={toolCall} />
    </details>
  );
}

function buildGroupSummary(toolCalls: ToolCallEntry[]): string {
  if (toolCalls.length === 1) {
    return humanizeToolName(toolCalls[0]!.toolName);
  }
  const uniqueNames = [...new Set(toolCalls.map((t) => t.toolName))];
  if (uniqueNames.length === 1) {
    return `${humanizeToolName(uniqueNames[0]!)} (${toolCalls.length})`;
  }
  const first = humanizeToolName(uniqueNames[0]!);
  return `${first} and ${toolCalls.length - 1} other tool${toolCalls.length > 2 ? "s" : ""}`;
}

export function ToolCallBlock({ toolCalls, pending }: ToolCallBlockProps) {
  if (!toolCalls.length) return null;

  if (pending) {
    return (
      <div className="tool-lines">
        {toolCalls.map((toolCall) => (
          <ToolCallRow key={toolCall.toolCallId} toolCall={toolCall} />
        ))}
      </div>
    );
  }

  const hasErrors = toolCalls.some((t) => t.status === "error");

  return (
    <details className="tool-group-details">
      <summary className={`tool-group-summary ${hasErrors ? "tool-group-summary-error" : ""}`}>
        {buildGroupSummary(toolCalls)}
        <span className="tool-line-chevron" aria-hidden="true" />
      </summary>
      <div className="tool-group-body">
        {toolCalls.map((toolCall) => (
          <ToolCallRow key={toolCall.toolCallId} toolCall={toolCall} />
        ))}
      </div>
    </details>
  );
}
