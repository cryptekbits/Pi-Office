import { useCallback } from "react";
import type { ToolCategory, ToolPermissionRequest } from "@pi-office/pi-office-pack/protocol";

const CATEGORY_LABELS: Record<ToolCategory, string> = {
  read: "Read (document)",
  "write-doc": "Write (document)",
  "escape-hatch": "Manual escape hatch",
  "read-external": "Read (workspace)",
  "write-external": "Write (workspace)",
  connector: "Connector",
  interaction: "Interaction",
};

interface ToolPermissionPopupProps {
  request: ToolPermissionRequest;
  onDecision: (requestId: string, allowed: boolean, scope: "once" | "session" | "workspace" | "always") => void;
}

function formatParams(params: Record<string, unknown>): string {
  const entries = Object.entries(params);
  if (!entries.length) return "No parameters";
  return entries
    .map(([key, value]) => {
      const display = typeof value === "string"
        ? value.length > 120 ? `${value.slice(0, 117)}...` : value
        : JSON.stringify(value)?.slice(0, 120) ?? String(value);
      return `${key}: ${display}`;
    })
    .join("\n");
}

export function ToolPermissionPopup({ request, onDecision }: ToolPermissionPopupProps) {
  const handleAllow = useCallback(
    (scope: "once" | "session" | "workspace" | "always") => {
      onDecision(request.requestId, true, scope);
    },
    [onDecision, request.requestId],
  );

  const handleDeny = useCallback(() => {
    onDecision(request.requestId, false, "once");
  }, [onDecision, request.requestId]);

  return (
    <div className="tool-permission-overlay">
      <div className="tool-permission-popup">
        <div className="tool-permission-header">
          <span className="tool-permission-title">Tool Permission Required</span>
          <span className="tool-permission-category">{CATEGORY_LABELS[request.toolCategory]}</span>
        </div>

        <div className="tool-permission-body">
          <div className="tool-permission-tool-name">
            Pi wants to use <code>{request.toolName}</code>
          </div>

          {Object.keys(request.params).length > 0 && (
            <pre className="tool-permission-params">{formatParams(request.params)}</pre>
          )}
        </div>

        <div className="tool-permission-actions">
          <div className="tool-permission-allow-group">
            <button type="button" className="button button-solid button-sm" onClick={() => handleAllow("once")}>
              Allow once
            </button>
            <button type="button" className="button button-sm" onClick={() => handleAllow("session")}>
              For session
            </button>
            <button type="button" className="button button-sm" onClick={() => handleAllow("workspace")}>
              For workspace
            </button>
            <button type="button" className="button button-sm" onClick={() => handleAllow("always")}>
              Always
            </button>
          </div>
          <button type="button" className="button button-danger button-sm" onClick={handleDeny}>
            Deny
          </button>
        </div>
      </div>
    </div>
  );
}
