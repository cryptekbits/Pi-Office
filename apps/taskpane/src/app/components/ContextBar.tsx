import { useState } from "react";
import type { OfficeStateUpdate, OfficeMode } from "@pi-office/pi-office-pack/protocol";
import { HOST_LABELS } from "@pi-office/pi-office-pack/defaults";
import { ChevronDownIcon, ChevronRightIcon } from "../../lib/icons";

interface ContextBarProps {
  officeState: OfficeStateUpdate | undefined;
  mode: OfficeMode;
  shouldAttachDraftVisuals: boolean;
}

export function ContextBar({ officeState, mode, shouldAttachDraftVisuals }: ContextBarProps) {
  const [expanded, setExpanded] = useState(false);

  const docTitle = officeState?.document.title ?? "Waiting for Office...";
  const selLabel = officeState?.selection.label ?? "No selection";
  const selDetails = officeState?.selection.details?.join(" · ");

  return (
    <section className="context-bar">
      <button
        type="button"
        className="context-bar-toggle"
        onClick={() => setExpanded((c) => !c)}
        aria-expanded={expanded}
      >
        <span className="context-bar-summary">
          <span className="context-bar-doc" title={docTitle}>{docTitle}</span>
          <span className="context-bar-sep">|</span>
          <span className="context-bar-sel">{selLabel}</span>
        </span>
        <span className="context-bar-chevron">
          {expanded ? <ChevronDownIcon /> : <ChevronRightIcon />}
        </span>
      </button>

      {expanded && (
        <div className="context-bar-details">
          <div className="context-cell">
            <span className="label">Document</span>
            <p className="context-primary">
              {docTitle}
              {officeState?.document.saved && officeState.document.workspaceDir
                ? ` · ${officeState.document.workspaceDir}`
                : ""}
            </p>
          </div>
          <div className="context-cell">
            <span className="label">Selection</span>
            <p className="context-primary">{selLabel}</p>
            {selDetails ? <p className="context-secondary">{selDetails}</p> : null}
          </div>
          <div className="detail-row">
            <span className="detail-tag">
              {officeState ? HOST_LABELS[officeState.host] : "Office"}
            </span>
            <span className="detail-tag">
              {mode === "workspace" ? "Workspace mode" : "Document-only mode"}
            </span>
            {(officeState?.selection.imageCount ?? 0) > 0 && (
              <span className="detail-tag">
                {officeState?.selection.imageCount} visual item(s)
              </span>
            )}
            {shouldAttachDraftVisuals && (
              <span className="detail-tag">Visual context ready</span>
            )}
          </div>
        </div>
      )}
    </section>
  );
}
