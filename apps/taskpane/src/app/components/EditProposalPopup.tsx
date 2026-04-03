import { useCallback, useState } from "react";
import type { OfficeEditProposal, OfficeProposedEdit, OfficeEditProposalDecision } from "@pi-office/pi-office-pack/protocol";
import { EDIT_REJECT_REASONS, EDIT_REJECT_REASON_LABELS, type EditRejectReason } from "@pi-office/pi-office-pack/protocol";

interface EditProposalPopupProps {
  proposal: OfficeEditProposal;
  onDecision: (decision: OfficeEditProposalDecision) => void;
}

interface EditFeedback {
  reason?: EditRejectReason | undefined;
  note?: string | undefined;
}

function EditCard({
  edit,
  accepted,
  feedback,
  onToggle,
  onFeedbackChange,
}: {
  edit: OfficeProposedEdit;
  accepted: boolean;
  feedback: EditFeedback;
  onToggle: () => void;
  onFeedbackChange: (editId: string, patch: Partial<EditFeedback>) => void;
}) {
  const [expanded, setExpanded] = useState(false);
  const [noteOpen, setNoteOpen] = useState(false);

  return (
    <div className={`edit-card ${accepted ? "edit-card-accepted" : "edit-card-rejected"} ${!edit.found ? "edit-card-notfound" : ""}`}>
      <div className="edit-card-header">
        <button
          type="button"
          className={`toggle-switch toggle-sm ${accepted ? "toggle-on" : ""}`}
          onClick={onToggle}
          role="switch"
          aria-checked={accepted}
          aria-label={accepted ? "Reject this edit" : "Accept this edit"}
        >
          <span className="toggle-thumb" />
        </button>
        <span className="edit-card-kind">{edit.kind}</span>
        {!edit.found && <span className="edit-card-notfound-badge">Not found</span>}
        <button
          type="button"
          className="edit-card-expand"
          onClick={() => setExpanded((c) => !c)}
          aria-label={expanded ? "Collapse" : "Expand"}
        >
          <svg viewBox="0 0 16 16" width="12" height="12" className={expanded ? "edit-card-chevron-open" : ""}>
            <path d="M6 4l4 4-4 4" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
        </button>
      </div>

      {edit.explanation && (
        <p className="edit-card-explanation">{edit.explanation}</p>
      )}

      <div className="edit-card-diff">
        {edit.oldText && (
          <div className="edit-card-diff-old">
            <span className="edit-card-diff-label">-</span>
            <span>{edit.oldText.length > 200 ? `${edit.oldText.slice(0, 197)}...` : edit.oldText}</span>
          </div>
        )}
        {edit.kind !== "delete" && edit.newText && (
          <div className="edit-card-diff-new">
            <span className="edit-card-diff-label">+</span>
            <span>{edit.newText.length > 200 ? `${edit.newText.slice(0, 197)}...` : edit.newText}</span>
          </div>
        )}
      </div>

      {expanded && edit.contextPreview && (
        <pre className="edit-card-context">{edit.contextPreview}</pre>
      )}

      {!accepted && (
        <div className="edit-card-feedback">
          <div className="edit-card-reject-chips">
            {EDIT_REJECT_REASONS.map((reason) => (
              <button
                key={reason}
                type="button"
                className={`edit-reject-chip ${feedback.reason === reason ? "edit-reject-chip-active" : ""}`}
                onClick={() => onFeedbackChange(edit.id, { reason: feedback.reason === reason ? undefined : reason })}
              >
                {EDIT_REJECT_REASON_LABELS[reason]}
              </button>
            ))}
          </div>
          <div className="edit-card-note-row">
            <button
              type="button"
              className="edit-card-note-toggle"
              onClick={() => setNoteOpen((c) => !c)}
            >
              {noteOpen ? "Hide note" : "+ Add note"}
            </button>
            {noteOpen && (
              <input
                type="text"
                className="edit-card-note-input"
                placeholder="Why reject this edit..."
                value={feedback.note ?? ""}
                onChange={(e) => onFeedbackChange(edit.id, { note: e.target.value || undefined })}
              />
            )}
          </div>
        </div>
      )}
    </div>
  );
}

export function EditProposalPopup({ proposal, onDecision }: EditProposalPopupProps) {
  const [decisions, setDecisions] = useState<Map<string, boolean>>(() => {
    const map = new Map<string, boolean>();
    for (const edit of proposal.edits) {
      map.set(edit.id, edit.found !== false);
    }
    return map;
  });
  const [feedbacks, setFeedbacks] = useState<Map<string, EditFeedback>>(() => new Map());
  const [globalFeedback, setGlobalFeedback] = useState("");
  const [globalFeedbackOpen, setGlobalFeedbackOpen] = useState(false);

  const acceptedCount = Array.from(decisions.values()).filter(Boolean).length;
  const rejectedCount = proposal.edits.length - acceptedCount;

  const toggleEdit = useCallback((editId: string) => {
    setDecisions((prev) => {
      const next = new Map(prev);
      next.set(editId, !next.get(editId));
      return next;
    });
  }, []);

  const updateFeedback = useCallback((editId: string, patch: Partial<EditFeedback>) => {
    setFeedbacks((prev) => {
      const next = new Map(prev);
      const existing = next.get(editId) ?? {};
      next.set(editId, { ...existing, ...patch });
      return next;
    });
  }, []);

  const handleAcceptAll = useCallback(() => {
    setDecisions((prev) => {
      const next = new Map(prev);
      for (const edit of proposal.edits) {
        if (edit.found !== false) next.set(edit.id, true);
      }
      return next;
    });
  }, [proposal.edits]);

  const handleRejectAll = useCallback(() => {
    setDecisions((prev) => {
      const next = new Map(prev);
      for (const key of next.keys()) next.set(key, false);
      return next;
    });
  }, []);

  const buildDecision = useCallback(
    (overrideAllRejected?: boolean): OfficeEditProposalDecision => {
      return {
        requestId: proposal.requestId,
        decisions: proposal.edits.map((edit) => {
          const accepted = overrideAllRejected ? false : (decisions.get(edit.id) ?? false);
          const fb = feedbacks.get(edit.id);
          return {
            editId: edit.id,
            accepted,
            ...(!accepted && fb?.reason ? { rejectReason: fb.reason } : {}),
            ...(!accepted && fb?.note ? { rejectNote: fb.note } : {}),
          };
        }),
        ...(globalFeedback.trim() ? { globalFeedback: globalFeedback.trim() } : {}),
      };
    },
    [decisions, feedbacks, globalFeedback, proposal],
  );

  const handleSubmit = useCallback(() => {
    onDecision(buildDecision());
  }, [buildDecision, onDecision]);

  const handleReject = useCallback(() => {
    onDecision(buildDecision(true));
  }, [buildDecision, onDecision]);

  return (
    <div className="edit-proposal-overlay">
      <div className="edit-proposal-popup">
        <div className="edit-proposal-header">
          <div className="edit-proposal-header-text">
            <span className="edit-proposal-title">Proposed Edits</span>
            <span className="edit-proposal-counter">
              {acceptedCount}/{proposal.edits.length} accepted
            </span>
          </div>
          <div className="edit-proposal-batch-actions">
            <button type="button" className="button button-sm" onClick={handleAcceptAll}>All</button>
            <button type="button" className="button button-sm" onClick={handleRejectAll}>None</button>
          </div>
        </div>

        <p className="edit-proposal-summary">{proposal.summary}</p>

        <div className="edit-proposal-list">
          {proposal.edits.map((edit) => (
            <EditCard
              key={edit.id}
              edit={edit}
              accepted={decisions.get(edit.id) ?? false}
              feedback={feedbacks.get(edit.id) ?? {}}
              onToggle={() => toggleEdit(edit.id)}
              onFeedbackChange={updateFeedback}
            />
          ))}
        </div>

        {rejectedCount > 0 && (
          <div className="edit-proposal-global-feedback">
            <button
              type="button"
              className="edit-proposal-global-toggle"
              onClick={() => setGlobalFeedbackOpen((c) => !c)}
            >
              <svg viewBox="0 0 16 16" width="12" height="12" className={globalFeedbackOpen ? "edit-card-chevron-open" : ""}>
                <path d="M6 4l4 4-4 4" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
              Add overall feedback
            </button>
            {globalFeedbackOpen && (
              <textarea
                className="edit-proposal-global-textarea"
                placeholder="General direction for rejected edits (e.g. &quot;keep the tone more technical&quot;)..."
                value={globalFeedback}
                onChange={(e) => setGlobalFeedback(e.target.value)}
                rows={2}
              />
            )}
          </div>
        )}

        <div className="edit-proposal-footer">
          <button type="button" className="button button-danger button-sm" onClick={handleReject}>
            Reject All
          </button>
          <button
            type="button"
            className="button button-solid"
            onClick={handleSubmit}
          >
            Apply {acceptedCount} Edit{acceptedCount !== 1 ? "s" : ""}
          </button>
        </div>
      </div>
    </div>
  );
}
