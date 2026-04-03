import { useCallback, useEffect, useRef, useState } from "react";
import { AUTONOMY_LEVELS, AUTONOMY_LEVEL_LABELS, type AutonomyLevel } from "@pi-office/pi-office-pack/protocol";

const LEVEL_DESCRIPTIONS: Record<AutonomyLevel, string> = {
  off: "Ask before every tool call",
  low: "Auto-approve reads only",
  medium: "Auto-approve reads + doc edits",
  high: "Auto-approve reads + doc + workspace reads",
  extreme: "Auto-approve everything",
};

const SHIELD_FILLS: Record<AutonomyLevel, number> = {
  off: 0,
  low: 1,
  medium: 2,
  high: 3,
  extreme: 4,
};

function ShieldIcon({ level }: { level: AutonomyLevel }) {
  const fill = SHIELD_FILLS[level];
  return (
    <svg viewBox="0 0 16 16" width="14" height="14" className="autonomy-shield-icon" aria-hidden="true">
      <path
        d="M8 1L2 4v4c0 3.3 2.6 6.4 6 7 3.4-.6 6-3.7 6-7V4L8 1z"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.2"
        strokeLinejoin="round"
      />
      {fill >= 1 && <rect x="4.5" y="10" width="7" height="2.5" rx="0.5" fill="currentColor" opacity="0.5" />}
      {fill >= 2 && <rect x="4.5" y="7" width="7" height="2.5" rx="0.5" fill="currentColor" opacity="0.6" />}
      {fill >= 3 && <rect x="4.5" y="4" width="7" height="2.5" rx="0.5" fill="currentColor" opacity="0.7" />}
      {fill >= 4 && <rect x="5.5" y="2" width="5" height="1.5" rx="0.5" fill="currentColor" opacity="0.85" />}
    </svg>
  );
}

interface AutonomyToggleProps {
  level: AutonomyLevel;
  onSetLevel: (level: AutonomyLevel) => void;
}

export function AutonomyToggle({ level, onSetLevel }: AutonomyToggleProps) {
  const [open, setOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  const handleToggle = useCallback(() => setOpen((c) => !c), []);

  const handleSelect = useCallback(
    (next: AutonomyLevel) => {
      onSetLevel(next);
      setOpen(false);
    },
    [onSetLevel],
  );

  useEffect(() => {
    if (!open) return;
    const handleClick = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClick);
    return () => document.removeEventListener("mousedown", handleClick);
  }, [open]);

  return (
    <div ref={containerRef} className="autonomy-toggle">
      <button
        type="button"
        className="autonomy-toggle-button"
        onClick={handleToggle}
        title={`Autonomy: ${AUTONOMY_LEVEL_LABELS[level]} - ${LEVEL_DESCRIPTIONS[level]}`}
        aria-haspopup="listbox"
        aria-expanded={open}
      >
        <ShieldIcon level={level} />
        <span className="autonomy-toggle-label">{AUTONOMY_LEVEL_LABELS[level]}</span>
      </button>

      {open && (
        <div className="autonomy-dropdown" role="listbox" aria-label="Autonomy level">
          {AUTONOMY_LEVELS.map((lvl) => (
            <button
              key={lvl}
              type="button"
              role="option"
              aria-selected={lvl === level}
              className={`autonomy-dropdown-item ${lvl === level ? "autonomy-dropdown-item-active" : ""}`}
              onClick={() => handleSelect(lvl)}
            >
              <ShieldIcon level={lvl} />
              <span className="autonomy-dropdown-item-content">
                <span className="autonomy-dropdown-item-label">{AUTONOMY_LEVEL_LABELS[lvl]}</span>
                <span className="autonomy-dropdown-item-desc">{LEVEL_DESCRIPTIONS[lvl]}</span>
              </span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
