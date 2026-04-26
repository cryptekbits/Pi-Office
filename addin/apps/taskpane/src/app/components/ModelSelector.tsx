import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { ProviderModelDescriptor, ThinkingLevel } from "@pi-office/pi-office-pack/protocol";
import { CheckIcon, SearchIcon, CloseIcon } from "../../lib/icons";
import { getRecentModels, pushRecentModel } from "../../hooks/usePreferences";
import { formatTokenCount } from "../../lib/helpers";

export interface ConfiguredModelEntry {
  provider: { provider: string; label: string };
  model: ProviderModelDescriptor;
  key: string;
}

interface ModelSelectorProps {
  configuredModels: ConfiguredModelEntry[];
  selectedModelKey: string;
  thinkingLevel: ThinkingLevel;
  availableThinkingLevels: ThinkingLevel[];
  onSelectModel: (key: string) => void;
  onSetThinkingLevel: (level: ThinkingLevel) => void;
}

export function ModelSelector({
  configuredModels,
  selectedModelKey,
  thinkingLevel,
  availableThinkingLevels,
  onSelectModel,
  onSetThinkingLevel,
}: ModelSelectorProps) {
  const [isOpen, setIsOpen] = useState(false);
  const [search, setSearch] = useState("");
  const [hoveredKey, setHoveredKey] = useState<string | null>(null);
  const [hoveredTop, setHoveredTop] = useState(0);
  const [detailPinned, setDetailPinned] = useState(false);
  const popupRef = useRef<HTMLDivElement>(null);
  const detailRef = useRef<HTMLDivElement>(null);
  const searchRef = useRef<HTMLInputElement>(null);
  const hoverTimerRef = useRef<ReturnType<typeof setTimeout>>(undefined);

  const selectedModel = useMemo(
    () => configuredModels.find((e) => e.key === selectedModelKey),
    [configuredModels, selectedModelKey],
  );

  const recentKeys = useMemo(() => getRecentModels(), [isOpen]);

  const filtered = useMemo(() => {
    if (!search.trim()) return configuredModels;
    const q = search.toLowerCase();
    return configuredModels.filter(
      (e) =>
        e.model.modelName.toLowerCase().includes(q) ||
        e.provider.label.toLowerCase().includes(q) ||
        e.model.modelId.toLowerCase().includes(q),
    );
  }, [configuredModels, search]);

  const recentModels = useMemo(
    () => recentKeys.map((k) => filtered.find((e) => e.key === k)).filter(Boolean) as ConfiguredModelEntry[],
    [filtered, recentKeys],
  );

  const groupedByProvider = useMemo(() => {
    const recentSet = new Set(recentModels.map((e) => e.key));
    const groups = new Map<string, ConfiguredModelEntry[]>();
    for (const entry of filtered) {
      if (recentSet.has(entry.key)) continue;
      const list = groups.get(entry.provider.provider) ?? [];
      list.push(entry);
      groups.set(entry.provider.provider, list);
    }
    return groups;
  }, [filtered, recentModels]);

  useEffect(() => {
    if (!isOpen) {
      setSearch("");
      setHoveredKey(null);
      setHoveredTop(0);
      setDetailPinned(false);
      return;
    }
    searchRef.current?.focus();

    const handler = (e: MouseEvent) => {
      const target = e.target as Node;
      if (
        popupRef.current && !popupRef.current.contains(target) &&
        detailRef.current && !detailRef.current.contains(target)
      ) {
        setIsOpen(false);
      }
    };
    const keyHandler = (e: KeyboardEvent) => {
      if (e.key === "Escape") setIsOpen(false);
    };
    document.addEventListener("mousedown", handler);
    document.addEventListener("keydown", keyHandler);
    return () => {
      document.removeEventListener("mousedown", handler);
      document.removeEventListener("keydown", keyHandler);
    };
  }, [isOpen]);

  const handleSelect = useCallback(
    (key: string) => {
      onSelectModel(key);
      if (key) pushRecentModel(key);
      setIsOpen(false);
    },
    [onSelectModel],
  );

  const handleRowMouseEnter = useCallback((key: string, el: HTMLElement) => {
    clearTimeout(hoverTimerRef.current);
    const popupEl = popupRef.current;
    if (popupEl) {
      const popupRect = popupEl.getBoundingClientRect();
      const rowRect = el.getBoundingClientRect();
      setHoveredTop(rowRect.top - popupRect.top);
    }
    setHoveredKey(key);
  }, []);

  const handleRowMouseLeave = useCallback(() => {
    if (detailPinned) return;
    hoverTimerRef.current = setTimeout(() => {
      setHoveredKey(null);
    }, 200);
  }, [detailPinned]);

  const handleDetailMouseEnter = useCallback(() => {
    clearTimeout(hoverTimerRef.current);
  }, []);

  const handleDetailMouseLeave = useCallback(() => {
    if (detailPinned) return;
    hoverTimerRef.current = setTimeout(() => {
      setHoveredKey(null);
    }, 200);
  }, [detailPinned]);

  const hoveredModel = useMemo(
    () => (hoveredKey ? configuredModels.find((e) => e.key === hoveredKey) : null),
    [configuredModels, hoveredKey],
  );

  const triggerLabel = selectedModel ? selectedModel.model.modelName : "Pi default";
  const triggerSublabel = selectedModel ? selectedModel.provider.label : undefined;
  const triggerAuthLabel = selectedModel ? authStateLabel(selectedModel.model) : undefined;

  const thinkingLabel = thinkingLevel === "off"
    ? "Off" : thinkingLevel === "xhigh"
    ? "XHigh" : thinkingLevel.charAt(0).toUpperCase() + thinkingLevel.slice(1);

  return (
    <div className="model-selector-root">
      <button
        type="button"
        className="model-selector-trigger"
        onClick={() => setIsOpen((c) => !c)}
        aria-haspopup="listbox"
        aria-expanded={isOpen}
      >
        <div className="model-selector-trigger-text">
          {triggerSublabel && (
            <span className="model-selector-provider">{triggerSublabel}</span>
          )}
          <span className="model-selector-label">{triggerLabel}</span>
        </div>
        {selectedModel?.model.supportsThinking && (
          <span className="model-selector-thinking-badge">{thinkingLabel}</span>
        )}
        {triggerAuthLabel && (
          <span className="model-selector-auth-badge">{triggerAuthLabel}</span>
        )}
        <svg viewBox="0 0 24 24" width="10" height="10" aria-hidden="true">
          <path d="M7 10l5 5 5-5z" fill="currentColor" />
        </svg>
      </button>

      {isOpen && (
          <div ref={popupRef} className="model-popup">
           <div className="model-popup-inner">
            <div className="model-popup-search">
              <SearchIcon />
              <input
                ref={searchRef}
                type="text"
                placeholder="Search all models"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
              />
              {search && (
                <button
                  type="button"
                  className="model-popup-search-clear"
                  onClick={() => setSearch("")}
                  aria-label="Clear search"
                >
                  <CloseIcon />
                </button>
              )}
            </div>

            <div className="model-popup-scroll">
              <ModelRow
                label="Pi default"
                modelKey=""
                isSelected={!selectedModelKey}
                onSelect={() => handleSelect("")}
                onMouseEnter={() => {
                  setHoveredKey(null);
                }}
                onMouseLeave={handleRowMouseLeave}
              />

              {recentModels.length > 0 && (
                <>
                  <div className="model-popup-group-label">Recently Used</div>
                  {recentModels.map((entry) => (
                    <ModelRow
                      key={`recent-${entry.key}`}
                      label={entry.model.modelName}
                      modelKey={entry.key}
                      costTier={entry.model.costTier}
                      warningLabel={entry.model.requiresUnrecommendedWarning ? "Advanced" : undefined}
                      authStateLabel={authStateLabel(entry.model)}
                      isSelected={entry.key === selectedModelKey}
                      onSelect={() => handleSelect(entry.key)}
                      onMouseEnter={handleRowMouseEnter}
                      onMouseLeave={handleRowMouseLeave}
                    />
                  ))}
                </>
              )}

              {Array.from(groupedByProvider.entries()).map(([providerKey, entries]) => (
                <div key={providerKey}>
                  <div className="model-popup-group-label">{entries[0]!.provider.label}</div>
                  {entries.map((entry) => (
                    <ModelRow
                      key={entry.key}
                      label={entry.model.modelName}
                      modelKey={entry.key}
                      costTier={entry.model.costTier}
                      warningLabel={entry.model.requiresUnrecommendedWarning ? "Advanced" : undefined}
                      authStateLabel={authStateLabel(entry.model)}
                      isSelected={entry.key === selectedModelKey}
                      onSelect={() => handleSelect(entry.key)}
                      onMouseEnter={handleRowMouseEnter}
                      onMouseLeave={handleRowMouseLeave}
                    />
                  ))}
                </div>
              ))}

              {filtered.length === 0 && (
                <div className="model-popup-empty">No models match &quot;{search}&quot;</div>
              )}
            </div>
           </div>
            {/* Floating detail popover - appears on hover, anchored to hovered row */}
            {hoveredModel && (
              <div
                ref={detailRef}
                className="model-detail-popover"
                style={{ top: `${hoveredTop}px` }}
              onMouseEnter={handleDetailMouseEnter}
              onMouseLeave={handleDetailMouseLeave}
            >
              <div className="model-detail-popover-header">
                <strong>{hoveredModel.model.modelName}</strong>
                <button
                  type="button"
                  className="model-detail-pin"
                  onClick={() => setDetailPinned((c) => !c)}
                  aria-label={detailPinned ? "Unpin" : "Pin"}
                  title={detailPinned ? "Unpin detail" : "Pin detail open"}
                >
                  <svg viewBox="0 0 24 24" width="12" height="12" aria-hidden="true">
                    <path d="M16 12V4h1V2H7v2h1v8l-2 2v2h5.2v6h1.6v-6H18v-2l-2-2z" fill="currentColor" opacity={detailPinned ? 1 : 0.4} />
                  </svg>
                </button>
              </div>
              <p className="model-detail-meta">
                {hoveredModel.model.contextWindow
                  ? `${formatTokenCount(hoveredModel.model.contextWindow)} context window`
                  : ""}
                {hoveredModel.model.contextWindow && hoveredModel.model.costTier ? " · " : ""}
                {hoveredModel.model.costTier ?? ""}
              </p>
              {authStateLabel(hoveredModel.model) && (
                <p className="model-detail-auth-note">
                  {authStateDescription(hoveredModel.model)}
                </p>
              )}
              {hoveredModel.model.recommendationReason && (
                <p className="model-detail-auth-note">
                  {hoveredModel.model.recommendationReason}
                </p>
              )}
              {hoveredModel.model.requiresUnrecommendedWarning && (
                <p className="model-detail-auth-note">
                  Advanced catalog model. Pi-Office will ask before using it unless the warning is suppressed.
                </p>
              )}

              {hoveredModel.model.supportsThinking && availableThinkingLevels.length > 0 && (
                <div className="model-detail-section">
                  <span className="model-detail-section-label">Reasoning Effort</span>
                  <div className="model-detail-levels">
                    {availableThinkingLevels.map((level) => (
                      <button
                        key={level}
                        type="button"
                        className={`model-detail-level-btn ${thinkingLevel === level ? "model-detail-level-active" : ""}`}
                        onClick={() => onSetThinkingLevel(level)}
                      >
                        {level === "off" ? "Off" : level === "xhigh" ? "XHigh" : level.charAt(0).toUpperCase() + level.slice(1)}
                      </button>
                    ))}
                  </div>
                </div>
              )}
            </div>
          )}
          </div>
      )}
    </div>
  );
}

function ModelRow({
  label,
  modelKey,
  costTier,
  warningLabel,
  authStateLabel,
  isSelected,
  onSelect,
  onMouseEnter,
  onMouseLeave,
}: {
  label: string;
  modelKey: string;
  costTier?: string | undefined;
  warningLabel?: string | undefined;
  authStateLabel?: string | undefined;
  isSelected: boolean;
  onSelect: () => void;
  onMouseEnter: (key: string, el: HTMLElement) => void;
  onMouseLeave: () => void;
}) {
  const rowRef = useRef<HTMLButtonElement>(null);

  return (
    <button
      ref={rowRef}
      type="button"
      className={`model-popup-item ${isSelected ? "model-popup-item-selected" : ""}`}
      onClick={onSelect}
      onMouseEnter={() => { if (rowRef.current) onMouseEnter(modelKey, rowRef.current); }}
      onMouseLeave={onMouseLeave}
    >
      <span className="model-popup-item-name">{label}</span>
      <span className="model-popup-item-meta">
        {warningLabel && <span className="model-popup-auth-state">{warningLabel}</span>}
        {authStateLabel && <span className="model-popup-auth-state">{authStateLabel}</span>}
        {costTier && <span className="model-popup-cost">{costTier}</span>}
        {isSelected && <CheckIcon />}
      </span>
    </button>
  );
}

function authStateLabel(model: ProviderModelDescriptor): string | undefined {
  if (model.verifiedUsable) return undefined;
  if (model.authState === "verification_failed") return "Auth failed";
  if (model.credentialStored) return "Unverified";
  return "Locked";
}

function authStateDescription(model: ProviderModelDescriptor): string {
  if (model.authState === "verification_failed") {
    return model.verificationError
      ? `Stored credential failed verification: ${model.verificationError}`
      : "Stored credential failed verification. Check the key and save it again.";
  }
  if (model.credentialStored) {
    return "Credential is stored but has not succeeded with this provider yet.";
  }
  return "Add provider credentials before using this model.";
}
