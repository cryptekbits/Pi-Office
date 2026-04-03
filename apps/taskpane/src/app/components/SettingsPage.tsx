import { useState, useCallback, useEffect, useRef } from "react";
import type {
  AuthStatusResponse,
  ConnectorAuditPreference,
  ConnectorCatalogItem,
  ConnectorDiagnosticsResponse,
  ConnectorExportBundle,
  ConnectorFavoriteRequest,
  ConnectorImportApplyResponse,
  ConnectorImportPreviewResponse,
  ConnectorLogResponse,
  ConnectorPrepareResponse,
  ConnectorScopeContext,
  ConnectorScopeUpdateRequest,
  ConnectorSetupRequest,
  ImageModelDescriptor,
  ImageModelCatalogResponse,
  ImageReasoningEffort,
  ConnectorSetupResponse,
  ConnectorStatus,
  ConnectorTestResponse,
  OfficeMode,
  OfficeStateUpdate,
  ProviderDescriptor,
  SessionStatsResponse,
  UserPreferences,
  ThinkingLevel,
} from "@pi-office/pi-office-pack/protocol";
import {
  AUTONOMY_LEVELS,
  AUTONOMY_LEVEL_AUTO_APPROVE,
  AUTONOMY_LEVEL_LABELS,
  IMAGE_REASONING_EFFORTS,
  OFFICE_TOOL_NAMES,
  THINKING_LEVELS,
  TOOL_CATEGORY_MAP,
  type AutonomyLevel,
  type ToolCategory,
  type ToolPermissionOverride,
} from "@pi-office/pi-office-pack/protocol";
import { fetchJson } from "../../lib/api";
import { HOST_LABELS } from "@pi-office/pi-office-pack/defaults";
import { BackArrowIcon, UserIcon, ProviderIcon, ModelIcon, ToolIcon, PrefsIcon, IntegrationIcon } from "../../lib/icons";
import { formatTokenCount } from "../../lib/helpers";
import { IntegrationsSection } from "./IntegrationsSection";

type SettingsTab = "profile" | "providers" | "models" | "integrations" | "tools" | "preferences";

const TABS: { key: SettingsTab; label: string; Icon: () => React.JSX.Element }[] = [
  { key: "profile", label: "Profile", Icon: UserIcon },
  { key: "providers", label: "AI Providers", Icon: ProviderIcon },
  { key: "models", label: "Models", Icon: ModelIcon },
  { key: "integrations", label: "Integrations", Icon: IntegrationIcon },
  { key: "tools", label: "Tools", Icon: ToolIcon },
  { key: "preferences", label: "Preferences", Icon: PrefsIcon },
];

const TOOL_DESCRIPTIONS: Record<string, string> = {
  office_get_context: "Read document content, selection, and metadata from the active Office document.",
  office_apply_edit: "Apply text edits, formatting changes, and insertions to the active document.",
  office_navigate: "Navigate to specific locations within the document (sections, pages, ranges).",
  office_capture_snapshot: "Capture a visual screenshot of the current document view.",
  office_capture_viewport: "Capture the visible viewport area of the document window.",
  office_read_section: "Read a paginated section of the document by paragraph index.",
  office_execute_js: "Execute arbitrary Office.js code in the document context.",
  office_propose_edits: "Propose batch edits for user review before applying.",
  read: "Read file contents from the workspace.",
  grep: "Search file contents with regex patterns.",
  find: "Find files by name or glob pattern.",
  ls: "List directory contents.",
  edit: "Edit file contents in the workspace.",
  write: "Write new file contents to the workspace.",
  bash: "Execute shell commands in the workspace.",
};

const CATEGORY_LABELS: Record<string, string> = {
  read: "Read",
  "write-doc": "Write (doc)",
  "read-external": "Read (workspace)",
  "write-external": "Write (workspace)",
  connector: "Connector",
  interaction: "Interaction",
};

const WORKSPACE_TOOL_NAMES = ["read", "grep", "find", "ls", "edit", "write", "bash"] as const;

interface SettingsPageProps {
  officeState: OfficeStateUpdate | undefined;
  mode: OfficeMode;
  providers: ProviderDescriptor[];
  authStatus: AuthStatusResponse | undefined;
  connectors: ConnectorCatalogItem[];
  connectorStatuses: ConnectorStatus[];
  connectorDiagnostics: ConnectorDiagnosticsResponse | undefined;
  connectorAuditPreference: ConnectorAuditPreference | undefined;
  connectorScopeContext: ConnectorScopeContext | undefined;
  sessionStats: SessionStatsResponse | undefined;
  preferences: UserPreferences;
  enabledModels: Set<string>;
  enabledProviders: Set<string>;
  onClose: () => void;
  onUpdatePreferences: (patch: Partial<UserPreferences>) => void;
  onToggleModel: (key: string) => void;
  onToggleProvider: (provider: string) => void;
  onSaveApiKey: (provider: string, key: string) => void;
  onStartOAuth: (provider: string) => void;
  onClearAuth: (provider: string) => void;
  onPrepareConnector: (connectorId: string, scopeContext?: ConnectorScopeContext) => Promise<ConnectorPrepareResponse>;
  onConnectConnector: (request: ConnectorSetupRequest) => Promise<ConnectorSetupResponse>;
  onTestConnector: (request: ConnectorSetupRequest) => Promise<ConnectorTestResponse>;
  onReverifyConnector: (connectorId: string, scopeContext?: ConnectorScopeContext) => Promise<ConnectorTestResponse>;
  onStartConnectorOAuth: (connectorId: string) => Promise<void>;
  onRemoveConnector: (storedConnectorId: string) => Promise<void>;
  onSetConnectorFavorite: (request: ConnectorFavoriteRequest) => Promise<void>;
  onUpdateConnectorScope: (request: ConnectorScopeUpdateRequest) => Promise<void>;
  onLoadConnectorLogs: (connectorId: string) => Promise<ConnectorLogResponse>;
  onExportConnectors: () => Promise<ConnectorExportBundle>;
  onPreviewConnectorImport: (bundle: ConnectorExportBundle) => Promise<ConnectorImportPreviewResponse>;
  onApplyConnectorImport: (bundle: ConnectorExportBundle, resolutions?: Record<string, "skip" | "replace">) => Promise<ConnectorImportApplyResponse>;
  onSetConnectorAuditPreference: (preference: ConnectorAuditPreference) => Promise<ConnectorAuditPreference>;
}

export function SettingsPage({
  officeState,
  mode,
  providers,
  authStatus,
  connectors,
  connectorStatuses,
  connectorDiagnostics,
  connectorAuditPreference,
  connectorScopeContext,
  sessionStats,
  preferences,
  enabledModels,
  enabledProviders,
  onClose,
  onUpdatePreferences,
  onToggleModel,
  onToggleProvider,
  onSaveApiKey,
  onStartOAuth,
  onClearAuth,
  onPrepareConnector,
  onConnectConnector,
  onTestConnector,
  onReverifyConnector,
  onStartConnectorOAuth,
  onRemoveConnector,
  onSetConnectorFavorite,
  onUpdateConnectorScope,
  onLoadConnectorLogs,
  onExportConnectors,
  onPreviewConnectorImport,
  onApplyConnectorImport,
  onSetConnectorAuditPreference,
}: SettingsPageProps) {
  const [activeTab, setActiveTab] = useState<SettingsTab>("profile");
  const pageRef = useRef<HTMLDivElement | null>(null);
  const previousFocusRef = useRef<HTMLElement | null>(null);

  useEffect(() => {
    previousFocusRef.current = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const focusTarget = pageRef.current?.querySelector<HTMLElement>("[data-settings-initial-focus]");
    focusTarget?.focus();

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "Tab" || !pageRef.current) {
        return;
      }
      const focusable = Array.from(
        pageRef.current.querySelectorAll<HTMLElement>(
          'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])',
        ),
      ).filter((entry) => !entry.hasAttribute("disabled") && entry.offsetParent !== null);
      if (!focusable.length) {
        return;
      }
      const first = focusable[0];
      const last = focusable[focusable.length - 1];
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last?.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first?.focus();
      }
    };

    document.addEventListener("keydown", handleKeyDown);
    return () => {
      document.removeEventListener("keydown", handleKeyDown);
      previousFocusRef.current?.focus();
    };
  }, []);

  return (
    <div ref={pageRef} className="settings-page" role="dialog" aria-modal="true" aria-label="Settings">
      <header className="settings-page-header">
        <button type="button" className="icon-button" aria-label="Back to chat" onClick={onClose} data-settings-initial-focus>
          <BackArrowIcon />
        </button>
        <h2>Settings</h2>
      </header>

      <div className="settings-page-body">
        <nav className="settings-tabs" role="tablist" aria-label="Settings sections">
          {TABS.map(({ key, label, Icon }) => (
            <button
              key={key}
              type="button"
              role="tab"
              aria-selected={activeTab === key}
              className={`settings-tab ${activeTab === key ? "settings-tab-active" : ""}`}
              onClick={() => setActiveTab(key)}
            >
              <Icon />
              <span>{label}</span>
            </button>
          ))}
        </nav>

        <div className="settings-content">
          {activeTab === "profile" && (
            <ProfileSection officeState={officeState} mode={mode} sessionStats={sessionStats} />
          )}
          {activeTab === "providers" && (
            <ProvidersSection
              providers={providers}
              authStatus={authStatus}
              enabledProviders={enabledProviders}
              onToggleProvider={onToggleProvider}
              onSaveApiKey={onSaveApiKey}
              onStartOAuth={onStartOAuth}
              onClearAuth={onClearAuth}
            />
          )}
          {activeTab === "models" && (
            <ModelsSection
              providers={providers}
              enabledModels={enabledModels}
              enabledProviders={enabledProviders}
              onToggleModel={onToggleModel}
            />
          )}
          {activeTab === "integrations" && (
            <IntegrationsSection
              host={officeState?.host}
              connectors={connectors}
              statuses={connectorStatuses}
              diagnostics={connectorDiagnostics}
              auditPreference={connectorAuditPreference}
              scopeContext={connectorScopeContext}
              onPrepareConnector={onPrepareConnector}
              onConnectConnector={onConnectConnector}
              onTestConnector={onTestConnector}
              onReverifyConnector={onReverifyConnector}
              onStartOAuth={onStartConnectorOAuth}
              onRemoveConnector={onRemoveConnector}
              onSetFavorite={onSetConnectorFavorite}
              onUpdateScope={onUpdateConnectorScope}
              onLoadLogs={onLoadConnectorLogs}
              onExportConnectors={onExportConnectors}
              onPreviewImport={onPreviewConnectorImport}
              onApplyImport={onApplyConnectorImport}
              onSetAuditPreference={onSetConnectorAuditPreference}
            />
          )}
          {activeTab === "tools" && (
            <ToolsSection
              mode={mode}
              preferences={preferences}
              onUpdatePreferences={onUpdatePreferences}
            />
          )}
          {activeTab === "preferences" && (
            <PreferencesSection preferences={preferences} onUpdate={onUpdatePreferences} />
          )}
        </div>
      </div>
    </div>
  );
}

function ProfileSection({
  officeState,
  mode,
  sessionStats,
}: {
  officeState: OfficeStateUpdate | undefined;
  mode: OfficeMode;
  sessionStats: SessionStatsResponse | undefined;
}) {
  return (
    <div className="settings-section">
      <h3>Session Info</h3>
      <div className="settings-card-grid">
        <div className="settings-card">
          <span className="label">Host</span>
          <p>{officeState ? HOST_LABELS[officeState.host] : "Not connected"}</p>
        </div>
        <div className="settings-card">
          <span className="label">Document</span>
          <p>{officeState?.document.title ?? "Unknown"}</p>
        </div>
        <div className="settings-card">
          <span className="label">Mode</span>
          <p>{mode === "workspace" ? "Workspace" : "Document-only"}</p>
        </div>
        {officeState?.document.workspaceDir && (
          <div className="settings-card">
            <span className="label">Workspace</span>
            <p>{officeState.document.workspaceDir}</p>
          </div>
        )}
      </div>

      {sessionStats && (
        <>
          <h3>Session Stats</h3>
          <div className="settings-card-grid">
            <div className="settings-card">
              <span className="label">Messages</span>
              <p>{sessionStats.userMessages} user / {sessionStats.assistantMessages} assistant</p>
            </div>
            <div className="settings-card">
              <span className="label">Tool Calls</span>
              <p>{sessionStats.toolCalls}</p>
            </div>
            <div className="settings-card">
              <span className="label">Tokens</span>
              <p>{formatTokenCount(sessionStats.tokens.total)} total</p>
            </div>
            <div className="settings-card">
              <span className="label">Cost</span>
              <p>${sessionStats.cost.toFixed(4)}</p>
            </div>
            {sessionStats.contextUsage && (
              <div className="settings-card">
                <span className="label">Context Usage</span>
                <p>
                  {formatTokenCount(sessionStats.contextUsage.tokens)} /{" "}
                  {formatTokenCount(sessionStats.contextUsage.contextWindow)}
                  {sessionStats.contextUsage.percent != null &&
                    ` (${sessionStats.contextUsage.percent.toFixed(1)}%)`}
                </p>
              </div>
            )}
          </div>
        </>
      )}
    </div>
  );
}

function ProvidersSection({
  providers,
  authStatus,
  enabledProviders,
  onToggleProvider,
  onSaveApiKey,
  onStartOAuth,
  onClearAuth,
}: {
  providers: ProviderDescriptor[];
  authStatus: AuthStatusResponse | undefined;
  enabledProviders: Set<string>;
  onToggleProvider: (provider: string) => void;
  onSaveApiKey: (provider: string, key: string) => void;
  onStartOAuth: (provider: string) => void;
  onClearAuth: (provider: string) => void;
}) {
  const stored = authStatus?.storedProviders ?? [];

  return (
    <div className="settings-section">
      <h3>AI Providers</h3>
      <div className="provider-list">
        {providers.map((provider) => (
          <ProviderCard
            key={provider.provider}
            isEnabled={enabledProviders.has(provider.provider)}
            onToggleEnabled={() => onToggleProvider(provider.provider)}
            provider={provider}
            hasAuth={stored.includes(provider.provider)}
            onSaveApiKey={(key) => onSaveApiKey(provider.provider, key)}
            onStartOAuth={() => onStartOAuth(provider.provider)}
            onClearAuth={() => onClearAuth(provider.provider)}
          />
        ))}
        {providers.length === 0 && (
          <p className="settings-note">No providers discovered. Ensure the companion server is running.</p>
        )}
      </div>
    </div>
  );
}

function ProviderCard({
  provider,
  hasAuth,
  isEnabled,
  onToggleEnabled,
  onSaveApiKey,
  onStartOAuth,
  onClearAuth,
}: {
  provider: ProviderDescriptor;
  hasAuth: boolean;
  isEnabled: boolean;
  onToggleEnabled: () => void;
  onSaveApiKey: (key: string) => void;
  onStartOAuth: () => void;
  onClearAuth: () => void;
}) {
  const [apiKey, setApiKey] = useState("");
  const [expanded, setExpanded] = useState(false);
  const readyCount = provider.models.filter((m) => m.configured).length;

  const handleSave = useCallback(() => {
    if (apiKey.trim()) {
      onSaveApiKey(apiKey.trim());
      setApiKey("");
    }
  }, [apiKey, onSaveApiKey]);

  return (
    <div className={`provider-card ${!isEnabled ? "provider-card-disabled" : ""}`}>
      <div className="provider-card-header">
        <button
          type="button"
          className="provider-card-header-left"
          onClick={() => setExpanded((c) => !c)}
        >
          <div className="provider-card-info">
            <strong>{provider.label}</strong>
            <span className={`provider-status ${provider.configured ? "provider-status-ready" : ""}`}>
              {provider.configured ? "Ready" : "Not configured"}
            </span>
          </div>
          <span className="provider-card-meta">
            {readyCount}/{provider.models.length} models
          </span>
        </button>
        <button
          type="button"
          className={`toggle-switch toggle-sm ${isEnabled ? "toggle-on" : ""}`}
          onClick={onToggleEnabled}
          role="switch"
          aria-checked={isEnabled}
          aria-label={`${isEnabled ? "Disable" : "Enable"} ${provider.label}`}
        >
          <span className="toggle-thumb" />
        </button>
      </div>

      {expanded && (
        <div className="provider-card-body">
          <div className="provider-auth-status">
            {hasAuth ? "Stored credentials found." : "No stored credentials."}
          </div>

          <div className="field">
            <span>API Key</span>
            <input
              type="password"
              value={apiKey}
              onChange={(e) => setApiKey(e.target.value)}
              placeholder={provider.oauthSupported ? "Optional key override" : "Paste provider API key"}
            />
          </div>

          <div className="settings-actions">
            <button type="button" className="button button-solid" onClick={handleSave}>
              Save Key
            </button>
            {provider.oauthSupported && (
              <button type="button" className="button" onClick={onStartOAuth}>
                Start OAuth
              </button>
            )}
            <button type="button" className="button" onClick={onClearAuth}>
              Clear Auth
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

function ModelsSection({
  providers,
  enabledModels,
  enabledProviders,
  onToggleModel,
}: {
  providers: ProviderDescriptor[];
  enabledModels: Set<string>;
  enabledProviders: Set<string>;
  onToggleModel: (key: string) => void;
}) {
  const activeProviders = providers.filter((p) => enabledProviders.has(p.provider));

  return (
    <div className="settings-section">
      <h3>Available Models</h3>
      <p className="settings-note">
        Toggle models on or off to control which appear in the model selector.
        Only models from enabled providers are shown.
      </p>
      {activeProviders.length === 0 && (
        <p className="settings-note">Enable a provider in the Providers tab first.</p>
      )}
      {activeProviders.map((provider) => (
        <div key={provider.provider} className="settings-model-group">
          <h4>{provider.label}</h4>
          <div className="settings-model-list">
            {provider.models.map((model) => {
              const modelKey = `${model.provider}::${model.modelId}`;
              const isEnabled = enabledModels.has(modelKey);
              const isReady = model.configured;
              return (
                <div
                  key={modelKey}
                  className={`settings-model-row ${!isEnabled ? "settings-model-row-disabled" : ""}`}
                >
                  <div className="settings-model-info">
                    <span>{model.modelName}</span>
                    <span className="settings-model-badges">
                      {model.contextWindow ? (
                        <span className="settings-model-ctx">
                          {Math.round(model.contextWindow / 1000)}K
                        </span>
                      ) : null}
                      {model.costTier ? (
                        <span className="settings-model-cost">{model.costTier}</span>
                      ) : null}
                      {model.supportsThinking ? (
                        <span className="settings-model-tag">Reasoning</span>
                      ) : null}
                    </span>
                  </div>
                  <div className="settings-model-actions">
                    {!isReady && (
                      <span className="settings-model-state">Locked</span>
                    )}
                    <button
                      type="button"
                      className={`toggle-switch toggle-sm ${isEnabled ? "toggle-on" : ""}`}
                      onClick={() => onToggleModel(modelKey)}
                      role="switch"
                      aria-checked={isEnabled}
                      aria-label={`${isEnabled ? "Disable" : "Enable"} ${model.modelName}`}
                    >
                      <span className="toggle-thumb" />
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      ))}
    </div>
  );
}

function ToolsSection({
  mode,
  preferences,
  onUpdatePreferences,
}: {
  mode: OfficeMode;
  preferences: UserPreferences;
  onUpdatePreferences: (patch: Partial<UserPreferences>) => void;
}) {
  const handleOverrideChange = useCallback(
    (toolName: string, level: AutonomyLevel | "default" | "disabled") => {
      const current = preferences.toolPermissionOverrides;
      if (level === "default") {
        onUpdatePreferences({
          toolPermissionOverrides: current.filter((o) => o.toolName !== toolName),
        });
      } else {
        const next = current.filter((o) => o.toolName !== toolName);
        next.push({ toolName, autoApproveAtLevel: level });
        onUpdatePreferences({ toolPermissionOverrides: next });
      }
    },
    [onUpdatePreferences, preferences.toolPermissionOverrides],
  );

  const getOverrideLevel = useCallback(
    (toolName: string): AutonomyLevel | "default" | "disabled" => {
      const override = preferences.toolPermissionOverrides.find((o) => o.toolName === toolName);
      return override ? override.autoApproveAtLevel : "default";
    },
    [preferences.toolPermissionOverrides],
  );

  return (
    <div className="settings-section">
      <div className="settings-section-header">
        <h3>Autonomy Level</h3>
        <button
          type="button"
          className="button button-sm button-reset"
          onClick={() => onUpdatePreferences({ autonomyLevel: "medium", toolPermissionOverrides: [] })}
        >
          Reset defaults
        </button>
      </div>
      <p className="settings-note">
        Current level: <strong>{AUTONOMY_LEVEL_LABELS[preferences.autonomyLevel]}</strong>.
        Change it from the composer bar or set per-tool overrides below.
      </p>
      <div className="segmented-control segmented-control-wide">
        {AUTONOMY_LEVELS.map((level) => (
          <button
            key={level}
            type="button"
            className={`segmented-item ${preferences.autonomyLevel === level ? "segmented-active" : ""}`}
            onClick={() => onUpdatePreferences({ autonomyLevel: level })}
          >
            {AUTONOMY_LEVEL_LABELS[level]}
          </button>
        ))}
      </div>

      <h3>Office Tools</h3>
      <div className="tool-list">
        {OFFICE_TOOL_NAMES.map((toolName) => (
          <ToolCardWithOverride
            key={toolName}
            toolName={toolName}
            overrideLevel={getOverrideLevel(toolName)}
            onOverrideChange={handleOverrideChange}
          />
        ))}
      </div>

      {mode === "workspace" && (
        <>
          <h3>Workspace Tools</h3>
          <p className="settings-note">
            Coding tools are available in workspace mode when the document is saved.
          </p>
          <div className="tool-list">
            {WORKSPACE_TOOL_NAMES.map((toolName) => (
              <ToolCardWithOverride
                key={toolName}
                toolName={toolName}
                overrideLevel={getOverrideLevel(toolName)}
                onOverrideChange={handleOverrideChange}
              />
            ))}
          </div>
        </>
      )}
    </div>
  );
}

function getDefaultAutoApproveLevel(category: ToolCategory): AutonomyLevel {
  for (const level of AUTONOMY_LEVELS) {
    if (AUTONOMY_LEVEL_AUTO_APPROVE[level].has(category)) return level;
  }
  return "extreme";
}

function ToolCardWithOverride({
  toolName,
  overrideLevel,
  onOverrideChange,
}: {
  toolName: string;
  overrideLevel: AutonomyLevel | "default" | "disabled";
  onOverrideChange: (toolName: string, level: AutonomyLevel | "default" | "disabled") => void;
}) {
  const category = (TOOL_CATEGORY_MAP[toolName] ?? "connector") as ToolCategory;
  const defaultLevel = getDefaultAutoApproveLevel(category);
  const isDisabled = overrideLevel === "disabled";
  const activeLevel = isDisabled ? defaultLevel : overrideLevel === "default" ? defaultLevel : overrideLevel;
  const activeIndex = AUTONOMY_LEVELS.indexOf(activeLevel);

  return (
    <div className={`tool-card ${isDisabled ? "tool-card-disabled" : ""}`}>
      <div className="tool-card-header">
        <code>{toolName}</code>
        <div className="tool-card-header-right">
          {isDisabled && <span className="tool-card-disabled-badge">Disabled</span>}
          <span className="tool-card-category">{CATEGORY_LABELS[category] ?? category}</span>
        </div>
      </div>
      <p className="tool-card-desc">
        {TOOL_DESCRIPTIONS[toolName] ?? "No description available."}
      </p>
      <div className="tool-card-override">
        <span className="tool-card-override-label">Auto-approve at:</span>
        <div className="tool-level-bar">
          {!isDisabled && (
            <div
              className="tool-level-bar-indicator"
              style={{ left: `${(activeIndex / AUTONOMY_LEVELS.length) * 100}%` }}
            />
          )}
          {AUTONOMY_LEVELS.map((level) => {
            const isDefault = level === defaultLevel;
            const isActive = !isDisabled && level === activeLevel;
            return (
              <button
                key={level}
                type="button"
                disabled={isDisabled}
                className={[
                  "tool-level-bar-item",
                  isActive ? "tool-level-bar-item-active" : "",
                  isDefault && !isActive && !isDisabled ? "tool-level-bar-item-default" : "",
                ].filter(Boolean).join(" ")}
                onClick={() => {
                  if (level === defaultLevel) {
                    onOverrideChange(toolName, "default");
                  } else {
                    onOverrideChange(toolName, level);
                  }
                }}
                title={isDefault ? `${AUTONOMY_LEVEL_LABELS[level]} (default for ${category})` : AUTONOMY_LEVEL_LABELS[level]}
              >
                {AUTONOMY_LEVEL_LABELS[level]}
              </button>
            );
          })}
        </div>
      </div>
      <div className="tool-card-disable-toggle">
        <button
          type="button"
          className={`toggle-switch toggle-sm ${!isDisabled ? "toggle-on" : ""}`}
          onClick={() => onOverrideChange(toolName, isDisabled ? "default" : "disabled")}
          role="switch"
          aria-checked={!isDisabled}
          aria-label={`${isDisabled ? "Enable" : "Disable"} ${toolName}`}
        >
          <span className="toggle-thumb" />
        </button>
        <span className="tool-card-disable-label">{isDisabled ? "Tool disabled" : "Enabled"}</span>
      </div>
    </div>
  );
}

function PreferencesSection({
  preferences,
  onUpdate,
}: {
  preferences: UserPreferences;
  onUpdate: (patch: Partial<UserPreferences>) => void;
}) {
  return (
    <div className="settings-section">
      <h3>Chat Behavior</h3>
      <div className="pref-list">
        <PrefToggle
          label="Show thinking traces"
          description="Display the model's internal reasoning steps in chat messages."
          value={preferences.showThinkingTraces}
          onChange={(v) => onUpdate({ showThinkingTraces: v })}
        />
        <PrefToggle
          label="Auto-attach visual context"
          description="Automatically capture and send visual context when your prompt mentions formatting, images, or layout."
          value={preferences.autoAttachVisuals}
          onChange={(v) => onUpdate({ autoAttachVisuals: v })}
        />
        <PrefToggle
          label="Show token usage"
          description="Display token counts and cost information in the session."
          value={preferences.showTokenUsage}
          onChange={(v) => onUpdate({ showTokenUsage: v })}
        />
        <PrefToggle
          label="Compact messages"
          description="Use a denser message layout to fit more conversation on screen."
          value={preferences.compactMessages}
          onChange={(v) => onUpdate({ compactMessages: v })}
        />
        <PrefToggle
          label="Persistent rewind snapshots (Beta)"
          description="Save document checkpoints to disk for multi-session rewind. Uses additional disk space."
          value={preferences.experimentalRewindSnapshots}
          onChange={(v) => onUpdate({ experimentalRewindSnapshots: v })}
        />
      </div>

      <h3>Default Thinking Level</h3>
      <p className="settings-note">
        Controls how much effort the model spends on thinking/reasoning for new conversations.
        &quot;Off&quot; disables thinking. Available levels depend on the active model.
      </p>
      <div className="segmented-control segmented-control-wide">
        {THINKING_LEVELS.map((level) => (
          <button
            key={level}
            type="button"
            className={`segmented-item ${preferences.defaultThinkingLevel === level ? "segmented-active" : ""}`}
            onClick={() => onUpdate({ defaultThinkingLevel: level })}
          >
            {level === "off" ? "Off" : level === "xhigh" ? "XHigh" : level.charAt(0).toUpperCase() + level.slice(1)}
          </button>
        ))}
      </div>

      <ImageGenerationSettings preferences={preferences} onUpdate={onUpdate} />
    </div>
  );
}

function ImageGenerationSettings({
  preferences,
  onUpdate,
}: {
  preferences: UserPreferences;
  onUpdate: (patch: Partial<UserPreferences>) => void;
}) {
  const [imageModels, setImageModels] = useState<ImageModelDescriptor[]>([]);
  const [defaultModelKey, setDefaultModelKey] = useState("");

  useEffect(() => {
    let active = true;
    void (async () => {
      try {
        const catalog = await fetchJson<ImageModelCatalogResponse>("/v1/image-models");
        if (!active) return;
        setImageModels(catalog.models);
        setDefaultModelKey(catalog.defaultModelKey);
      } catch {
        // non-critical
      }
    })();
    return () => { active = false; };
  }, []);

  const configuredModels = imageModels.filter((m) => m.configured);
  const selectedModel = imageModels.find(
    (m) => `${m.provider}::${m.modelId}` === (preferences.defaultImageModel || defaultModelKey),
  );

  return (
    <>
      <h3>Image Generation</h3>
      <div className="pref-list">
        <PrefToggle
          label="Enable image generation"
          description="Allow Pi to generate images using AI models and insert them into your documents."
          value={preferences.imageGenerationEnabled}
          onChange={(v) => onUpdate({ imageGenerationEnabled: v })}
        />
      </div>

      {preferences.imageGenerationEnabled && (
        <>
          <p className="settings-note">
            Choose the default model for image generation. Only providers with stored credentials are available.
          </p>

          {configuredModels.length > 0 ? (
            <div className="field" style={{ marginBottom: "12px" }}>
              <span>Default Image Model</span>
              <select
                value={preferences.defaultImageModel || defaultModelKey}
                onChange={(e) => onUpdate({ defaultImageModel: e.target.value })}
              >
                <option value="">Auto (best available)</option>
                {configuredModels.map((model) => {
                  const key = `${model.provider}::${model.modelId}`;
                  return (
                    <option key={key} value={key}>
                      {model.modelName}
                    </option>
                  );
                })}
              </select>
            </div>
          ) : (
            <p className="settings-note">
              No image models available. Configure an AI provider (OpenAI, Google, or OpenRouter) in the Providers tab.
            </p>
          )}

          {selectedModel?.supportsReasoningEffort !== false && (
            <>
              <p className="settings-note">
                Controls how much effort the image model spends on reasoning before generating.
                Higher effort produces better results but takes longer.
              </p>
              <div className="segmented-control">
                {IMAGE_REASONING_EFFORTS.map((level) => (
                  <button
                    key={level}
                    type="button"
                    className={`segmented-item ${preferences.imageReasoningEffort === level ? "segmented-active" : ""}`}
                    onClick={() => onUpdate({ imageReasoningEffort: level })}
                  >
                    {level.charAt(0).toUpperCase() + level.slice(1)}
                  </button>
                ))}
              </div>
            </>
          )}
        </>
      )}
    </>
  );
}

function PrefToggle({
  label,
  description,
  value,
  onChange,
}: {
  label: string;
  description: string;
  value: boolean;
  onChange: (next: boolean) => void;
}) {
  return (
    <div className="pref-row">
      <div className="pref-text">
        <strong>{label}</strong>
        <p>{description}</p>
      </div>
      <button
        type="button"
        className={`toggle-switch ${value ? "toggle-on" : ""}`}
        onClick={() => onChange(!value)}
        role="switch"
        aria-checked={value}
      >
        <span className="toggle-thumb" />
      </button>
    </div>
  );
}
