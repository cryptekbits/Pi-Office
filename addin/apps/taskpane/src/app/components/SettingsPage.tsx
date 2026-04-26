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
  ConnectorOAuthStartResponse,
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
  CompanionState,
  OfficeDocumentState,
  OfficeStateUpdate,
  ProviderAuthMethod,
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
import { fetchJson, type RuntimeRequestFailureDiagnostic } from "../../lib/api";
import { HOST_LABELS } from "@pi-office/pi-office-pack/defaults";
import {
  BackArrowIcon,
  UserIcon,
  ProviderIcon,
  ModelIcon,
  ToolIcon,
  PrefsIcon,
  IntegrationIcon,
  DiagnosticsIcon,
  ShieldIcon,
} from "../../lib/icons";
import { formatTokenCount } from "../../lib/helpers";
import { IntegrationsSection } from "./IntegrationsSection";

type SettingsTab = "profile" | "companion" | "providers" | "models" | "integrations" | "privacy" | "tools" | "preferences" | "diagnostics";

const TABS: { key: SettingsTab; label: string; Icon: () => React.JSX.Element }[] = [
  { key: "profile", label: "Profile", Icon: UserIcon },
  { key: "companion", label: "Companion", Icon: IntegrationIcon },
  { key: "providers", label: "AI Providers", Icon: ProviderIcon },
  { key: "models", label: "Models", Icon: ModelIcon },
  { key: "integrations", label: "Integrations", Icon: IntegrationIcon },
  { key: "privacy", label: "Privacy", Icon: ShieldIcon },
  { key: "diagnostics", label: "Diagnostics", Icon: DiagnosticsIcon },
  { key: "tools", label: "Tools", Icon: ToolIcon },
  { key: "preferences", label: "Preferences", Icon: PrefsIcon },
];

const TOOL_DESCRIPTIONS: Record<string, string> = {
  office_get_context: "Read document content, selection, and metadata from the active Office document.",
  office_apply_edit: "Apply text edits, formatting changes, and insertions to the active document.",
  office_navigate: "Navigate to specific locations within the document (sections, pages, ranges).",
  office_capture_snapshot: "Capture selection and document context metadata from the current document state.",
  office_capture_viewport: "Capture Word viewport metadata from Office.js context (not a pixel-perfect OS/window screenshot).",
  office_read_section: "Read a paginated section of the document by paragraph index.",
  office_execute_js: "Manual-only escape hatch for direct Office.js snippets when structured tools are insufficient.",
  office_propose_edits: "Propose batch edits for user review before applying.",
  read: "Read file contents from the saved document folder through the optional companion.",
  grep: "Search file contents from the saved document folder through the optional companion.",
  find: "Find files by name inside the saved document folder through the optional companion.",
  ls: "List directory contents from the saved document folder through the optional companion.",
  mcp: "Execute a verified read-only local MCP tool through the optional companion.",
  bash: "Execute a companion-sandboxed shell command only after isolation detection and destructive probes pass.",
};

const CATEGORY_LABELS: Record<string, string> = {
  read: "Read",
  "write-doc": "Write (doc)",
  "escape-hatch": "Manual only",
  "read-external": "Read (workspace)",
  "write-external": "Write (workspace)",
  connector: "Connector",
  interaction: "Interaction",
};

const COMPANION_FILE_TOOL_NAMES = ["read", "grep", "find", "ls"] as const;

interface SettingsPageProps {
  officeState: OfficeStateUpdate | undefined;
  documentState: OfficeDocumentState;
  companion: CompanionState;
  providers: ProviderDescriptor[];
  authStatus: AuthStatusResponse | undefined;
  connectors: ConnectorCatalogItem[];
  connectorStatuses: ConnectorStatus[];
  connectorDiagnostics: ConnectorDiagnosticsResponse | undefined;
  connectorAuditPreference: ConnectorAuditPreference | undefined;
  connectorScopeContext: ConnectorScopeContext | undefined;
  runtimeDiagnostics: RuntimeRequestFailureDiagnostic[];
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
  onStartConnectorOAuth: (connectorId: string) => Promise<ConnectorOAuthStartResponse>;
  onRemoveConnector: (storedConnectorId: string) => Promise<void>;
  onSetConnectorFavorite: (request: ConnectorFavoriteRequest) => Promise<void>;
  onUpdateConnectorScope: (request: ConnectorScopeUpdateRequest) => Promise<void>;
  onLoadConnectorLogs: (connectorId: string) => Promise<ConnectorLogResponse>;
  onExportConnectors: () => Promise<ConnectorExportBundle>;
  onPreviewConnectorImport: (bundle: ConnectorExportBundle) => Promise<ConnectorImportPreviewResponse>;
  onApplyConnectorImport: (bundle: ConnectorExportBundle, resolutions?: Record<string, "skip" | "replace">) => Promise<ConnectorImportApplyResponse>;
  onSetConnectorAuditPreference: (preference: ConnectorAuditPreference) => Promise<ConnectorAuditPreference>;
  onClearAllProviderAuth: () => Promise<void>;
  onClearConnectorData: () => Promise<void>;
  onClearChatHistory: () => void;
  onRetryCompanion: () => Promise<void> | void;
  onSaveCompanionEndpoint: (endpoint: string) => Promise<void> | void;
  onClearRuntimeDiagnostics: () => void;
}

export function SettingsPage({
  officeState,
  documentState,
  companion,
  providers,
  authStatus,
  connectors,
  connectorStatuses,
  connectorDiagnostics,
  connectorAuditPreference,
  connectorScopeContext,
  runtimeDiagnostics,
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
  onClearAllProviderAuth,
  onClearConnectorData,
  onClearChatHistory,
  onRetryCompanion,
  onSaveCompanionEndpoint,
  onClearRuntimeDiagnostics,
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
            <ProfileSection officeState={officeState} documentState={documentState} companion={companion} sessionStats={sessionStats} />
          )}
          {activeTab === "companion" && (
            <CompanionSection
              companion={companion}
              onRetryCompanion={onRetryCompanion}
              onSaveCompanionEndpoint={onSaveCompanionEndpoint}
            />
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
          {activeTab === "privacy" && (
            <PrivacySection
              onClearAllProviderAuth={onClearAllProviderAuth}
              onClearConnectorData={onClearConnectorData}
              onClearChatHistory={onClearChatHistory}
            />
          )}
          {activeTab === "diagnostics" && (
            <DiagnosticsSection
              diagnostics={runtimeDiagnostics}
              onClear={onClearRuntimeDiagnostics}
            />
          )}
          {activeTab === "tools" && (
            <ToolsSection
              documentState={documentState}
              companion={companion}
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

const DIAGNOSTIC_SOURCE_LABELS: Record<RuntimeRequestFailureDiagnostic["source"], string> = {
  route: "Route",
  auth: "Auth",
  connector: "Connector",
};

function PrivacySection({
  onClearAllProviderAuth,
  onClearConnectorData,
  onClearChatHistory,
}: {
  onClearAllProviderAuth: () => Promise<void>;
  onClearConnectorData: () => Promise<void>;
  onClearChatHistory: () => void;
}) {
  const confirmAndRun = useCallback((message: string, action: () => Promise<void> | void) => {
    if (window.confirm(message)) {
      void action();
    }
  }, []);

  return (
    <section className="settings-section privacy-section">
      <h3>Privacy & Storage</h3>
      <p className="settings-note">
        Pi-Office keeps the Office bridge local to the taskpane. Provider and connector requests leave the taskpane only when a configured runtime or connector is used.
      </p>

      <div className="privacy-disclosure-list">
        <article className="privacy-disclosure-row">
          <div>
            <strong>Provider calls</strong>
            <p>Prompts, selected Office context, generated images, and tool results are sent to the selected AI provider when a request runs. Provider credentials are stored in this browser origin.</p>
          </div>
          <button
            type="button"
            className="button"
            onClick={() => confirmAndRun("Clear all stored provider credentials from this taskpane?", onClearAllProviderAuth)}
          >
            Clear provider auth
          </button>
        </article>

        <article className="privacy-disclosure-row">
          <div>
            <strong>Connectors</strong>
            <p>Connector calls can contact external services or the optional local companion. Connector config, secrets, OAuth handoffs, scope settings, and redacted audit entries are stored locally.</p>
          </div>
          <button
            type="button"
            className="button"
            onClick={() => confirmAndRun("Remove all connector configuration, secrets, OAuth state, scopes, and logs?", onClearConnectorData)}
          >
            Clear connectors
          </button>
        </article>

        <article className="privacy-disclosure-row">
          <div>
            <strong>Chat history</strong>
            <p>Saved chats may include prompts, document snippets, model responses, and metadata for the current Office host. They are stored in localStorage for this taskpane origin.</p>
          </div>
          <button
            type="button"
            className="button"
            onClick={() => confirmAndRun("Clear saved local chat history from this taskpane?", onClearChatHistory)}
          >
            Clear saved chats
          </button>
        </article>

        <article className="privacy-disclosure-row privacy-disclosure-row-static">
          <div>
            <strong>Telemetry</strong>
            <p>Pi-Office does not enable product analytics or telemetry by default. AI providers, connectors, and hosted services may keep their own request logs under their policies.</p>
          </div>
        </article>

        <article className="privacy-disclosure-row privacy-disclosure-row-static">
          <div>
            <strong>Local encryption limit</strong>
            <p>Provider and connector envelopes use AES-GCM, but the encryption keys are also stored in localStorage. Treat this as local obfuscation, not protection from same-origin script access.</p>
          </div>
        </article>
      </div>
    </section>
  );
}

function formatDiagnosticTimestamp(timestamp: number): string {
  return new Date(timestamp).toLocaleTimeString([], { hour12: false });
}

function DiagnosticsSection({
  diagnostics,
  onClear,
}: {
  diagnostics: RuntimeRequestFailureDiagnostic[];
  onClear: () => void;
}) {
  return (
    <section className="settings-section">
      <div className="settings-section-header">
        <h3>Runtime diagnostics</h3>
        {!!diagnostics.length && (
          <button type="button" className="settings-inline-action" onClick={onClear}>
            Clear
          </button>
        )}
      </div>
      <p className="settings-note">
        Captures recent route/auth/connector request failures from the in-process runtime.
      </p>
      {!diagnostics.length ? (
        <p className="settings-note">No runtime failures recorded in this session.</p>
      ) : (
        <div className="settings-card-grid diagnostics-card-grid">
          {diagnostics.map((diagnostic) => (
            <article key={diagnostic.id} className="settings-card diagnostics-card">
              <h4>
                {diagnostic.method} {diagnostic.path}
              </h4>
              <p>{diagnostic.message}</p>
              <p className="settings-meta">
                {DIAGNOSTIC_SOURCE_LABELS[diagnostic.source]} • {formatDiagnosticTimestamp(diagnostic.timestamp)}
              </p>
            </article>
          ))}
        </div>
      )}
    </section>
  );
}

function CompanionSection({
  companion,
  onRetryCompanion,
  onSaveCompanionEndpoint,
}: {
  companion: CompanionState;
  onRetryCompanion: () => Promise<void> | void;
  onSaveCompanionEndpoint: (endpoint: string) => Promise<void> | void;
}) {
  const [manualEndpoint, setManualEndpoint] = useState("");

  useEffect(() => {
    setManualEndpoint(companion.manualEndpoint ?? companion.endpoint ?? companion.lastSuccessfulEndpoint ?? "");
  }, [companion.endpoint, companion.lastSuccessfulEndpoint, companion.manualEndpoint]);

  const statusLabel =
    companion.status === "connected"
      ? "Connected"
      : companion.status === "discovering"
        ? "Discovering"
        : companion.status === "error"
          ? "Error"
          : "Unavailable";

  return (
    <section className="settings-section">
      <div className="settings-section-header">
        <h3>Optional Companion</h3>
        <button type="button" className="button" onClick={() => void onRetryCompanion()}>
          Retry discovery
        </button>
      </div>
      <p className="settings-note">
        Pi-Office works without the companion. The companion adds read-only local file access and read-only local MCP execution. Sandboxed shell stays hidden unless isolation detection and destructive probes pass.
      </p>

      <div className="settings-card-grid">
        <div className="settings-card">
          <span className="label">Status</span>
          <p>{statusLabel}</p>
        </div>
        <div className="settings-card">
          <span className="label">Endpoint</span>
          <p>{companion.endpoint ?? companion.lastSuccessfulEndpoint ?? "Not discovered yet"}</p>
        </div>
        <div className="settings-card">
          <span className="label">Capabilities</span>
          <p>
            Files: {companion.capabilities.fileRead ? "Read-only ready" : "Unavailable"}
            <br />
            Local MCP: {companion.capabilities.localMcp ? "Read-only ready" : "Unavailable"}
            <br />
            Shell: {companion.capabilities.shell?.state === "available"
              ? "Sandbox ready"
              : companion.capabilities.shell?.state === "degraded"
                ? "Sandbox blocked"
                : "Unavailable"}
          </p>
        </div>
      </div>

      {companion.capabilities.shell?.reason && (
        <div className="settings-card">
          <span className="label">Shell sandbox</span>
          <p>{companion.capabilities.shell.reason}</p>
        </div>
      )}

      {companion.lastError && (
        <div className="settings-card">
          <span className="label">Last error</span>
          <p>{companion.lastError}</p>
        </div>
      )}

      <div className="field">
        <span>Manual endpoint override</span>
        <input
          type="text"
          value={manualEndpoint}
          onChange={(event) => setManualEndpoint(event.target.value)}
          placeholder="https://localhost:3444"
        />
      </div>
      <div className="settings-actions">
        <button type="button" className="button button-solid" onClick={() => void onSaveCompanionEndpoint(manualEndpoint)}>
          Save endpoint
        </button>
        <button type="button" className="button" onClick={() => void onSaveCompanionEndpoint("")}>
          Clear override
        </button>
      </div>

      <h3>Setup</h3>
      <div className="settings-card-grid">
        <div className="settings-card">
          <span className="label">1. Install</span>
          <p>Install the companion package when the desktop bundle, zip, or npm package is published.</p>
          <p className="settings-note">
            <a href="#" onClick={(event) => event.preventDefault()}>Download ZIP (coming soon)</a>
            {" · "}
            <a href="#" onClick={(event) => event.preventDefault()}>Download binaries (coming soon)</a>
            {" · "}
            <a href="#" onClick={(event) => event.preventDefault()}>npm package (coming soon)</a>
          </p>
        </div>
        <div className="settings-card">
          <span className="label">2. Start</span>
          <p>Run the companion on your machine and keep it listening on `https://localhost:3444` or your chosen loopback endpoint.</p>
        </div>
        <div className="settings-card">
          <span className="label">3. Use</span>
          <p>When discovery succeeds, saved documents gain read-only file tools and local stdio MCP execution. Sandboxed shell appears only after the companion policy passes. Without it, the add-in still works normally.</p>
        </div>
      </div>
    </section>
  );
}

function ProfileSection({
  officeState,
  documentState,
  companion,
  sessionStats,
}: {
  officeState: OfficeStateUpdate | undefined;
  documentState: OfficeDocumentState;
  companion: CompanionState;
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
          <span className="label">Document State</span>
          <p>{documentState === "saved" ? "Saved" : "Unsaved"}</p>
        </div>
        <div className="settings-card">
          <span className="label">Companion</span>
          <p>{companion.status === "connected" ? "Connected" : companion.status === "discovering" ? "Discovering" : companion.status === "error" ? "Error" : "Unavailable"}</p>
        </div>
        {officeState?.document.workspaceDir && (
          <div className="settings-card">
            <span className="label">Document Folder</span>
            <p>{officeState.document.workspaceDir}</p>
          </div>
        )}
        {companion.endpoint && (
          <div className="settings-card">
            <span className="label">Companion Endpoint</span>
            <p>{companion.endpoint}</p>
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
  const providerStates = new Map((authStatus?.providerStates ?? []).map((entry) => [entry.provider, entry]));

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
            lastVerificationError={providerStates.get(provider.provider)?.lastVerificationError}
            onSaveApiKey={(key) => onSaveApiKey(provider.provider, key)}
            onStartOAuth={() => onStartOAuth(provider.provider)}
            onClearAuth={() => onClearAuth(provider.provider)}
          />
        ))}
        {providers.length === 0 && (
          <p className="settings-note">No providers discovered yet. Pi-Office uses provider credentials directly from the taskpane, so no companion is required for this section.</p>
        )}
      </div>
    </div>
  );
}

function ProviderCard({
  provider,
  hasAuth,
  lastVerificationError,
  isEnabled,
  onToggleEnabled,
  onSaveApiKey,
  onStartOAuth,
  onClearAuth,
}: {
  provider: ProviderDescriptor;
  hasAuth: boolean;
  lastVerificationError?: string | undefined;
  isEnabled: boolean;
  onToggleEnabled: () => void;
  onSaveApiKey: (key: string) => void;
  onStartOAuth: () => void;
  onClearAuth: () => void;
}) {
  const [apiKey, setApiKey] = useState("");
  const [expanded, setExpanded] = useState(false);
  const verifiedCount = provider.models.filter((m) => m.verifiedUsable).length;
  const providerCallable = provider.browserCallable;
  const statusLabel = !providerCallable
    ? providerSupportLabel(provider)
    : provider.verifiedUsable
    ? "Verified"
    : provider.authState === "verification_failed"
      ? "Auth failed"
      : provider.credentialStored
        ? "Unverified"
        : "Not configured";
  const statusClass = !providerCallable
    ? provider.supportStatus === "blocked" || provider.supportStatus === "research_only"
      ? "provider-status-error"
      : "provider-status-pending"
    : provider.verifiedUsable
    ? "provider-status-ready"
    : provider.authState === "verification_failed"
      ? "provider-status-error"
      : provider.credentialStored
        ? "provider-status-pending"
        : "";
  const authStatusText = !providerCallable
    ? provider.capabilityNote ?? "This provider is not callable from the browser taskpane yet."
    : provider.verifiedUsable
    ? "Verified credentials found."
    : provider.authState === "verification_failed"
      ? `Stored credential failed verification${lastVerificationError ? `: ${lastVerificationError}` : "."}`
      : hasAuth
        ? "Credential stored. It will be marked verified after the first successful provider request."
        : "No stored credentials.";
  const providerMeta = providerCallable
    ? `${verifiedCount}/${provider.models.length} verified`
    : provider.companionRequired
      ? "Companion required"
      : "Not available";
  const authMethodSummary = provider.authMethods.length
    ? provider.authMethods.map(authMethodLabel).join(" + ")
    : "No supported auth path";

  const handleSave = useCallback(() => {
    if (provider.apiKeySupported && apiKey.trim()) {
      onSaveApiKey(apiKey.trim());
      setApiKey("");
    }
  }, [apiKey, onSaveApiKey, provider.apiKeySupported]);

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
            <span className={`provider-status ${statusClass}`}>
              {statusLabel}
            </span>
          </div>
          <span className="provider-card-meta">
            {providerMeta}
          </span>
        </button>
        <button
          type="button"
          className={`toggle-switch toggle-sm ${isEnabled ? "toggle-on" : ""}`}
          onClick={onToggleEnabled}
          disabled={!providerCallable}
          role="switch"
          aria-checked={providerCallable && isEnabled}
          aria-label={`${isEnabled ? "Disable" : "Enable"} ${provider.label}`}
          title={providerCallable ? undefined : provider.capabilityNote ?? "This provider is not callable yet."}
        >
          <span className="toggle-thumb" />
        </button>
      </div>

      {expanded && (
        <div className="provider-card-body">
          <div className="provider-auth-status">
            {authStatusText}
          </div>

          <div className="provider-capability-row">
            <span className="settings-model-tag">{providerCallable ? "Browser callable" : providerSupportLabel(provider)}</span>
            {provider.companionRequired && <span className="settings-model-tag">Companion</span>}
            {provider.subscriptionBacked && <span className="settings-model-tag">Subscription</span>}
            <span className="settings-model-tag">{authMethodSummary}</span>
            {provider.imageGenerationSupported && <span className="settings-model-tag">Images</span>}
          </div>

          {provider.apiKeySupported ? (
            <div className="field">
              <span>API Key</span>
              <input
                type="password"
                value={apiKey}
                onChange={(e) => setApiKey(e.target.value)}
                placeholder="Paste provider API key"
              />
            </div>
          ) : (
            <p className="settings-note">
              {provider.companionRequired
                ? "This provider needs companion-owned auth before setup can be enabled here."
                : "Pi-Office does not support direct browser auth for this provider yet."}
            </p>
          )}

          <div className="settings-actions">
            <button type="button" className="button button-solid" onClick={handleSave} disabled={!provider.apiKeySupported}>
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

function providerSupportLabel(provider: ProviderDescriptor): string {
  if (provider.supportStatus === "supported") return "Supported";
  if (provider.supportStatus === "planned") return "Planned";
  if (provider.supportStatus === "blocked") return "Blocked";
  return "Research";
}

function authMethodLabel(method: ProviderAuthMethod): string {
  if (method === "api_key") return "API key";
  if (method === "oauth") return "OAuth";
  if (method === "manual_token") return "Manual token";
  if (method === "cloud_identity") return "Cloud identity";
  if (method === "aws_credentials") return "AWS credentials";
  return method;
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
  const activeProviders = providers.filter((p) => enabledProviders.has(p.provider) && p.browserCallable);

  return (
    <div className="settings-section">
      <h3>Available Models</h3>
      <p className="settings-note">
        Toggle models on or off to control which appear in the model selector.
        Only models from enabled providers are shown.
      </p>
      {activeProviders.length === 0 && (
        <p className="settings-note">Enable a browser-callable provider in the Providers tab first.</p>
      )}
      {activeProviders.map((provider) => (
        <div key={provider.provider} className="settings-model-group">
          <h4>{provider.label}</h4>
          <div className="settings-model-list">
            {provider.models.map((model) => {
              const modelKey = `${model.provider}::${model.modelId}`;
              const isEnabled = enabledModels.has(modelKey);
              const stateLabel = model.verifiedUsable
                ? ""
                : model.authState === "verification_failed"
                  ? "Auth failed"
                  : model.credentialStored
                    ? "Unverified"
                    : "Locked";
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
                    {stateLabel && (
                      <span className="settings-model-state">{stateLabel}</span>
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
  documentState,
  companion,
  preferences,
  onUpdatePreferences,
}: {
  documentState: OfficeDocumentState;
  companion: CompanionState;
  preferences: UserPreferences;
  onUpdatePreferences: (patch: Partial<UserPreferences>) => void;
}) {
  const showCompanionFileTools =
    documentState === "saved" &&
    companion.status === "connected" &&
    companion.capabilities.fileRead;
  const showCompanionMcp =
    companion.status === "connected" &&
    (companion.connectorToolNames?.length ?? 0) > 0;
  const showCompanionShell =
    documentState === "saved" &&
    companion.status === "connected" &&
    companion.capabilities.shell?.state === "available";

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

      <h3>Optional Companion Tools</h3>
      {!showCompanionFileTools && !showCompanionMcp && !showCompanionShell && (
        <p className="settings-note">
          {documentState !== "saved"
            ? "Save the document first to expose document-folder context. Local file tools stay disabled until the optional companion also connects."
            : companion.status === "connected" && companion.capabilities.localMcp
              ? "The companion is connected, but no read-only local MCP tools are verified for this session yet."
              : "The add-in is running without companion-backed local tools. Read-only local files and local MCP execution become available only when the optional companion connects."}
        </p>
      )}
      {showCompanionFileTools && (
        <>
          <div className="tool-list">
            {COMPANION_FILE_TOOL_NAMES.map((toolName) => (
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
      {showCompanionMcp && (
        <div className="tool-list">
          <ToolCardWithOverride
            toolName="mcp"
            overrideLevel={getOverrideLevel("mcp")}
            onOverrideChange={handleOverrideChange}
          />
        </div>
      )}
      {showCompanionShell && (
        <div className="tool-list">
          <ToolCardWithOverride
            toolName="bash"
            overrideLevel={getOverrideLevel("bash")}
            onOverrideChange={handleOverrideChange}
          />
        </div>
      )}
    </div>
  );
}

function getDefaultAutoApproveLevel(category: ToolCategory): AutonomyLevel {
  for (const level of AUTONOMY_LEVELS) {
    if (AUTONOMY_LEVEL_AUTO_APPROVE[level].has(category)) return level;
  }
  return "off";
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
  const manualOnly = category === "escape-hatch";
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
      {manualOnly && (
        <div className="tool-card-note">
          This tool is never auto-approved. Each call must be explicitly reviewed.
        </div>
      )}
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
          label="Next-prompt suggestions"
          description="Show concise prompt chips after useful Pi responses."
          value={preferences.nextPromptSuggestionsEnabled}
          onChange={(v) => onUpdate({ nextPromptSuggestionsEnabled: v })}
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
            Choose the default model for image generation. The browser taskpane currently exposes OpenAI image models only.
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
              No image models available. Configure OpenAI in the Providers tab.
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
