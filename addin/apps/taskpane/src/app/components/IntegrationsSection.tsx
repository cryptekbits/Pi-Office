import { useDeferredValue, useEffect, useMemo, useRef, useState } from "react";
import type {
  ConnectorAuditPreference,
  ConnectorCatalogItem,
  ConnectorDiagnostic,
  ConnectorDiagnosticsResponse,
  ConnectorExportBundle,
  ConnectorFavoriteRequest,
  ConnectorImportApplyResponse,
  ConnectorImportPreviewResponse,
  ConnectorLogResponse,
  ConnectorOAuthStartResponse,
  ConnectorPrepareResponse,
  ConnectorScopeContext,
  ConnectorScopeTarget,
  ConnectorScopeUpdateRequest,
  ConnectorSetupRequest,
  ConnectorSetupResponse,
  ConnectorStatus,
  ConnectorTestResponse,
  OfficeHost,
} from "@pi-office/pi-office-pack/protocol";
import {
  CheckIcon,
  ConnectorBrandIcon,
  DiagnosticsIcon,
  DownloadIcon,
  ExternalLinkIcon,
  IntegrationIcon,
  LinkIcon,
  MagicIcon,
  PinIcon,
  ScopeIcon,
  ShieldIcon,
  UploadIcon,
} from "../../lib/icons";

type IntegrationsView = "library" | "connected" | "custom" | "diagnostics";
type WizardStep = 1 | 2 | 3 | 4;
type CredentialMode = "detected" | "env" | "manual" | "none";

interface ConnectorDraftState {
  connectorId: string;
  existingId?: string | undefined;
  name: string;
  enabled: boolean;
  favorite: boolean;
  scopeTarget: ConnectorScopeTarget;
  authMethod: ConnectorSetupRequest["authMethod"];
  transport: ConnectorSetupRequest["transport"];
  credentialMode: CredentialMode;
  detectedEnvKey: string;
  secretEnvKey: string;
  secret: string;
  preserveStoredSecret: boolean;
  url: string;
  command: string;
  argsText: string;
  cwd: string;
  envText: string;
  advanced: boolean;
}

interface PendingOAuthState {
  state: string;
  expiresAt: string;
  url?: string | undefined;
}

interface IntegrationsSectionProps {
  host: OfficeHost | undefined;
  connectors: ConnectorCatalogItem[];
  statuses: ConnectorStatus[];
  diagnostics: ConnectorDiagnosticsResponse | undefined;
  auditPreference: ConnectorAuditPreference | undefined;
  scopeContext: ConnectorScopeContext | undefined;
  onPrepareConnector: (connectorId: string, scopeContext?: ConnectorScopeContext) => Promise<ConnectorPrepareResponse>;
  onConnectConnector: (request: ConnectorSetupRequest) => Promise<ConnectorSetupResponse>;
  onTestConnector: (request: ConnectorSetupRequest) => Promise<ConnectorTestResponse>;
  onReverifyConnector: (connectorId: string, scopeContext?: ConnectorScopeContext) => Promise<ConnectorTestResponse>;
  onStartOAuth: (connectorId: string) => Promise<ConnectorOAuthStartResponse>;
  onRemoveConnector: (storedConnectorId: string) => Promise<void>;
  onSetFavorite: (request: ConnectorFavoriteRequest) => Promise<void>;
  onUpdateScope: (request: ConnectorScopeUpdateRequest) => Promise<void>;
  onLoadLogs: (connectorId: string) => Promise<ConnectorLogResponse>;
  onExportConnectors: () => Promise<ConnectorExportBundle>;
  onPreviewImport: (bundle: ConnectorExportBundle) => Promise<ConnectorImportPreviewResponse>;
  onApplyImport: (bundle: ConnectorExportBundle, resolutions?: Record<string, "skip" | "replace">) => Promise<ConnectorImportApplyResponse>;
  onSetAuditPreference: (preference: ConnectorAuditPreference) => Promise<ConnectorAuditPreference>;
}

const VIEW_OPTIONS: Array<{ key: IntegrationsView; label: string; Icon: () => React.JSX.Element }> = [
  { key: "library", label: "Library", Icon: IntegrationIcon },
  { key: "connected", label: "Connected", Icon: LinkIcon },
  { key: "custom", label: "Add Custom", Icon: MagicIcon },
  { key: "diagnostics", label: "Diagnostics", Icon: DiagnosticsIcon },
];

const CUSTOM_CONNECTOR_CARD: ConnectorCatalogItem = {
  id: "custom",
  name: "Custom MCP",
  vendor: "Custom",
  iconKey: "custom",
  category: "knowledge",
  maturity: "custom_mcp_only",
  setupKind: "local_executable_or_docker",
  authMethod: "none",
  transport: "local_stdio",
  readOnly: true,
  customOnly: true,
  summary: "Connect an existing MCP that already runs on your PC or behind a hosted MCP URL.",
  officeValue: "Bring internal search, research, or file systems into Office without exposing transport jargon to users.",
  tags: ["custom", "mcp", "internal"],
  capabilityHints: ["verify safe tools", "reuse local credentials", "connect hosted MCPs"],
  requirements: [],
  envHints: [],
  readPolicy: {
    mode: "hard-read-only",
    allowResources: true,
    allowPrompts: false,
    allowToolPatterns: [],
    blockToolPatterns: [],
  },
  setupNotes: [
    "Use this when your team already has an MCP server that is not in the library yet.",
    "The add-in still keeps the connector in read-only mode.",
  ],
  recommendedHosts: ["word", "excel", "powerpoint"],
  setupDifficulty: "advanced",
  catalogRevision: "custom-v1",
};

function formatArgs(args?: string[]): string {
  return args?.join("\n") ?? "";
}

function parseArgs(value: string): string[] | undefined {
  const args = value
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  return args.length ? args : undefined;
}

function formatEnv(env?: Record<string, string>): string {
  if (!env) return "";
  return Object.entries(env)
    .map(([key, value]) => `${key}=${value}`)
    .join("\n");
}

function parseEnv(value: string): Record<string, string> | undefined {
  const result: Record<string, string> = {};
  for (const line of value.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    const separator = trimmed.indexOf("=");
    if (separator <= 0) continue;
    const key = trimmed.slice(0, separator).trim();
    const entry = trimmed.slice(separator + 1).trim();
    if (key && entry) {
      result[key] = entry;
    }
  }
  return Object.keys(result).length ? result : undefined;
}

function badgeLabel(connector: ConnectorCatalogItem): string {
  if (connector.maturity === "ready") return "Ready";
  if (connector.maturity === "beta") return "Beta";
  return "Custom";
}

function authLabel(method: ConnectorSetupRequest["authMethod"]): string {
  switch (method) {
    case "api_key":
      return "API key";
    case "bearer_token":
      return "Access token";
    case "oauth":
      return "Browser sign-in";
    default:
      return "No auth";
  }
}

function runtimeLabel(transport: ConnectorSetupRequest["transport"]): string {
  return transport === "local_stdio" ? "Runs on this PC" : "Connects online";
}

function difficultyLabel(level: ConnectorCatalogItem["setupDifficulty"]): string {
  return level === "easy" ? "Easy setup" : level === "guided" ? "Guided setup" : "Advanced setup";
}

function healthLabel(status: ConnectorStatus): string {
  switch (status.healthState) {
    case "ready":
      return "Live";
    case "stale":
      return "Needs recheck";
    case "auth_required":
      return "Needs sign-in";
    case "auth_expired":
      return "Sign-in expired";
    case "offline":
      return "Offline";
    case "degraded":
      return "Degraded";
    default:
      return "Unverified";
  }
}

function healthMessage(status: ConnectorStatus): string {
  if (status.healthState === "ready") {
    return `${status.capabilities?.allowedTools.length ?? 0} read-safe tools are available.`;
  }
  return status.staleReason || status.lastError || "Needs attention before the connector will activate in chat.";
}

function deriveSetupKind(draft: ConnectorDraftState): ConnectorSetupRequest["setupKind"] {
  if (draft.transport === "remote_http") {
    if (draft.authMethod === "oauth") return "remote_oauth";
    if (draft.authMethod === "api_key") return "remote_api_key";
    return "remote_url_token";
  }

  const command = draft.command.trim().toLowerCase();
  if (/(^|\\)(python|uv|uvx)(\.exe)?$/.test(command) || command.includes("python") || command.includes("uv")) {
    return "local_python";
  }
  if (/(^|\\)(node|npx|npm|pnpm)(\.cmd|\.exe)?$/.test(command) || command.includes("node") || command.includes("npx")) {
    return "local_node";
  }
  return "local_executable_or_docker";
}

function buildDraft(
  connector: ConnectorCatalogItem,
  prepare: ConnectorPrepareResponse | undefined,
  diagnostics: ConnectorDiagnosticsResponse | undefined,
  status: ConnectorStatus | undefined,
): ConnectorDraftState {
  const envSuggestions = prepare?.envSuggestions ?? diagnostics?.envSuggestions ?? [];
  const existingDraft = prepare?.draft;
  const detectedKey = existingDraft?.useDetectedEnvKey ?? envSuggestions.find((item) => item.present)?.key ?? "";
  const credentialMode: CredentialMode =
    connector.authMethod === "none"
      ? "none"
      : detectedKey
        ? "detected"
        : existingDraft?.secretEnvKey
          ? "env"
          : "manual";

  return {
    connectorId: existingDraft?.connectorId ?? connector.id,
    existingId: existingDraft?.existingId ?? status?.id,
    name: existingDraft?.name ?? connector.name,
    enabled: existingDraft?.enabled ?? status?.enabled ?? true,
    favorite: existingDraft?.favorite ?? status?.favorite ?? false,
    scopeTarget: existingDraft?.scopeTarget ?? status?.activeScope ?? "global",
    authMethod: existingDraft?.authMethod ?? connector.authMethod,
    transport: existingDraft?.transport ?? connector.transport,
    credentialMode,
    detectedEnvKey: detectedKey,
    secretEnvKey: existingDraft?.secretEnvKey ?? connector.envHints[0]?.key ?? "",
    secret: "",
    preserveStoredSecret: status?.credentialSource === "manual",
    url: existingDraft?.url ?? connector.template?.url ?? "",
    command: existingDraft?.command ?? connector.template?.command ?? "",
    argsText: formatArgs(existingDraft?.args ?? connector.template?.args),
    cwd: existingDraft?.cwd ?? connector.template?.cwd ?? "",
    envText: formatEnv(existingDraft?.env ?? connector.template?.env),
    advanced: false,
  };
}

function buildRequest(draft: ConnectorDraftState, scopeContext: ConnectorScopeContext | undefined): ConnectorSetupRequest {
  const credentialSource =
    draft.credentialMode === "detected"
      ? "detected_env"
      : draft.credentialMode === "env"
        ? "env"
        : draft.credentialMode === "manual"
          ? draft.authMethod === "oauth"
            ? "oauth"
            : "manual"
          : "none";

  return {
    connectorId: draft.connectorId,
    existingId: draft.existingId,
    name: draft.name,
    enabled: draft.enabled,
    favorite: draft.favorite,
    scopeTarget: draft.scopeTarget,
    scopeContext,
    setupKind: deriveSetupKind(draft),
    authMethod: draft.authMethod,
    transport: draft.transport,
    credentialSource,
    preserveStoredSecret: draft.preserveStoredSecret && !draft.secret,
    useDetectedEnvKey: draft.credentialMode === "detected" ? draft.detectedEnvKey || undefined : undefined,
    secretEnvKey:
      draft.credentialMode === "env" || (draft.credentialMode === "manual" && draft.transport === "local_stdio")
        ? draft.secretEnvKey || undefined
        : undefined,
    secret: draft.credentialMode === "manual" ? draft.secret || undefined : undefined,
    url: draft.transport === "remote_http" ? draft.url || undefined : undefined,
    command: draft.transport === "local_stdio" ? draft.command || undefined : undefined,
    args: draft.transport === "local_stdio" ? parseArgs(draft.argsText) : undefined,
    cwd: draft.transport === "local_stdio" ? draft.cwd || undefined : undefined,
    env: parseEnv(draft.envText),
  };
}

function renderDiagnostics(items: ConnectorDiagnostic[] | undefined): React.JSX.Element | null {
  if (!items?.length) return null;
  return (
    <div className="integration-diagnostic-notes">
      {items.map((item) => (
        <p key={`${item.code}-${item.message}`} className={`settings-note integration-note integration-note-${item.level}`}>
          <strong>{item.title}</strong> {item.message}
        </p>
      ))}
    </div>
  );
}

function scopeLabel(target: ConnectorScopeTarget): string {
  return target === "document" ? "This document" : target === "workspace" ? "This folder" : "Everywhere";
}

function connectorStatusMap(statuses: ConnectorStatus[]): Map<string, ConnectorStatus> {
  return new Map(statuses.map((status) => [status.source === "custom" ? status.id : status.connectorId, status]));
}

function sortConnectors(connectors: ConnectorCatalogItem[], statuses: ConnectorStatus[], host: OfficeHost | undefined, query: string): ConnectorCatalogItem[] {
  const statusMap = connectorStatusMap(statuses);
  const filtered = query.trim()
    ? connectors.filter((connector) =>
        [connector.name, connector.vendor, connector.summary, connector.officeValue, ...connector.tags]
          .join(" ")
          .toLowerCase()
          .includes(query.trim().toLowerCase()),
      )
    : connectors;

  return [...filtered].sort((left, right) => {
    const leftStatus = statusMap.get(left.id);
    const rightStatus = statusMap.get(right.id);
    const leftPinned = leftStatus?.favorite ? 1 : 0;
    const rightPinned = rightStatus?.favorite ? 1 : 0;
    if (leftPinned !== rightPinned) return rightPinned - leftPinned;
    const leftRecommended = host && left.recommendedHosts.includes(host) ? 1 : 0;
    const rightRecommended = host && right.recommendedHosts.includes(host) ? 1 : 0;
    if (leftRecommended !== rightRecommended) return rightRecommended - leftRecommended;
    const leftConfigured = leftStatus?.configured ? 1 : 0;
    const rightConfigured = rightStatus?.configured ? 1 : 0;
    if (leftConfigured !== rightConfigured) return rightConfigured - leftConfigured;
    if (left.maturity !== right.maturity) {
      const order = { ready: 0, beta: 1, custom_mcp_only: 2 } as const;
      return order[left.maturity] - order[right.maturity];
    }
    return left.name.localeCompare(right.name);
  });
}

export function IntegrationsSection({
  host,
  connectors,
  statuses,
  diagnostics,
  auditPreference,
  scopeContext,
  onPrepareConnector,
  onConnectConnector,
  onTestConnector,
  onReverifyConnector,
  onStartOAuth,
  onRemoveConnector,
  onSetFavorite,
  onUpdateScope,
  onLoadLogs,
  onExportConnectors,
  onPreviewImport,
  onApplyImport,
  onSetAuditPreference,
}: IntegrationsSectionProps) {
  const [view, setView] = useState<IntegrationsView>("library");
  const [search, setSearch] = useState("");
  const [selectedKey, setSelectedKey] = useState<string>();
  const [wizardStep, setWizardStep] = useState<WizardStep>(1);
  const [prepare, setPrepare] = useState<ConnectorPrepareResponse>();
  const [draft, setDraft] = useState<ConnectorDraftState>();
  const [loadingPrepare, setLoadingPrepare] = useState(false);
  const [working, setWorking] = useState<"testing" | "saving" | "oauth" | "removing" | "reverify" | "export" | "import" | undefined>();
  const [result, setResult] = useState<ConnectorTestResponse | ConnectorSetupResponse>();
  const [logs, setLogs] = useState<ConnectorLogResponse>();
  const [importPreview, setImportPreview] = useState<ConnectorImportPreviewResponse>();
  const [pendingOAuth, setPendingOAuth] = useState<Record<string, PendingOAuthState>>({});
  const [error, setError] = useState<string>();
  const [liveMessage, setLiveMessage] = useState("");
  const importRef = useRef<HTMLInputElement | null>(null);

  function resetSelection() {
    setSelectedKey(undefined);
    setPrepare(undefined);
    setDraft(undefined);
    setError(undefined);
    setResult(undefined);
    setLogs(undefined);
    setWizardStep(1);
  }
  const deferredSearch = useDeferredValue(search);
  const scopeContextKey = useMemo(
    () =>
      scopeContext
        ? [
            scopeContext.host,
            scopeContext.documentId ?? "",
            scopeContext.documentSaved ? "saved" : "unsaved",
            scopeContext.documentUrl ?? "",
            scopeContext.workspaceId ?? "",
          ].join("::")
        : "none",
    [
      scopeContext?.documentId,
      scopeContext?.documentSaved,
      scopeContext?.documentUrl,
      scopeContext?.host,
      scopeContext?.workspaceId,
    ],
  );

  const sortedConnectors = useMemo(
    () => sortConnectors(connectors, statuses, host, deferredSearch),
    [connectors, deferredSearch, host, statuses],
  );

  const connectedItems = useMemo(
    () => statuses.filter((status) => status.configured || status.enabled || status.connected),
    [statuses],
  );

  const selectedStatus = useMemo(
    () => statuses.find((status) => status.id === selectedKey || status.connectorId === selectedKey),
    [selectedKey, statuses],
  );

  const selectedConnector = useMemo(() => {
    if (prepare?.connector) return prepare.connector;
    if (selectedKey === "custom" || selectedStatus?.source === "custom") {
      return CUSTOM_CONNECTOR_CARD;
    }
    if (!selectedKey) return undefined;
    return connectors.find((connector) => connector.id === selectedKey);
  }, [connectors, prepare?.connector, selectedKey, selectedStatus?.source]);

  useEffect(() => {
    if (!selectedKey) return;
    let active = true;
    setLoadingPrepare(true);
    setError(undefined);
    setResult(undefined);
    setPrepare(undefined);
    setDraft(undefined);

    void onPrepareConnector(selectedKey, scopeContext)
      .then((nextPrepare) => {
        if (!active) return;
        const connector = nextPrepare.connector.id === "custom" ? CUSTOM_CONNECTOR_CARD : nextPrepare.connector;
        setPrepare(nextPrepare);
        setDraft(buildDraft(connector, nextPrepare, diagnostics, selectedStatus));
      })
      .catch((reason) => {
        if (active) setError(reason instanceof Error ? reason.message : String(reason));
      })
      .finally(() => {
        if (active) setLoadingPrepare(false);
      });

    return () => {
      active = false;
    };
  }, [diagnostics, onPrepareConnector, scopeContext, scopeContextKey, selectedKey, selectedStatus]);

  useEffect(() => {
    if (!selectedStatus?.id) {
      setLogs(undefined);
      return;
    }

    let active = true;
    void onLoadLogs(selectedStatus.id)
      .then((response) => {
        if (active) {
          setLogs(response);
        }
      })
      .catch(() => undefined);

    return () => {
      active = false;
    };
  }, [onLoadLogs, selectedStatus?.id]);

  async function handleTest() {
    if (!draft) return;
    setWorking("testing");
    setError(undefined);
    try {
      const response = await onTestConnector(buildRequest(draft, scopeContext));
      setResult(response);
      setWizardStep(4);
      setLiveMessage(response.ok ? "Connector verification finished." : "Connector verification failed.");
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason));
    } finally {
      setWorking(undefined);
    }
  }

  async function handleSave() {
    if (!draft) return;
    setWorking("saving");
    setError(undefined);
    try {
      const response = await onConnectConnector(buildRequest(draft, scopeContext));
      setResult(response);
      setSelectedKey(response.status.source === "custom" ? response.status.id : response.status.connectorId);
      setView("connected");
      setWizardStep(4);
      setLiveMessage("Connector saved.");
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason));
    } finally {
      setWorking(undefined);
    }
  }

  async function handleReverify(targetId?: string) {
    const target = targetId ? statuses.find((status) => status.id === targetId) : selectedStatus;
    if (!target) return;
    setWorking("reverify");
    setError(undefined);
    try {
      const response = await onReverifyConnector(target.id, scopeContext);
      setResult(response);
      setLogs(await onLoadLogs(target.id));
      setLiveMessage(response.ok ? "Connector re-verified." : "Connector re-verification failed.");
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason));
    } finally {
      setWorking(undefined);
    }
  }

  async function handleStartOAuth(targetId?: string) {
    const connectorId = targetId ?? selectedStatus?.id;
    if (!connectorId) {
      setError("Save this connector first, then start browser sign-in.");
      return;
    }
    setWorking("oauth");
    setError(undefined);
    try {
      const started = await onStartOAuth(connectorId);
      setPendingOAuth((current) => ({
        ...current,
        [connectorId]: {
          state: started.state,
          expiresAt: started.expiresAt,
          url: started.url,
        },
      }));
      setLiveMessage("OAuth sign-in started. Complete sign-in after authenticating in the browser.");
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason));
    } finally {
      setWorking(undefined);
    }
  }

  async function handleRemove(statusId: string) {
    setWorking("removing");
    setError(undefined);
    try {
      await onRemoveConnector(statusId);
      if (selectedStatus?.id === statusId) {
        setSelectedKey(undefined);
        setPrepare(undefined);
        setDraft(undefined);
        setLogs(undefined);
      }
      setPendingOAuth((current) => {
        const next = { ...current };
        delete next[statusId];
        return next;
      });
      setLiveMessage("Connector removed.");
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason));
    } finally {
      setWorking(undefined);
    }
  }

  async function handleExport() {
    setWorking("export");
    try {
      const bundle = await onExportConnectors();
      const blob = new Blob([JSON.stringify(bundle, null, 2)], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = `pi-office-connectors-${bundle.exportedAt.slice(0, 10)}.json`;
      anchor.click();
      URL.revokeObjectURL(url);
      setLiveMessage("Connector export downloaded.");
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason));
    } finally {
      setWorking(undefined);
    }
  }

  async function handleImportFile(file: File | null) {
    if (!file) return;
    setWorking("import");
    try {
      const text = await file.text();
      const parsed = JSON.parse(text) as ConnectorExportBundle;
      const preview = await onPreviewImport(parsed);
      setImportPreview(preview);
      setView("diagnostics");
      setLiveMessage("Import preview ready.");
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : "Import file is not valid JSON.");
    } finally {
      setWorking(undefined);
      if (importRef.current) {
        importRef.current.value = "";
      }
    }
  }

  useEffect(() => {
    setPendingOAuth((current) => {
      let changed = false;
      const statusMap = new Map(statuses.map((status) => [status.id, status]));
      const next: Record<string, PendingOAuthState> = {};
      for (const [connectorId, pending] of Object.entries(current)) {
        const status = statusMap.get(connectorId);
        if (!status || status.healthState === "ready") {
          changed = true;
          continue;
        }
        if (Date.parse(pending.expiresAt) <= Date.now()) {
          changed = true;
          continue;
        }
        next[connectorId] = pending;
      }
      return changed ? next : current;
    });
  }, [statuses]);

  const selectedPendingOAuth = selectedStatus ? pendingOAuth[selectedStatus.id] : undefined;
  const panelOpen = Boolean(selectedKey && selectedConnector);

  return (
    <div className="settings-section integrations-shell">
      <div className="integrations-hero">
        <div>
          <div className="integrations-kicker">Read-only sources</div>
          <h3>Connect the information behind your Office work</h3>
          <p className="settings-note integrations-hero-copy">
            Users see plain language like “Connects online” or “Runs on this PC”. The runtime handles secure local state,
            scope-aware enablement, verification snapshots, and hard read-only filtering.
          </p>
        </div>
        <div className="integrations-hero-badges">
          <span className="integration-pill integration-pill-strong"><ShieldIcon /> Hard read-only</span>
          <span className="integration-pill"><ScopeIcon /> Scope aware</span>
          <span className="integration-pill"><DiagnosticsIcon /> Verified</span>
        </div>
      </div>

      <div className="integrations-toolbar">
        <div className="integrations-segmented segmented-control" role="tablist" aria-label="Integrations views">
          {VIEW_OPTIONS.map(({ key, label, Icon }) => (
            <button
              key={key}
              type="button"
              role="tab"
              aria-selected={view === key}
              className={`segmented-item ${view === key ? "segmented-active" : ""}`}
              onClick={() => {
                setView(key);
                if (key === "custom") {
                  resetSelection();
                  setSelectedKey("custom");
                  setWizardStep(1);
                } else {
                  resetSelection();
                }
              }}
            >
              <Icon />
              <span>{label}</span>
            </button>
          ))}
        </div>
        {view === "library" && (
          <div className="integrations-search field">
            <span>Find an integration</span>
            <input
              type="text"
              value={search}
              onChange={(event) => setSearch(event.target.value)}
              placeholder="Search Jira, SharePoint, Stripe, research, files..."
            />
          </div>
        )}
      </div>

      <div className={`integrations-layout ${panelOpen ? "integrations-layout-panel" : ""}`}>
        <div className="integrations-main">
          {view === "library" && (
            <div className="integration-card-grid">
              {sortedConnectors.map((connector) => {
                const status = statuses.find((entry) => entry.connectorId === connector.id);
                return (
                  <button
                    key={connector.id}
                    type="button"
                    className={`integration-card ${selectedKey === connector.id ? "integration-card-active" : ""}`}
                    onClick={() => {
                      setSelectedKey(connector.id);
                      setWizardStep(1);
                    }}
                  >
                    <div className="integration-card-top">
                      <ConnectorBrandIcon iconKey={connector.iconKey} label={connector.name} />
                      <div className="integration-card-heading">
                        <strong>{connector.name}</strong>
                        <span>{connector.vendor}</span>
                      </div>
                      <span className={`integration-state integration-state-${connector.maturity}`}>{badgeLabel(connector)}</span>
                    </div>
                    <p className="integration-card-copy">{connector.officeValue}</p>
                    <div className="integration-badges">
                      <span className="integration-pill">{runtimeLabel(connector.transport)}</span>
                      <span className="integration-pill">{authLabel(connector.authMethod)}</span>
                      <span className="integration-pill">{difficultyLabel(connector.setupDifficulty)}</span>
                    </div>
                    {status && (
                      <div className="integration-status-row">
                        <span className={`integration-status-dot integration-status-dot-${status.healthState}`} />
                        <span>{healthLabel(status)}</span>
                      </div>
                    )}
                  </button>
                );
              })}
            </div>
          )}

          {view === "connected" && (
            <div className="integration-connected-list">
              {connectedItems.length === 0 && <div className="settings-card"><p className="settings-note">No connectors are configured yet.</p></div>}
              {connectedItems.map((status) => (
                <div key={status.id} className="settings-card integration-connected-card">
                  <div className="integration-connected-top">
                    <div>
                      <strong>{status.name}</strong>
                      <p className="settings-note">{healthMessage(status)}</p>
                    </div>
                    <div className="integration-inline-actions">
                      <button type="button" className={`icon-button icon-button-sm ${status.favorite ? "icon-button-active" : ""}`} aria-label={status.favorite ? "Unpin connector" : "Pin connector"} onClick={() => void onSetFavorite({ connectorId: status.id, favorite: !status.favorite })}>
                        <PinIcon />
                      </button>
                      <span className={`integration-state integration-state-${status.healthState === "ready" ? "ready" : "beta"}`}>{healthLabel(status)}</span>
                    </div>
                  </div>
                  <div className="integration-badges">
                    <span className="integration-pill">{scopeLabel(status.activeScope)}</span>
                    <span className="integration-pill">{runtimeLabel(status.transport)}</span>
                    <span className="integration-pill">{authLabel(status.authMethod)}</span>
                  </div>
                  <div className="settings-actions">
                    <button type="button" className="button" onClick={() => { setView(status.source === "custom" ? "custom" : "library"); setSelectedKey(status.source === "custom" ? status.id : status.connectorId); setWizardStep(1); }}>Open setup</button>
                    {status.authMethod === "oauth" && (
                      <div className="integration-oauth-actions">
                        <button type="button" className="button" onClick={() => void handleStartOAuth(status.id)} disabled={working === "oauth"}>
                          {working === "oauth" ? "Opening..." : (status.healthState === "auth_expired" || status.healthState === "auth_required" ? "Sign in" : "Re-auth")}
                        </button>
                        {pendingOAuth[status.id] && (
                          <span className="settings-note">Waiting for a verified OAuth callback.</span>
                        )}
                      </div>
                    )}
                    <button type="button" className="button" onClick={() => void handleReverify(status.id)}>Re-verify</button>
                    <button type="button" className="button" onClick={() => void handleRemove(status.id)} disabled={working === "removing"}>Remove</button>
                  </div>
                </div>
              ))}
            </div>
          )}

          {view === "custom" && (
            <div className="integration-custom-grid">
              <div className="settings-card integration-custom-card"><strong>Hosted MCP</strong><p className="settings-note">Use a remote URL when your team already exposes an MCP endpoint behind company auth.</p></div>
              <div className="settings-card integration-custom-card"><strong>Local MCP</strong><p className="settings-note">Use a command when the connector runs on this PC through Node, Python, Docker, or a native binary.</p></div>
              <div className="settings-card integration-custom-card"><strong>Safer by default</strong><p className="settings-note">Custom connectors keep prompts blocked and only expose verified read-safe tools and resources.</p></div>
            </div>
          )}

          {view === "diagnostics" && (
            <div className="integration-diagnostics-grid">
              <div className="settings-card integration-diagnostics-card">
                <div className="integration-diagnostics-header">
                  <strong>Runtime checks</strong>
                  <button type="button" className="button" onClick={() => importRef.current?.click()}><UploadIcon /> Import</button>
                </div>
                <div className="integration-diagnostics-list">
                  {(diagnostics?.runtimes ?? []).map((check) => (
                    <div key={check.key} className={`integration-diagnostic-row ${check.ok ? "integration-diagnostic-ok" : ""}`}>
                      <span>{check.label}</span>
                      <span>{check.ok ? "Ready" : "Missing"}</span>
                    </div>
                  ))}
                </div>
              </div>
              <div className="settings-card integration-diagnostics-card">
                <strong>Audit logging</strong>
                <p className="settings-note">Redacted local audit entries record connector usage without prompt text or result content.</p>
                <div className="settings-actions">
                  <button type="button" className={`button ${auditPreference?.enabled ? "button-solid" : ""}`} onClick={() => void onSetAuditPreference({ enabled: !(auditPreference?.enabled ?? false) })}>
                    {auditPreference?.enabled ? "Disable audit log" : "Enable audit log"}
                  </button>
                  <button type="button" className="button" onClick={() => void handleExport()} disabled={working === "export"}><DownloadIcon /> Export</button>
                </div>
              </div>
              {importPreview && (
                <div className="settings-card integration-diagnostics-card">
                  <strong>Import preview</strong>
                  <p className="settings-note">{importPreview.bundle.connectors.length} connector definitions detected. Secrets are never imported.</p>
                  {importPreview.conflicts.length > 0 && (
                    <div className="integration-import-conflicts">
                      {importPreview.conflicts.map((conflict) => (
                        <p key={conflict.key} className="settings-note"><strong>{conflict.title}</strong> {conflict.message}</p>
                      ))}
                    </div>
                  )}
                  <div className="settings-actions">
                    <button
                      type="button"
                      className="button button-solid"
                      onClick={() => {
                        setWorking("import");
                        void onApplyImport(
                          importPreview.bundle,
                          Object.fromEntries(importPreview.conflicts.map((conflict) => [conflict.key, conflict.resolution === "replace" ? "replace" : "skip"])),
                        )
                          .then(() => {
                            setImportPreview(undefined);
                            setLiveMessage("Connector import applied.");
                          })
                          .catch((reason) => setError(reason instanceof Error ? reason.message : String(reason)))
                          .finally(() => setWorking(undefined));
                      }}
                    >
                      Apply import
                    </button>
                    <button type="button" className="button" onClick={() => setImportPreview(undefined)}>Clear preview</button>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>

        {panelOpen && selectedConnector && (
          <aside className="integration-panel">
            <div className="integration-panel-header">
              <div className="integration-panel-title">
                <ConnectorBrandIcon iconKey={selectedConnector.iconKey} label={draft?.name ?? selectedConnector.name} />
                <div>
                  <strong>{draft?.name ?? selectedConnector.name}</strong>
                  <p className="settings-note">{selectedConnector.summary}</p>
                </div>
              </div>
              <div className="integration-stepper" role="tablist" aria-label="Connector setup steps">
                {[1, 2, 3, 4].map((step) => (
                  <button key={step} type="button" role="tab" aria-selected={wizardStep === step} className={`integration-step ${wizardStep === step ? "integration-step-active" : ""}`} onClick={() => setWizardStep(step as WizardStep)}>{step}</button>
                ))}
              </div>
            </div>

            {loadingPrepare || !draft ? (
              <div className="settings-card"><p className="settings-note">Checking local prerequisites and saved hints...</p></div>
            ) : (
              <>
                {wizardStep === 1 && (
                  <div className="integration-panel-section">
                    <h4>Check</h4>
                    <p className="settings-note">{selectedConnector.officeValue}</p>
                    <div className="integration-badges">
                      <span className="integration-pill">{runtimeLabel(draft.transport)}</span>
                      <span className="integration-pill">{difficultyLabel(selectedConnector.setupDifficulty)}</span>
                      <span className="integration-pill">{selectedConnector.recommendedHosts.join(" / ")}</span>
                    </div>
                    {selectedConnector.docsUrl && <a className="integration-link" href={selectedConnector.docsUrl} target="_blank" rel="noreferrer"><ExternalLinkIcon /> Setup docs</a>}
                    {selectedConnector.authUrl && <a className="integration-link" href={selectedConnector.authUrl} target="_blank" rel="noreferrer"><ExternalLinkIcon /> Sign-in help</a>}
                    {renderDiagnostics(prepare?.diagnostics ?? diagnostics?.diagnostics)}
                  </div>
                )}

                {wizardStep === 2 && (
                  <div className="integration-panel-section">
                    <h4>Connect</h4>
                    <div className="field"><span>Name</span><input type="text" value={draft.name} onChange={(event) => setDraft((current) => current ? { ...current, name: event.target.value } : current)} /></div>
                    <div className="field"><span>Where should this be enabled?</span><div className="segmented-control segmented-control-wide">{(["global", "workspace", "document"] as ConnectorScopeTarget[]).map((target) => (<button key={target} type="button" className={`segmented-item ${draft.scopeTarget === target ? "segmented-active" : ""}`} onClick={() => setDraft((current) => current ? { ...current, scopeTarget: target } : current)}>{scopeLabel(target)}</button>))}</div></div>
                    <div className="field"><span>Where does it run?</span><div className="segmented-control segmented-control-wide"><button type="button" className={`segmented-item ${draft.transport === "remote_http" ? "segmented-active" : ""}`} onClick={() => setDraft((current) => current ? { ...current, transport: "remote_http" } : current)}>Connects online</button><button type="button" className={`segmented-item ${draft.transport === "local_stdio" ? "segmented-active" : ""}`} onClick={() => setDraft((current) => current ? { ...current, transport: "local_stdio" } : current)}>Runs on this PC</button></div></div>
                    {selectedConnector.id === "custom" && <div className="field"><span>How should it sign in?</span><div className="segmented-control segmented-control-wide">{(["none", "api_key", "bearer_token", "oauth"] as const).map((method) => (<button key={method} type="button" className={`segmented-item ${draft.authMethod === method ? "segmented-active" : ""}`} onClick={() => setDraft((current) => current ? { ...current, authMethod: method, credentialMode: method === "none" ? "none" : current.credentialMode === "none" ? "manual" : current.credentialMode } : current)}>{authLabel(method)}</button>))}</div></div>}
                    {draft.authMethod !== "none" && <div className="field"><span>Credentials</span><div className="integration-choice-row"><button type="button" className={`button ${draft.credentialMode === "detected" ? "button-solid" : ""}`} onClick={() => setDraft((current) => current ? { ...current, credentialMode: "detected" } : current)}>Use detected key</button><button type="button" className={`button ${draft.credentialMode === "env" ? "button-solid" : ""}`} onClick={() => setDraft((current) => current ? { ...current, credentialMode: "env" } : current)}>Use env var</button><button type="button" className={`button ${draft.credentialMode === "manual" ? "button-solid" : ""}`} onClick={() => setDraft((current) => current ? { ...current, credentialMode: "manual" } : current)}>Enter my own</button></div></div>}
                    {draft.credentialMode === "env" && <div className="field"><span>Environment variable name</span><input type="text" value={draft.secretEnvKey} onChange={(event) => setDraft((current) => current ? { ...current, secretEnvKey: event.target.value } : current)} /></div>}
                    {draft.credentialMode === "manual" && draft.authMethod !== "none" && <div className="field"><span>{draft.authMethod === "api_key" ? "API key" : "Access token"}</span><input type="password" value={draft.secret} onChange={(event) => setDraft((current) => current ? { ...current, secret: event.target.value, preserveStoredSecret: false } : current)} placeholder={draft.preserveStoredSecret ? "Using stored secret unless you replace it" : "Paste credential"} /></div>}
                    {draft.authMethod === "oauth" && (
                      <div className="settings-actions">
                        <button
                          type="button"
                          className="button"
                          onClick={() => void handleStartOAuth()}
                          disabled={!selectedStatus || working === "oauth"}
                        >
                          {working === "oauth" ? "Opening..." : "Open browser sign-in"}
                        </button>
                        {!selectedStatus && (
                          <p className="settings-note">
                            Save this connector first, then start OAuth sign-in.
                          </p>
                        )}
                        {selectedPendingOAuth && (
                          <p className="settings-note">
                            Waiting for a verified OAuth callback. Manual completion is disabled; sign-in expires at {new Date(selectedPendingOAuth.expiresAt).toLocaleTimeString()}.
                          </p>
                        )}
                      </div>
                    )}
                    {draft.transport === "remote_http" && <div className="field"><span>MCP service URL</span><input type="text" value={draft.url} onChange={(event) => setDraft((current) => current ? { ...current, url: event.target.value } : current)} placeholder="https://..." /></div>}
                    {draft.transport === "local_stdio" && (<><div className="field"><span>Launch command</span><input type="text" value={draft.command} onChange={(event) => setDraft((current) => current ? { ...current, command: event.target.value } : current)} /></div><div className="field"><span>Arguments</span><textarea rows={4} value={draft.argsText} onChange={(event) => setDraft((current) => current ? { ...current, argsText: event.target.value } : current)} /></div><div className="field"><span>Working folder</span><input type="text" value={draft.cwd} onChange={(event) => setDraft((current) => current ? { ...current, cwd: event.target.value } : current)} /></div></>)}
                    <button type="button" className="button" onClick={() => setDraft((current) => current ? { ...current, advanced: !current.advanced } : current)}>{draft.advanced ? "Hide advanced" : "Show advanced"}</button>
                    {draft.advanced && <div className="field"><span>Extra environment entries</span><textarea rows={5} value={draft.envText} onChange={(event) => setDraft((current) => current ? { ...current, envText: event.target.value } : current)} /></div>}
                  </div>
                )}

                {wizardStep === 3 && (
                  <div className="integration-panel-section">
                    <h4>Verify</h4>
                    <p className="settings-note">Only verified read-safe tools and resource helpers become available to the model.</p>
                    <div className="settings-actions">
                      <button type="button" className="button button-solid" onClick={() => void handleTest()} disabled={working === "testing"}>{working === "testing" ? "Checking..." : "Check connection"}</button>
                      {selectedStatus && <button type="button" className="button" onClick={() => void handleReverify()} disabled={working === "reverify"}>{working === "reverify" ? "Rechecking..." : "Re-verify saved connector"}</button>}
                    </div>
                    {result?.status.capabilities && (
                      <div className="integration-verification">
                        <div className="integration-diagnostic-row integration-diagnostic-ok"><span>Read-safe tools</span><span>{result.status.capabilities.allowedTools.length}</span></div>
                        <div className="integration-diagnostic-row"><span>Blocked tools</span><span>{result.status.capabilities.blockedTools.length}</span></div>
                        <div className="integration-diagnostic-row"><span>Resources</span><span>{result.status.capabilities.resourceCount}</span></div>
                        <div className="integration-diagnostic-row"><span>Prompts</span><span>{result.status.capabilities.allowedPrompts.length}</span></div>
                      </div>
                    )}
                    {logs?.logs?.length ? <div className="integration-log-list">{logs.logs.slice(-4).reverse().map((entry) => <p key={entry.id} className="settings-note"><strong>{entry.kind}</strong> {entry.message}</p>)}</div> : null}
                  </div>
                )}

                {wizardStep === 4 && (
                  <div className="integration-panel-section">
                    <h4>Finish</h4>
                    <div className="settings-card"><strong>{draft.name}</strong><p className="settings-note">{selectedStatus ? healthMessage(selectedStatus) : result?.status ? healthMessage(result.status) : "Save the connector when you are ready."}</p></div>
                    <div className="settings-actions">
                      <button type="button" className="button button-solid" onClick={() => void handleSave()} disabled={working === "saving"}>{working === "saving" ? "Saving..." : "Save connector"}</button>
                      {selectedStatus && <button type="button" className={`button ${selectedStatus.favorite ? "button-solid" : ""}`} onClick={() => void onSetFavorite({ connectorId: selectedStatus.id, favorite: !selectedStatus.favorite })}><PinIcon /> {selectedStatus.favorite ? "Pinned" : "Pin connector"}</button>}
                    </div>
                    {selectedStatus?.scopeStates && <div className="integration-scope-grid">{selectedStatus.scopeStates.map((scope) => (<button key={scope.target} type="button" className={`integration-scope-card ${scope.applies ? "integration-scope-card-active" : ""}`} onClick={() => void onUpdateScope({ connectorId: selectedStatus.id, enabled: !scope.enabled, scopeTarget: scope.target, scopeContext })} disabled={Boolean(scope.reason && scope.target !== "global")}><strong>{scope.label}</strong><span>{scope.enabled ? "Enabled" : "Disabled"}</span><span>{scope.reason ?? (scope.applies ? "Active scope" : "Available")}</span></button>))}</div>}
                  </div>
                )}
              </>
            )}

            {error && <div className="message-error"><div className="message-body">{error}</div></div>}
          </aside>
        )}
      </div>

      <div className="sr-only" aria-live="polite">{liveMessage}</div>
      <input ref={importRef} type="file" accept=".json,application/json" className="sr-only" onChange={(event) => void handleImportFile(event.target.files?.[0] ?? null)} />
    </div>
  );
}
