import { useCallback, useDeferredValue, useEffect, useMemo, useRef, useState } from "react";
import type {
  CompanionState,
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
  CloseIcon,
  ConnectorBrandIcon,
  DiagnosticsIcon,
  DownloadIcon,
  ExternalLinkIcon,
  IntegrationIcon,
  LinkIcon,
  MagicIcon,
  PinIcon,
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
  stdioEnvPassthroughText: string;
  remoteHttpHeaders: Array<{ name: string; value: string }>;
  remoteHttpHeadersFromEnv: Array<{ name: string; envVarName: string }>;
  advanced: boolean;
}

interface PendingOAuthState {
  state: string;
  expiresAt: string;
  url?: string | undefined;
}

interface IntegrationsSectionProps {
  host: OfficeHost | undefined;
  companion: CompanionState | undefined;
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

const WIZARD_STEPS: Array<{ step: WizardStep; label: string }> = [
  { step: 1, label: "Check" },
  { step: 2, label: "Connect" },
  { step: 3, label: "Verify" },
  { step: 4, label: "Finish" },
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

function formatPassthroughLines(keys?: string[]): string {
  return keys?.length ? keys.join("\n") : "";
}

function parsePassthroughLines(text: string): string[] {
  return text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
}

function mcpTransportNeedsCompanion(transport: ConnectorSetupRequest["transport"] | undefined): boolean {
  return transport === "local_stdio" || transport === "remote_http";
}

function companionIsOnline(companion: CompanionState | undefined): boolean {
  return companion?.status === "connected";
}

function connectionTypeUserLabel(transport: ConnectorSetupRequest["transport"]): string {
  return transport === "local_stdio" ? "Local command (stdio)" : "Hosted URL (Streamable HTTP)";
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
    stdioEnvPassthroughText: formatPassthroughLines(existingDraft?.stdioEnvPassthrough),
    remoteHttpHeaders:
      existingDraft?.remoteHttpHeaders?.length ? [...existingDraft.remoteHttpHeaders] : [{ name: "", value: "" }],
    remoteHttpHeadersFromEnv:
      existingDraft?.remoteHttpHeadersFromEnv?.length
        ? [...existingDraft.remoteHttpHeadersFromEnv]
        : [{ name: "", envVarName: "" }],
    advanced: false,
  };
}

function connectStepValidationMessage(d: ConnectorDraftState, scopeContext: ConnectorScopeContext | undefined): string | undefined {
  if (!d.name.trim()) return "Enter a name for this connector.";
  if ((d.scopeTarget === "workspace" || d.scopeTarget === "document") && scopeContext && !scopeContext.documentSaved) {
    return "Save the Office file to use folder or document scope.";
  }
  if (d.transport === "remote_http" && !d.url.trim()) return "Enter the MCP service URL.";
  if (d.transport === "local_stdio" && !d.command.trim()) return "Enter the launch command.";
  if (d.authMethod !== "none") {
    if (d.credentialMode === "detected" && !d.detectedEnvKey.trim()) {
      return "No detected environment key is available. Choose another credential option.";
    }
    if (d.credentialMode === "env" && !d.secretEnvKey.trim()) return "Enter the environment variable name.";
    if (d.credentialMode === "manual" && d.authMethod !== "oauth" && !d.secret && !d.preserveStoredSecret) {
      return "Enter a credential or keep the stored secret.";
    }
  }
  return undefined;
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
    stdioEnvPassthrough: parsePassthroughLines(draft.stdioEnvPassthroughText),
    remoteHttpHeaders: draft.remoteHttpHeaders
      .map((row) => ({ name: row.name.trim(), value: row.value.trim() }))
      .filter((row) => row.name && row.value),
    remoteHttpHeadersFromEnv: draft.remoteHttpHeadersFromEnv
      .map((row) => ({ name: row.name.trim(), envVarName: row.envVarName.trim() }))
      .filter((row) => row.name && row.envVarName),
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
  companion,
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
  const [prepareNonce, setPrepareNonce] = useState(0);
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

  const resetSelection = useCallback(() => {
    setSelectedKey(undefined);
    setPrepare(undefined);
    setDraft(undefined);
    setError(undefined);
    setResult(undefined);
    setLogs(undefined);
    setWizardStep(1);
  }, []);
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
  }, [diagnostics, onPrepareConnector, prepareNonce, scopeContext, scopeContextKey, selectedKey, selectedStatus]);

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
    const message = connectStepValidationMessage(draft, scopeContext);
    if (message) {
      setError(message);
      return;
    }
    if (mcpTransportNeedsCompanion(draft.transport) && !companionIsOnline(companion)) {
      setError("Connect the optional companion on the Companion tab, then try again.");
      return;
    }
    setWorking("testing");
    setError(undefined);
    try {
      const response = await onTestConnector(buildRequest(draft, scopeContext));
      setResult(response);
      setLiveMessage(response.ok ? "Connector verification finished." : "Connector verification failed.");
      if (response.ok) {
        setWizardStep(4);
      }
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason));
    } finally {
      setWorking(undefined);
    }
  }

  async function handleSaveConnector() {
    if (!draft) return;
    const message = connectStepValidationMessage(draft, scopeContext);
    if (message) {
      setError(message);
      return;
    }
    setWorking("saving");
    setError(undefined);
    try {
      const response = await onConnectConnector(buildRequest(draft, scopeContext));
      setSelectedKey(response.status.id);
      setPrepareNonce((n) => n + 1);
      setResult(undefined);
      setLiveMessage("Connector saved.");
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason));
    } finally {
      setWorking(undefined);
    }
  }

  function handleWizardDone() {
    setView("connected");
    resetSelection();
  }

  async function handleReverify(targetId?: string) {
    const target = targetId ? statuses.find((status) => status.id === targetId) : selectedStatus;
    if (!target) return;
    if (mcpTransportNeedsCompanion(target.transport) && !companionIsOnline(companion)) {
      setError("Connect the optional companion on the Companion tab, then try again.");
      return;
    }
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

  useEffect(() => {
    if (!panelOpen) return;
    function onKey(event: KeyboardEvent) {
      if (event.key === "Escape") resetSelection();
    }
    document.addEventListener("keydown", onKey);
    return () => document.removeEventListener("keydown", onKey);
  }, [panelOpen, resetSelection]);

  function handleWizardFooterPrimary() {
    if (loadingPrepare || !draft) return;
    if (wizardStep === 1) {
      setError(undefined);
      setWizardStep(2);
      return;
    }
    if (wizardStep === 2) {
      const message = connectStepValidationMessage(draft, scopeContext);
      if (message) {
        setError(message);
        return;
      }
      setError(undefined);
      setWizardStep(3);
      return;
    }
    if (wizardStep === 4) {
      if (selectedStatus) {
        handleWizardDone();
        return;
      }
      void handleSaveConnector();
    }
  }

  const wizardFooterPrimaryLabel =
    wizardStep === 1 || wizardStep === 2
      ? "Continue"
      : wizardStep === 4
        ? selectedStatus
          ? "Done"
          : working === "saving"
            ? "Saving..."
            : "Save connector"
        : "";

  const wizardFooterPrimaryDisabled =
    loadingPrepare || !draft || (wizardStep === 4 && working === "saving");

  const checkConnectionDisabled =
    !draft ||
    working === "testing" ||
    (mcpTransportNeedsCompanion(draft.transport) && !companionIsOnline(companion));

  const saveConnectorDisabled = !draft || working === "saving" || working === "testing";

  return (
    <div className="settings-section integrations-shell">
      <div className="integrations-hero">
        <div>
          <div className="integrations-kicker">Read-only sources</div>
          <h3>Connect trusted sources</h3>
        </div>
        <div className="integrations-hero-badges" aria-label="Connector safety">
          <span className="integration-pill integration-pill-strong"><ShieldIcon /> Read-only</span>
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

      <div className="integrations-layout">
        <div className="integrations-main">
          {view === "library" && (
            <div className="integration-library-list">
              {sortedConnectors.map((connector) => {
                const status = statuses.find((entry) => entry.connectorId === connector.id);
                return (
                  <button
                    key={connector.id}
                    type="button"
                    className={`integration-library-row ${selectedKey === connector.id ? "integration-library-row-active" : ""}`}
                    onClick={() => {
                      setSelectedKey(connector.id);
                      setWizardStep(1);
                    }}
                  >
                    <ConnectorBrandIcon iconKey={connector.iconKey} label={connector.name} />
                    <div className="integration-library-content">
                      <div className="integration-library-mainline">
                        <div className="integration-library-heading">
                          <strong>{connector.name}</strong>
                          <span>{connector.vendor}</span>
                        </div>
                        <span className={`integration-state integration-state-${connector.maturity}`}>{badgeLabel(connector)}</span>
                      </div>
                      <p className="integration-library-description">{connector.officeValue}</p>
                      <div className="integration-library-meta">
                        <div className="integration-badges">
                          <span className="integration-pill">{connectionTypeUserLabel(connector.transport)}</span>
                          <span className="integration-pill">{authLabel(connector.authMethod)}</span>
                          <span className="integration-pill">{difficultyLabel(connector.setupDifficulty)}</span>
                        </div>
                        {status && (
                          <div className="integration-status-row">
                            <span className={`integration-status-dot integration-status-dot-${status.healthState}`} />
                            <span>{healthLabel(status)}</span>
                          </div>
                        )}
                      </div>
                    </div>
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
                    <span className="integration-pill">{connectionTypeUserLabel(status.transport)}</span>
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
                      <span>
                        {check.label}
                        <small>{check.detail}</small>
                      </span>
                      <span>{check.ok ? "Ready" : "Missing"}</span>
                    </div>
                  ))}
                </div>
              </div>
              <div className="settings-card integration-diagnostics-card integration-audit-card">
                <div className="integration-audit-copy">
                  <strong>Audit logging</strong>
                  <p className="settings-note">Redacted local audit entries record connector usage without prompt text or result content.</p>
                </div>
                <div className="settings-actions integration-diagnostics-actions">
                  <button type="button" className={`button ${auditPreference?.enabled ? "button-solid" : ""}`} onClick={() => void onSetAuditPreference({ enabled: !(auditPreference?.enabled ?? false) })}>
                    {auditPreference?.enabled ? "Disable audit log" : "Enable audit log"}
                  </button>
                  <button type="button" className="button" onClick={() => void handleExport()} disabled={working === "export"}><DownloadIcon /> Export audit log</button>
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
      </div>

      {panelOpen && selectedConnector && (
        <div
          className="modal-backdrop connector-wizard-backdrop"
          role="presentation"
          onMouseDown={(event) => {
            if (event.target === event.currentTarget) resetSelection();
          }}
        >
          <div
            className="connector-wizard"
            role="dialog"
            aria-modal="true"
            aria-labelledby="connector-wizard-title"
          >
            <div className="connector-wizard-header">
              <div className="connector-wizard-title">
                <ConnectorBrandIcon iconKey={selectedConnector.iconKey} label={draft?.name ?? selectedConnector.name} />
                <div className="connector-wizard-title-text">
                  <strong id="connector-wizard-title">{draft?.name ?? selectedConnector.name}</strong>
                  <p className="settings-note">{selectedConnector.summary}</p>
                </div>
              </div>
              <button type="button" className="icon-button connector-wizard-close" aria-label="Close setup" onClick={resetSelection}>
                <CloseIcon />
              </button>
            </div>

            <div className="connector-wizard-progress" role="tablist" aria-label="Connector setup steps">
              {WIZARD_STEPS.map(({ step, label }) => (
                <button
                  key={step}
                  type="button"
                  role="tab"
                  aria-current={wizardStep === step ? "step" : undefined}
                  className={`connector-wizard-step-chip ${wizardStep === step ? "connector-wizard-step-chip-active" : ""} ${step < wizardStep ? "connector-wizard-step-chip-done" : ""}`}
                  disabled={step > wizardStep}
                  onClick={() => {
                    if (step < wizardStep) setWizardStep(step);
                  }}
                >
                  {label}
                </button>
              ))}
            </div>

            <div className="connector-wizard-body">
              {loadingPrepare || !draft ? (
                <div className="settings-card">
                  <p className="settings-note">Checking local prerequisites and saved hints...</p>
                </div>
              ) : (
                <>
                  {wizardStep === 1 && (
                    <div className="connector-wizard-section">
                      <h4>Check</h4>
                      <p className="settings-note">{selectedConnector.officeValue}</p>
                      <div className="integration-badges">
                        <span className="integration-pill">{connectionTypeUserLabel(draft.transport)}</span>
                        <span className="integration-pill">{difficultyLabel(selectedConnector.setupDifficulty)}</span>
                        <span className="integration-pill">{selectedConnector.recommendedHosts.join(" / ")}</span>
                      </div>
                      <div className="connector-wizard-doc-links">
                        {selectedConnector.docsUrl && (
                          <a className="integration-link" href={selectedConnector.docsUrl} target="_blank" rel="noreferrer">
                            <ExternalLinkIcon /> Setup docs
                          </a>
                        )}
                        {selectedConnector.authUrl && (
                          <a className="integration-link" href={selectedConnector.authUrl} target="_blank" rel="noreferrer">
                            <ExternalLinkIcon /> Sign-in help
                          </a>
                        )}
                      </div>
                      {renderDiagnostics(prepare?.diagnostics ?? diagnostics?.diagnostics)}
                    </div>
                  )}

                  {wizardStep === 2 && (
                    <div className="connector-wizard-section">
                      <h4>Connect</h4>
                      {mcpTransportNeedsCompanion(draft.transport) ? (
                        <div className={`settings-note integration-note ${companionIsOnline(companion) ? "integration-note-info" : "integration-note-warning"}`}>
                          <strong>Optional companion</strong>{" "}
                          {companionIsOnline(companion)
                            ? "Connected. You can verify MCP connectors and expose read-safe tools to the model."
                            : "Not connected. Save settings anytime; use the Companion tab to start discovery, then run Check connection."}
                          {draft.transport === "local_stdio" ? " Local stdio MCP always runs on this PC through the companion process." : " Hosted MCP URLs are called from the companion so credentials stay off the Office webview where possible."}
                        </div>
                      ) : null}
                      <div className="field">
                        <span>Name</span>
                        <input
                          type="text"
                          value={draft.name}
                          onChange={(event) => setDraft((current) => (current ? { ...current, name: event.target.value } : current))}
                        />
                      </div>
                      <div className="field">
                        <span>Use this connector in</span>
                        <p className="settings-note connector-wizard-field-hint">
                          Settings are saved on this device for Pi-Office. Folder and document choices apply only after the Office file is saved.
                        </p>
                        <div className="segmented-control segmented-control-wide">
                          {(["global", "workspace", "document"] as ConnectorScopeTarget[]).map((target) => {
                            const blocked = target !== "global" && !scopeContext?.documentSaved;
                            return (
                              <button
                                key={target}
                                type="button"
                                title={blocked ? "Save the document to enable this scope." : undefined}
                                className={`segmented-item ${draft.scopeTarget === target ? "segmented-active" : ""}`}
                                disabled={blocked}
                                onClick={() => setDraft((current) => (current ? { ...current, scopeTarget: target } : current))}
                              >
                                {scopeLabel(target)}
                              </button>
                            );
                          })}
                        </div>
                      </div>
                      {selectedConnector.id === "custom" ? (
                        <div className="field">
                          <span>Connection type</span>
                          <div className="segmented-control segmented-control-wide">
                            <button
                              type="button"
                              className={`segmented-item ${draft.transport === "local_stdio" ? "segmented-active" : ""}`}
                              onClick={() => setDraft((current) => (current ? { ...current, transport: "local_stdio" } : current))}
                            >
                              Local command (stdio)
                            </button>
                            <button
                              type="button"
                              className={`segmented-item ${draft.transport === "remote_http" ? "segmented-active" : ""}`}
                              onClick={() => setDraft((current) => (current ? { ...current, transport: "remote_http" } : current))}
                            >
                              Hosted URL (HTTP)
                            </button>
                          </div>
                        </div>
                      ) : (
                        <div className="field">
                          <span>Connection</span>
                          <p className="settings-note">
                            {connectionTypeUserLabel(draft.transport)} — verified and executed through the optional Pi-Office companion on this device.
                          </p>
                        </div>
                      )}
                      {selectedConnector.id === "custom" && (
                        <div className="field">
                          <span>How should it sign in?</span>
                          <div className="segmented-control segmented-control-wide">
                            {(["none", "api_key", "bearer_token", "oauth"] as const).map((method) => (
                              <button
                                key={method}
                                type="button"
                                className={`segmented-item ${draft.authMethod === method ? "segmented-active" : ""}`}
                                onClick={() =>
                                  setDraft((current) =>
                                    current
                                      ? {
                                          ...current,
                                          authMethod: method,
                                          credentialMode: method === "none" ? "none" : current.credentialMode === "none" ? "manual" : current.credentialMode,
                                        }
                                      : current,
                                  )
                                }
                              >
                                {authLabel(method)}
                              </button>
                            ))}
                          </div>
                        </div>
                      )}
                      {draft.authMethod !== "none" && (
                        <div className="field">
                          <span>Credentials</span>
                          <p className="settings-note connector-wizard-field-hint">
                            Prefer environment variables on the companion machine. Values you paste here are stored in this add-in&apos;s local profile (see Privacy), not the OS keychain.
                          </p>
                          <div className="integration-choice-row connector-wizard-credential-row">
                            <button
                              type="button"
                              className={`button ${draft.credentialMode === "detected" ? "button-solid" : ""}`}
                              onClick={() => setDraft((current) => (current ? { ...current, credentialMode: "detected" } : current))}
                            >
                              Use detected key
                            </button>
                            <button
                              type="button"
                              className={`button ${draft.credentialMode === "env" ? "button-solid" : ""}`}
                              onClick={() => setDraft((current) => (current ? { ...current, credentialMode: "env" } : current))}
                            >
                              Use env var
                            </button>
                            <button
                              type="button"
                              className={`button ${draft.credentialMode === "manual" ? "button-solid" : ""}`}
                              onClick={() => setDraft((current) => (current ? { ...current, credentialMode: "manual" } : current))}
                            >
                              Enter my own
                            </button>
                          </div>
                        </div>
                      )}
                      {draft.credentialMode === "env" && (
                        <div className="field">
                          <span>{draft.transport === "remote_http" ? "Bearer token environment variable" : "Credential environment variable"}</span>
                          <input
                            type="text"
                            value={draft.secretEnvKey}
                            onChange={(event) => setDraft((current) => (current ? { ...current, secretEnvKey: event.target.value } : current))}
                            placeholder={draft.transport === "remote_http" ? "MCP_BEARER_TOKEN" : "EXAMPLE_API_KEY"}
                          />
                        </div>
                      )}
                      {draft.credentialMode === "manual" && draft.authMethod !== "none" && (
                        <div className="field">
                          <span>{draft.authMethod === "api_key" ? "API key" : "Access token"}</span>
                          <input
                            type="password"
                            value={draft.secret}
                            onChange={(event) =>
                              setDraft((current) => (current ? { ...current, secret: event.target.value, preserveStoredSecret: false } : current))
                            }
                            placeholder={draft.preserveStoredSecret ? "Using stored secret unless you replace it" : "Paste credential"}
                          />
                        </div>
                      )}
                      {draft.authMethod === "oauth" && (
                        <div className="connector-wizard-oauth-block">
                          <div className="settings-actions connector-wizard-oauth-actions">
                            <button type="button" className="button" onClick={() => void handleStartOAuth()} disabled={!selectedStatus || working === "oauth"}>
                              {working === "oauth" ? "Opening..." : "Open browser sign-in"}
                            </button>
                          </div>
                          {!selectedStatus && (
                            <p className="settings-note">Save this connector first, then start OAuth sign-in.</p>
                          )}
                          {selectedPendingOAuth && (
                            <p className="settings-note">
                              Waiting for a verified OAuth callback. Manual completion is disabled; sign-in expires at{" "}
                              {new Date(selectedPendingOAuth.expiresAt).toLocaleTimeString()}.
                            </p>
                          )}
                        </div>
                      )}
                      {draft.transport === "remote_http" && (
                        <>
                          <div className="field">
                            <span>MCP service URL</span>
                            <input
                              type="text"
                              value={draft.url}
                              onChange={(event) => setDraft((current) => (current ? { ...current, url: event.target.value } : current))}
                              placeholder="https://..."
                            />
                          </div>
                          <div className="field">
                            <span>HTTP headers</span>
                            <p className="settings-note connector-wizard-field-hint">Static headers sent on every MCP request (companion).</p>
                            {draft.remoteHttpHeaders.map((row, index) => (
                              <div key={`hdr-${index}`} className="connector-wizard-kv-row">
                                <input
                                  type="text"
                                  aria-label="Header name"
                                  placeholder="Name"
                                  value={row.name}
                                  onChange={(event) =>
                                    setDraft((current) => {
                                      if (!current) return current;
                                      const next = [...current.remoteHttpHeaders];
                                      const prev = next[index] ?? { name: "", value: "" };
                                      next[index] = { name: event.target.value, value: prev.value };
                                      return { ...current, remoteHttpHeaders: next };
                                    })
                                  }
                                />
                                <input
                                  type="text"
                                  aria-label="Header value"
                                  placeholder="Value"
                                  value={row.value}
                                  onChange={(event) =>
                                    setDraft((current) => {
                                      if (!current) return current;
                                      const next = [...current.remoteHttpHeaders];
                                      const prev = next[index] ?? { name: "", value: "" };
                                      next[index] = { name: prev.name, value: event.target.value };
                                      return { ...current, remoteHttpHeaders: next };
                                    })
                                  }
                                />
                                <button
                                  type="button"
                                  className="icon-button"
                                  aria-label="Remove header"
                                  onClick={() =>
                                    setDraft((current) => {
                                      if (!current) return current;
                                      const next = current.remoteHttpHeaders.filter((_, i) => i !== index);
                                      return {
                                        ...current,
                                        remoteHttpHeaders: next.length ? next : [{ name: "", value: "" }],
                                      };
                                    })
                                  }
                                >
                                  ×
                                </button>
                              </div>
                            ))}
                            <button
                              type="button"
                              className="button connector-wizard-add-row"
                              onClick={() =>
                                setDraft((current) =>
                                  current ? { ...current, remoteHttpHeaders: [...current.remoteHttpHeaders, { name: "", value: "" }] } : current,
                                )
                              }
                            >
                              + Add header
                            </button>
                          </div>
                          <div className="field">
                            <span>Headers from environment variables</span>
                            <p className="settings-note connector-wizard-field-hint">Header values are read from the companion process environment.</p>
                            {draft.remoteHttpHeadersFromEnv.map((row, index) => (
                              <div key={`hdre-${index}`} className="connector-wizard-kv-row">
                                <input
                                  type="text"
                                  aria-label="Header name"
                                  placeholder="Header name"
                                  value={row.name}
                                  onChange={(event) =>
                                    setDraft((current) => {
                                      if (!current) return current;
                                      const next = [...current.remoteHttpHeadersFromEnv];
                                      const prev = next[index] ?? { name: "", envVarName: "" };
                                      next[index] = { name: event.target.value, envVarName: prev.envVarName };
                                      return { ...current, remoteHttpHeadersFromEnv: next };
                                    })
                                  }
                                />
                                <input
                                  type="text"
                                  aria-label="Environment variable name"
                                  placeholder="Env var name"
                                  value={row.envVarName}
                                  onChange={(event) =>
                                    setDraft((current) => {
                                      if (!current) return current;
                                      const next = [...current.remoteHttpHeadersFromEnv];
                                      const prev = next[index] ?? { name: "", envVarName: "" };
                                      next[index] = { name: prev.name, envVarName: event.target.value };
                                      return { ...current, remoteHttpHeadersFromEnv: next };
                                    })
                                  }
                                />
                                <button
                                  type="button"
                                  className="icon-button"
                                  aria-label="Remove header from env"
                                  onClick={() =>
                                    setDraft((current) => {
                                      if (!current) return current;
                                      const next = current.remoteHttpHeadersFromEnv.filter((_, i) => i !== index);
                                      return {
                                        ...current,
                                        remoteHttpHeadersFromEnv: next.length ? next : [{ name: "", envVarName: "" }],
                                      };
                                    })
                                  }
                                >
                                  ×
                                </button>
                              </div>
                            ))}
                            <button
                              type="button"
                              className="button connector-wizard-add-row"
                              onClick={() =>
                                setDraft((current) =>
                                  current
                                    ? { ...current, remoteHttpHeadersFromEnv: [...current.remoteHttpHeadersFromEnv, { name: "", envVarName: "" }] }
                                    : current,
                                )
                              }
                            >
                              + Add header from env
                            </button>
                          </div>
                        </>
                      )}
                      {draft.transport === "local_stdio" && (
                        <>
                          <div className="field">
                            <span>Command to launch</span>
                            <input
                              type="text"
                              value={draft.command}
                              onChange={(event) => setDraft((current) => (current ? { ...current, command: event.target.value } : current))}
                              placeholder="e.g. npx or full path to executable"
                            />
                          </div>
                          <div className="field">
                            <span>Arguments</span>
                            <textarea
                              rows={4}
                              value={draft.argsText}
                              onChange={(event) => setDraft((current) => (current ? { ...current, argsText: event.target.value } : current))}
                              placeholder="One argument per line"
                            />
                          </div>
                          <div className="field">
                            <span>Working directory</span>
                            <input
                              type="text"
                              value={draft.cwd}
                              onChange={(event) => setDraft((current) => (current ? { ...current, cwd: event.target.value } : current))}
                              placeholder="Optional folder for the process"
                            />
                          </div>
                          {selectedConnector.id === "custom" ? (
                            <>
                              <div className="field">
                                <span>Environment variables</span>
                                <textarea
                                  rows={4}
                                  value={draft.envText}
                                  onChange={(event) => setDraft((current) => (current ? { ...current, envText: event.target.value } : current))}
                                  placeholder={"KEY=value per line"}
                                />
                              </div>
                              <div className="field">
                                <span>Environment variable passthrough</span>
                                <p className="settings-note connector-wizard-field-hint">
                                  Names of variables to copy from the companion host into the MCP process (one per line), in addition to safe inherited defaults.
                                </p>
                                <textarea
                                  rows={3}
                                  value={draft.stdioEnvPassthroughText}
                                  onChange={(event) =>
                                    setDraft((current) => (current ? { ...current, stdioEnvPassthroughText: event.target.value } : current))
                                  }
                                  placeholder={"PATH_EXTRA\nCUSTOM_VAR"}
                                />
                              </div>
                            </>
                          ) : null}
                        </>
                      )}
                      {selectedConnector.id !== "custom" && draft.transport === "local_stdio" ? (
                        <>
                          <button
                            type="button"
                            className="button"
                            onClick={() => setDraft((current) => (current ? { ...current, advanced: !current.advanced } : current))}
                          >
                            {draft.advanced ? "Hide advanced" : "Show advanced"}
                          </button>
                          {draft.advanced && (
                            <>
                              <div className="field">
                                <span>Environment variables</span>
                                <textarea
                                  rows={5}
                                  value={draft.envText}
                                  onChange={(event) => setDraft((current) => (current ? { ...current, envText: event.target.value } : current))}
                                />
                              </div>
                              <div className="field">
                                <span>Environment variable passthrough</span>
                                <p className="settings-note connector-wizard-field-hint">
                                  Names of variables to copy from the companion host into the MCP process (one per line).
                                </p>
                                <textarea
                                  rows={3}
                                  value={draft.stdioEnvPassthroughText}
                                  onChange={(event) =>
                                    setDraft((current) => (current ? { ...current, stdioEnvPassthroughText: event.target.value } : current))
                                  }
                                  placeholder={"PATH_EXTRA\nCUSTOM_VAR"}
                                />
                              </div>
                            </>
                          )}
                        </>
                      ) : null}
                    </div>
                  )}

                  {wizardStep === 3 && (
                    <div className="connector-wizard-section">
                      <h4>Verify</h4>
                      <p className="settings-note">
                        Save persists configuration on this device. Check connection runs a read-safe verification through the companion and activates the connector for chat when it succeeds.
                      </p>
                      {draft && mcpTransportNeedsCompanion(draft.transport) && !companionIsOnline(companion) ? (
                        <div className="settings-note integration-note integration-note-warning">
                          Companion is offline. Start discovery on the Companion tab, then use Check connection.
                        </div>
                      ) : null}
                      <p className="settings-note">Only verified read-safe tools and resource helpers become available to the model.</p>
                      {selectedStatus && (
                        <div className="settings-actions">
                          <button type="button" className="button" onClick={() => void handleReverify()} disabled={working === "reverify"}>
                            {working === "reverify" ? "Rechecking..." : "Re-verify saved connector"}
                          </button>
                        </div>
                      )}
                      {result?.status.capabilities && (
                        <div className="integration-verification">
                          <div className="integration-diagnostic-row integration-diagnostic-ok">
                            <span>Read-safe tools</span>
                            <span>{result.status.capabilities.allowedTools.length}</span>
                          </div>
                          <div className="integration-diagnostic-row">
                            <span>Blocked tools</span>
                            <span>{result.status.capabilities.blockedTools.length}</span>
                          </div>
                          <div className="integration-diagnostic-row">
                            <span>Resources</span>
                            <span>{result.status.capabilities.resourceCount}</span>
                          </div>
                          <div className="integration-diagnostic-row">
                            <span>Prompts</span>
                            <span>{result.status.capabilities.allowedPrompts.length}</span>
                          </div>
                        </div>
                      )}
                      {logs?.logs?.length ? (
                        <div className="integration-log-list">
                          {logs.logs
                            .slice(-4)
                            .reverse()
                            .map((entry) => (
                              <p key={entry.id} className="settings-note">
                                <strong>{entry.kind}</strong> {entry.message}
                              </p>
                            ))}
                        </div>
                      ) : null}
                    </div>
                  )}

                  {wizardStep === 4 && (
                    <div className="connector-wizard-section">
                      <h4>Finish</h4>
                      {!selectedStatus && result?.ok ? (
                        <p className="settings-note integration-note integration-note-info">
                          Verification succeeded. Use <strong>Save connector</strong> below to persist settings and enable this connector for chat.
                        </p>
                      ) : null}
                      <div className="settings-card">
                        <strong>{draft.name}</strong>
                        <p className="settings-note">
                          {selectedStatus
                            ? healthMessage(selectedStatus)
                            : result?.status
                              ? healthMessage(result.status)
                              : "Save the connector when you are ready."}
                        </p>
                      </div>
                      {selectedStatus && (
                        <div className="settings-actions">
                          <button
                            type="button"
                            className={`button ${selectedStatus.favorite ? "button-solid" : ""}`}
                            onClick={() => void onSetFavorite({ connectorId: selectedStatus.id, favorite: !selectedStatus.favorite })}
                          >
                            <PinIcon /> {selectedStatus.favorite ? "Pinned" : "Pin connector"}
                          </button>
                        </div>
                      )}
                      {selectedStatus?.scopeStates && (
                        <div className="integration-scope-grid">
                          {selectedStatus.scopeStates.map((scope) => (
                            <button
                              key={scope.target}
                              type="button"
                              className={`integration-scope-card ${scope.applies ? "integration-scope-card-active" : ""}`}
                              onClick={() =>
                                void onUpdateScope({
                                  connectorId: selectedStatus.id,
                                  enabled: !scope.enabled,
                                  scopeTarget: scope.target,
                                  scopeContext,
                                })
                              }
                              disabled={Boolean(scope.reason && scope.target !== "global")}
                            >
                              <strong>{scope.label}</strong>
                              <span>{scope.enabled ? "Enabled" : "Disabled"}</span>
                              <span>{scope.reason ?? (scope.applies ? "Active scope" : "Available")}</span>
                            </button>
                          ))}
                        </div>
                      )}
                    </div>
                  )}
                </>
              )}
            </div>

            {error && (
              <div className="message-error connector-wizard-error">
                <div className="message-body">{error}</div>
              </div>
            )}

            <div className="connector-wizard-footer">
              {wizardStep > 1 ? (
                <button type="button" className="button" onClick={() => setWizardStep((step) => (step - 1) as WizardStep)}>
                  Back
                </button>
              ) : null}
              <span className="connector-wizard-footer-grow" />
              {wizardStep === 3 && draft ? (
                <div className="connector-wizard-footer-actions">
                  <button
                    type="button"
                    className="button"
                    onClick={() => void handleSaveConnector()}
                    disabled={saveConnectorDisabled}
                  >
                    {working === "saving" ? "Saving..." : "Save connector"}
                  </button>
                  <button
                    type="button"
                    className="button button-solid"
                    onClick={() => void handleTest()}
                    disabled={checkConnectionDisabled}
                  >
                    {working === "testing" ? "Checking..." : "Check connection"}
                  </button>
                </div>
              ) : (
                <button
                  type="button"
                  className="button button-solid"
                  onClick={handleWizardFooterPrimary}
                  disabled={wizardFooterPrimaryDisabled || !wizardFooterPrimaryLabel}
                >
                  {wizardFooterPrimaryLabel}
                </button>
              )}
            </div>
          </div>
        </div>
      )}

      <div className="sr-only" aria-live="polite">{liveMessage}</div>
      <input ref={importRef} type="file" accept=".json,application/json" className="sr-only" onChange={(event) => void handleImportFile(event.target.files?.[0] ?? null)} />
    </div>
  );
}
