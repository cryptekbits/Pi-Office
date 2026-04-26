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
  ConnectorSetupProfile,
  ConnectorStatus,
  ConnectorTestResponse,
  ConnectorToolClassification,
  ConnectorToolInventoryItem,
  ConnectorToolPolicyUpdateRequest,
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
type CredentialMode = "detected" | "env" | "manual" | "oauth" | "none";
type ToolFilter = "all" | "enabled" | "read" | "sensitive" | "disabled" | "advanced";

interface ConnectorDraftState {
  connectorId: string;
  existingId?: string | undefined;
  setupProfileId?: string | undefined;
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
  onUpdateToolPolicy: (request: ConnectorToolPolicyUpdateRequest) => Promise<void>;
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
  setupProfiles: [
    {
      id: "custom-local-stdio",
      label: "Use local companion",
      description: "Run a local MCP command through the optional companion.",
      transport: "local_stdio",
      setupKind: "local_executable_or_docker",
      authMethod: "none",
      requiresCompanion: true,
      officialness: "community",
      availability: "needs_companion",
      browserDirect: "unsupported",
      docsUrl: "https://modelcontextprotocol.io/docs/concepts/transports",
      endpointEvidenceUrl: "https://modelcontextprotocol.io/docs/concepts/transports",
      checkedAt: "2026-04-27",
      riskNotes: ["Custom local commands require the companion and should be reviewed before enabling."],
      defaultWhenCompanionPresent: true,
      simpleFields: ["Local companion", "Launch command"],
      advancedFields: ["Arguments", "Working directory", "Environment variables", "Tool policy"],
    },
    {
      id: "custom-hosted-http",
      label: "Connect online",
      description: "Connect to a hosted MCP endpoint through the optional companion.",
      transport: "remote_http",
      setupKind: "remote_url_token",
      authMethod: "bearer_token",
      requiresCompanion: true,
      officialness: "community",
      availability: "advanced",
      browserDirect: "unknown",
      docsUrl: "https://modelcontextprotocol.io/docs/concepts/transports",
      endpointEvidenceUrl: "https://modelcontextprotocol.io/docs/concepts/transports",
      authEvidenceUrl: "https://modelcontextprotocol.io/docs/concepts/transports",
      checkedAt: "2026-04-27",
      riskNotes: ["Custom hosted MCP endpoints are advanced and should be reviewed before enabling."],
      defaultWhenCompanionAbsent: true,
      simpleFields: ["Connector URL", "Access token"],
      advancedFields: ["HTTP headers", "Environment-backed headers", "Tool policy"],
    },
  ],
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

function companionIsOnline(companion: CompanionState | undefined): boolean {
  return companion?.status === "connected";
}

function setupProfileNeedsCompanion(profile: ConnectorSetupProfile | undefined, transport: ConnectorSetupRequest["transport"] | undefined): boolean {
  if (profile?.requiresCompanion === true) return true;
  if (transport === "local_stdio" || profile?.transport === "local_stdio") return true;
  return false;
}

function setupProfileIsBrowserDirect(profile: ConnectorSetupProfile | undefined): boolean {
  return profile?.transport === "remote_http" && profile.browserDirect === "supported" && profile.requiresCompanion !== true && profile.setupDisabled !== true;
}

function setupProfileDisabled(profile: ConnectorSetupProfile | undefined): boolean {
  return profile?.setupDisabled === true || profile?.availability === "planned" || profile?.officialness === "planned";
}

function profileBadgeLabel(profile: ConnectorSetupProfile | undefined): string {
  switch (profile?.officialness) {
    case "official":
      return "Official";
    case "official_preview":
      return "Official preview";
    case "community":
      return "Community";
    case "provider_reference":
      return "Reference";
    case "deprecated":
      return "Deprecated";
    case "experimental":
      return "Advanced";
    case "planned":
      return "Planned";
    default:
      return "MCP";
  }
}

function profileNeedsCommunityWarning(profile: ConnectorSetupProfile | undefined): boolean {
  return profile?.officialness === "community" ||
    profile?.officialness === "provider_reference" ||
    profile?.officialness === "deprecated" ||
    profile?.officialness === "experimental";
}

function selectSetupProfile(
  connector: ConnectorCatalogItem,
  profileId: string | undefined,
  companion: CompanionState | undefined,
) {
  const profiles = connector.setupProfiles ?? [];
  if (!profiles.length) return undefined;
  if (profileId) {
    const explicit = profiles.find((profile) => profile.id === profileId);
    if (explicit) return explicit;
  }
  const companionOnline = companionIsOnline(companion);
  return companionOnline
    ? profiles.find((profile) => profile.defaultWhenCompanionPresent)
        ?? profiles.find((profile) => profile.transport === "local_stdio")
        ?? profiles[0]
    : profiles.find((profile) => profile.defaultWhenCompanionAbsent)
        ?? profiles.find((profile) => profile.transport === "remote_http")
        ?? profiles[0];
}

function connectorIsLocalOnly(connector: ConnectorCatalogItem): boolean {
  const profiles = connector.setupProfiles ?? [];
  if (!profiles.length) return connector.transport === "local_stdio";
  return profiles.every((profile) => profile.transport === "local_stdio");
}

function connectionTypeUserLabel(transport: ConnectorSetupRequest["transport"]): string {
  return transport === "local_stdio" ? "Local command (stdio)" : "Hosted URL (Streamable HTTP)";
}

function connectionTypeChipLabel(transport: ConnectorSetupRequest["transport"]): string {
  return transport === "local_stdio" ? "STDIO" : "HTTP";
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

function classificationLabel(classification: ConnectorToolClassification): string {
  switch (classification) {
    case "read":
      return "Read";
    case "sensitive_read":
      return "Sensitive";
    case "costly_read":
      return "Costly read";
    case "write":
      return "Write";
    case "destructive":
      return "Destructive";
    default:
      return "Unknown";
  }
}

function isAdvancedTool(tool: ConnectorToolInventoryItem): boolean {
  return tool.classification === "write" || tool.classification === "destructive" || tool.classification === "unknown";
}

function visibleToolsForFilter(tools: ConnectorToolInventoryItem[], filter: ToolFilter): ConnectorToolInventoryItem[] {
  switch (filter) {
    case "enabled":
      return tools.filter((tool) => tool.enabled);
    case "read":
      return tools.filter((tool) => tool.classification === "read");
    case "sensitive":
      return tools.filter((tool) => tool.classification === "sensitive_read" || tool.classification === "costly_read");
    case "disabled":
      return tools.filter((tool) => !tool.enabled);
    case "advanced":
      return tools.filter(isAdvancedTool);
    default:
      return tools;
  }
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
  companion: CompanionState | undefined,
): ConnectorDraftState {
  const envSuggestions = prepare?.envSuggestions ?? diagnostics?.envSuggestions ?? [];
  const existingDraft = prepare?.draft;
  const selectedProfile = selectSetupProfile(connector, existingDraft?.setupProfileId ?? status?.setupProfileId, companion);
  const canUseCompanionCredentialHints = companionIsOnline(companion);
  const detectedKey = canUseCompanionCredentialHints
    ? (existingDraft?.useDetectedEnvKey ?? envSuggestions.find((item) => item.present)?.key ?? "")
    : "";
  const authMethod = existingDraft?.authMethod ?? selectedProfile?.authMethod ?? connector.authMethod;
  const credentialMode: CredentialMode =
    authMethod === "none"
      ? "none"
      : authMethod === "oauth"
        ? "oauth"
        : detectedKey
        ? "detected"
        : existingDraft?.secretEnvKey && canUseCompanionCredentialHints
          ? "env"
          : "manual";

  return {
    connectorId: existingDraft?.connectorId ?? connector.id,
    existingId: existingDraft?.existingId ?? status?.id,
    setupProfileId: existingDraft?.setupProfileId ?? status?.setupProfileId ?? selectedProfile?.id,
    name: existingDraft?.name ?? connector.name,
    enabled: existingDraft?.enabled ?? status?.enabled ?? true,
    favorite: existingDraft?.favorite ?? status?.favorite ?? false,
    scopeTarget: existingDraft?.scopeTarget ?? status?.activeScope ?? "global",
    authMethod,
    transport: existingDraft?.transport ?? selectedProfile?.transport ?? connector.transport,
    credentialMode,
    detectedEnvKey: detectedKey,
    secretEnvKey: canUseCompanionCredentialHints
      ? (existingDraft?.secretEnvKey ?? selectedProfile?.credentialEnvKey ?? connector.envHints[0]?.key ?? "")
      : "",
    secret: "",
    preserveStoredSecret: status?.credentialSource === "manual",
    url: existingDraft?.url ?? selectedProfile?.endpoint ?? connector.template?.url ?? "",
    command: existingDraft?.command ?? selectedProfile?.command ?? connector.template?.command ?? "",
    argsText: formatArgs(existingDraft?.args ?? selectedProfile?.args ?? connector.template?.args),
    cwd: existingDraft?.cwd ?? selectedProfile?.cwd ?? connector.template?.cwd ?? "",
    envText: formatEnv(existingDraft?.env ?? selectedProfile?.env ?? connector.template?.env),
    stdioEnvPassthroughText: formatPassthroughLines(existingDraft?.stdioEnvPassthrough),
    remoteHttpHeaders:
      existingDraft?.remoteHttpHeaders?.length
        ? [...existingDraft.remoteHttpHeaders]
        : selectedProfile?.remoteHttpHeaders?.length
          ? [...selectedProfile.remoteHttpHeaders]
          : [{ name: "", value: "" }],
    remoteHttpHeadersFromEnv:
      existingDraft?.remoteHttpHeadersFromEnv?.length
        ? [...existingDraft.remoteHttpHeadersFromEnv]
        : selectedProfile?.remoteHttpHeadersFromEnv?.length
          ? [...selectedProfile.remoteHttpHeadersFromEnv]
        : [{ name: "", envVarName: "" }],
    advanced: false,
  };
}

function connectStepValidationMessage(
  d: ConnectorDraftState,
  scopeContext: ConnectorScopeContext | undefined,
  profile: ConnectorSetupProfile | undefined,
): string | undefined {
  if (setupProfileDisabled(profile)) return profile?.riskNotes?.[0] ?? "This connector profile is not available yet.";
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
        : draft.credentialMode === "oauth"
          ? "oauth"
          : draft.credentialMode === "manual"
            ? "manual"
          : "none";

  return {
    connectorId: draft.connectorId,
    existingId: draft.existingId,
    setupProfileId: draft.setupProfileId,
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

function renderDiagnostics(
  items: ConnectorDiagnostic[] | undefined,
  transport?: ConnectorSetupRequest["transport"],
  companion?: CompanionState | undefined,
): React.JSX.Element | null {
  if (!items?.length) return null;
  const coalesceLocalCompanionWarnings =
    transport === "local_stdio" &&
    !companionIsOnline(companion) &&
    items.some((item) => item.code === "local_stdio_requires_companion" || item.code === "local_stdio_companion_unavailable");

  const visibleItems = coalesceLocalCompanionWarnings
    ? [
        ...items.filter((item) => item.code !== "local_stdio_requires_companion" && item.code !== "local_stdio_companion_unavailable"),
        {
          level: "warning" as const,
          code: "local_stdio_companion_required",
          title: "Companion required",
          message: "STDIO connectors verify and execute through the optional companion. Start it from the Companion tab, then run Check connection.",
        },
      ]
    : items;

  if (!visibleItems.length) return null;
  return (
    <div className="integration-diagnostic-notes">
      {visibleItems.map((item) => (
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
  onUpdateToolPolicy,
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
  const [expandedConnectorId, setExpandedConnectorId] = useState<string>();
  const [toolFilter, setToolFilter] = useState<ToolFilter>("all");
  const [pendingToolEnable, setPendingToolEnable] = useState<{ status: ConnectorStatus; tool: ConnectorToolInventoryItem }>();
  const [suppressToolWarning, setSuppressToolWarning] = useState(false);
  const [pendingCommunityConnector, setPendingCommunityConnector] = useState<{ connector: ConnectorCatalogItem; profile: ConnectorSetupProfile }>();
  const [suppressCommunityWarning, setSuppressCommunityWarning] = useState(() => {
    try {
      return localStorage.getItem("pi-office-community-connector-warning-suppressed") === "true";
    } catch {
      return false;
    }
  });
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

  const selectedDraftProfile = useMemo(
    () => selectedConnector ? selectSetupProfile(selectedConnector, draft?.setupProfileId ?? selectedStatus?.setupProfileId, companion) : undefined,
    [companion, draft?.setupProfileId, selectedConnector, selectedStatus?.setupProfileId],
  );
  const selectedDraftNeedsCompanion = setupProfileNeedsCompanion(selectedDraftProfile, draft?.transport);

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
        setDraft(buildDraft(connector, nextPrepare, diagnostics, selectedStatus, companion));
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
  }, [companion, diagnostics, onPrepareConnector, prepareNonce, scopeContext, scopeContextKey, selectedKey, selectedStatus]);

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

  function openConnectorWizard(connector: ConnectorCatalogItem): void {
    setSelectedKey(connector.id);
    setWizardStep(1);
  }

  function acceptCommunityConnector(): void {
    if (!pendingCommunityConnector) return;
    if (suppressCommunityWarning) {
      try {
        localStorage.setItem("pi-office-community-connector-warning-suppressed", "true");
      } catch {
        // Ignore storage errors in Office webview mode.
      }
    }
    openConnectorWizard(pendingCommunityConnector.connector);
    setPendingCommunityConnector(undefined);
  }

  async function handleTest() {
    if (!draft) return;
    const message = connectStepValidationMessage(draft, scopeContext, selectedDraftProfile);
    if (message) {
      setError(message);
      return;
    }
    if (selectedDraftNeedsCompanion && !companionIsOnline(companion)) {
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
    const message = connectStepValidationMessage(draft, scopeContext, selectedDraftProfile);
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

  function applySetupProfile(profileId: string) {
    if (!selectedConnector) return;
    const profile = selectSetupProfile(selectedConnector, profileId, companion);
    if (!profile) return;
    setDraft((current) => {
      if (!current) return current;
      const nextAuthMethod = profile.authMethod;
      return {
        ...current,
        setupProfileId: profile.id,
        authMethod: nextAuthMethod,
        transport: profile.transport,
        credentialMode: nextAuthMethod === "none"
          ? "none"
          : nextAuthMethod === "oauth"
            ? "oauth"
            : current.credentialMode === "none"
              ? "manual"
              : current.credentialMode,
        secretEnvKey: companionIsOnline(companion) ? (profile.credentialEnvKey ?? current.secretEnvKey) : "",
        url: profile.transport === "remote_http" ? (profile.endpoint ?? current.url) : current.url,
        command: profile.transport === "local_stdio" ? (profile.command ?? current.command) : current.command,
        argsText: profile.transport === "local_stdio" ? formatArgs(profile.args) : current.argsText,
        cwd: profile.transport === "local_stdio" ? (profile.cwd ?? current.cwd) : current.cwd,
        envText: profile.env ? formatEnv(profile.env) : current.envText,
        remoteHttpHeaders: profile.remoteHttpHeaders?.length ? [...profile.remoteHttpHeaders] : current.remoteHttpHeaders,
        remoteHttpHeadersFromEnv: profile.remoteHttpHeadersFromEnv?.length
          ? [...profile.remoteHttpHeadersFromEnv]
          : current.remoteHttpHeadersFromEnv,
      };
    });
  }

  async function handleToolPolicyToggle(status: ConnectorStatus, tool: ConnectorToolInventoryItem, enabled: boolean, warningAcknowledged = false) {
    try {
      await onUpdateToolPolicy({
        connectorId: status.id,
        toolName: tool.name,
        enabled,
        warningAcknowledged,
        suppressWarning: warningAcknowledged && suppressToolWarning,
        scopeContext,
      });
      setPendingToolEnable(undefined);
      setSuppressToolWarning(false);
      setLiveMessage(`${enabled ? "Enabled" : "Disabled"} ${tool.name}.`);
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason));
    }
  }

  function handleWizardDone() {
    setView("connected");
    resetSelection();
  }

  async function handleReverify(targetId?: string) {
    const target = targetId ? statuses.find((status) => status.id === targetId) : selectedStatus;
    if (!target) return;
    const catalog = connectors.find((connector) => connector.id === target.connectorId);
    const profile = catalog ? selectSetupProfile(catalog, target.setupProfileId, companion) : undefined;
    if (setupProfileNeedsCompanion(profile, target.transport) && !companionIsOnline(companion)) {
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
    let connectorId = targetId ?? selectedStatus?.id;
    if (!connectorId) {
      if (!draft) {
        setError("Choose a connector before starting sign-in.");
        return;
      }
      const message = connectStepValidationMessage(draft, scopeContext, selectedDraftProfile);
      if (message) {
        setError(message);
        return;
      }
      try {
        const saved = await onConnectConnector(buildRequest(draft, scopeContext));
        connectorId = saved.status.id;
        setSelectedKey(saved.status.id);
        setPrepareNonce((n) => n + 1);
      } catch (reason) {
        setError(reason instanceof Error ? reason.message : String(reason));
        return;
      }
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
      const message = connectStepValidationMessage(draft, scopeContext, selectedDraftProfile);
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
    (selectedDraftNeedsCompanion && !companionIsOnline(companion)) ||
    setupProfileDisabled(selectedDraftProfile);

  const saveConnectorDisabled = !draft || working === "saving" || working === "testing" || setupProfileDisabled(selectedDraftProfile);

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
                const localOnlyDisabled = connectorIsLocalOnly(connector) && !companionIsOnline(companion);
                const profile = selectSetupProfile(connector, status?.setupProfileId, companion);
                const profileDisabled = setupProfileDisabled(profile);
                const disabledReason = profileDisabled
                  ? (profile?.riskNotes?.[0] ?? "Setup is planned but not available yet.")
                  : localOnlyDisabled
                    ? "Disabled until the local companion is connected because this connector requires a local command/stdio MCP server."
                    : undefined;
                return (
                  <button
                    key={connector.id}
                    type="button"
                    className={`integration-library-row ${selectedKey === connector.id ? "integration-library-row-active" : ""} ${localOnlyDisabled || profileDisabled ? "integration-library-row-disabled" : ""}`}
                    title={disabledReason}
                    aria-disabled={localOnlyDisabled || profileDisabled}
                    onClick={() => {
                      if (localOnlyDisabled || profileDisabled) return;
                      if (profileNeedsCommunityWarning(profile) && !suppressCommunityWarning && profile) {
                        setPendingCommunityConnector({ connector, profile });
                        return;
                      }
                      openConnectorWizard(connector);
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
                          <span className="integration-pill">{connectionTypeChipLabel(profile?.transport ?? connector.transport)}</span>
                          <span className="integration-pill">{authLabel(profile?.authMethod ?? connector.authMethod)}</span>
                          <span className={`integration-pill integration-profile-${profile?.officialness ?? "unknown"}`}>{profileBadgeLabel(profile)}</span>
                          {setupProfileIsBrowserDirect(profile) && <span className="integration-pill">Browser-direct</span>}
                          <span className="integration-pill">{difficultyLabel(connector.setupDifficulty)}</span>
                          {localOnlyDisabled && <span className="integration-pill integration-pill-warning">Companion required</span>}
                          {profileDisabled && <span className="integration-pill integration-pill-warning">Planned</span>}
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
                    <span className="integration-pill">{connectionTypeChipLabel(status.transport)}</span>
                    <span className="integration-pill">{authLabel(status.authMethod)}</span>
                    {status.setupProfileId && <span className="integration-pill">{status.setupProfileId}</span>}
                  </div>
                  {expandedConnectorId === status.id && (
                    <div className="integration-connected-details">
                      <div className="integration-config-grid">
                        <div>
                          <span>Runtime</span>
                          <strong>{status.executionEnvironment === "companion" ? "Companion" : "Taskpane"}</strong>
                        </div>
                        <div>
                          <span>Execution</span>
                          <strong>{status.executionAvailable ? "Available" : "Not available"}</strong>
                        </div>
                        <div>
                          <span>Last checked</span>
                          <strong>{status.lastTestedAt ? new Date(status.lastTestedAt).toLocaleString() : "Never"}</strong>
                        </div>
                        <div>
                          <span>Credential</span>
                          <strong>{status.credentialSource === "manual" ? "Stored locally" : status.credentialSource.replace("_", " ")}</strong>
                        </div>
                      </div>
                      <div className="integration-tool-panel">
                        <div className="integration-tool-panel-header">
                          <strong>Tools</strong>
                          <div className="integration-tool-filters" role="tablist" aria-label="Tool filters">
                            {(["all", "enabled", "read", "sensitive", "disabled", "advanced"] as ToolFilter[]).map((filter) => (
                              <button
                                key={filter}
                                type="button"
                                className={`integration-tool-filter ${toolFilter === filter ? "integration-tool-filter-active" : ""}`}
                                onClick={() => setToolFilter(filter)}
                              >
                                {filter}
                              </button>
                            ))}
                          </div>
                        </div>
                        {(status.capabilities?.toolInventory?.length ?? 0) === 0 ? (
                          <p className="settings-note">No tool inventory has been verified yet. Run Re-verify after the companion is connected.</p>
                        ) : (
                          <div className="integration-tool-list">
                            {visibleToolsForFilter(status.capabilities?.toolInventory ?? [], toolFilter).map((tool) => (
                              <div key={tool.name} className={`integration-tool-row integration-tool-row-${tool.classification}`}>
                                <div className="integration-tool-copy">
                                  <strong>{tool.name}</strong>
                                  <span>{classificationLabel(tool.classification)} - {tool.reason}</span>
                                </div>
                                <button
                                  type="button"
                                  className={`button ${tool.enabled ? "" : "button-solid"}`}
                                  onClick={() => {
                                    if (!tool.enabled && isAdvancedTool(tool) && !status.suppressNonReadToolWarning) {
                                      setPendingToolEnable({ status, tool });
                                      return;
                                    }
                                    void handleToolPolicyToggle(status, tool, !tool.enabled, !tool.enabled);
                                  }}
                                >
                                  {tool.enabled ? "Disable" : "Enable"}
                                </button>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    </div>
                  )}
                  <div className="settings-actions">
                    <button type="button" className="button" onClick={() => setExpandedConnectorId((current) => (current === status.id ? undefined : status.id))}>
                      {expandedConnectorId === status.id ? "Hide details" : "Details"}
                    </button>
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
                        <span className="integration-pill">{connectionTypeChipLabel(draft.transport)}</span>
                        <span className="integration-pill">{difficultyLabel(selectedConnector.setupDifficulty)}</span>
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
                      {renderDiagnostics(prepare?.diagnostics ?? diagnostics?.diagnostics, draft.transport, companion)}
                    </div>
                  )}

                  {wizardStep === 2 && (
                    <div className="connector-wizard-section">
                      <h4>Connect</h4>
                      {selectedDraftNeedsCompanion ? (
                        <div className={`settings-note integration-note ${companionIsOnline(companion) ? "integration-note-info" : "integration-note-warning"}`}>
                          <strong>Optional companion</strong>{" "}
                          {companionIsOnline(companion)
                            ? "Connected. You can verify MCP connectors and expose read-safe tools to the model."
                            : "Not connected. Save settings anytime; use the Companion tab to start discovery, then run Check connection."}
                          {draft.transport === "local_stdio" ? " Local stdio MCP always runs on this PC through the companion process." : " This profile needs companion-side verification before it can run."}
                        </div>
                      ) : setupProfileIsBrowserDirect(selectedDraftProfile) ? (
                        <div className="settings-note integration-note integration-note-info">
                          <strong>Browser sign-in</strong> This hosted MCP can be verified and used directly from the taskpane after sign-in.
                        </div>
                      ) : null}
                      {(selectedConnector.setupProfiles?.length ?? 0) > 1 && (
                        <div className="field">
                          <span>Setup path</span>
                          <p className="settings-note connector-wizard-field-hint">
                            Choose the simple online path, or use the local companion path when you want a connector command to run on this PC.
                          </p>
                          <div className="integration-profile-grid">
                            {selectedConnector.setupProfiles!.map((profile) => (
                              <button
                                key={profile.id}
                                type="button"
                                className={`integration-profile-option ${draft.setupProfileId === profile.id ? "integration-profile-option-active" : ""} ${setupProfileDisabled(profile) ? "integration-profile-option-disabled" : ""}`}
                                disabled={setupProfileDisabled(profile)}
                                title={setupProfileDisabled(profile) ? profile.riskNotes?.[0] : undefined}
                                onClick={() => applySetupProfile(profile.id)}
                              >
                                <strong>{profile.label}</strong>
                                <span className="integration-profile-badge-line">{profileBadgeLabel(profile)}{setupProfileIsBrowserDirect(profile) ? " - Browser-direct" : ""}</span>
                                <span>{profile.description}</span>
                              </button>
                            ))}
                          </div>
                        </div>
                      )}
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
                            {connectionTypeUserLabel(draft.transport)} — {setupProfileIsBrowserDirect(selectedDraftProfile)
                              ? "verified and executed directly from the taskpane after sign-in."
                              : selectedDraftNeedsCompanion
                                ? "verified and executed through the optional Pi-Office companion on this device."
                                : "saved in the taskpane profile."}
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
                      {draft.authMethod !== "none" && draft.authMethod !== "oauth" && (
                        <div className="field">
                          <span>Credentials</span>
                          <p className="settings-note connector-wizard-field-hint">
                            {companionIsOnline(companion)
                              ? "You can use a companion environment variable or paste a credential. Pasted values are stored in this add-in's local profile, not the OS keychain."
                              : "Paste the API key or token. Environment variable detection is shown only when the companion is connected."}
                          </p>
                          <div className="integration-choice-row connector-wizard-credential-row">
                            {companionIsOnline(companion) && (
                              <>
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
                              </>
                            )}
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
                            <button type="button" className="button button-solid" onClick={() => void handleStartOAuth()} disabled={working === "oauth" || setupProfileDisabled(selectedDraftProfile)}>
                              {working === "oauth" ? "Opening..." : `Sign in with ${selectedConnector.vendor}`}
                            </button>
                          </div>
                          {selectedPendingOAuth && (
                            <p className="settings-note">
                              Waiting for a verified OAuth callback. Manual completion is disabled; sign-in expires at{" "}
                              {new Date(selectedPendingOAuth.expiresAt).toLocaleTimeString()}.
                            </p>
                          )}
                        </div>
                      )}
                      {selectedConnector.id !== "custom" && draft.transport === "remote_http" && (
                        <button
                          type="button"
                          className="button"
                          onClick={() => setDraft((current) => (current ? { ...current, advanced: !current.advanced } : current))}
                        >
                          {draft.advanced ? "Hide advanced" : "Show advanced"}
                        </button>
                      )}
                      {draft.transport === "remote_http" && (selectedConnector.id === "custom" || draft.advanced) && (
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
                            <p className="settings-note connector-wizard-field-hint">Static headers sent on every MCP request.</p>
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
                          {companionIsOnline(companion) && (
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
                          )}
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
                        Save persists configuration on this device. Check connection discovers tools, disables unsafe ones, and activates the connector for chat when verification succeeds.
                      </p>
                      {draft && selectedDraftNeedsCompanion && !companionIsOnline(companion) ? (
                        <div className="settings-note integration-note integration-note-warning">
                          Companion is offline. Start discovery on the Companion tab, then use Check connection.
                        </div>
                      ) : setupProfileIsBrowserDirect(selectedDraftProfile) ? (
                        <div className="settings-note integration-note integration-note-info">
                          Verification will use the hosted MCP directly from the taskpane.
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

      {pendingToolEnable && (
        <div className="modal-backdrop connector-wizard-backdrop" role="presentation">
          <div className="connector-tool-warning" role="dialog" aria-modal="true" aria-labelledby="connector-tool-warning-title">
            <div className="connector-wizard-header">
              <div>
                <strong id="connector-tool-warning-title">Enable advanced connector tool?</strong>
                <p className="settings-note">
                  {pendingToolEnable.tool.name} is classified as {classificationLabel(pendingToolEnable.tool.classification).toLowerCase()}.
                  It will stay behind Pi-Office permission prompts, but it may change external systems or expose data outside Office.
                </p>
              </div>
              <button type="button" className="icon-button connector-wizard-close" aria-label="Close warning" onClick={() => setPendingToolEnable(undefined)}>
                <CloseIcon />
              </button>
            </div>
            <div className="connector-wizard-body">
              <label className="checkbox-row">
                <input
                  type="checkbox"
                  checked={suppressToolWarning}
                  onChange={(event) => setSuppressToolWarning(event.target.checked)}
                />
                <span>Do not show this again for this connector</span>
              </label>
            </div>
            <div className="connector-wizard-footer">
              <button type="button" className="button" onClick={() => setPendingToolEnable(undefined)}>Cancel</button>
              <span className="connector-wizard-footer-grow" />
              <button
                type="button"
                className="button button-solid"
                onClick={() => void handleToolPolicyToggle(pendingToolEnable.status, pendingToolEnable.tool, true, true)}
              >
                Enable tool
              </button>
            </div>
          </div>
        </div>
      )}

      {pendingCommunityConnector && (
        <div className="modal-backdrop connector-wizard-backdrop" role="presentation">
          <div className="connector-tool-warning" role="dialog" aria-modal="true" aria-labelledby="connector-community-warning-title">
            <div className="connector-wizard-header">
              <div>
                <strong id="connector-community-warning-title">Use community connector?</strong>
                <p className="settings-note">
                  {pendingCommunityConnector.connector.name} is marked {profileBadgeLabel(pendingCommunityConnector.profile).toLowerCase()}.
                  Pi-Office will keep it read-only by default, but you should trust the package or endpoint before continuing.
                </p>
              </div>
              <button type="button" className="icon-button connector-wizard-close" aria-label="Close warning" onClick={() => setPendingCommunityConnector(undefined)}>
                <CloseIcon />
              </button>
            </div>
            <div className="connector-wizard-body">
              <label className="checkbox-row">
                <input
                  type="checkbox"
                  checked={suppressCommunityWarning}
                  onChange={(event) => setSuppressCommunityWarning(event.target.checked)}
                />
                <span>Do not show this again</span>
              </label>
            </div>
            <div className="connector-wizard-footer">
              <button type="button" className="button" onClick={() => setPendingCommunityConnector(undefined)}>Cancel</button>
              <span className="connector-wizard-footer-grow" />
              <button type="button" className="button button-solid" onClick={acceptCommunityConnector}>
                Continue
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="sr-only" aria-live="polite">{liveMessage}</div>
      <input ref={importRef} type="file" accept=".json,application/json" className="sr-only" onChange={(event) => void handleImportFile(event.target.files?.[0] ?? null)} />
    </div>
  );
}
