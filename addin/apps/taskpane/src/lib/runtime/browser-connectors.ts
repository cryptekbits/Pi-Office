import type {
  CompanionConnectorDefinition,
  ConnectorAuditPreference,
  ConnectorAuditPreferenceResponse,
  ConnectorAuthMethod,
  ConnectorCapabilitySummary,
  ConnectorCatalogItem,
  ConnectorCatalogResponse,
  ConnectorCategory,
  ConnectorConflict,
  ConnectorCredentialSource,
  ConnectorDiagnostic,
  ConnectorDiagnosticsResponse,
  ConnectorEnvSuggestion,
  ConnectorExportBundle,
  ConnectorExportItem,
  ConnectorExportScopeOverride,
  ConnectorFavoriteRequest,
  ConnectorHealthState,
  ConnectorImportApplyRequest,
  ConnectorImportApplyResponse,
  ConnectorImportPreviewRequest,
  ConnectorImportPreviewResponse,
  ConnectorLogEntry,
  ConnectorLogResponse,
  ConnectorMaturity,
  ConnectorOAuthCallbackRequest,
  ConnectorOAuthCallbackResponse,
  ConnectorOAuthStartResponse,
  ConnectorPrepareResponse,
  ConnectorRemoteHttpHeader,
  ConnectorRemoteHttpHeaderFromEnv,
  ConnectorRuntimeCheck,
  ConnectorScopeContext,
  ConnectorScopeState,
  ConnectorScopeTarget,
  ConnectorScopeUpdateRequest,
  ConnectorSetupRequest,
  ConnectorSetupResponse,
  ConnectorSetupProfile,
  ConnectorSetupKind,
  ConnectorStatus,
  ConnectorStatusResponse,
  ConnectorTestResponse,
  ConnectorToolClassification,
  ConnectorToolInventoryItem,
  ConnectorToolPolicyOverride,
  ConnectorToolPolicyUpdateRequest,
  ConnectorToolPolicyUpdateResponse,
  ConnectorTransport,
  ConnectorVerificationSnapshot,
} from "@pi-office/pi-office-pack/protocol";
import { executeBrowserMcpTool, probeBrowserMcpConnector, type BrowserMcpConnectorConfig } from "./browser-mcp-client";
import { getConnectorCatalogItem, listConnectorCatalog } from "./connector-catalog";

const CONNECTOR_STORAGE_KEY = "pi-office-connectors";
const CONNECTOR_STORAGE_VERSION = 1;
const CONNECTOR_CRYPTO_KEY_STORAGE_KEY = "pi-office-connectors-key-v1";
const MAX_CONNECTOR_LOG_ENTRIES = 300;

interface EncryptedConnectorEnvelope {
  version: number;
  algorithm: "AES-GCM";
  iv: string;
  ciphertext: string;
}

interface StoredConnectorRecord {
  id: string;
  connectorId: string;
  source: "library" | "custom";
  name: string;
  category: ConnectorCategory;
  maturity: ConnectorMaturity;
  setupKind: ConnectorSetupKind;
  authMethod: ConnectorAuthMethod;
  transport: ConnectorTransport;
  credentialSource: ConnectorCredentialSource;
  secret?: string | undefined;
  secretEnvKey?: string | undefined;
  useDetectedEnvKey?: string | undefined;
  url?: string | undefined;
  command?: string | undefined;
  args?: string[] | undefined;
  cwd?: string | undefined;
  setupProfileId?: string | undefined;
  env?: Record<string, string> | undefined;
  stdioEnvPassthrough?: string[] | undefined;
  remoteHttpHeaders?: ConnectorRemoteHttpHeader[] | undefined;
  remoteHttpHeadersFromEnv?: ConnectorRemoteHttpHeaderFromEnv[] | undefined;
  defaultEnabled: boolean;
  favorite: boolean;
  createdAt: string;
  updatedAt: string;
  lastTestedAt?: string | undefined;
  lastHealthyAt?: string | undefined;
  lastError?: string | undefined;
  verification?: ConnectorVerificationSnapshot | undefined;
  capabilities?: ConnectorCapabilitySummary | undefined;
  toolPolicyOverrides?: ConnectorToolPolicyOverride[] | undefined;
  suppressNonReadToolWarning?: boolean | undefined;
  oauthConnected?: boolean | undefined;
  oauthExpiresAt?: string | undefined;
  oauthLastAuthAt?: string | undefined;
  oauthLastAuthError?: string | undefined;
  oauthRefreshToken?: string | undefined;
  oauthTokenType?: string | undefined;
  oauthScope?: string | undefined;
  oauthClientId?: string | undefined;
  oauthTokenEndpoint?: string | undefined;
}

interface StoredScopeOverride {
  connectorId: string;
  scopeTarget: Exclude<ConnectorScopeTarget, "global">;
  scopeKey: string;
  enabled: boolean;
  updatedAt: string;
}

interface StoredOAuthFlow {
  connectorId: string;
  state: string;
  createdAt: string;
  expiresAt: string;
  url?: string | undefined;
  redirectUri?: string | undefined;
  codeVerifier?: string | undefined;
  tokenEndpoint?: string | undefined;
  clientId?: string | undefined;
  clientSecret?: string | undefined;
  scope?: string | undefined;
}

interface StoredConnectorState {
  version: number;
  updatedAt: string;
  connectors: StoredConnectorRecord[];
  scopeOverrides: StoredScopeOverride[];
  oauthFlows: StoredOAuthFlow[];
  logs: ConnectorLogEntry[];
  auditPreference: ConnectorAuditPreference;
}

interface VerificationResult {
  ok: boolean;
  diagnostics: ConnectorDiagnostic[];
  next: StoredConnectorRecord;
}

function nowIso(): string {
  return new Date().toISOString();
}

function trimString(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed.length ? trimmed : undefined;
}

function normalizeUrl(value: string | undefined): string | undefined {
  const trimmed = trimString(value);
  if (!trimmed) return undefined;
  try {
    const parsed = new URL(trimmed);
    parsed.hash = "";
    if ((parsed.protocol === "https:" && parsed.port === "443") || (parsed.protocol === "http:" && parsed.port === "80")) {
      parsed.port = "";
    }
    return parsed.toString().replace(/\/$/, "");
  } catch {
    return trimmed;
  }
}

function normalizeArgs(value: unknown): string[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const args = value
    .map((entry) => trimString(entry))
    .filter((entry): entry is string => Boolean(entry));
  return args.length ? args : undefined;
}

function normalizeEnv(value: unknown): Record<string, string> | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const result: Record<string, string> = {};
  for (const [key, entry] of Object.entries(value)) {
    const trimmed = trimString(entry);
    if (trimmed) result[key] = trimmed;
  }
  return Object.keys(result).length ? result : undefined;
}

function normalizeRemoteHttpHeaders(value: unknown): ConnectorRemoteHttpHeader[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const out: ConnectorRemoteHttpHeader[] = [];
  for (const item of value) {
    if (!item || typeof item !== "object") continue;
    const rec = item as { name?: string; value?: string };
    const name = trimString(rec.name);
    const headerValue = trimString(rec.value);
    if (name && headerValue) out.push({ name, value: headerValue });
  }
  return out.length ? out : undefined;
}

function normalizeRemoteHttpHeadersFromEnv(value: unknown): ConnectorRemoteHttpHeaderFromEnv[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const out: ConnectorRemoteHttpHeaderFromEnv[] = [];
  for (const item of value) {
    if (!item || typeof item !== "object") continue;
    const rec = item as { name?: string; envVarName?: string };
    const name = trimString(rec.name);
    const envVarName = trimString(rec.envVarName);
    if (name && envVarName) out.push({ name, envVarName });
  }
  return out.length ? out : undefined;
}

function normalizeStdioEnvPassthrough(value: unknown): string[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const keys = value.map((entry) => trimString(entry)).filter((entry): entry is string => Boolean(entry));
  return keys.length ? keys : undefined;
}

function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    const slice = bytes.subarray(index, index + chunkSize);
    binary += String.fromCharCode(...slice);
  }
  return btoa(binary);
}

function base64ToUint8Array(value: string): Uint8Array {
  const binary = atob(value);
  const bytes = new Uint8Array(new ArrayBuffer(binary.length));
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(new ArrayBuffer(bytes.byteLength));
  copy.set(bytes);
  return copy.buffer;
}

function isEncryptedConnectorEnvelope(value: unknown): value is EncryptedConnectorEnvelope {
  if (!value || typeof value !== "object") return false;
  const envelope = value as Partial<EncryptedConnectorEnvelope>;
  return (
    envelope.version === CONNECTOR_STORAGE_VERSION &&
    envelope.algorithm === "AES-GCM" &&
    typeof envelope.iv === "string" &&
    typeof envelope.ciphertext === "string"
  );
}

function createConnectorState(): StoredConnectorState {
  return {
    version: CONNECTOR_STORAGE_VERSION,
    updatedAt: nowIso(),
    connectors: [],
    scopeOverrides: [],
    oauthFlows: [],
    logs: [],
    auditPreference: { enabled: true },
  };
}

function createRandomId(prefix: string): string {
  if (typeof crypto.randomUUID === "function") {
    return `${prefix}${crypto.randomUUID()}`;
  }
  const bytes = new Uint8Array(16);
  crypto.getRandomValues(bytes);
  return `${prefix}${Array.from(bytes, (byte) => byte.toString(16).padStart(2, "0")).join("")}`;
}

function createOAuthState(): string {
  const bytes = new Uint8Array(24);
  crypto.getRandomValues(bytes);
  return arrayBufferToBase64(bytes.buffer).replace(/[+/=]/g, "").slice(0, 40);
}

function base64UrlEncode(buffer: ArrayBuffer): string {
  return arrayBufferToBase64(buffer)
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

function createPkceVerifier(): string {
  const bytes = new Uint8Array(48);
  crypto.getRandomValues(bytes);
  return base64UrlEncode(bytes.buffer);
}

async function createPkceChallenge(verifier: string): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier));
  return base64UrlEncode(digest);
}

function oauthCallbackUrl(): string {
  if (typeof window !== "undefined" && window.location?.origin) {
    return `${window.location.origin}/connector-oauth-callback`;
  }
  return "https://localhost:3443/connector-oauth-callback";
}

function endpointOrigin(url: string | undefined): string | undefined {
  const normalized = normalizeUrl(url);
  if (!normalized) return undefined;
  try {
    return new URL(normalized).origin;
  } catch {
    return undefined;
  }
}

function metadataUrlForEndpoint(url: string | undefined): string | undefined {
  const origin = endpointOrigin(url);
  return origin ? `${origin}/.well-known/oauth-authorization-server` : undefined;
}

function profileIsBrowserDirect(profile: ConnectorSetupProfile | undefined): boolean {
  return profile?.transport === "remote_http" && profile.browserDirect === "supported" && profile.requiresCompanion !== true && profile.setupDisabled !== true;
}

function profileNeedsCompanion(profile: ConnectorSetupProfile | undefined, transport: ConnectorTransport | undefined): boolean {
  if (profile?.requiresCompanion === true) return true;
  if (transport === "local_stdio") return true;
  return false;
}

function profileSetupDisabled(profile: ConnectorSetupProfile | undefined): boolean {
  return profile?.setupDisabled === true || profile?.availability === "planned" || profile?.officialness === "planned";
}

function customCommandSignature(record: Pick<StoredConnectorRecord, "command" | "args" | "cwd">): string {
  const command = trimString(record.command)?.toLowerCase() ?? "";
  const args = (record.args ?? []).map((entry) => entry.toLowerCase()).join("\u001f");
  const cwd = trimString(record.cwd)?.toLowerCase() ?? "";
  return `${command}\u001e${args}\u001e${cwd}`;
}

function resolveWorkspaceScopeKey(context: ConnectorScopeContext | undefined): string | undefined {
  return trimString(context?.workspaceId)?.toLowerCase();
}

function resolveDocumentScopeKey(context: ConnectorScopeContext | undefined): string | undefined {
  if (!context?.documentSaved) return undefined;
  return trimString(context.documentUrl ?? context.documentId)?.toLowerCase();
}

function connectorLabel(target: ConnectorScopeTarget): string {
  if (target === "workspace") return "This folder";
  if (target === "document") return "This document";
  return "Everywhere";
}

function sanitizeHintToToolName(value: string): string {
  let next = value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .replace(/_+/g, "_");
  if (!next) next = "read";
  if (/^\d/.test(next)) next = `tool_${next}`;
  return next;
}

function normalizeToolPolicyOverrides(value: unknown): ConnectorToolPolicyOverride[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const out: ConnectorToolPolicyOverride[] = [];
  for (const item of value) {
    if (!item || typeof item !== "object") continue;
    const rec = item as { toolName?: string; enabled?: boolean; warningAcknowledged?: boolean; updatedAt?: string };
    const toolName = trimString(rec.toolName);
    if (!toolName || typeof rec.enabled !== "boolean") continue;
    out.push({
      toolName,
      enabled: rec.enabled,
      warningAcknowledged: rec.warningAcknowledged === true,
      updatedAt: trimString(rec.updatedAt),
    });
  }
  return out.length ? out : undefined;
}

function profileForConnector(connector: ConnectorCatalogItem, profileId: string | undefined) {
  if (!connector.setupProfiles?.length) return undefined;
  const explicit = trimString(profileId);
  if (explicit) {
    const profile = connector.setupProfiles.find((entry) => entry.id === explicit);
    if (profile) return profile;
  }
  return connector.setupProfiles.find((entry) => entry.defaultWhenCompanionAbsent)
    ?? connector.setupProfiles.find((entry) => entry.defaultWhenCompanionPresent)
    ?? connector.setupProfiles[0];
}

function classificationIsEnabledByDefault(classification: ConnectorToolClassification): boolean {
  return classification === "read" || classification === "sensitive_read" || classification === "costly_read";
}

function applyToolPolicyOverrides(
  inventory: ConnectorToolInventoryItem[],
  overrides: ConnectorToolPolicyOverride[] | undefined,
): ConnectorToolInventoryItem[] {
  if (!overrides?.length) return inventory;
  const byName = new Map(overrides.map((entry) => [entry.toolName, entry]));
  return inventory.map((tool) => {
    const override = byName.get(tool.name);
    return override ? { ...tool, enabled: override.enabled } : tool;
  });
}

function makeInventoryHash(values: string[]): string {
  return values.join("|").slice(0, 240) || "empty";
}

function buildCustomConnectorCatalogItem(name = "Custom MCP"): ConnectorCatalogItem {
  return {
    id: "custom",
    name,
    vendor: "Custom",
    iconKey: "custom",
    category: "knowledge",
    maturity: "custom_mcp_only",
    setupKind: "local_executable_or_docker",
    authMethod: "none",
    transport: "local_stdio",
    readOnly: true,
    customOnly: true,
    summary: "Connect an internal or self-hosted MCP through the guided setup flow.",
    officeValue: "Use this for internal knowledge, search, or data connectors that are not in the built-in library yet.",
    tags: ["custom", "mcp", "internal"],
    capabilityHints: ["verify safe tools", "reuse local credentials", "connect hosted MCPs"],
    requirements: [],
    envHints: [],
    readPolicy: {
      mode: "hard-read-only",
      allowResources: true,
      allowPrompts: false,
      allowToolPatterns: [
        "^(get|list|search|find|read|fetch|query|retrieve|describe|preview|inspect|lookup|resolve|show|view|count|whoami|stat)",
      ],
      blockToolPatterns: [
        "^(create|update|delete|remove|write|edit|insert|append|upload|post|put|patch|merge|commit|apply|send|message|reply|comment|close|reopen|archive|move|rename|grant|revoke|start|stop|restart|approve|reject|refund|charge|pay)",
      ],
      allowPromptPatterns: [],
      blockPromptPatterns: [".*"],
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
      "Local stdio connectors require a local runtime and are not available in browser-only mode.",
      "Remote MCP endpoints over HTTP can still be configured.",
    ],
    recommendedHosts: ["word", "excel", "powerpoint"],
    setupDifficulty: "advanced",
    catalogRevision: "custom-v1",
  };
}

export class BrowserConnectorRuntime {
  private state: StoredConnectorState = createConnectorState();
  readonly ready: Promise<void>;

  constructor() {
    this.ready = this.load();
  }

  getCatalogResponse(): ConnectorCatalogResponse {
    return {
      connectors: listConnectorCatalog(),
    };
  }

  buildCompanionSessionConnectors(scopeContext?: ConnectorScopeContext): CompanionConnectorDefinition[] {
    this.dropExpiredOAuthFlows();
    return this.state.connectors
      .filter((record) => this.computeScope(record, scopeContext).enabled)
      .map((record) => this.toCompanionConnectorDefinition(record))
      .filter((record): record is CompanionConnectorDefinition => Boolean(record));
  }

  buildCompanionConnectorDefinitionFromSetup(request: ConnectorSetupRequest): CompanionConnectorDefinition | undefined {
    const record = this.buildRecordFromRequest(request, false);
    return this.toCompanionConnectorDefinition(record);
  }

  buildCompanionConnectorDefinition(storedConnectorId: string): CompanionConnectorDefinition | undefined {
    const record = this.state.connectors.find((entry) => entry.id === storedConnectorId);
    if (!record) {
      return undefined;
    }
    return this.toCompanionConnectorDefinition(record);
  }

  getBrowserConnectorToolNames(scopeContext?: ConnectorScopeContext): string[] {
    this.dropExpiredOAuthFlows();
    const names = new Set<string>();
    for (const record of this.state.connectors) {
      if (!this.computeScope(record, scopeContext).enabled) continue;
      if (!this.toBrowserMcpConfig(record)) continue;
      for (const toolName of record.capabilities?.allowedTools ?? []) {
        names.add(toolName);
      }
    }
    return [...names].sort((left, right) => left.localeCompare(right));
  }

  async executeBrowserMcpTool(toolName: string, params: Record<string, unknown>, scopeContext?: ConnectorScopeContext): Promise<unknown> {
    const candidates = this.state.connectors.filter((record) =>
      this.computeScope(record, scopeContext).enabled &&
      Boolean(this.toBrowserMcpConfig(record)) &&
      Boolean(record.capabilities?.allowedTools.includes(toolName))
    );
    const record = candidates[0];
    if (!record) {
      throw new Error(`Connector tool "${toolName}" is not enabled for browser-direct execution.`);
    }
    const config = this.toBrowserMcpConfig(record);
    if (!config) {
      throw new Error(`Connector tool "${toolName}" requires the companion.`);
    }
    return executeBrowserMcpTool(config, toolName, params);
  }

  getStatusResponse(scopeContext?: ConnectorScopeContext): ConnectorStatusResponse {
    this.dropExpiredOAuthFlows();
    const connectors = this.state.connectors
      .map((record) => this.toStatus(record, scopeContext))
      .sort((left, right) => left.name.localeCompare(right.name));
    return { connectors };
  }

  getDiagnostics(): ConnectorDiagnosticsResponse {
    this.dropExpiredOAuthFlows();
    const runtimes: ConnectorRuntimeCheck[] = [
      { key: "node", label: "Node.js", ok: false, detail: "Unavailable in browser-only runtime." },
      { key: "npm", label: "npm", ok: false, detail: "Unavailable in browser-only runtime." },
      { key: "npx", label: "npx", ok: false, detail: "Unavailable in browser-only runtime." },
      { key: "python", label: "Python", ok: false, detail: "Unavailable in browser-only runtime." },
      { key: "uv", label: "uv", ok: false, detail: "Unavailable in browser-only runtime." },
      { key: "docker", label: "Docker", ok: false, detail: "Unavailable in browser-only runtime." },
      { key: "git", label: "Git", ok: false, detail: "Unavailable in browser-only runtime." },
    ];

    const envKeyUsage = new Set<string>();
    for (const record of this.state.connectors) {
      if (record.secretEnvKey) envKeyUsage.add(record.secretEnvKey);
      if (record.useDetectedEnvKey) envKeyUsage.add(record.useDetectedEnvKey);
    }

    const envSuggestions = this.collectEnvSuggestions(undefined).map((entry) => ({
      ...entry,
      present: envKeyUsage.has(entry.key),
    }));

    const diagnostics: ConnectorDiagnostic[] = [
      {
        level: "info",
        code: "connector_runtime_browser_mode",
        title: "Browser runtime",
        message: "Connector setup and state persistence are enabled in the taskpane runtime.",
      },
    ];

    if (this.state.connectors.some((record) => record.transport === "local_stdio")) {
      diagnostics.push({
        level: "warning",
        code: "local_stdio_requires_companion",
        title: "Local connector runtime unavailable",
        message: "Local stdio connectors need a local runtime bridge and cannot execute directly in browser-only mode.",
      });
    }

    return {
      generatedAt: nowIso(),
      runtimes,
      envSuggestions,
      diagnostics,
    };
  }

  getAuditPreferenceResponse(): ConnectorAuditPreferenceResponse {
    return { preference: this.state.auditPreference };
  }

  async setAuditPreference(preference: ConnectorAuditPreference): Promise<ConnectorAuditPreferenceResponse> {
    this.state.auditPreference = { enabled: preference.enabled !== false };
    this.state.updatedAt = nowIso();
    await this.persist();
    return { preference: this.state.auditPreference };
  }

  getExportBundle(): ConnectorExportBundle {
    const connectors: ConnectorExportItem[] = this.state.connectors.map((record) => ({
      storedId: record.id,
      connectorId: record.connectorId,
      name: record.name,
      source: record.source,
      category: record.category,
      maturity: record.maturity,
      setupKind: record.setupKind,
      authMethod: record.authMethod,
      transport: record.transport,
      setupProfileId: record.setupProfileId,
      credentialSource: record.credentialSource,
      secretEnvKey: record.secretEnvKey,
      useDetectedEnvKey: record.useDetectedEnvKey,
      url: record.url,
      command: record.command,
      args: record.args,
      cwd: record.cwd,
      env: record.env,
      stdioEnvPassthrough: record.stdioEnvPassthrough,
      remoteHttpHeaders: record.remoteHttpHeaders,
      remoteHttpHeadersFromEnv: record.remoteHttpHeadersFromEnv,
      toolPolicyOverrides: record.toolPolicyOverrides,
      suppressNonReadToolWarning: record.suppressNonReadToolWarning,
      defaultEnabled: record.defaultEnabled,
    }));

    const scopeOverrides: ConnectorExportScopeOverride[] = this.state.scopeOverrides.map((entry) => ({
      storedId: entry.connectorId,
      connectorId: this.state.connectors.find((record) => record.id === entry.connectorId)?.connectorId ?? entry.connectorId,
      scopeTarget: entry.scopeTarget,
      scopeKey: entry.scopeKey,
      enabled: entry.enabled,
    }));

    const favorites = this.state.connectors.filter((record) => record.favorite).map((record) => record.id);

    return {
      version: 1,
      exportedAt: nowIso(),
      connectors,
      scopeOverrides,
      favorites,
      auditPreference: this.state.auditPreference,
    };
  }

  previewImport(request: ConnectorImportPreviewRequest): ConnectorImportPreviewResponse {
    const conflicts = this.detectImportConflicts(request.bundle);
    return {
      ok: true,
      bundle: request.bundle,
      conflicts,
    };
  }

  async applyImport(request: ConnectorImportApplyRequest): Promise<ConnectorImportApplyResponse> {
    const bundle = request.bundle;
    const conflicts = this.detectImportConflicts(bundle);
    const conflictByKey = new Map(conflicts.map((conflict) => [conflict.key, conflict]));
    const importedConnectorIds: string[] = [];
    const skippedConflictKeys: string[] = [];
    const importedRecordIds = new Set<string>();

    for (const [index, item] of bundle.connectors.entries()) {
      const key = this.importConflictKey(item, index);
      const conflict = conflictByKey.get(key);
      const resolution = request.resolutions?.[key] ?? conflict?.resolution ?? "replace";
      if (conflict && resolution === "skip") {
        skippedConflictKeys.push(key);
        continue;
      }

      const existingById = trimString(item.storedId)
        ? this.state.connectors.find((record) => record.id === item.storedId)
        : undefined;
      const replaceTargetId = conflict?.existingConnectorId ?? existingById?.id;
      if (replaceTargetId) {
        this.state.connectors = this.state.connectors.filter((record) => record.id !== replaceTargetId);
        this.state.scopeOverrides = this.state.scopeOverrides.filter((entry) => entry.connectorId !== replaceTargetId);
      }

      const recordId = trimString(item.storedId)
        ?? (item.source === "library" ? item.connectorId : createRandomId("custom-"));
      const timestamp = nowIso();
      const next: StoredConnectorRecord = {
        id: recordId,
        connectorId: item.connectorId,
        source: item.source,
        name: trimString(item.name) ?? item.connectorId,
        category: item.category,
        maturity: item.maturity,
        setupKind: item.setupKind,
        authMethod: item.authMethod,
        transport: item.transport,
        credentialSource: item.credentialSource,
        setupProfileId: trimString(item.setupProfileId),
        secretEnvKey: trimString(item.secretEnvKey),
        useDetectedEnvKey: trimString(item.useDetectedEnvKey),
        url: normalizeUrl(item.url),
        command: trimString(item.command),
        args: normalizeArgs(item.args),
        cwd: trimString(item.cwd),
        env: normalizeEnv(item.env),
        stdioEnvPassthrough: normalizeStdioEnvPassthrough(item.stdioEnvPassthrough),
        remoteHttpHeaders: normalizeRemoteHttpHeaders(item.remoteHttpHeaders),
        remoteHttpHeadersFromEnv: normalizeRemoteHttpHeadersFromEnv(item.remoteHttpHeadersFromEnv),
        toolPolicyOverrides: normalizeToolPolicyOverrides(item.toolPolicyOverrides),
        suppressNonReadToolWarning: item.suppressNonReadToolWarning === true,
        defaultEnabled: item.defaultEnabled !== false,
        favorite: false,
        createdAt: timestamp,
        updatedAt: timestamp,
        oauthConnected: false,
        oauthExpiresAt: undefined,
        oauthLastAuthAt: undefined,
        oauthLastAuthError: item.credentialSource === "oauth"
          ? "OAuth credentials are not included in connector imports. Sign in again before use."
          : undefined,
        lastError: item.credentialSource === "oauth"
          ? "OAuth credentials are not included in connector imports. Sign in again before use."
          : undefined,
      };

      this.state.connectors.push(next);
      importedConnectorIds.push(next.id);
      importedRecordIds.add(next.id);
    }

    for (const override of bundle.scopeOverrides) {
      const recordId = trimString(override.storedId)
        ?? this.state.connectors.find((record) => record.source === "library" && record.connectorId === override.connectorId)?.id;
      if (!recordId || !importedRecordIds.has(recordId)) continue;
      this.upsertScopeOverride({
        connectorId: recordId,
        scopeTarget: override.scopeTarget,
        scopeKey: override.scopeKey,
        enabled: override.enabled,
      });
    }

    const favoriteIds = new Set(bundle.favorites);
    this.state.connectors = this.state.connectors.map((record) => ({
      ...record,
      favorite: favoriteIds.has(record.id),
    }));

    this.state.auditPreference = { enabled: bundle.auditPreference?.enabled !== false };
    this.appendLog({
      connectorId: undefined,
      connectorName: "Import",
      kind: "import",
      level: skippedConflictKeys.length ? "warning" : "info",
      message: skippedConflictKeys.length
        ? `Imported ${importedConnectorIds.length} connectors. ${skippedConflictKeys.length} conflict entries were skipped.`
        : `Imported ${importedConnectorIds.length} connectors.`,
    });
    await this.persist();
    return {
      ok: true,
      importedConnectorIds,
      skippedConflictKeys,
    };
  }

  prepareConnector(connectorId: string, scopeContext?: ConnectorScopeContext): ConnectorPrepareResponse {
    const connector = connectorId === "custom"
      ? buildCustomConnectorCatalogItem()
      : getConnectorCatalogItem(connectorId);
    if (!connector) {
      throw new Error("Unknown connector.");
    }

    const existing = connector.id === "custom"
      ? undefined
      : this.state.connectors.find((record) => record.source === "library" && record.connectorId === connector.id);
    const diagnostics: ConnectorDiagnostic[] = [];
    const profiles = connector.setupProfiles ?? [];
    const allProfilesNeedLocalCommand = profiles.length > 0 && profiles.every((profile) => profile.transport === "local_stdio");
    if (connector.transport === "local_stdio" || allProfilesNeedLocalCommand) {
      diagnostics.push({
        level: "warning",
        code: "local_stdio_requires_companion",
        title: "Local runtime unavailable",
        message: allProfilesNeedLocalCommand
          ? "This connector only supports a local command today. It needs the optional companion before setup can be verified."
          : "Local stdio connectors cannot execute directly in browser-only mode.",
        connectorId: connector.id,
      });
    }

    return {
      connector,
      existing: existing ? this.toStatus(existing, scopeContext) : undefined,
      draft: existing ? this.toSetupDraft(existing, scopeContext) : this.defaultDraft(connector),
      runtimes: this.runtimeChecksFor(connector),
      envSuggestions: this.collectEnvSuggestions(connector),
      diagnostics,
      auditPreference: this.state.auditPreference,
    };
  }

  async connectConnector(request: ConnectorSetupRequest): Promise<ConnectorSetupResponse> {
    const record = this.buildRecordFromRequest(request, true);
    const diagnostics = this.setupDiagnostics(record);
    this.upsertRecord(record);
    this.state.oauthFlows = this.state.oauthFlows.filter((entry) => entry.connectorId !== record.id);
    this.appendLog({
      connectorId: record.id,
      connectorName: record.name,
      kind: "setup",
      level: diagnostics.some((entry) => entry.level === "error") ? "warning" : "info",
      message: `Saved connector ${record.name}.`,
      scopeTarget: request.scopeTarget,
    });
    await this.persist();
    return {
      ok: true,
      status: this.toStatus(record, request.scopeContext),
      diagnostics,
    };
  }

  async testConnector(request: ConnectorSetupRequest): Promise<ConnectorTestResponse> {
    const record = this.buildRecordFromRequest(request, false);
    const verified = await this.verifyRecord(record);
    return {
      ok: verified.ok,
      status: this.toStatus(verified.next, request.scopeContext),
      diagnostics: verified.diagnostics,
    };
  }

  async reverifyConnector(connectorId: string, scopeContext?: ConnectorScopeContext): Promise<ConnectorTestResponse> {
    const record = this.state.connectors.find((entry) => entry.id === connectorId);
    if (!record) {
      throw new Error("Unknown connector.");
    }
    const verified = await this.verifyRecord(record);
    this.upsertRecord(verified.next);
    this.appendLog({
      connectorId: record.id,
      connectorName: record.name,
      kind: "reverify",
      level: verified.ok ? "info" : "warning",
      message: verified.ok ? `Connector ${record.name} verified.` : `Connector ${record.name} verification failed.`,
      healthState: this.toStatus(verified.next, scopeContext).healthState,
    });
    await this.persist();
    return {
      ok: verified.ok,
      status: this.toStatus(verified.next, scopeContext),
      diagnostics: verified.diagnostics,
    };
  }

  async startOAuth(connectorId: string): Promise<ConnectorOAuthStartResponse> {
    const stored = this.state.connectors.find((entry) => entry.id === connectorId);
    if (!stored) {
      throw new Error("Save this connector first, then start browser sign-in.");
    }
    const catalog = stored
      ? (stored.source === "library" ? getConnectorCatalogItem(stored.connectorId) : buildCustomConnectorCatalogItem(stored.name))
      : getConnectorCatalogItem(connectorId);
    if (stored.authMethod !== "oauth") {
      throw new Error("This connector does not use OAuth sign-in.");
    }

    const profile = catalog ? profileForConnector(catalog, stored.setupProfileId) : undefined;
    const canUseMcpOAuthDiscovery = profile?.transport === "remote_http" && profile.authMethod === "oauth" && !profileSetupDisabled(profile);
    if (!canUseMcpOAuthDiscovery) {
      const url = trimString(catalog?.authUrl);
      const state = createOAuthState();
      const issuedAt = nowIso();
      const expiresAt = new Date(Date.now() + 10 * 60 * 1000).toISOString();
      this.state.oauthFlows = this.state.oauthFlows.filter((entry) => entry.connectorId !== connectorId && Date.parse(entry.expiresAt) > Date.now());
      this.state.oauthFlows.push({
        connectorId,
        state,
        createdAt: issuedAt,
        expiresAt,
        url,
      });
      stored.oauthConnected = false;
      stored.oauthLastAuthError = "OAuth sign-in in progress.";
      stored.updatedAt = issuedAt;
      this.upsertRecord(stored);
      this.appendLog({
        connectorId,
        connectorName: stored.name ?? catalog?.name ?? connectorId,
        kind: "setup",
        level: "info",
        message: "Started OAuth sign-in flow.",
      });
      await this.persist();
      return { ok: true, connectorId, url, state, expiresAt };
    }
    const endpoint = normalizeUrl(stored.url ?? profile?.endpoint);
    const metadataUrl = metadataUrlForEndpoint(endpoint);
    if (!metadataUrl || typeof fetch !== "function") {
      throw new Error("This connector does not expose browser-readable OAuth metadata.");
    }
    const metadataResponse = await fetch(metadataUrl, { headers: { accept: "application/json" } });
    if (!metadataResponse.ok) {
      throw new Error(`OAuth metadata discovery failed with HTTP ${metadataResponse.status}.`);
    }
    const metadata = await metadataResponse.json() as {
      authorization_endpoint?: string;
      token_endpoint?: string;
      registration_endpoint?: string;
      scopes_supported?: string[];
    };
    const authorizationEndpoint = trimString(metadata.authorization_endpoint);
    const tokenEndpoint = trimString(metadata.token_endpoint);
    if (!authorizationEndpoint || !tokenEndpoint) {
      throw new Error("OAuth metadata did not include authorization and token endpoints.");
    }
    const redirectUri = oauthCallbackUrl();
    let clientId: string | undefined;
    let clientSecret: string | undefined;
    if (metadata.registration_endpoint) {
      const registrationResponse = await fetch(metadata.registration_endpoint, {
        method: "POST",
        headers: {
          "content-type": "application/json",
          accept: "application/json",
        },
        body: JSON.stringify({
          client_name: "Pi-Office",
          redirect_uris: [redirectUri],
          grant_types: ["authorization_code", "refresh_token"],
          response_types: ["code"],
          token_endpoint_auth_method: "none",
          application_type: "web",
        }),
      });
      if (!registrationResponse.ok) {
        throw new Error(`OAuth dynamic client registration failed with HTTP ${registrationResponse.status}.`);
      }
      const registration = await registrationResponse.json() as { client_id?: string; client_secret?: string };
      clientId = trimString(registration.client_id);
      clientSecret = trimString(registration.client_secret);
    }
    if (!clientId) {
      throw new Error("OAuth metadata did not provide Dynamic Client Registration for Pi-Office.");
    }

    const state = createOAuthState();
    const codeVerifier = createPkceVerifier();
    const codeChallenge = await createPkceChallenge(codeVerifier);
    const issuedAt = nowIso();
    const expiresAt = new Date(Date.now() + 10 * 60 * 1000).toISOString();
    const url = new URL(authorizationEndpoint);
    url.searchParams.set("response_type", "code");
    url.searchParams.set("client_id", clientId);
    url.searchParams.set("redirect_uri", redirectUri);
    url.searchParams.set("state", state);
    url.searchParams.set("code_challenge", codeChallenge);
    url.searchParams.set("code_challenge_method", "S256");
    this.state.oauthFlows = this.state.oauthFlows.filter((entry) => entry.connectorId !== connectorId && Date.parse(entry.expiresAt) > Date.now());
    this.state.oauthFlows.push({
      connectorId,
      state,
      createdAt: issuedAt,
      expiresAt,
      url: url.toString(),
      redirectUri,
      codeVerifier,
      tokenEndpoint,
      clientId,
      clientSecret,
    });

    stored.oauthConnected = false;
    stored.oauthLastAuthError = "OAuth sign-in in progress.";
    stored.updatedAt = issuedAt;
    this.upsertRecord(stored);

    this.appendLog({
      connectorId,
      connectorName: stored.name ?? catalog?.name ?? connectorId,
      kind: "setup",
      level: "info",
      message: "Started OAuth sign-in flow.",
    });
    await this.persist();
    return { ok: true, connectorId, url: url.toString(), callbackUrl: redirectUri, state, expiresAt };
  }

  async completeOAuth(request: ConnectorOAuthCallbackRequest, scopeContext?: ConnectorScopeContext): Promise<ConnectorOAuthCallbackResponse> {
    const state = trimString(request.state);
    if (!state) {
      throw new Error("OAuth state is required.");
    }
    const requestedConnectorId = trimString(request.connectorId);
    const pending = this.state.oauthFlows.find((entry) =>
      entry.state === state && (!requestedConnectorId || entry.connectorId === requestedConnectorId)
    );
    if (!pending) {
      throw new Error("OAuth callback state was not found or already completed.");
    }
    const connectorId = pending.connectorId;
    if (Date.parse(pending.expiresAt) <= Date.now()) {
      this.state.oauthFlows = this.state.oauthFlows.filter((entry) => !(entry.connectorId === connectorId && entry.state === state));
      await this.markOAuthIncomplete(connectorId, "OAuth callback state expired. Start sign-in again.");
      await this.persist();
      throw new Error("OAuth callback state expired. Start sign-in again.");
    }

    const record = this.state.connectors.find((entry) => entry.id === connectorId);
    if (!record) {
      throw new Error("Unknown connector.");
    }
    if (record.authMethod !== "oauth") {
      throw new Error("Connector is not configured for OAuth.");
    }

    const now = nowIso();
    let credential = request.credential;
    const code = trimString(request.code);
    const callbackError = trimString(request.error);
    if (!callbackError && !credential && code) {
      if (!pending.tokenEndpoint || !pending.clientId || !pending.redirectUri || !pending.codeVerifier) {
        await this.markOAuthIncomplete(connectorId, "OAuth callback could not be exchanged because the saved PKCE flow is incomplete.");
        throw new Error("OAuth callback could not be exchanged because the saved PKCE flow is incomplete.");
      }
      const form = new URLSearchParams();
      form.set("grant_type", "authorization_code");
      form.set("code", code);
      form.set("redirect_uri", pending.redirectUri);
      form.set("client_id", pending.clientId);
      form.set("code_verifier", pending.codeVerifier);
      if (pending.clientSecret) {
        form.set("client_secret", pending.clientSecret);
      }
      const tokenResponse = await fetch(pending.tokenEndpoint, {
        method: "POST",
        headers: {
          "content-type": "application/x-www-form-urlencoded",
          accept: "application/json",
        },
        body: form.toString(),
      });
      if (!tokenResponse.ok) {
        const bodyText = await tokenResponse.text().catch(() => "");
        await this.markOAuthIncomplete(connectorId, `OAuth token exchange failed with HTTP ${tokenResponse.status}.`);
        throw new Error(bodyText.trim() || `OAuth token exchange failed with HTTP ${tokenResponse.status}.`);
      }
      const token = await tokenResponse.json() as {
        access_token?: string;
        refresh_token?: string;
        token_type?: string;
        scope?: string;
        expires_in?: number;
      };
      const accessToken = trimString(token.access_token);
      if (accessToken) {
        credential = {
          accessToken,
          refreshToken: trimString(token.refresh_token),
          tokenType: trimString(token.token_type),
          scope: trimString(token.scope),
          expiresInSeconds: typeof token.expires_in === "number" ? token.expires_in : undefined,
        };
      }
    }
    const accessToken = trimString(credential?.accessToken);
    if (!accessToken || callbackError) {
      record.credentialSource = "oauth";
      record.secret = undefined;
      record.oauthConnected = false;
      record.oauthRefreshToken = undefined;
      record.oauthTokenType = undefined;
      record.oauthScope = undefined;
      record.oauthClientId = undefined;
      record.oauthTokenEndpoint = undefined;
      record.oauthLastAuthError = callbackError
        ?? "OAuth callback did not include a verified credential handoff. Manual completion is not supported.";
      record.oauthExpiresAt = undefined;
      record.lastError = record.oauthLastAuthError;
      record.updatedAt = now;
      this.state.oauthFlows = this.state.oauthFlows.filter((entry) => !(entry.connectorId === connectorId && entry.state === state));

      this.upsertRecord(record);
      const diagnostics: ConnectorDiagnostic[] = [{
        level: "warning",
        code: "oauth_not_completed",
        title: "OAuth incomplete",
        message: record.oauthLastAuthError,
        connectorId: record.connectorId,
      }];

      this.appendLog({
        connectorId: record.id,
        connectorName: record.name,
        kind: "setup",
        level: "warning",
        message: "OAuth sign-in incomplete.",
        healthState: this.toStatus(record, scopeContext).healthState,
      });
      await this.persist();
      return {
        ok: true,
        status: this.toStatus(record, scopeContext),
        diagnostics,
      };
    }

    record.credentialSource = "oauth";
    record.secret = accessToken;
    record.oauthRefreshToken = trimString(credential?.refreshToken);
    record.oauthTokenType = trimString(credential?.tokenType);
    record.oauthScope = trimString(credential?.scope);
    record.oauthClientId = pending.clientId;
    record.oauthTokenEndpoint = pending.tokenEndpoint;
    record.oauthConnected = true;
    record.oauthLastAuthAt = now;
    record.oauthLastAuthError = undefined;
    record.oauthExpiresAt = this.resolveOAuthExpiry({
      expiresAt: credential?.expiresAt ?? request.expiresAt,
      expiresInSeconds: credential?.expiresInSeconds ?? request.expiresInSeconds,
    }, now);
    record.lastError = undefined;
    record.lastTestedAt = now;
    record.lastHealthyAt = now;
    record.updatedAt = now;
    this.state.oauthFlows = this.state.oauthFlows.filter((entry) => !(entry.connectorId === connectorId && entry.state === state));

    this.upsertRecord(record);
    const diagnostics: ConnectorDiagnostic[] = [{
      level: "info",
      code: "oauth_completed",
      title: "OAuth complete",
      message: "OAuth callback included a verified credential handoff.",
      connectorId: record.connectorId,
    }];

    this.appendLog({
      connectorId: record.id,
      connectorName: record.name,
      kind: "setup",
      level: "info",
      message: "OAuth sign-in completed.",
      healthState: this.toStatus(record, scopeContext).healthState,
    });
    await this.persist();
    return {
      ok: true,
      status: this.toStatus(record, scopeContext),
      diagnostics,
    };
  }

  async removeConnector(storedConnectorId: string): Promise<{ ok: true }> {
    const existing = this.state.connectors.find((entry) => entry.id === storedConnectorId);
    this.state.connectors = this.state.connectors.filter((entry) => entry.id !== storedConnectorId);
    this.state.scopeOverrides = this.state.scopeOverrides.filter((entry) => entry.connectorId !== storedConnectorId);
    this.state.oauthFlows = this.state.oauthFlows.filter((entry) => entry.connectorId !== storedConnectorId);
    this.state.logs = this.state.logs.filter((entry) => entry.connectorId !== storedConnectorId);
    if (existing) {
      this.appendLog({
        connectorId: undefined,
        connectorName: existing.name,
        kind: "remove",
        level: "info",
        message: `Removed connector ${existing.name}.`,
      });
    }
    await this.persist();
    return { ok: true };
  }

  async clearAll(): Promise<{ ok: true }> {
    this.state = createConnectorState();
    try {
      localStorage.removeItem(CONNECTOR_STORAGE_KEY);
      localStorage.removeItem(CONNECTOR_CRYPTO_KEY_STORAGE_KEY);
    } catch {
      // Ignore storage errors in browser sandbox.
    }
    return { ok: true };
  }

  async setFavorite(request: ConnectorFavoriteRequest, scopeContext?: ConnectorScopeContext): Promise<{ ok: true; status: ConnectorStatus }> {
    const record = this.state.connectors.find((entry) => entry.id === request.connectorId);
    if (!record) {
      throw new Error("Unknown connector.");
    }
    record.favorite = request.favorite === true;
    record.updatedAt = nowIso();
    this.upsertRecord(record);
    this.appendLog({
      connectorId: record.id,
      connectorName: record.name,
      kind: "favorite",
      level: "info",
      message: request.favorite ? `Pinned connector ${record.name}.` : `Unpinned connector ${record.name}.`,
    });
    await this.persist();
    return { ok: true, status: this.toStatus(record, scopeContext) };
  }

  async updateScope(request: ConnectorScopeUpdateRequest): Promise<{ ok: true; status: ConnectorStatus }> {
    const record = this.state.connectors.find((entry) => entry.id === request.connectorId);
    if (!record) {
      throw new Error("Unknown connector.");
    }

    if (request.scopeTarget === "global") {
      record.defaultEnabled = request.enabled;
      record.updatedAt = nowIso();
      this.upsertRecord(record);
    } else {
      const scopeKey = request.scopeTarget === "workspace"
        ? resolveWorkspaceScopeKey(request.scopeContext)
        : resolveDocumentScopeKey(request.scopeContext);
      if (!scopeKey) {
        throw new Error(request.scopeTarget === "workspace"
          ? "Workspace scope requires a saved document or workspace context."
          : "Document scope requires a saved document context.");
      }
      this.upsertScopeOverride({
        connectorId: record.id,
        scopeTarget: request.scopeTarget,
        scopeKey,
        enabled: request.enabled,
      });
    }

    this.appendLog({
      connectorId: record.id,
      connectorName: record.name,
      kind: "scope",
      level: "info",
      message: `${connectorLabel(request.scopeTarget)} scope ${request.enabled ? "enabled" : "disabled"} for ${record.name}.`,
      scopeTarget: request.scopeTarget,
    });
    await this.persist();
    const refreshed = this.state.connectors.find((entry) => entry.id === record.id) ?? record;
    return { ok: true, status: this.toStatus(refreshed, request.scopeContext) };
  }

  async updateToolPolicy(request: ConnectorToolPolicyUpdateRequest): Promise<ConnectorToolPolicyUpdateResponse> {
    const record = this.state.connectors.find((entry) => entry.id === request.connectorId);
    if (!record) {
      throw new Error("Unknown connector.");
    }
    const toolName = trimString(request.toolName);
    if (!toolName) {
      throw new Error("toolName is required.");
    }

    const existing = record.toolPolicyOverrides ?? [];
    const nextOverride: ConnectorToolPolicyOverride = {
      toolName,
      enabled: request.enabled === true,
      warningAcknowledged: request.warningAcknowledged === true,
      updatedAt: nowIso(),
    };
    record.toolPolicyOverrides = [
      ...existing.filter((entry) => entry.toolName !== toolName),
      nextOverride,
    ];
    if (request.suppressWarning === true) {
      record.suppressNonReadToolWarning = true;
    }
    record.capabilities = record.capabilities
      ? {
          ...record.capabilities,
          toolInventory: record.capabilities.toolInventory
            ? applyToolPolicyOverrides(record.capabilities.toolInventory, record.toolPolicyOverrides)
            : undefined,
          allowedTools: (record.capabilities.toolInventory
            ? applyToolPolicyOverrides(record.capabilities.toolInventory, record.toolPolicyOverrides)
            : record.capabilities.allowedTools.map((name): ConnectorToolInventoryItem => ({
                name,
                classification: "read",
                defaultEnabled: true,
                enabled: true,
                reason: "Previously verified read-safe tool.",
                source: "catalog_hint",
              })))
            .filter((tool) => tool.enabled)
            .map((tool) => tool.name),
          blockedTools: (record.capabilities.toolInventory
            ? applyToolPolicyOverrides(record.capabilities.toolInventory, record.toolPolicyOverrides)
            : [])
            .filter((tool) => !tool.enabled)
            .map((tool) => tool.name),
        }
      : this.buildCapabilitySummary(record);
    if (record.verification) {
      const inventory = record.capabilities.toolInventory;
      record.verification = {
        ...record.verification,
        toolInventory: inventory,
        allowedTools: record.capabilities.allowedTools,
        blockedTools: record.capabilities.blockedTools,
      };
    }
    record.updatedAt = nowIso();
    this.upsertRecord(record);
    this.appendLog({
      connectorId: record.id,
      connectorName: record.name,
      kind: "scope",
      level: "info",
      message: `${request.enabled ? "Enabled" : "Disabled"} connector tool ${toolName}.`,
    });
    await this.persist();
    return { ok: true, status: this.toStatus(record, request.scopeContext) };
  }

  getLogs(connectorId: string): ConnectorLogResponse {
    return {
      connectorId,
      auditPreference: this.state.auditPreference,
      logs: this.state.logs
        .filter((entry) => entry.connectorId === connectorId)
        .sort((left, right) => right.timestamp.localeCompare(left.timestamp)),
    };
  }

  private toCompanionConnectorDefinition(record: StoredConnectorRecord): CompanionConnectorDefinition | undefined {
    const catalog = record.source === "library"
      ? getConnectorCatalogItem(record.connectorId)
      : buildCustomConnectorCatalogItem(record.name);
    if (!catalog) {
      return undefined;
    }
    const profile = profileForConnector(catalog, record.setupProfileId);
    if (!profileNeedsCompanion(profile, record.transport)) {
      return undefined;
    }

    return {
      id: record.id,
      connectorId: record.connectorId,
      name: record.name,
      source: record.source,
      category: record.category,
      maturity: record.maturity,
      setupKind: record.setupKind,
      authMethod: record.authMethod,
      transport: record.transport,
      credentialSource: record.credentialSource,
      setupProfileId: record.setupProfileId,
      url: record.url,
      command: record.command,
      args: record.args,
      cwd: record.cwd,
      env: record.env,
      stdioEnvPassthrough: record.stdioEnvPassthrough,
      remoteHttpHeaders: record.remoteHttpHeaders,
      remoteHttpHeadersFromEnv: record.remoteHttpHeadersFromEnv,
      secret: record.secret,
      secretEnvKey: record.secretEnvKey,
      useDetectedEnvKey: record.useDetectedEnvKey,
      readOnly: catalog.readOnly,
      readPolicy: catalog.readPolicy,
      toolPolicyOverrides: record.toolPolicyOverrides,
      catalogRevision: catalog.catalogRevision,
    };
  }

  private toBrowserMcpConfig(record: StoredConnectorRecord): BrowserMcpConnectorConfig | undefined {
    const catalog = record.source === "library"
      ? getConnectorCatalogItem(record.connectorId)
      : buildCustomConnectorCatalogItem(record.name);
    if (!catalog) return undefined;
    const profile = profileForConnector(catalog, record.setupProfileId);
    if (!profileIsBrowserDirect(profile)) return undefined;
    if (record.transport !== "remote_http") return undefined;
    const url = normalizeUrl(record.url ?? profile?.endpoint);
    if (!url) return undefined;
    if (record.remoteHttpHeadersFromEnv?.length) return undefined;
    if (record.credentialSource === "env" || record.credentialSource === "detected_env") return undefined;
    if (record.authMethod === "oauth" && this.needsCredential(record)) return undefined;
    if ((record.authMethod === "api_key" || record.authMethod === "bearer_token") && this.needsCredential(record)) return undefined;
    return {
      connectorId: record.connectorId,
      name: record.name,
      url,
      accessToken: trimString(record.secret),
      headers: record.remoteHttpHeaders,
      readPolicy: catalog.readPolicy,
      toolPolicyOverrides: record.toolPolicyOverrides,
      catalogRevision: catalog.catalogRevision,
    };
  }

  private toStatus(record: StoredConnectorRecord, scopeContext?: ConnectorScopeContext): ConnectorStatus {
    const { activeScope, enabled, scopeStates } = this.computeScope(record, scopeContext);
    const needsCredential = this.needsCredential(record);
    const healthState = this.computeHealth(record, enabled, needsCredential);
    const pendingOAuth = this.state.oauthFlows.find((entry) => entry.connectorId === record.id && Date.parse(entry.expiresAt) > Date.now());
    const logEntries = this.state.logs.filter((entry) => entry.connectorId === record.id);
    const lastLog = [...logEntries].sort((left, right) => right.timestamp.localeCompare(left.timestamp))[0];
    const staleReason =
      healthState === "stale"
        ? "Re-run connector verification."
        : healthState === "auth_required"
          ? (pendingOAuth ? "OAuth sign-in started. Complete sign-in to activate this connector." : "Complete OAuth sign-in to activate this connector.")
          : healthState === "auth_expired"
            ? "OAuth session expired. Re-authenticate to continue."
            : undefined;
    return {
      id: record.id,
      connectorId: record.connectorId,
      name: record.name,
      enabled,
      configured: true,
      connected: enabled && healthState === "ready",
      source: record.source,
      category: record.category,
      maturity: record.maturity,
      setupKind: record.setupKind,
      authMethod: record.authMethod,
      transport: record.transport,
      setupProfileId: record.setupProfileId,
      credentialSource: record.credentialSource,
      detectedEnvKey: record.useDetectedEnvKey,
      usesDetectedCredential: record.credentialSource === "detected_env" && Boolean(record.useDetectedEnvKey),
      lastTestedAt: record.lastTestedAt,
      lastError: record.lastError,
      lastHealthyAt: record.lastHealthyAt,
      healthState,
      staleReason,
      activeScope,
      scopeStates,
      favorite: record.favorite,
      recommendedHosts: (getConnectorCatalogItem(record.connectorId)?.recommendedHosts ?? ["word", "excel", "powerpoint"]),
      setupDifficulty: getConnectorCatalogItem(record.connectorId)?.setupDifficulty ?? "advanced",
      needsCredential,
      verification: record.verification,
      capabilities: record.capabilities,
      logSummary: {
        total: logEntries.length,
        lastAt: lastLog?.timestamp,
        lastLevel: lastLog?.level,
      },
      suppressNonReadToolWarning: record.suppressNonReadToolWarning === true,
    };
  }

  private computeScope(record: StoredConnectorRecord, scopeContext?: ConnectorScopeContext): {
    activeScope: ConnectorScopeTarget;
    enabled: boolean;
    scopeStates: ConnectorScopeState[];
  } {
    const workspaceKey = resolveWorkspaceScopeKey(scopeContext);
    const documentKey = resolveDocumentScopeKey(scopeContext);
    const workspaceOverride = workspaceKey
      ? this.state.scopeOverrides.find((entry) =>
        entry.connectorId === record.id && entry.scopeTarget === "workspace" && entry.scopeKey === workspaceKey)
      : undefined;
    const documentOverride = documentKey
      ? this.state.scopeOverrides.find((entry) =>
        entry.connectorId === record.id && entry.scopeTarget === "document" && entry.scopeKey === documentKey)
      : undefined;

    let activeScope: ConnectorScopeTarget = "global";
    let enabled = record.defaultEnabled;
    if (documentOverride) {
      activeScope = "document";
      enabled = documentOverride.enabled;
    } else if (workspaceOverride) {
      activeScope = "workspace";
      enabled = workspaceOverride.enabled;
    }

    const scopeStates: ConnectorScopeState[] = [
      {
        target: "global",
        enabled: record.defaultEnabled,
        applies: activeScope === "global",
        label: connectorLabel("global"),
      },
      {
        target: "workspace",
        enabled: workspaceOverride?.enabled ?? record.defaultEnabled,
        applies: activeScope === "workspace",
        label: connectorLabel("workspace"),
        scopeKey: workspaceKey,
        reason: workspaceKey ? undefined : "Save the document to enable folder scope.",
      },
      {
        target: "document",
        enabled: documentOverride?.enabled ?? record.defaultEnabled,
        applies: activeScope === "document",
        label: connectorLabel("document"),
        scopeKey: documentKey,
        reason: documentKey ? undefined : "Save the document to enable document scope.",
      },
    ];

    return { activeScope, enabled, scopeStates };
  }

  private needsCredential(record: StoredConnectorRecord): boolean {
    if (record.authMethod === "none") return false;
    if (record.credentialSource === "manual") return !trimString(record.secret);
    if (record.credentialSource === "env") return !trimString(record.secretEnvKey);
    if (record.credentialSource === "detected_env") return !trimString(record.useDetectedEnvKey);
    if (record.credentialSource === "oauth") {
      if (!trimString(record.secret)) return true;
      if (record.oauthConnected !== true) return true;
      if (record.oauthExpiresAt && Date.parse(record.oauthExpiresAt) <= Date.now()) return true;
      return false;
    }
    return true;
  }

  private resolveOAuthExpiry(
    request: Pick<ConnectorOAuthCallbackRequest, "expiresAt" | "expiresInSeconds">,
    fallbackIso: string,
  ): string {
    const explicitExpiry = trimString(request.expiresAt);
    if (explicitExpiry) {
      const parsed = Date.parse(explicitExpiry);
      if (Number.isFinite(parsed)) {
        return new Date(parsed).toISOString();
      }
    }
    const seconds = Number.isFinite(request.expiresInSeconds)
      ? Math.max(60, Math.min(60 * 60 * 24 * 30, Math.floor(request.expiresInSeconds as number)))
      : 60 * 60 * 24;
    return new Date(Date.parse(fallbackIso) + seconds * 1000).toISOString();
  }

  private async markOAuthIncomplete(connectorId: string, message: string): Promise<void> {
    const record = this.state.connectors.find((entry) => entry.id === connectorId);
    if (!record || record.authMethod !== "oauth") {
      return;
    }
    if (record.oauthConnected === true && trimString(record.secret)) {
      record.oauthLastAuthError = message;
      record.lastError = message;
    } else {
      record.credentialSource = "oauth";
      record.secret = undefined;
      record.oauthConnected = false;
      record.oauthExpiresAt = undefined;
      record.oauthLastAuthError = message;
      record.lastError = message;
    }
    record.updatedAt = nowIso();
    this.upsertRecord(record);
    this.appendLog({
      connectorId: record.id,
      connectorName: record.name,
      kind: "setup",
      level: "warning",
      message,
      healthState: this.toStatus(record).healthState,
    });
    await this.persist();
  }

  private async refreshOAuthIfNeeded(record: StoredConnectorRecord): Promise<StoredConnectorRecord> {
    if (
      record.credentialSource !== "oauth" ||
      !record.oauthRefreshToken ||
      !record.oauthTokenEndpoint ||
      !record.oauthClientId ||
      !record.oauthExpiresAt ||
      Date.parse(record.oauthExpiresAt) > Date.now() + 60_000
    ) {
      return record;
    }
    const form = new URLSearchParams();
    form.set("grant_type", "refresh_token");
    form.set("refresh_token", record.oauthRefreshToken);
    form.set("client_id", record.oauthClientId);
    try {
      const response = await fetch(record.oauthTokenEndpoint, {
        method: "POST",
        headers: {
          "content-type": "application/x-www-form-urlencoded",
          accept: "application/json",
        },
        body: form.toString(),
      });
      if (!response.ok) {
        return {
          ...record,
          oauthConnected: false,
          oauthLastAuthError: `OAuth token refresh failed with HTTP ${response.status}.`,
          lastError: `OAuth token refresh failed with HTTP ${response.status}.`,
        };
      }
      const token = await response.json() as {
        access_token?: string;
        refresh_token?: string;
        token_type?: string;
        scope?: string;
        expires_in?: number;
      };
      const accessToken = trimString(token.access_token);
      if (!accessToken) {
        return {
          ...record,
          oauthConnected: false,
          oauthLastAuthError: "OAuth token refresh did not return an access token.",
          lastError: "OAuth token refresh did not return an access token.",
        };
      }
      const now = nowIso();
      return {
        ...record,
        secret: accessToken,
        oauthRefreshToken: trimString(token.refresh_token) ?? record.oauthRefreshToken,
        oauthTokenType: trimString(token.token_type) ?? record.oauthTokenType,
        oauthScope: trimString(token.scope) ?? record.oauthScope,
        oauthExpiresAt: this.resolveOAuthExpiry({ expiresInSeconds: token.expires_in }, now),
        oauthConnected: true,
        oauthLastAuthAt: now,
        oauthLastAuthError: undefined,
        lastError: undefined,
        updatedAt: now,
      };
    } catch (error) {
      return {
        ...record,
        oauthConnected: false,
        oauthLastAuthError: error instanceof Error ? error.message : "OAuth token refresh failed.",
        lastError: error instanceof Error ? error.message : "OAuth token refresh failed.",
      };
    }
  }

  private computeHealth(record: StoredConnectorRecord, enabled: boolean, needsCredential: boolean): ConnectorHealthState {
    if (!enabled) {
      return record.lastHealthyAt ? "stale" : "unverified";
    }
    if (record.credentialSource === "oauth" && record.oauthExpiresAt && Date.parse(record.oauthExpiresAt) <= Date.now()) {
      return "auth_expired";
    }
    if (needsCredential) return "auth_required";
    if (record.lastError && !record.lastHealthyAt) return "offline";
    if (record.lastError) return "degraded";
    if (record.lastHealthyAt) return "ready";
    if (record.lastTestedAt) return "stale";
    return "unverified";
  }

  private toSetupDraft(record: StoredConnectorRecord, scopeContext?: ConnectorScopeContext): ConnectorSetupRequest {
    return {
      connectorId: record.connectorId,
      existingId: record.id,
      name: record.name,
      enabled: record.defaultEnabled,
      favorite: record.favorite,
      scopeTarget: this.computeScope(record, scopeContext).activeScope,
      setupKind: record.setupKind,
      authMethod: record.authMethod,
      transport: record.transport,
      credentialSource: record.credentialSource,
      secretEnvKey: record.secretEnvKey,
      useDetectedEnvKey: record.useDetectedEnvKey,
      url: record.url,
      command: record.command,
      args: record.args,
      cwd: record.cwd,
      env: record.env,
      stdioEnvPassthrough: record.stdioEnvPassthrough,
      remoteHttpHeaders: record.remoteHttpHeaders,
      remoteHttpHeadersFromEnv: record.remoteHttpHeadersFromEnv,
      preserveStoredSecret: Boolean(record.secret),
    };
  }

  private defaultDraft(connector: ConnectorCatalogItem): ConnectorSetupRequest {
    const profile = profileForConnector(connector, undefined);
    return {
      connectorId: connector.id,
      name: connector.name,
      enabled: true,
      favorite: false,
      scopeTarget: "global",
      setupProfileId: profile?.id,
      setupKind: profile?.setupKind ?? connector.setupKind,
      authMethod: profile?.authMethod ?? connector.authMethod,
      transport: profile?.transport ?? connector.transport,
      credentialSource: (profile?.authMethod ?? connector.authMethod) === "none"
        ? "none"
        : (profile?.authMethod ?? connector.authMethod) === "oauth"
          ? "oauth"
          : "manual",
      secretEnvKey: profile?.credentialEnvKey ?? connector.envHints[0]?.key,
      url: profile?.transport === "remote_http" ? (profile.endpoint ?? connector.template?.url) : connector.template?.url,
      command: profile?.transport === "local_stdio" ? (profile.command ?? connector.template?.command) : connector.template?.command,
      args: profile?.transport === "local_stdio" ? (profile.args ?? connector.template?.args) : connector.template?.args,
      cwd: profile?.transport === "local_stdio" ? (profile.cwd ?? connector.template?.cwd) : connector.template?.cwd,
      env: profile?.env ?? connector.template?.env,
      remoteHttpHeaders: profile?.remoteHttpHeaders,
      remoteHttpHeadersFromEnv: profile?.remoteHttpHeadersFromEnv,
    };
  }

  private runtimeChecksFor(connector: ConnectorCatalogItem): ConnectorRuntimeCheck[] {
    const profile = profileForConnector(connector, undefined);
    if (profile?.setupDisabled || profile?.availability === "planned") {
      return [
        { key: "git", label: "Connector setup", ok: false, detail: profile.riskNotes?.[0] ?? "This profile is visible as roadmap but not available yet." },
      ];
    }
    if (connector.transport === "local_stdio" || profile?.transport === "local_stdio") {
      return [
        { key: "node", label: "Node.js", ok: false, detail: "Local runtime unavailable in browser-only mode." },
        { key: "python", label: "Python", ok: false, detail: "Local runtime unavailable in browser-only mode." },
      ];
    }
    if (profileIsBrowserDirect(profile)) {
      return [
        {
          key: "git",
          label: "Browser-direct MCP",
          ok: true,
          detail: "This hosted MCP can be verified and executed directly from the taskpane after sign-in.",
        },
      ];
    }
    return [
      {
        key: "git",
        label: "Remote connector execution",
        ok: false,
        detail: "Remote HTTP MCP connectors are verified and executed through the optional companion.",
      },
    ];
  }

  private collectEnvSuggestions(connector: ConnectorCatalogItem | undefined): ConnectorEnvSuggestion[] {
    const keys = new Set<string>();
    const source = connector ? [connector] : listConnectorCatalog();
    for (const item of source) {
      for (const hint of item.envHints) {
        if (hint.key) keys.add(hint.key);
      }
    }

    return [...keys]
      .sort((left, right) => left.localeCompare(right))
      .map((key) => ({ key, present: false }));
  }

  private buildRecordFromRequest(request: ConnectorSetupRequest, persistTarget: boolean): StoredConnectorRecord {
    const connectorId = trimString(request.connectorId);
    if (!connectorId) {
      throw new Error("connectorId is required.");
    }

    const catalogItem = connectorId === "custom"
      ? buildCustomConnectorCatalogItem(trimString(request.name) ?? "Custom MCP")
      : getConnectorCatalogItem(connectorId);
    const source: "library" | "custom" = catalogItem && connectorId !== "custom" ? "library" : "custom";
    const base = catalogItem ?? buildCustomConnectorCatalogItem(trimString(request.name) ?? "Custom MCP");
    const profile = profileForConnector(base, request.setupProfileId);

    const existing = trimString(request.existingId)
      ? this.state.connectors.find((entry) => entry.id === request.existingId)
      : source === "library"
        ? this.state.connectors.find((entry) => entry.source === "library" && entry.connectorId === connectorId)
        : undefined;

    const id = persistTarget
      ? (trimString(request.existingId)
          ?? existing?.id
          ?? (source === "library" ? connectorId : createRandomId("custom-")))
      : (trimString(request.existingId) ?? existing?.id ?? (source === "library" ? connectorId : createRandomId("preview-")));

    const authMethod = request.authMethod ?? profile?.authMethod ?? base.authMethod;
    const credentialSource = request.credentialSource
      ?? (authMethod === "none" ? "none" : authMethod === "oauth" ? "oauth" : "manual");
    const preserveStoredSecret = request.preserveStoredSecret === true;
    const nextSecret = credentialSource === "manual"
      ? (trimString(request.secret) ?? (preserveStoredSecret ? existing?.secret : undefined))
      : credentialSource === "oauth" && existing?.credentialSource === "oauth"
        ? trimString(existing.secret)
      : undefined;
    const hasOAuthCredential = credentialSource === "oauth" && Boolean(nextSecret);

    return {
      id,
      connectorId: source === "library" ? connectorId : "custom",
      source,
      name: trimString(request.name) ?? existing?.name ?? base.name,
      category: source === "library" ? base.category : "knowledge",
      maturity: source === "library" ? base.maturity : "custom_mcp_only",
      setupKind: request.setupKind ?? profile?.setupKind ?? existing?.setupKind ?? base.setupKind,
      authMethod,
      transport: request.transport ?? profile?.transport ?? existing?.transport ?? base.transport,
      setupProfileId: trimString(request.setupProfileId) ?? profile?.id ?? existing?.setupProfileId,
      credentialSource,
      secret: nextSecret,
      secretEnvKey: trimString(request.secretEnvKey) ?? profile?.credentialEnvKey,
      useDetectedEnvKey: trimString(request.useDetectedEnvKey),
      url: normalizeUrl(request.url ?? profile?.endpoint),
      command: trimString(request.command) ?? profile?.command,
      args: normalizeArgs(request.args) ?? profile?.args,
      cwd: trimString(request.cwd) ?? profile?.cwd,
      env: normalizeEnv(request.env) ?? profile?.env,
      stdioEnvPassthrough:
        request.stdioEnvPassthrough !== undefined
          ? normalizeStdioEnvPassthrough(request.stdioEnvPassthrough)
          : existing?.stdioEnvPassthrough,
      remoteHttpHeaders:
        request.remoteHttpHeaders !== undefined
          ? normalizeRemoteHttpHeaders(request.remoteHttpHeaders)
          : existing?.remoteHttpHeaders ?? profile?.remoteHttpHeaders,
      remoteHttpHeadersFromEnv:
        request.remoteHttpHeadersFromEnv !== undefined
          ? normalizeRemoteHttpHeadersFromEnv(request.remoteHttpHeadersFromEnv)
          : existing?.remoteHttpHeadersFromEnv ?? profile?.remoteHttpHeadersFromEnv,
      defaultEnabled: request.enabled !== false,
      favorite: request.favorite === true,
      createdAt: existing?.createdAt ?? nowIso(),
      updatedAt: nowIso(),
      lastTestedAt: existing?.lastTestedAt,
      lastHealthyAt: existing?.lastHealthyAt,
      lastError: existing?.lastError,
      verification: existing?.verification,
      capabilities: existing?.capabilities,
      toolPolicyOverrides: existing?.toolPolicyOverrides,
      suppressNonReadToolWarning: existing?.suppressNonReadToolWarning,
      oauthConnected: credentialSource === "oauth" ? (hasOAuthCredential && existing?.oauthConnected === true) : undefined,
      oauthExpiresAt: credentialSource === "oauth" && hasOAuthCredential ? existing?.oauthExpiresAt : undefined,
      oauthLastAuthAt: credentialSource === "oauth" && hasOAuthCredential ? existing?.oauthLastAuthAt : undefined,
      oauthLastAuthError: credentialSource === "oauth" ? existing?.oauthLastAuthError : undefined,
      oauthRefreshToken: credentialSource === "oauth" && hasOAuthCredential ? existing?.oauthRefreshToken : undefined,
      oauthTokenType: credentialSource === "oauth" && hasOAuthCredential ? existing?.oauthTokenType : undefined,
      oauthScope: credentialSource === "oauth" && hasOAuthCredential ? existing?.oauthScope : undefined,
      oauthClientId: credentialSource === "oauth" && hasOAuthCredential ? existing?.oauthClientId : undefined,
      oauthTokenEndpoint: credentialSource === "oauth" && hasOAuthCredential ? existing?.oauthTokenEndpoint : undefined,
    };
  }

  private setupDiagnostics(record: StoredConnectorRecord): ConnectorDiagnostic[] {
    const diagnostics: ConnectorDiagnostic[] = [];
    const catalog = record.source === "library"
      ? getConnectorCatalogItem(record.connectorId)
      : buildCustomConnectorCatalogItem(record.name);
    const profile = profileForConnector(catalog ?? buildCustomConnectorCatalogItem(record.name), record.setupProfileId);
    if (profileSetupDisabled(profile)) {
      diagnostics.push({
        level: "error",
        code: "connector_profile_not_available",
        title: "Setup not available",
        message: profile?.riskNotes?.[0] ?? "This connector profile is visible as roadmap but is not available for setup yet.",
        connectorId: record.connectorId,
      });
    }
    if (record.transport === "local_stdio") {
      diagnostics.push({
        level: "warning",
        code: "local_stdio_requires_companion",
        title: "Local runtime unavailable",
        message: "Local stdio connectors can be saved but cannot execute directly in browser-only mode.",
        connectorId: record.connectorId,
      });
    }
    if (record.transport === "remote_http" && !trimString(record.url)) {
      diagnostics.push({
        level: "error",
        code: "remote_url_required",
        title: "Remote URL required",
        message: "Remote HTTP connectors require a URL.",
        connectorId: record.connectorId,
      });
    }
    if (record.credentialSource === "oauth") {
      if (record.oauthExpiresAt && Date.parse(record.oauthExpiresAt) <= Date.now()) {
        diagnostics.push({
          level: "warning",
          code: "oauth_expired",
          title: "OAuth session expired",
          message: "Re-authenticate this connector to continue.",
          connectorId: record.connectorId,
        });
      } else if (record.oauthConnected !== true) {
        diagnostics.push({
          level: "warning",
          code: "oauth_required",
          title: "OAuth sign-in required",
          message: "Complete browser sign-in before verifying this connector.",
          connectorId: record.connectorId,
        });
      }
      if (record.oauthLastAuthError) {
        diagnostics.push({
          level: "warning",
          code: "oauth_last_error",
          title: "Last sign-in issue",
          message: record.oauthLastAuthError,
          connectorId: record.connectorId,
        });
      }
    } else if (this.needsCredential(record)) {
      diagnostics.push({
        level: "warning",
        code: "credential_required",
        title: "Credential required",
        message: "This connector needs credentials before it can become ready.",
        connectorId: record.connectorId,
      });
    }
    return diagnostics;
  }

  private async verifyRecord(record: StoredConnectorRecord): Promise<VerificationResult> {
    record = await this.refreshOAuthIfNeeded(record);
    const diagnostics = this.setupDiagnostics(record);
    const verifiedAt = nowIso();
    const browserConfig = this.toBrowserMcpConfig(record);
    if (!diagnostics.some((entry) => entry.level === "error") && browserConfig) {
      const probe = await probeBrowserMcpConnector(browserConfig);
      const probeDiagnostics: ConnectorDiagnostic[] = probe.diagnostics.map((message): ConnectorDiagnostic => ({
        level: probe.ok ? "info" : "warning",
        code: probe.ok ? "browser_mcp_verified" : "browser_mcp_failed",
        title: probe.ok ? "Browser MCP verified" : "Browser MCP verification failed",
        message,
        connectorId: record.connectorId,
      }));
      const capabilities = probe.capabilities ?? this.buildCapabilitySummary(record);
      const next: StoredConnectorRecord = {
        ...record,
        lastTestedAt: verifiedAt,
        lastHealthyAt: probe.ok ? verifiedAt : record.lastHealthyAt,
        lastError: probe.ok ? undefined : (probeDiagnostics[0]?.message ?? "Browser MCP verification failed."),
        oauthLastAuthError: probe.ok ? undefined : record.oauthLastAuthError,
        capabilities,
        verification: probe.verification ?? {
          verifiedAt,
          catalogRevision: getConnectorCatalogItem(record.connectorId)?.catalogRevision ?? "browser-direct-v1",
          inventoryHash: capabilities.inventoryHash ?? makeInventoryHash(capabilities.tools),
          toolNames: capabilities.tools,
          allowedTools: capabilities.allowedTools,
          blockedTools: capabilities.blockedTools,
          toolInventory: capabilities.toolInventory,
          resourceToolNames: [],
          promptNames: [],
          allowedPrompts: [],
          blockedPrompts: [],
          resourceCount: 0,
          promptCount: 0,
          staleReason: probe.ok ? undefined : (probeDiagnostics[0]?.message ?? "Browser MCP verification failed."),
          lastFailure: probe.ok ? undefined : probeDiagnostics.map((entry) => entry.message).join("; "),
        },
      };
      return { ok: probe.ok, diagnostics: [...diagnostics, ...probeDiagnostics], next };
    }

    const ok = !diagnostics.some((entry) => entry.level === "error") && !this.needsCredential(record);
    const capabilities = this.buildCapabilitySummary(record);
    const tools = capabilities.tools;
    const next: StoredConnectorRecord = {
      ...record,
      lastTestedAt: verifiedAt,
      lastHealthyAt: ok ? verifiedAt : record.lastHealthyAt,
      lastError: ok ? undefined : (record.oauthLastAuthError ?? diagnostics[0]?.message ?? "Verification failed."),
      oauthLastAuthError: ok ? undefined : record.oauthLastAuthError,
      capabilities,
      verification: {
        verifiedAt,
        catalogRevision: getConnectorCatalogItem(record.connectorId)?.catalogRevision ?? "browser-v1",
        inventoryHash: makeInventoryHash(tools),
        toolNames: tools,
        allowedTools: capabilities.allowedTools,
        blockedTools: capabilities.blockedTools,
        toolInventory: capabilities.toolInventory,
        resourceToolNames: [],
        promptNames: [],
        allowedPrompts: [],
        blockedPrompts: [],
        resourceCount: 0,
        promptCount: 0,
        staleReason: ok ? undefined : (diagnostics[0]?.message ?? "Verification failed."),
        lastFailure: ok ? undefined : diagnostics.map((entry) => entry.message).join("; "),
      },
    };

    return { ok, diagnostics, next };
  }

  private buildCapabilitySummary(record: StoredConnectorRecord): ConnectorCapabilitySummary {
    const catalog = getConnectorCatalogItem(record.connectorId);
    const hints = catalog?.capabilityHints ?? [];
    const tools = [...new Set(hints.map(sanitizeHintToToolName))];
    const inventory = applyToolPolicyOverrides(
      tools.map((toolName): ConnectorToolInventoryItem => ({
        name: toolName,
        rawName: toolName,
        classification: "read",
        defaultEnabled: classificationIsEnabledByDefault("read"),
        enabled: classificationIsEnabledByDefault("read"),
        reason: "Catalog capability hint is read-only.",
        source: "catalog_hint",
      })),
      record.toolPolicyOverrides,
    );
    return {
      tools,
      allowedTools: inventory.filter((tool) => tool.enabled).map((tool) => tool.name),
      blockedTools: inventory.filter((tool) => !tool.enabled).map((tool) => tool.name),
      toolInventory: inventory,
      resourceToolNames: [],
      promptNames: [],
      allowedPrompts: [],
      blockedPrompts: [],
      resourceCount: 0,
      promptCount: 0,
      inventoryHash: makeInventoryHash(tools),
    };
  }

  private upsertRecord(record: StoredConnectorRecord): void {
    const index = this.state.connectors.findIndex((entry) => entry.id === record.id);
    if (index >= 0) {
      this.state.connectors[index] = record;
    } else {
      this.state.connectors.push(record);
    }
    this.state.updatedAt = nowIso();
  }

  private upsertScopeOverride(override: {
    connectorId: string;
    scopeTarget: Exclude<ConnectorScopeTarget, "global">;
    scopeKey: string;
    enabled: boolean;
  }): void {
    const next: StoredScopeOverride = {
      connectorId: override.connectorId,
      scopeTarget: override.scopeTarget,
      scopeKey: override.scopeKey,
      enabled: override.enabled,
      updatedAt: nowIso(),
    };
    const index = this.state.scopeOverrides.findIndex((entry) =>
      entry.connectorId === next.connectorId &&
      entry.scopeTarget === next.scopeTarget &&
      entry.scopeKey === next.scopeKey,
    );
    if (index >= 0) {
      this.state.scopeOverrides[index] = next;
    } else {
      this.state.scopeOverrides.push(next);
    }
    this.state.updatedAt = nowIso();
  }

  private appendLog(entry: {
    connectorId?: string | undefined;
    connectorName?: string | undefined;
    kind: ConnectorLogEntry["kind"];
    level: ConnectorLogEntry["level"];
    message: string;
    scopeTarget?: ConnectorScopeTarget | undefined;
    healthState?: ConnectorHealthState | undefined;
  }): void {
    const log: ConnectorLogEntry = {
      id: createRandomId("log-"),
      connectorId: entry.connectorId,
      connectorName: entry.connectorName,
      kind: entry.kind,
      channel: this.state.auditPreference.enabled ? "audit" : "diagnostic",
      level: entry.level,
      message: entry.message,
      timestamp: nowIso(),
      scopeTarget: entry.scopeTarget,
      healthState: entry.healthState,
      redacted: true,
    };
    this.state.logs.unshift(log);
    if (this.state.logs.length > MAX_CONNECTOR_LOG_ENTRIES) {
      this.state.logs = this.state.logs.slice(0, MAX_CONNECTOR_LOG_ENTRIES);
    }
    this.state.updatedAt = nowIso();
  }

  private detectImportConflicts(bundle: ConnectorExportBundle): ConnectorConflict[] {
    const conflicts: ConnectorConflict[] = [];
    for (const [index, item] of bundle.connectors.entries()) {
      const key = this.importConflictKey(item, index);
      const existing = this.findImportConflictTarget(item);
      if (!existing) continue;

      let kind: ConnectorConflict["kind"] = "import_collision";
      if (item.source === "library" && existing.source === "library" && existing.connectorId === item.connectorId) {
        kind = "duplicate_connector";
      } else if (item.transport === "remote_http" && normalizeUrl(item.url) && normalizeUrl(item.url) === normalizeUrl(existing.url)) {
        kind = "duplicate_remote_url";
      } else if (item.transport === "local_stdio" && customCommandSignature(item) === customCommandSignature(existing)) {
        kind = "duplicate_local_runtime";
      }

      conflicts.push({
        key,
        kind,
        existingConnectorId: existing.id,
        existingName: existing.name,
        incomingConnectorId: trimString(item.storedId) ?? item.connectorId,
        title: "Connector import conflict",
        message: `Incoming connector conflicts with existing connector "${existing.name}".`,
        resolution: kind === "duplicate_connector" ? "replace" : "skip",
      });
    }
    return conflicts;
  }

  private importConflictKey(item: ConnectorExportItem, index: number): string {
    return `import-${index}-${trimString(item.storedId) ?? item.connectorId}`;
  }

  private findImportConflictTarget(item: ConnectorExportItem): StoredConnectorRecord | undefined {
    const byStoredId = trimString(item.storedId)
      ? this.state.connectors.find((record) => record.id === item.storedId)
      : undefined;
    if (byStoredId) return byStoredId;
    if (item.source === "library") {
      return this.state.connectors.find((record) => record.source === "library" && record.connectorId === item.connectorId);
    }
    if (item.transport === "remote_http") {
      const normalized = normalizeUrl(item.url);
      if (!normalized) return undefined;
      return this.state.connectors.find((record) => record.source === "custom" && normalizeUrl(record.url) === normalized);
    }
    return this.state.connectors.find((record) =>
      record.source === "custom" &&
      record.transport === "local_stdio" &&
      customCommandSignature(record) === customCommandSignature(item),
    );
  }

  private async load(): Promise<void> {
    try {
      const raw = localStorage.getItem(CONNECTOR_STORAGE_KEY);
      if (!raw) {
        this.state = createConnectorState();
        return;
      }

      const parsed = JSON.parse(raw) as unknown;
      if (this.isPlainState(parsed)) {
        this.state = this.normalizeState(parsed);
        await this.persist();
        return;
      }

      if (!isEncryptedConnectorEnvelope(parsed)) {
        this.state = createConnectorState();
        return;
      }

      this.state = this.normalizeState(await this.decryptState(parsed));
    } catch {
      this.state = createConnectorState();
    }
  }

  private async persist(): Promise<void> {
    try {
      const envelope = await this.encryptState(this.state);
      localStorage.setItem(CONNECTOR_STORAGE_KEY, JSON.stringify(envelope));
    } catch {
      // Ignore storage errors in browser sandbox.
    }
  }

  private isPlainState(value: unknown): value is StoredConnectorState {
    if (!value || typeof value !== "object") return false;
    const parsed = value as Partial<StoredConnectorState>;
    return (
      parsed.version === CONNECTOR_STORAGE_VERSION &&
      Array.isArray(parsed.connectors) &&
      Array.isArray(parsed.scopeOverrides) &&
      (parsed.oauthFlows === undefined || Array.isArray(parsed.oauthFlows)) &&
      Array.isArray(parsed.logs) &&
      !!parsed.auditPreference &&
      typeof parsed.auditPreference.enabled === "boolean"
    );
  }

  private normalizeState(state: StoredConnectorState): StoredConnectorState {
    return {
      ...state,
      connectors: state.connectors.map((record) => {
        const normalized = {
          ...record,
          toolPolicyOverrides: normalizeToolPolicyOverrides(record.toolPolicyOverrides),
          suppressNonReadToolWarning: record.suppressNonReadToolWarning === true,
        };
        if (record.credentialSource !== "oauth") {
          return {
            ...normalized,
            oauthConnected: undefined,
            oauthExpiresAt: undefined,
            oauthLastAuthAt: undefined,
            oauthLastAuthError: undefined,
            oauthRefreshToken: undefined,
            oauthTokenType: undefined,
            oauthScope: undefined,
            oauthClientId: undefined,
            oauthTokenEndpoint: undefined,
          };
        }

        const hasCredential = Boolean(trimString(record.secret));
        const connected = hasCredential && record.oauthConnected === true;
        return {
          ...normalized,
          secret: hasCredential ? record.secret : undefined,
          oauthConnected: connected,
          oauthExpiresAt: connected ? record.oauthExpiresAt : undefined,
          oauthLastAuthAt: connected ? record.oauthLastAuthAt : undefined,
          oauthRefreshToken: connected ? record.oauthRefreshToken : undefined,
          oauthTokenType: connected ? record.oauthTokenType : undefined,
          oauthScope: connected ? record.oauthScope : undefined,
          oauthClientId: connected ? record.oauthClientId : undefined,
          oauthTokenEndpoint: connected ? record.oauthTokenEndpoint : undefined,
          oauthLastAuthError: connected
            ? record.oauthLastAuthError
            : (record.oauthLastAuthError ?? "OAuth credential missing. Sign in again before use."),
          lastError: connected ? record.lastError : (record.lastError ?? "OAuth credential missing. Sign in again before use."),
        };
      }),
      oauthFlows: (state.oauthFlows ?? []).filter((entry) =>
        Boolean(entry.connectorId) &&
        Boolean(entry.state) &&
        Boolean(entry.expiresAt),
      ),
      logs: state.logs.slice(0, MAX_CONNECTOR_LOG_ENTRIES),
      auditPreference: { enabled: state.auditPreference.enabled !== false },
    };
  }

  private dropExpiredOAuthFlows(): void {
    const now = Date.now();
    const expired = this.state.oauthFlows.filter((entry) => Date.parse(entry.expiresAt) <= now);
    this.state.oauthFlows = this.state.oauthFlows.filter((entry) => Date.parse(entry.expiresAt) > now);
    if (!expired.length) {
      return;
    }
    for (const flow of expired) {
      const record = this.state.connectors.find((entry) => entry.id === flow.connectorId);
      if (!record || record.authMethod !== "oauth" || record.oauthConnected === true) {
        continue;
      }
      record.oauthLastAuthError = "OAuth sign-in timed out. Start sign-in again.";
      record.lastError = record.oauthLastAuthError;
      record.updatedAt = nowIso();
      this.upsertRecord(record);
    }
  }

  private async getOrCreateCryptoKey(): Promise<CryptoKey> {
    let keyBase64 = localStorage.getItem(CONNECTOR_CRYPTO_KEY_STORAGE_KEY);
    if (!keyBase64) {
      const keyBytes = new Uint8Array(32);
      crypto.getRandomValues(keyBytes);
      keyBase64 = arrayBufferToBase64(keyBytes.buffer);
      localStorage.setItem(CONNECTOR_CRYPTO_KEY_STORAGE_KEY, keyBase64);
    }

    const rawKey = base64ToUint8Array(keyBase64);
    return crypto.subtle.importKey("raw", toArrayBuffer(rawKey), { name: "AES-GCM" }, false, ["encrypt", "decrypt"]);
  }

  private async encryptState(state: StoredConnectorState): Promise<EncryptedConnectorEnvelope> {
    const key = await this.getOrCreateCryptoKey();
    const iv = new Uint8Array(12);
    crypto.getRandomValues(iv);
    const plaintext = new TextEncoder().encode(JSON.stringify(state));
    const ciphertext = await crypto.subtle.encrypt({ name: "AES-GCM", iv: toArrayBuffer(iv) }, key, toArrayBuffer(plaintext));
    return {
      version: CONNECTOR_STORAGE_VERSION,
      algorithm: "AES-GCM",
      iv: arrayBufferToBase64(iv.buffer),
      ciphertext: arrayBufferToBase64(ciphertext),
    };
  }

  private async decryptState(envelope: EncryptedConnectorEnvelope): Promise<StoredConnectorState> {
    const key = await this.getOrCreateCryptoKey();
    const iv = base64ToUint8Array(envelope.iv);
    const ciphertext = base64ToUint8Array(envelope.ciphertext);
    const plaintext = await crypto.subtle.decrypt(
      { name: "AES-GCM", iv: toArrayBuffer(iv) },
      key,
      toArrayBuffer(ciphertext),
    );
    const decoded = new TextDecoder().decode(plaintext);
    const parsed = JSON.parse(decoded) as unknown;
    return this.isPlainState(parsed) ? this.normalizeState(parsed) : createConnectorState();
  }
}
