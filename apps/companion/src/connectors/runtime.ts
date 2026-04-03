import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { createRequire } from "node:module";
import { spawnSync } from "node:child_process";
import { randomUUID } from "node:crypto";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import { SSEClientTransport } from "@modelcontextprotocol/sdk/client/sse.js";
import type {
  ConnectorAuditPreference,
  ConnectorCatalogItem,
  ConnectorConflict,
  ConnectorDiagnostic,
  ConnectorDiagnosticsResponse,
  ConnectorEnvSuggestion,
  ConnectorExportBundle,
  ConnectorFavoriteRequest,
  ConnectorHealthState,
  ConnectorImportApplyRequest,
  ConnectorImportApplyResponse,
  ConnectorImportPreviewRequest,
  ConnectorImportPreviewResponse,
  ConnectorLogEntry,
  ConnectorLogResponse,
  ConnectorPrepareResponse,
  ConnectorRuntimeCheck,
  ConnectorScopeContext,
  ConnectorScopeTarget,
  ConnectorScopeUpdateRequest,
  ConnectorSetupRequest,
  ConnectorSetupResponse,
  ConnectorStatus,
  ConnectorStatusResponse,
  ConnectorTestResponse,
  ConnectorTransport,
  ConnectorVerificationSnapshot,
} from "@pi-office/pi-office-pack";
import type { CompanionConfig } from "../config.js";
import { getConnectorCatalogItem, listConnectorCatalog } from "./catalog.js";
import { ConnectorSecretStore } from "./secret-store.js";
import {
  computeConnectorFingerprint,
  ConnectorStore,
  resolveDocumentScopeKey,
  resolveWorkspaceScopeKey,
  type StoredConnectorDefinition,
  type StoredConnectorDraft,
  type StoredConnectorLogEntry,
  type StoredConnectorVerification,
} from "./store.js";

const require = createRequire(import.meta.url);

type ProbeTransport = StdioClientTransport | StreamableHTTPClientTransport | SSEClientTransport;
type ProbeTool = { name: string; description?: string | undefined; inputSchema?: unknown };
type ProbeResource = { name: string; uri: string; description?: string | undefined };

interface ResolvedRuntimeDefinition {
  transport: ConnectorTransport;
  command?: string | undefined;
  args?: string[] | undefined;
  env?: Record<string, string> | undefined;
  cwd?: string | undefined;
  url?: string | undefined;
  headers?: Record<string, string> | undefined;
  bearerToken?: string | undefined;
  bearerTokenEnv?: string | undefined;
}

export interface ConnectorGuardServerPolicy {
  serverName: string;
  connectorId: string;
  connectorName: string;
  healthState: ConnectorHealthState;
  allowedTools: string[];
  blockedTools: string[];
  allowedPrompts: string[];
  blockedPrompts: string[];
}

export interface ConnectorGuardSnapshot {
  scopeContext?: ConnectorScopeContext | undefined;
  servers: ConnectorGuardServerPolicy[];
}

interface PreparedConnectorRuntime {
  sessionId: string;
  configPath: string;
  guardSnapshot: ConnectorGuardSnapshot;
  statuses: ConnectorStatus[];
}

interface VerificationProbeResult {
  ok: boolean;
  snapshot: ConnectorVerificationSnapshot;
  error?: string | undefined;
}

function nowIso(): string {
  return new Date().toISOString();
}

function unique<T>(values: T[]): T[] {
  return [...new Set(values)];
}

function trimString(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed ? trimmed : undefined;
}

function normalizeUrl(value: string | undefined): string | undefined {
  const trimmed = trimString(value);
  if (!trimmed) {
    return undefined;
  }
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

function sanitizeServerName(input: string): string {
  return input
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "") || "connector";
}

function isCustomConnectorId(value: string): boolean {
  return value === "custom" || value.startsWith("custom-");
}

function formatToolName(serverName: string, toolName: string): string {
  return `${serverName.replace(/-/g, "_")}_${toolName}`;
}

function resourceNameToToolName(name: string): string {
  let result = name
    .replace(/[^a-zA-Z0-9]/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+/, "")
    .replace(/_+$/, "")
    .toLowerCase();
  if (!result || /^\d/.test(result)) {
    result = `resource${result ? `_${result}` : ""}`;
  }
  return result;
}

function maskValue(value: string): string {
  if (!value) return "";
  if (value.length <= 8) {
    return `${value.slice(0, 2)}***${value.slice(-1)}`;
  }
  return `${value.slice(0, 4)}***${value.slice(-4)}`;
}

function firstLine(value: string | undefined): string {
  return value?.split(/\r?\n/, 1)[0]?.trim() || "";
}

function withTimeout<T>(promise: Promise<T>, timeoutMs: number, message: string): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(message)), timeoutMs);
    promise.then(
      (value) => {
        clearTimeout(timer);
        resolve(value);
      },
      (error) => {
        clearTimeout(timer);
        reject(error);
      },
    );
  });
}

function discoverCommand(binary: string): ConnectorRuntimeCheck {
  const locator = process.platform === "win32" ? "where.exe" : "which";
  const found = spawnSync(locator, [binary], {
    encoding: "utf8",
    windowsHide: true,
  });
  if (found.status !== 0) {
    return {
      key: binary as ConnectorRuntimeCheck["key"],
      label: binary,
      ok: false,
      detail: `${binary} was not found on this PC.`,
    };
  }

  const version = spawnSync(binary, ["--version"], {
    encoding: "utf8",
    windowsHide: true,
  });
  const pathLine = firstLine(found.stdout);
  const versionLine = firstLine(version.stdout || version.stderr);
  return {
    key: binary as ConnectorRuntimeCheck["key"],
    label: binary,
    ok: true,
    detail: versionLine ? `${versionLine} (${pathLine})` : pathLine,
  };
}

function matchesAnyPattern(value: string, patterns: string[]): boolean {
  return patterns.some((pattern) => new RegExp(pattern, "i").test(value));
}

function redactText(value: string | undefined): string | undefined {
  if (!value) {
    return value;
  }
  return value
    .replace(/[A-Za-z0-9_-]{20,}/g, "[redacted]")
    .replace(/https?:\/\/[^\s]+/g, "[redacted-url]");
}

function getAllEnvSuggestions(catalog: ConnectorCatalogItem[]): ConnectorEnvSuggestion[] {
  const seen = new Set<string>();
  const suggestions: ConnectorEnvSuggestion[] = [];
  for (const item of catalog) {
    for (const hint of item.envHints) {
      if (seen.has(hint.key)) continue;
      seen.add(hint.key);
      const value = process.env[hint.key];
      suggestions.push({
        key: hint.key,
        present: Boolean(value),
        maskedValue: value ? maskValue(value) : undefined,
      });
    }
  }
  return suggestions.sort((left, right) => left.key.localeCompare(right.key));
}

function buildCustomCatalogItem(name = "Custom MCP"): ConnectorCatalogItem {
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
        "(^|_)(search|list|get|read|find|fetch|query|describe|inspect|preview|lookup|retrieve)(_|$)",
      ],
      blockToolPatterns: [
        "^(create|update|delete|remove|write|edit|insert|append|upload|post|put|patch|merge|commit|apply|send|message|reply|comment|close|reopen|archive|move|rename|grant|revoke|start|stop|restart|approve|reject|refund|charge|pay)",
        "(^|_)(create|update|delete|remove|write|edit|insert|append|upload|post|put|patch|merge|commit|apply|send|message|reply|comment|close|reopen|archive|move|rename|grant|revoke|start|stop|restart|approve|reject|refund|charge|pay)(_|$)",
      ],
      allowPromptPatterns: [],
      blockPromptPatterns: [".*"],
    },
    setupNotes: [
      "The companion keeps custom connectors in read-only mode.",
      "Only verified read-safe tools and resources are exposed to the model.",
    ],
    recommendedHosts: ["word", "excel", "powerpoint"],
    setupDifficulty: "advanced",
    catalogRevision: "custom-v1",
  };
}

async function createRemoteTransport(definition: ResolvedRuntimeDefinition): Promise<ProbeTransport> {
  const headers = { ...(definition.headers ?? {}) };
  if (definition.bearerToken) {
    headers.Authorization = `Bearer ${definition.bearerToken}`;
  } else if (definition.bearerTokenEnv) {
    const envValue = process.env[definition.bearerTokenEnv];
    if (envValue) {
      headers.Authorization = `Bearer ${envValue}`;
    }
  }

  const requestInit = Object.keys(headers).length ? ({ headers } satisfies RequestInit) : undefined;
  const url = new URL(definition.url!);
  const streamOptions = requestInit ? { requestInit } : {};
  const probe = new StreamableHTTPClientTransport(url, streamOptions);
  try {
    const client = new Client({ name: "pi-office-mcp-probe", version: "1.0.0" });
    await client.connect(probe as never);
    await client.close().catch(() => {});
    await probe.close().catch(() => {});
    return new StreamableHTTPClientTransport(url, streamOptions);
  } catch {
    await probe.close().catch(() => {});
    const sseOptions = requestInit ? { requestInit } : {};
    return new SSEClientTransport(url, sseOptions);
  }
}

async function fetchAllTools(client: Client): Promise<ProbeTool[]> {
  const tools: ProbeTool[] = [];
  let cursor: string | undefined;
  do {
    const result = await client.listTools(cursor ? { cursor } : undefined);
    tools.push(...(result.tools ?? []).map((tool) => ({
      name: tool.name,
      description: tool.description,
      inputSchema: tool.inputSchema,
    })));
    cursor = result.nextCursor;
  } while (cursor);
  return tools;
}

async function fetchAllResources(client: Client): Promise<ProbeResource[]> {
  try {
    const resources: ProbeResource[] = [];
    let cursor: string | undefined;
    do {
      const result = await client.listResources(cursor ? { cursor } : undefined);
      resources.push(...(result.resources ?? []).map((resource) => ({
        name: resource.name,
        uri: resource.uri,
        description: resource.description,
      })));
      cursor = result.nextCursor;
    } while (cursor);
    return resources;
  } catch {
    return [];
  }
}

async function fetchAllPrompts(client: Client): Promise<Array<{ name: string }>> {
  const listPrompts = (client as unknown as {
    listPrompts?: (input?: { cursor?: string }) => Promise<{ prompts?: Array<{ name: string }>; nextCursor?: string }>;
  }).listPrompts;
  if (!listPrompts) {
    return [];
  }

  try {
    const prompts: Array<{ name: string }> = [];
    let cursor: string | undefined;
    do {
      const result = await listPrompts.call(client, cursor ? { cursor } : undefined);
      prompts.push(...(result.prompts ?? []));
      cursor = result.nextCursor;
    } while (cursor);
    return prompts;
  } catch {
    return [];
  }
}

export class ConnectorService {
  private readonly secretStore: ConnectorSecretStore;
  private readonly store: ConnectorStore;
  private readonly preparedRuntimes = new Map<string, PreparedConnectorRuntime>();

  constructor(private readonly config: CompanionConfig) {
    this.secretStore = new ConnectorSecretStore(config.connectorsSecretsDir, config.connectorsVaultPath);
    this.store = new ConnectorStore(config.connectorsStatePath, this.secretStore);
  }

  listCatalog(): ConnectorCatalogItem[] {
    return listConnectorCatalog();
  }

  getAdapterExtensionPath(): string {
    return require.resolve("pi-mcp-adapter/index.ts");
  }

  listStatuses(scopeContext?: ConnectorScopeContext): ConnectorStatus[] {
    return this.store.listDefinitions().map((definition) => this.toStatus(definition, scopeContext));
  }

  getStatusResponse(scopeContext?: ConnectorScopeContext): ConnectorStatusResponse {
    return {
      connectors: this.listStatuses(scopeContext),
    };
  }

  getDiagnostics(): ConnectorDiagnosticsResponse {
    const catalog = this.listCatalog();
    const runtimes = this.checkRuntimes();
    const diagnostics: ConnectorDiagnostic[] = [];

    for (const check of runtimes) {
      if (!check.ok && (check.key === "node" || check.key === "npx")) {
        diagnostics.push({
          level: "warning",
          code: `${check.key}_missing`,
          title: `${check.label} not detected`,
          message: `${check.label} is needed for local MCP connectors.`,
        });
      }
    }

    return {
      generatedAt: nowIso(),
      runtimes,
      envSuggestions: getAllEnvSuggestions(catalog),
      diagnostics,
    };
  }

  prepare(connectorId: string, scopeContext?: ConnectorScopeContext): ConnectorPrepareResponse {
    const customExisting = isCustomConnectorId(connectorId) ? this.store.getDefinition(connectorId) : undefined;
    const connector =
      connectorId === "custom"
        ? buildCustomCatalogItem()
        : customExisting
          ? buildCustomCatalogItem(customExisting.name)
          : getConnectorCatalogItem(connectorId);
    if (!connector) {
      throw new Error(`Unknown connector: ${connectorId}`);
    }

    const existing = customExisting ?? (connectorId === "custom" ? undefined : this.store.getByLibraryConnectorId(connectorId));
    const runtimes = this.checkRuntimes(connector.requirements.map((item) => item.key));
    const draft = existing ? this.buildDraftFromStored(existing) : this.buildDraftFromRequest({ connectorId }, connector, existing);
    const diagnostics = this.validateDraft(draft, connector, false);
    const conflicts = existing ? [] : this.store.findConflicts({
      connectorId: draft.connectorId,
      source: draft.source,
      name: draft.name,
      transport: draft.transport,
      url: draft.url,
      command: draft.command,
      args: draft.args,
      cwd: draft.cwd,
    });

    return {
      connector,
      existing: existing ? this.toStatus(existing, scopeContext) : undefined,
      draft: this.toSetupRequest(draft, scopeContext),
      runtimes,
      envSuggestions: connector.envHints.map((hint) => ({
        key: hint.key,
        present: Boolean(process.env[hint.key]),
        maskedValue: process.env[hint.key] ? maskValue(process.env[hint.key]!) : undefined,
      })),
      diagnostics,
      conflicts,
      auditPreference: this.store.getAuditPreference(),
    };
  }

  async connect(request: ConnectorSetupRequest): Promise<ConnectorSetupResponse> {
    const { draft, connector, diagnostics, conflicts } = this.prepareDraft(request, true);
    if (diagnostics.some((item) => item.level === "error") || (conflicts.length > 0 && !request.replaceExisting)) {
      return {
        ok: true,
        status: this.toStatus({ ...draft, id: request.existingId ?? draft.id ?? draft.connectorId }, request.scopeContext, conflicts),
        diagnostics: diagnostics.concat(conflicts.map((item) => ({
          level: "error" as const,
          code: "conflict",
          title: item.title,
          message: item.message,
          connectorId: draft.connectorId,
        }))),
        conflicts,
      };
    }

    for (const conflict of conflicts) {
      if (request.replaceExisting && conflict.resolution === "replace") {
        this.store.removeDefinition(conflict.existingConnectorId);
      }
    }

    const persistedDraft = this.persistDraftSecret(draft, request);
    const saved = this.store.upsertDefinition(persistedDraft);
    if (typeof request.favorite === "boolean") {
      this.store.setFavorite(saved.id, request.favorite);
    }
    this.store.updateScope(saved.id, request.scopeTarget ?? "global", request.scopeContext, request.enabled ?? persistedDraft.defaultEnabled);

    const probe = await this.probeDefinition(saved, connector);
    this.store.setVerificationSnapshot(saved.id, probe.snapshot);
    this.store.updateLastHealth(saved.id, {
      lastHealthyAt: probe.ok ? probe.snapshot.verifiedAt : saved.lastHealthyAt,
      lastError: probe.error,
    });
    this.appendLog({
      connectorId: saved.id,
      connectorName: saved.name,
      kind: "setup",
      channel: "diagnostic",
      level: probe.ok ? "info" : "warning",
      message: probe.ok ? "Connector saved and verified." : redactText(probe.error) ?? "Connector saved but verification failed.",
      healthState: probe.ok ? "ready" : this.inferHealthState(saved, connector, probe.snapshot),
      redacted: true,
    });

    return {
      ok: true,
      status: this.toStatus(this.store.getDefinition(saved.id) ?? saved, request.scopeContext),
      diagnostics: this.buildStatusDiagnostics(saved, connector),
      conflicts: [],
    };
  }

  async test(request: ConnectorSetupRequest): Promise<ConnectorTestResponse> {
    const { draft, connector, diagnostics, conflicts } = this.prepareDraft(request, true);
    if (diagnostics.some((item) => item.level === "error") || (conflicts.length > 0 && !request.replaceExisting)) {
      return {
        ok: false,
        status: this.toStatus({ ...draft, id: request.existingId ?? draft.id ?? draft.connectorId }, request.scopeContext, conflicts),
        diagnostics,
        conflicts,
      };
    }

    const transient = this.persistDraftSecret(draft, request, false);
    const probe = await this.probeDefinition(transient, connector);
    if (request.secret && transient.secretRef) {
      this.secretStore.remove(transient.secretRef);
    }
    const status = this.toStatus(
      {
        ...transient,
        id: transient.id ?? transient.connectorId,
        lastHealthyAt: probe.ok ? probe.snapshot.verifiedAt : transient.lastHealthyAt,
        lastError: probe.error,
      } as StoredConnectorDefinition,
      request.scopeContext,
      conflicts,
      probe.snapshot,
    );
    this.appendLog({
      connectorId: transient.id,
      connectorName: transient.name,
      kind: "test",
      channel: "diagnostic",
      level: probe.ok ? "info" : "warning",
      message: probe.ok ? "Verification completed." : redactText(probe.error) ?? "Verification failed.",
      healthState: status.healthState,
      redacted: true,
    });

    return {
      ok: probe.ok,
      status,
      diagnostics: this.buildStatusDiagnostics(transient as StoredConnectorDefinition, connector, probe.snapshot),
      conflicts,
    };
  }

  async validateCustom(request: ConnectorSetupRequest): Promise<ConnectorTestResponse> {
    return this.test({ ...request, connectorId: isCustomConnectorId(request.connectorId) ? request.connectorId : "custom" });
  }

  async reverify(connectorId: string, scopeContext?: ConnectorScopeContext): Promise<ConnectorTestResponse> {
    const definition = this.store.getDefinition(connectorId);
    if (!definition) {
      throw new Error(`Unknown connector: ${connectorId}`);
    }
    const connector = definition.source === "custom" ? buildCustomCatalogItem(definition.name) : getConnectorCatalogItem(definition.connectorId);
    if (!connector) {
      throw new Error(`Unknown connector catalog entry: ${definition.connectorId}`);
    }

    const probe = await this.probeDefinition(definition, connector);
    this.store.setVerificationSnapshot(definition.id, probe.snapshot);
    this.store.updateLastHealth(definition.id, {
      lastHealthyAt: probe.ok ? probe.snapshot.verifiedAt : definition.lastHealthyAt,
      lastError: probe.error,
    });
    this.appendLog({
      connectorId: definition.id,
      connectorName: definition.name,
      kind: "reverify",
      channel: "diagnostic",
      level: probe.ok ? "info" : "warning",
      message: probe.ok ? "Connector re-verified." : redactText(probe.error) ?? "Connector re-verification failed.",
      healthState: probe.ok ? "ready" : this.inferHealthState(definition, connector, probe.snapshot),
      redacted: true,
    });

    return {
      ok: probe.ok,
      status: this.toStatus(this.store.getDefinition(definition.id) ?? definition, scopeContext, [], probe.snapshot),
      diagnostics: this.buildStatusDiagnostics(definition, connector, probe.snapshot),
      conflicts: [],
    };
  }

  remove(id: string): boolean {
    const existing = this.store.getDefinition(id);
    const removed = this.store.removeDefinition(id);
    if (removed && existing) {
      this.appendLog({
        connectorId: id,
        connectorName: existing.name,
        kind: "remove",
        channel: "diagnostic",
        level: "info",
        message: "Connector removed.",
        redacted: true,
      });
    }
    return removed;
  }

  setFavorite(request: ConnectorFavoriteRequest): ConnectorStatus | undefined {
    this.store.setFavorite(request.connectorId, request.favorite);
    const definition = this.store.getDefinition(request.connectorId);
    if (!definition) {
      return undefined;
    }
    this.appendLog({
      connectorId: definition.id,
      connectorName: definition.name,
      kind: "favorite",
      channel: "diagnostic",
      level: "info",
      message: request.favorite ? "Connector pinned." : "Connector unpinned.",
      redacted: true,
    });
    return this.toStatus(definition);
  }

  updateScope(request: ConnectorScopeUpdateRequest): ConnectorStatus | undefined {
    const definition = this.store.getDefinition(request.connectorId);
    if (!definition) {
      return undefined;
    }
    this.store.updateScope(request.connectorId, request.scopeTarget, request.scopeContext, request.enabled);
    this.appendLog({
      connectorId: definition.id,
      connectorName: definition.name,
      kind: "scope",
      channel: "diagnostic",
      level: "info",
      message: `${request.scopeTarget} scope updated.`,
      scopeTarget: request.scopeTarget,
      redacted: true,
    });
    return this.toStatus(definition, request.scopeContext);
  }

  getLogs(connectorId: string): ConnectorLogResponse {
    const definition = this.store.getDefinition(connectorId);
    return {
      connectorId,
      auditPreference: this.store.getAuditPreference(),
      logs: this.store.listLogs(connectorId).map((entry) => this.toLogEntry(entry, definition?.name)),
    };
  }

  getAuditPreference(): ConnectorAuditPreference {
    return this.store.getAuditPreference();
  }

  setAuditPreference(preference: ConnectorAuditPreference): ConnectorAuditPreference {
    return this.store.setAuditPreference(preference);
  }

  exportBundle(): ConnectorExportBundle {
    this.appendLog({
      kind: "export",
      channel: "diagnostic",
      level: "info",
      message: "Connector export bundle created.",
      redacted: true,
    });
    return this.store.buildExportBundle();
  }

  previewImport(request: ConnectorImportPreviewRequest): ConnectorImportPreviewResponse {
    return this.store.previewImport(request.bundle);
  }

  applyImport(request: ConnectorImportApplyRequest): ConnectorImportApplyResponse {
    const result = this.store.applyImport(request);
    this.appendLog({
      kind: "import",
      channel: "diagnostic",
      level: "info",
      message: `Imported ${result.importedConnectorIds.length} connector${result.importedConnectorIds.length === 1 ? "" : "s"}.`,
      redacted: true,
    });
    return result;
  }

  async prepareSessionRuntime(sessionId: string, scopeContext: ConnectorScopeContext): Promise<string> {
    const activeDefinitions = this.store.listDefinitions().filter((definition) => {
      const scopeState = this.store.resolveScope(definition, scopeContext).find((entry) => entry.applies);
      return scopeState?.enabled === true;
    });

    for (const definition of activeDefinitions) {
      const connector = definition.source === "custom" ? buildCustomCatalogItem(definition.name) : getConnectorCatalogItem(definition.connectorId);
      if (!connector || !this.isConfigured(definition, connector)) {
        continue;
      }
      const snapshot = this.store.getVerificationSnapshot(definition.id);
      const staleReason = this.getStaleReason(definition, connector, snapshot);
      if (staleReason) {
        this.store.markVerificationStale(definition.id, staleReason);
        await this.reverify(definition.id, scopeContext).catch(() => undefined);
      }
    }

    const statuses = this.listStatuses(scopeContext);
    const configPath = this.writeSessionConfig(sessionId, statuses);
    const guardSnapshot: ConnectorGuardSnapshot = {
      scopeContext,
      servers: statuses
        .filter((status) => status.healthState === "ready" && status.capabilities)
        .map((status) => ({
          serverName: sanitizeServerName(status.id),
          connectorId: status.id,
          connectorName: status.name,
          healthState: status.healthState,
          allowedTools: status.capabilities?.allowedTools ?? [],
          blockedTools: status.capabilities?.blockedTools ?? [],
          allowedPrompts: status.capabilities?.allowedPrompts ?? [],
          blockedPrompts: status.capabilities?.blockedPrompts ?? [],
        })),
    };

    this.preparedRuntimes.set(sessionId, {
      sessionId,
      configPath,
      guardSnapshot,
      statuses,
    });

    return configPath;
  }

  getGuardSnapshotForSession(sessionId: string): ConnectorGuardSnapshot {
    return this.preparedRuntimes.get(sessionId)?.guardSnapshot ?? { servers: [] };
  }

  clearPreparedSession(sessionId: string): void {
    this.preparedRuntimes.delete(sessionId);
  }

  recordSessionToolUse(sessionId: string, toolName: string, blocked: boolean): void {
    if (!this.store.getAuditPreference().enabled) {
      return;
    }
    const runtime = this.preparedRuntimes.get(sessionId);
    const connector = runtime?.statuses.find((status) => status.capabilities?.allowedTools.includes(toolName) || status.capabilities?.blockedTools.includes(toolName));
    this.appendLog({
      connectorId: connector?.id,
      connectorName: connector?.name,
      kind: "use",
      channel: "audit",
      level: blocked ? "warning" : "info",
      message: blocked ? `Blocked ${toolName}.` : `Used ${toolName}.`,
      healthState: connector?.healthState,
      redacted: true,
    });
  }

  private toLogEntry(entry: StoredConnectorLogEntry, fallbackName?: string | undefined): ConnectorLogEntry {
    return {
      ...entry,
      connectorName: entry.connectorName ?? fallbackName,
    };
  }

  private toSetupRequest(draft: StoredConnectorDraft, scopeContext?: ConnectorScopeContext): ConnectorSetupRequest {
    return {
      connectorId: draft.source === "custom" ? draft.id ?? draft.connectorId : draft.connectorId,
      existingId: draft.id,
      name: draft.name,
      enabled: draft.defaultEnabled,
      setupKind: draft.setupKind,
      authMethod: draft.authMethod,
      transport: draft.transport,
      credentialSource: draft.credentialSource,
      secretEnvKey: draft.secretEnvKey,
      useDetectedEnvKey: draft.useDetectedEnvKey,
      url: draft.url,
      command: draft.command,
      args: draft.args,
      cwd: draft.cwd,
      env: draft.env,
      scopeContext,
    };
  }

  private buildDraftFromStored(stored: StoredConnectorDefinition): StoredConnectorDraft {
    return {
      id: stored.id,
      connectorId: stored.connectorId,
      name: stored.name,
      source: stored.source,
      category: stored.category,
      maturity: stored.maturity,
      setupKind: stored.setupKind,
      authMethod: stored.authMethod,
      transport: stored.transport,
      credentialSource: stored.credentialSource,
      secretRef: stored.secretRef,
      secretEnvKey: stored.secretEnvKey,
      useDetectedEnvKey: stored.useDetectedEnvKey,
      url: stored.url,
      command: stored.command,
      args: stored.args,
      cwd: stored.cwd,
      env: stored.env,
      defaultEnabled: stored.defaultEnabled,
      lastHealthyAt: stored.lastHealthyAt,
      lastError: stored.lastError,
      oauthExpiresAt: stored.oauthExpiresAt,
    };
  }

  private buildDraftFromRequest(
    request: ConnectorSetupRequest,
    connector: ConnectorCatalogItem | undefined,
    existing: StoredConnectorDefinition | undefined,
  ): StoredConnectorDraft {
    const custom = isCustomConnectorId(request.connectorId);
    return {
      id: existing?.id ?? (custom && request.connectorId !== "custom" ? request.connectorId : request.existingId),
      connectorId: custom ? "custom" : request.connectorId,
      name: trimString(request.name) ?? existing?.name ?? connector?.name ?? "Custom Connector",
      source: custom ? "custom" : "library",
      category: existing?.category ?? connector?.category ?? "knowledge",
      maturity: existing?.maturity ?? connector?.maturity ?? "custom_mcp_only",
      setupKind: request.setupKind ?? existing?.setupKind ?? connector?.setupKind ?? "local_executable_or_docker",
      authMethod: request.authMethod ?? existing?.authMethod ?? connector?.authMethod ?? "none",
      transport: request.transport ?? existing?.transport ?? connector?.transport ?? "local_stdio",
      credentialSource:
        request.credentialSource ??
        (request.useDetectedEnvKey ? "detected_env" : undefined) ??
        existing?.credentialSource ??
        "none",
      secretRef: existing?.secretRef,
      secretEnvKey: trimString(request.secretEnvKey) ?? existing?.secretEnvKey,
      useDetectedEnvKey: trimString(request.useDetectedEnvKey) ?? existing?.useDetectedEnvKey,
      url: normalizeUrl(trimString(request.url) ?? existing?.url ?? connector?.template?.url),
      command: trimString(request.command) ?? existing?.command ?? connector?.template?.command,
      args: request.args?.filter(Boolean) ?? existing?.args ?? connector?.template?.args,
      cwd: trimString(request.cwd) ?? existing?.cwd ?? connector?.template?.cwd,
      env: request.env ?? existing?.env ?? connector?.template?.env,
      defaultEnabled: request.enabled ?? existing?.defaultEnabled ?? true,
      lastHealthyAt: existing?.lastHealthyAt,
      lastError: existing?.lastError,
      oauthExpiresAt: existing?.oauthExpiresAt,
    };
  }

  private prepareDraft(request: ConnectorSetupRequest, requireRuntimeShape: boolean): {
    draft: StoredConnectorDraft;
    connector: ConnectorCatalogItem;
    diagnostics: ConnectorDiagnostic[];
    conflicts: ConnectorConflict[];
  } {
    const custom = isCustomConnectorId(request.connectorId);
    const connector = custom ? buildCustomCatalogItem(request.name) : getConnectorCatalogItem(request.connectorId);
    if (!connector) {
      throw new Error(`Unknown connector: ${request.connectorId}`);
    }

    const existing = request.existingId
      ? this.store.getDefinition(request.existingId)
      : custom
        ? (request.connectorId === "custom" ? undefined : this.store.getDefinition(request.connectorId))
        : this.store.getByLibraryConnectorId(request.connectorId);
    const draft = this.buildDraftFromRequest(request, connector, existing);
    const diagnostics = this.validateDraft(draft, connector, requireRuntimeShape);
    const conflicts = this.store.findConflicts({
      id: draft.id,
      connectorId: draft.connectorId,
      source: draft.source,
      name: draft.name,
      transport: draft.transport,
      url: draft.url,
      command: draft.command,
      args: draft.args,
      cwd: draft.cwd,
    });

    return { draft, connector, diagnostics, conflicts };
  }

  private persistDraftSecret(draft: StoredConnectorDraft, request: ConnectorSetupRequest, persist = true): StoredConnectorDraft {
    if (!persist) {
      if (request.secret && !request.preserveStoredSecret) {
        return {
          ...draft,
          secretRef: this.secretStore.store(request.secret, `probe-${draft.id ?? draft.connectorId}-${randomUUID()}`),
        };
      }
      return draft;
    }

    if (request.secret && !request.preserveStoredSecret) {
      return {
        ...draft,
        secretRef: this.store.persistSecret(request.secret, draft.id ?? draft.connectorId),
      };
    }

    if (!request.secret && request.preserveStoredSecret === false) {
      return {
        ...draft,
        secretRef: undefined,
      };
    }

    return draft;
  }

  private toStatus(
    stored: StoredConnectorDefinition | (StoredConnectorDraft & { id?: string | undefined }),
    scopeContext?: ConnectorScopeContext,
    conflicts: ConnectorConflict[] = [],
    overrideVerification?: ConnectorVerificationSnapshot | undefined,
  ): ConnectorStatus {
    const connector = stored.source === "custom" ? buildCustomCatalogItem(stored.name) : getConnectorCatalogItem(stored.connectorId);
    const snapshot = overrideVerification ?? (stored.id ? this.toSnapshot(this.store.getVerificationSnapshot(stored.id)) : undefined);
    const capabilities = snapshot
      ? {
          tools: snapshot.toolNames,
          allowedTools: snapshot.allowedTools,
          blockedTools: snapshot.blockedTools,
          resourceToolNames: snapshot.resourceToolNames,
          promptNames: snapshot.promptNames,
          allowedPrompts: snapshot.allowedPrompts,
          blockedPrompts: snapshot.blockedPrompts,
          resourceCount: snapshot.resourceCount,
          promptCount: snapshot.promptCount,
          inventoryHash: snapshot.inventoryHash,
        }
      : undefined;
    const scopeStates =
      stored.id && "defaultEnabled" in stored
        ? this.store.resolveScope(stored as StoredConnectorDefinition, scopeContext)
        : [
            { target: "global" as const, enabled: stored.defaultEnabled ?? true, applies: true, label: "Everywhere" },
            { target: "workspace" as const, enabled: stored.defaultEnabled ?? true, applies: false, label: "This folder" },
            { target: "document" as const, enabled: stored.defaultEnabled ?? true, applies: false, label: "This document" },
          ];
    const activeScope = scopeStates.find((entry) => entry.applies)?.target ?? "global";
    const favorite = stored.id ? this.store.getFavoriteIds().has(stored.id) : false;
    const configured = this.isConfigured(stored as StoredConnectorDefinition, connector);
    const needsCredential = this.needsCredential(stored as StoredConnectorDefinition, connector);
    const healthState = connector ? this.inferHealthState(stored as StoredConnectorDefinition, connector, snapshot) : "unverified";
    const staleReason = connector ? this.getStaleReason(stored as StoredConnectorDefinition, connector, this.store.getVerificationSnapshot(stored.id ?? "")) : undefined;

    return {
      id: stored.id ?? stored.connectorId,
      connectorId: stored.connectorId,
      name: stored.name,
      enabled: scopeStates.find((entry) => entry.applies)?.enabled ?? stored.defaultEnabled ?? true,
      configured,
      connected: healthState === "ready",
      source: stored.source,
      category: stored.category,
      maturity: stored.maturity,
      setupKind: stored.setupKind,
      authMethod: stored.authMethod,
      transport: stored.transport,
      credentialSource: stored.credentialSource,
      detectedEnvKey: stored.useDetectedEnvKey ?? stored.secretEnvKey,
      usesDetectedCredential: Boolean(stored.useDetectedEnvKey),
      lastTestedAt: snapshot?.verifiedAt,
      lastError: stored.lastError,
      lastHealthyAt: stored.lastHealthyAt,
      healthState,
      staleReason,
      activeScope,
      scopeStates,
      favorite,
      recommendedHosts: connector?.recommendedHosts ?? ["word", "excel", "powerpoint"],
      setupDifficulty: connector?.setupDifficulty ?? "advanced",
      needsCredential,
      verification: snapshot,
      conflicts,
      logSummary: stored.id ? this.store.getLogSummary(stored.id) : undefined,
      capabilities,
    };
  }

  private toSnapshot(snapshot: StoredConnectorVerification | undefined): ConnectorVerificationSnapshot | undefined {
    if (!snapshot) {
      return undefined;
    }
    return {
      verifiedAt: snapshot.verifiedAt,
      catalogRevision: snapshot.catalogRevision,
      inventoryHash: snapshot.inventoryHash,
      toolNames: snapshot.toolNames,
      allowedTools: snapshot.allowedTools,
      blockedTools: snapshot.blockedTools,
      resourceToolNames: snapshot.resourceToolNames,
      promptNames: snapshot.promptNames,
      allowedPrompts: snapshot.allowedPrompts,
      blockedPrompts: snapshot.blockedPrompts,
      resourceCount: snapshot.resourceCount,
      promptCount: snapshot.promptCount,
      staleReason: snapshot.staleReason,
      lastFailure: snapshot.lastFailure,
    };
  }

  private checkRuntimes(keys?: ConnectorRuntimeCheck["key"][]): ConnectorRuntimeCheck[] {
    const requested = keys ?? ["node", "npm", "npx", "python", "uv", "docker", "git"];
    return unique(requested).map((key) => discoverCommand(key));
  }

  private validateDraft(
    draft: StoredConnectorDraft,
    connector: ConnectorCatalogItem | undefined,
    requireRuntimeShape: boolean,
  ): ConnectorDiagnostic[] {
    const diagnostics: ConnectorDiagnostic[] = [];
    if (!draft.name) {
      diagnostics.push({
        level: "error",
        code: "name_required",
        title: "Name required",
        message: "Give this connector a short name so users can recognize it later.",
        connectorId: draft.connectorId,
      });
    }

    if (draft.transport === "local_stdio" && requireRuntimeShape && !draft.command) {
      diagnostics.push({
        level: "error",
        code: "command_required",
        title: "Local launch command required",
        message: "This connector runs on the PC, so it needs a command to start the MCP server.",
        connectorId: draft.connectorId,
      });
    }

    if (draft.transport === "remote_http" && requireRuntimeShape && !draft.url) {
      diagnostics.push({
        level: "error",
        code: "url_required",
        title: "Service URL required",
        message: "This connector needs the MCP service URL before it can be tested.",
        connectorId: draft.connectorId,
      });
    }

    if (draft.authMethod !== "none" && !this.resolveCredential(draft as StoredConnectorDefinition, connector).value && !this.resolveCredential(draft as StoredConnectorDefinition, connector).envKey) {
      diagnostics.push({
        level: "warning",
        code: "credential_missing",
        title: "Credentials still needed",
        message: "No stored key or detected environment variable is available yet.",
        connectorId: draft.connectorId,
      });
    }

    if (draft.transport === "remote_http" && draft.url && !/^https?:\/\//i.test(draft.url)) {
      diagnostics.push({
        level: "error",
        code: "url_invalid",
        title: "Enter a full URL",
        message: "Use a full MCP URL such as https://example.com/mcp.",
        connectorId: draft.connectorId,
      });
    }

    return diagnostics;
  }

  private resolveCredential(
    stored: Pick<StoredConnectorDefinition, "useDetectedEnvKey" | "secretEnvKey" | "secretRef" | "credentialSource">,
    connector: ConnectorCatalogItem | undefined,
  ): { value?: string | undefined; envKey?: string | undefined } {
    const detected = trimString(stored.useDetectedEnvKey);
    if (detected && process.env[detected]) {
      return { value: process.env[detected], envKey: detected };
    }

    const explicitEnv = trimString(stored.secretEnvKey);
    if (explicitEnv && process.env[explicitEnv]) {
      return { value: process.env[explicitEnv], envKey: explicitEnv };
    }

    const hintKey = connector?.envHints[0]?.key;
    if ((stored.credentialSource === "env" || stored.credentialSource === "detected_env") && hintKey && process.env[hintKey]) {
      return { value: process.env[hintKey], envKey: hintKey };
    }

    if ("secretRef" in stored && stored.secretRef) {
      return { value: this.secretStore.read(stored.secretRef), envKey: explicitEnv };
    }

    return { value: undefined, envKey: explicitEnv ?? detected };
  }

  private needsCredential(stored: StoredConnectorDefinition, connector: ConnectorCatalogItem | undefined): boolean {
    if (stored.authMethod === "none") {
      return false;
    }
    const credential = this.resolveCredential(stored, connector);
    return !credential.value && !credential.envKey;
  }

  private isConfigured(stored: StoredConnectorDefinition, connector: ConnectorCatalogItem | undefined): boolean {
    if (stored.transport === "local_stdio" && !stored.command) {
      return false;
    }
    if (stored.transport === "remote_http" && !stored.url) {
      return false;
    }
    return !this.needsCredential(stored, connector);
  }

  private inferHealthState(
    stored: Pick<StoredConnectorDefinition, "authMethod" | "transport" | "lastError" | "oauthExpiresAt" | "useDetectedEnvKey" | "secretEnvKey" | "secretRef" | "credentialSource">,
    connector: ConnectorCatalogItem,
    snapshot: ConnectorVerificationSnapshot | undefined,
  ): ConnectorHealthState {
    if (stored.oauthExpiresAt && Date.parse(stored.oauthExpiresAt) <= Date.now()) {
      return "auth_expired";
    }
    if (this.needsCredential(stored as StoredConnectorDefinition, connector)) {
      return "auth_required";
    }
    if (!snapshot) {
      return stored.lastError ? (this.isAuthError(stored.lastError) ? "auth_expired" : "offline") : "unverified";
    }
    if (snapshot.staleReason) {
      return "stale";
    }
    if (snapshot.lastFailure) {
      return this.isAuthError(snapshot.lastFailure) ? "auth_expired" : snapshot.allowedTools.length ? "degraded" : "offline";
    }
    return "ready";
  }

  private getStaleReason(
    stored: StoredConnectorDefinition,
    connector: ConnectorCatalogItem,
    snapshot: StoredConnectorVerification | undefined,
  ): string | undefined {
    if (!snapshot) {
      return "Connector has not been verified yet.";
    }
    if (snapshot.staleReason) {
      return snapshot.staleReason;
    }
    if (connector.catalogRevision && snapshot.catalogRevision !== connector.catalogRevision) {
      return "The connector library entry changed. Re-verify before using it again.";
    }
    if (Date.parse(stored.updatedAt) > Date.parse(snapshot.verifiedAt)) {
      return "The connector settings changed. Re-verify before using it again.";
    }
    const ttlMs = stored.transport === "remote_http" ? 24 * 60 * 60 * 1000 : 7 * 24 * 60 * 60 * 1000;
    if (Date.now() - Date.parse(snapshot.verifiedAt) > ttlMs) {
      return stored.transport === "remote_http"
        ? "Remote verification is older than 24 hours."
        : "Local verification is older than 7 days.";
    }
    if (stored.oauthExpiresAt && Date.parse(stored.oauthExpiresAt) <= Date.now()) {
      return "Stored sign-in has expired.";
    }
    return undefined;
  }

  private isAuthError(message: string | undefined): boolean {
    const normalized = message?.toLowerCase() ?? "";
    return normalized.includes("401") || normalized.includes("403") || normalized.includes("expired") || normalized.includes("unauthor");
  }

  private async probeDefinition(
    stored: StoredConnectorDefinition | StoredConnectorDraft,
    connector: ConnectorCatalogItem,
  ): Promise<VerificationProbeResult> {
    const definition = this.resolveRuntimeDefinition(stored as StoredConnectorDefinition, connector);
    const emptySnapshot: ConnectorVerificationSnapshot = {
      verifiedAt: nowIso(),
      catalogRevision: connector.catalogRevision ?? "catalog-v1",
      inventoryHash: "",
      toolNames: [],
      allowedTools: [],
      blockedTools: [],
      resourceToolNames: [],
      promptNames: [],
      allowedPrompts: [],
      blockedPrompts: [],
      resourceCount: 0,
      promptCount: 0,
      lastFailure: undefined,
    };
    if (!definition) {
      return {
        ok: false,
        snapshot: {
          ...emptySnapshot,
          staleReason: "Connector configuration is incomplete.",
          lastFailure: "Connector configuration is incomplete.",
        },
        error: "Connector configuration is incomplete.",
      };
    }

    const serverName = sanitizeServerName(stored.id ?? stored.connectorId ?? stored.name);
    let client: Client | undefined;
    let transport: ProbeTransport | undefined;

    try {
      const runtime = await withTimeout(
        (async () => {
          client = new Client({ name: `pi-office-${serverName}`, version: "1.0.0" });
          if (definition.transport === "local_stdio") {
            transport = new StdioClientTransport({
              command: definition.command!,
              args: definition.args ?? [],
              stderr: "ignore",
              ...(definition.env ? { env: definition.env } : {}),
              ...(definition.cwd ? { cwd: definition.cwd } : {}),
            });
          } else {
            transport = await createRemoteTransport(definition);
          }

          await client.connect(transport as never);
          const [tools, resources, prompts] = await Promise.all([
            fetchAllTools(client),
            fetchAllResources(client),
            fetchAllPrompts(client),
          ]);
          return { tools, resources, prompts };
        })(),
        25_000,
        "Connector probe timed out.",
      );

      const allowedTools = runtime.tools
        .filter((tool) => this.isToolAllowed(tool.name, connector))
        .map((tool) => formatToolName(serverName, tool.name));
      const blockedTools = runtime.tools
        .filter((tool) => !this.isToolAllowed(tool.name, connector))
        .map((tool) => formatToolName(serverName, tool.name));
      const resourceToolNames = connector.readPolicy.allowResources
        ? runtime.resources.map((resource) => formatToolName(serverName, `get_${resourceNameToToolName(resource.name)}`))
        : [];
      const promptNames = runtime.prompts.map((prompt) => prompt.name);
      const allowedPrompts = connector.customOnly
        ? []
        : promptNames.filter((prompt) => this.isPromptAllowed(prompt, connector));
      const blockedPrompts = promptNames.filter((prompt) => !allowedPrompts.includes(prompt));
      const toolNames = [
        ...runtime.tools.map((tool) => formatToolName(serverName, tool.name)),
        ...resourceToolNames,
      ];
      const inventoryHash = `${computeConnectorFingerprint({
        source: stored.source,
        connectorId: stored.connectorId,
        transport: stored.transport,
        url: stored.url,
        command: stored.command,
        args: stored.args,
        cwd: stored.cwd,
      })}:${toolNames.join("|")}:${promptNames.join("|")}`;
      const snapshot: ConnectorVerificationSnapshot = {
        verifiedAt: nowIso(),
        catalogRevision: connector.catalogRevision ?? "catalog-v1",
        inventoryHash,
        toolNames,
        allowedTools: [...allowedTools, ...resourceToolNames],
        blockedTools,
        resourceToolNames,
        promptNames,
        allowedPrompts,
        blockedPrompts,
        resourceCount: runtime.resources.length,
        promptCount: promptNames.length,
      };

      if (snapshot.allowedTools.length === 0 && snapshot.allowedPrompts.length === 0) {
        return {
          ok: false,
          snapshot: {
            ...snapshot,
            lastFailure: "The connector did not expose any read-safe tools or prompts.",
          },
          error: "The connector did not expose any read-safe tools or prompts.",
        };
      }

      return {
        ok: true,
        snapshot,
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      return {
        ok: false,
        snapshot: {
          ...emptySnapshot,
          lastFailure: message,
        },
        error: message,
      };
    } finally {
      await client?.close().catch(() => {});
      await transport?.close().catch(() => {});
    }
  }

  private isToolAllowed(toolName: string, connector: ConnectorCatalogItem): boolean {
    if (matchesAnyPattern(toolName, connector.readPolicy.blockToolPatterns)) {
      return false;
    }
    return matchesAnyPattern(toolName, connector.readPolicy.allowToolPatterns);
  }

  private isPromptAllowed(promptName: string, connector: ConnectorCatalogItem): boolean {
    if (!connector.readPolicy.allowPrompts || connector.customOnly) {
      return false;
    }
    if (connector.readPolicy.blockPromptPatterns && matchesAnyPattern(promptName, connector.readPolicy.blockPromptPatterns)) {
      return false;
    }
    const patterns = connector.readPolicy.allowPromptPatterns ?? [".*"];
    return matchesAnyPattern(promptName, patterns);
  }

  private resolveRuntimeDefinition(
    stored: StoredConnectorDefinition,
    connector: ConnectorCatalogItem | undefined,
  ): ResolvedRuntimeDefinition | undefined {
    const credential = this.resolveCredential(stored, connector);
    if (stored.transport === "local_stdio") {
      if (!stored.command) return undefined;
      const env = { ...(stored.env ?? {}) };
      const templateEnv = connector?.template?.env ?? {};
      for (const [key, value] of Object.entries(templateEnv)) {
        const placeholderMatch = value.match(/^\$\{(.+)\}$/);
        if (placeholderMatch?.[1]) {
          const sourceKey = credential.envKey ?? placeholderMatch[1];
          const resolvedValue = credential.value ?? process.env[sourceKey];
          if (resolvedValue) {
            env[key] = resolvedValue;
          }
        } else {
          env[key] = value;
        }
      }

      if (connector?.envHints[0]?.key && credential.value && !env[connector.envHints[0].key]) {
        env[connector.envHints[0].key] = credential.value;
      }

      return {
        transport: stored.transport,
        command: stored.command,
        ...(stored.args?.length ? { args: stored.args } : {}),
        ...(Object.keys(env).length ? { env } : {}),
        ...(stored.cwd ? { cwd: stored.cwd } : {}),
      };
    }

    if (!stored.url) return undefined;
    return {
      transport: stored.transport,
      url: stored.url,
      ...(credential.value ? { bearerToken: credential.value } : {}),
      ...(credential.envKey && !credential.value ? { bearerTokenEnv: credential.envKey } : {}),
    };
  }

  private buildStatusDiagnostics(
    stored: StoredConnectorDefinition,
    connector: ConnectorCatalogItem,
    overrideVerification?: ConnectorVerificationSnapshot | undefined,
  ): ConnectorDiagnostic[] {
    const diagnostics = this.validateDraft(this.buildDraftFromStored(stored), connector, true);
    const snapshot = overrideVerification ?? this.toSnapshot(this.store.getVerificationSnapshot(stored.id));
    if (snapshot?.lastFailure) {
      diagnostics.push({
        level: "warning",
        code: "probe_failed",
        title: "Verification failed",
        message: redactText(snapshot.lastFailure) ?? "Verification failed.",
        connectorId: stored.id,
      });
    } else if (snapshot) {
      diagnostics.push({
        level: "info",
        code: "probe_ok",
        title: "Connector verified",
        message: `${snapshot.allowedTools.length} read-safe tools are available.`,
        connectorId: stored.id,
      });
    }
    return diagnostics;
  }

  private writeSessionConfig(sessionId: string, statuses: ConnectorStatus[]): string {
    const servers: Record<string, Record<string, unknown>> = {};
    for (const status of statuses) {
      if (status.healthState !== "ready") {
        continue;
      }

      const stored = this.store.getDefinition(status.id);
      if (!stored) {
        continue;
      }

      const connector = stored.source === "custom" ? buildCustomCatalogItem(stored.name) : getConnectorCatalogItem(stored.connectorId);
      const definition = connector ? this.resolveRuntimeDefinition(stored, connector) : undefined;
      if (!definition) {
        continue;
      }

      const serverName = sanitizeServerName(stored.id);
      if (definition.transport === "local_stdio") {
        servers[serverName] = {
          command: definition.command,
          args: definition.args ?? [],
          env: definition.env,
          cwd: definition.cwd,
          lifecycle: "lazy",
          exposeResources: true,
          directTools: false,
        };
      } else {
        servers[serverName] = {
          url: definition.url,
          auth: "bearer",
          bearerToken: definition.bearerToken,
          bearerTokenEnv: definition.bearerTokenEnv,
          lifecycle: "lazy",
          exposeResources: true,
          directTools: false,
        };
      }
    }

    const configPath = join(this.config.connectorsRuntimeDir, `${sessionId}.json`);
    mkdirSync(dirname(configPath), { recursive: true });
    writeFileSync(
      configPath,
      `${JSON.stringify({ settings: { toolPrefix: "server", directTools: false, idleTimeout: 10 }, mcpServers: servers }, null, 2)}\n`,
      "utf8",
    );
    return configPath;
  }

  private appendLog(entry: Omit<StoredConnectorLogEntry, "id" | "timestamp">): void {
    this.store.appendLog(entry);
  }
}
