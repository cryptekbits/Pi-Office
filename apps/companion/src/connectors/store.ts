import { createHash, randomUUID } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import type {
  ConnectorAuditPreference,
  ConnectorAuthMethod,
  ConnectorCategory,
  ConnectorConflict,
  ConnectorCredentialSource,
  ConnectorDiagnosticLevel,
  ConnectorExportBundle,
  ConnectorExportItem,
  ConnectorExportScopeOverride,
  ConnectorHealthState,
  ConnectorImportApplyRequest,
  ConnectorImportApplyResponse,
  ConnectorImportPreviewResponse,
  ConnectorMaturity,
  ConnectorScopeContext,
  ConnectorScopeState,
  ConnectorScopeTarget,
  ConnectorSetupKind,
  ConnectorTransport,
  ConnectorVerificationSnapshot,
} from "@pi-office/pi-office-pack";
import { ConnectorSecretStore, type StoredSecretReference } from "./secret-store.js";

export const CONNECTOR_STORE_VERSION = 2;
const MAX_LOG_ENTRIES = 300;

export interface LegacyStoredConnectorTestResult {
  ok: boolean;
  checkedAt: string;
  error?: string | undefined;
  tools: string[];
  allowedTools: string[];
  blockedTools: string[];
  resourceCount: number;
  promptCount: number;
}

export interface LegacyStoredConnector {
  id: string;
  connectorId: string;
  name: string;
  source: "library" | "custom";
  category: ConnectorCategory;
  maturity: ConnectorMaturity;
  enabled: boolean;
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
  env?: Record<string, string> | undefined;
  createdAt: string;
  updatedAt: string;
  lastTest?: LegacyStoredConnectorTestResult | undefined;
}

export interface StoredConnectorDefinition {
  id: string;
  connectorId: string;
  name: string;
  source: "library" | "custom";
  category: ConnectorCategory;
  maturity: ConnectorMaturity;
  setupKind: ConnectorSetupKind;
  authMethod: ConnectorAuthMethod;
  transport: ConnectorTransport;
  credentialSource: ConnectorCredentialSource;
  secretRef?: StoredSecretReference | undefined;
  secretEnvKey?: string | undefined;
  useDetectedEnvKey?: string | undefined;
  url?: string | undefined;
  command?: string | undefined;
  args?: string[] | undefined;
  cwd?: string | undefined;
  env?: Record<string, string> | undefined;
  createdAt: string;
  updatedAt: string;
  defaultEnabled: boolean;
  fingerprint: string;
  lastHealthyAt?: string | undefined;
  lastError?: string | undefined;
  oauthExpiresAt?: string | undefined;
}

export interface StoredConnectorScopeOverride {
  id: string;
  connectorId: string;
  scopeTarget: Exclude<ConnectorScopeTarget, "global">;
  scopeKey: string;
  enabled: boolean;
  updatedAt: string;
}

export interface StoredConnectorVerification {
  connectorId: string;
  verifiedAt: string;
  catalogRevision: string;
  inventoryHash: string;
  toolNames: string[];
  allowedTools: string[];
  blockedTools: string[];
  resourceToolNames: string[];
  promptNames: string[];
  allowedPrompts: string[];
  blockedPrompts: string[];
  resourceCount: number;
  promptCount: number;
  staleReason?: string | undefined;
  lastFailure?: string | undefined;
}

export interface StoredConnectorLogEntry {
  id: string;
  connectorId?: string | undefined;
  connectorName?: string | undefined;
  kind: "setup" | "test" | "reverify" | "scope" | "favorite" | "remove" | "import" | "export" | "use";
  channel: "diagnostic" | "audit";
  level: ConnectorDiagnosticLevel;
  message: string;
  timestamp: string;
  scopeTarget?: ConnectorScopeTarget | undefined;
  healthState?: ConnectorHealthState | undefined;
  durationMs?: number | undefined;
  errorCode?: string | undefined;
  redacted?: boolean | undefined;
}

interface ConnectorPreferences {
  favorites: string[];
  auditPreference: ConnectorAuditPreference;
}

interface ConnectorStoreFile {
  version: number;
  updatedAt: string;
  definitions: StoredConnectorDefinition[];
  scopeOverrides: StoredConnectorScopeOverride[];
  verificationSnapshots: StoredConnectorVerification[];
  preferences: ConnectorPreferences;
  logs: StoredConnectorLogEntry[];
}

interface LegacyConnectorStoreFile {
  connectors?: LegacyStoredConnector[] | undefined;
}

export interface StoredConnectorDraft
  extends Omit<StoredConnectorDefinition, "id" | "createdAt" | "updatedAt" | "fingerprint"> {
  id?: string | undefined;
}

function nowIso(): string {
  return new Date().toISOString();
}

function trimString(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  const trimmed = value.trim();
  return trimmed.length ? trimmed : undefined;
}

function trimStringArray(value: unknown): string[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value
    .map((entry) => trimString(entry))
    .filter((entry): entry is string => Boolean(entry));
}

function trimStringMap(value: unknown): Record<string, string> | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return undefined;
  }

  const result: Record<string, string> = {};
  for (const [key, entry] of Object.entries(value)) {
    const trimmed = trimString(entry);
    if (trimmed) {
      result[key] = trimmed;
    }
  }

  return Object.keys(result).length ? result : undefined;
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

function normalizeCommandFingerprint(command: string | undefined, args: string[] | undefined, cwd: string | undefined): string | undefined {
  const normalizedCommand = trimString(command)?.toLowerCase();
  if (!normalizedCommand) {
    return undefined;
  }
  const normalizedArgs = (args ?? []).map((entry) => entry.trim()).filter(Boolean).join("\u001f");
  const normalizedCwd = trimString(cwd)?.toLowerCase() ?? "";
  return `${normalizedCommand}\u001e${normalizedArgs}\u001e${normalizedCwd}`;
}

export function computeConnectorFingerprint(input: {
  source: "library" | "custom";
  connectorId: string;
  transport: ConnectorTransport;
  url?: string | undefined;
  command?: string | undefined;
  args?: string[] | undefined;
  cwd?: string | undefined;
}): string {
  if (input.source === "library") {
    return `library:${input.connectorId}`;
  }

  if (input.transport === "remote_http") {
    return `custom:remote:${normalizeUrl(input.url) ?? "missing-url"}`;
  }

  return `custom:local:${normalizeCommandFingerprint(input.command, input.args, input.cwd) ?? "missing-command"}`;
}

function isSavedDocumentContext(context: ConnectorScopeContext | undefined): boolean {
  return Boolean(context?.documentSaved && (context.documentUrl || context.documentId));
}

export function resolveWorkspaceScopeKey(context: ConnectorScopeContext | undefined): string | undefined {
  return trimString(context?.workspaceId)?.toLowerCase();
}

export function resolveDocumentScopeKey(context: ConnectorScopeContext | undefined): string | undefined {
  return isSavedDocumentContext(context)
    ? trimString(context?.documentUrl ?? context?.documentId)?.toLowerCase()
    : undefined;
}

function ensureDir(filePath: string): void {
  mkdirSync(dirname(filePath), { recursive: true });
}

function sha256(input: string): string {
  return createHash("sha256").update(input).digest("hex");
}

function toVerificationSnapshot(stored: StoredConnectorVerification): ConnectorVerificationSnapshot {
  return {
    verifiedAt: stored.verifiedAt,
    catalogRevision: stored.catalogRevision,
    inventoryHash: stored.inventoryHash,
    toolNames: stored.toolNames,
    allowedTools: stored.allowedTools,
    blockedTools: stored.blockedTools,
    resourceToolNames: stored.resourceToolNames,
    promptNames: stored.promptNames,
    allowedPrompts: stored.allowedPrompts,
    blockedPrompts: stored.blockedPrompts,
    resourceCount: stored.resourceCount,
    promptCount: stored.promptCount,
    staleReason: stored.staleReason,
    lastFailure: stored.lastFailure,
  };
}

export class ConnectorStore {
  constructor(
    private readonly filePath: string,
    private readonly secretStore: ConnectorSecretStore,
  ) {}

  listDefinitions(): StoredConnectorDefinition[] {
    return this.read().definitions;
  }

  listVerificationSnapshots(): StoredConnectorVerification[] {
    return this.read().verificationSnapshots;
  }

  listScopeOverrides(connectorId?: string): StoredConnectorScopeOverride[] {
    const overrides = this.read().scopeOverrides;
    return connectorId ? overrides.filter((entry) => entry.connectorId === connectorId) : overrides;
  }

  listLogs(connectorId?: string): StoredConnectorLogEntry[] {
    const logs = this.read().logs;
    return connectorId ? logs.filter((entry) => entry.connectorId === connectorId) : logs;
  }

  getDefinition(id: string): StoredConnectorDefinition | undefined {
    return this.listDefinitions().find((entry) => entry.id === id);
  }

  getByLibraryConnectorId(connectorId: string): StoredConnectorDefinition | undefined {
    return this.listDefinitions().find((entry) => entry.source === "library" && entry.connectorId === connectorId);
  }

  getVerificationSnapshot(connectorId: string): StoredConnectorVerification | undefined {
    return this.listVerificationSnapshots().find((entry) => entry.connectorId === connectorId);
  }

  getFavoriteIds(): Set<string> {
    return new Set(this.read().preferences.favorites);
  }

  getAuditPreference(): ConnectorAuditPreference {
    return this.read().preferences.auditPreference;
  }

  setAuditPreference(preference: ConnectorAuditPreference): ConnectorAuditPreference {
    const store = this.read();
    store.preferences.auditPreference = { enabled: preference.enabled === true };
    this.write(store);
    return store.preferences.auditPreference;
  }

  setFavorite(connectorId: string, favorite: boolean): void {
    const store = this.read();
    const next = new Set(store.preferences.favorites);
    if (favorite) {
      next.add(connectorId);
    } else {
      next.delete(connectorId);
    }
    store.preferences.favorites = [...next].sort((left, right) => left.localeCompare(right));
    this.write(store);
  }

  upsertDefinition(draft: StoredConnectorDraft): StoredConnectorDefinition {
    const store = this.read();
    const timestamp = nowIso();
    const nextId = draft.id?.trim() || (draft.source === "library" ? draft.connectorId : `custom-${randomUUID()}`);
    const existingIndex = store.definitions.findIndex((entry) => entry.id === nextId);
    const previous = existingIndex >= 0 ? store.definitions[existingIndex] : undefined;

    const next: StoredConnectorDefinition = {
      ...draft,
      id: nextId,
      createdAt: previous?.createdAt ?? timestamp,
      updatedAt: timestamp,
      fingerprint: computeConnectorFingerprint({
        source: draft.source,
        connectorId: draft.connectorId,
        transport: draft.transport,
        url: draft.url,
        command: draft.command,
        args: draft.args,
        cwd: draft.cwd,
      }),
    };

    if (existingIndex >= 0) {
      store.definitions[existingIndex] = next;
    } else {
      store.definitions.push(next);
    }

    this.write(store);
    return next;
  }

  removeDefinition(connectorId: string): boolean {
    const store = this.read();
    const existing = store.definitions.find((entry) => entry.id === connectorId);
    if (!existing) {
      return false;
    }

    if (existing.secretRef) {
      this.secretStore.remove(existing.secretRef);
    }

    store.definitions = store.definitions.filter((entry) => entry.id !== connectorId);
    store.scopeOverrides = store.scopeOverrides.filter((entry) => entry.connectorId !== connectorId);
    store.verificationSnapshots = store.verificationSnapshots.filter((entry) => entry.connectorId !== connectorId);
    store.preferences.favorites = store.preferences.favorites.filter((entry) => entry !== connectorId);
    store.logs = store.logs.filter((entry) => entry.connectorId !== connectorId);
    this.write(store);
    return true;
  }

  setVerificationSnapshot(connectorId: string, snapshot: ConnectorVerificationSnapshot): ConnectorVerificationSnapshot {
    const store = this.read();
    const next: StoredConnectorVerification = {
      connectorId,
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
    const existingIndex = store.verificationSnapshots.findIndex((entry) => entry.connectorId === connectorId);
    if (existingIndex >= 0) {
      store.verificationSnapshots[existingIndex] = next;
    } else {
      store.verificationSnapshots.push(next);
    }
    this.write(store);
    return toVerificationSnapshot(next);
  }

  clearVerificationSnapshot(connectorId: string): void {
    const store = this.read();
    store.verificationSnapshots = store.verificationSnapshots.filter((entry) => entry.connectorId !== connectorId);
    this.write(store);
  }

  markVerificationStale(connectorId: string, reason: string): void {
    const snapshot = this.getVerificationSnapshot(connectorId);
    if (!snapshot || snapshot.staleReason === reason) {
      return;
    }

    this.setVerificationSnapshot(connectorId, {
      ...toVerificationSnapshot(snapshot),
      staleReason: reason,
    });
  }

  updateLastHealth(connectorId: string, input: { lastHealthyAt?: string | undefined; lastError?: string | undefined }): void {
    const store = this.read();
    const definition = store.definitions.find((entry) => entry.id === connectorId);
    if (!definition) {
      return;
    }
    definition.lastHealthyAt = input.lastHealthyAt ?? definition.lastHealthyAt;
    definition.lastError = input.lastError;
    definition.updatedAt = nowIso();
    this.write(store);
  }

  updateScope(connectorId: string, scopeTarget: ConnectorScopeTarget, context: ConnectorScopeContext | undefined, enabled: boolean): ConnectorScopeState[] {
    const store = this.read();
    const definition = store.definitions.find((entry) => entry.id === connectorId);
    if (!definition) {
      return [];
    }

    if (scopeTarget === "global") {
      definition.defaultEnabled = enabled;
      definition.updatedAt = nowIso();
      this.write(store);
      return this.resolveScope(definition, context);
    }

    const scopeKey = scopeTarget === "workspace" ? resolveWorkspaceScopeKey(context) : resolveDocumentScopeKey(context);
    if (!scopeKey) {
      return this.resolveScope(definition, context);
    }

    const existingIndex = store.scopeOverrides.findIndex(
      (entry) => entry.connectorId === connectorId && entry.scopeTarget === scopeTarget && entry.scopeKey === scopeKey,
    );
    const next: StoredConnectorScopeOverride = {
      id: existingIndex >= 0 ? store.scopeOverrides[existingIndex]!.id : randomUUID(),
      connectorId,
      scopeTarget,
      scopeKey,
      enabled,
      updatedAt: nowIso(),
    };

    if (existingIndex >= 0) {
      store.scopeOverrides[existingIndex] = next;
    } else {
      store.scopeOverrides.push(next);
    }

    this.write(store);
    return this.resolveScope(definition, context);
  }

  resolveScope(definition: StoredConnectorDefinition, context: ConnectorScopeContext | undefined): ConnectorScopeState[] {
    const workspaceKey = resolveWorkspaceScopeKey(context);
    const documentKey = resolveDocumentScopeKey(context);
    const overrides = this.listScopeOverrides(definition.id);
    const workspaceOverride = workspaceKey
      ? overrides.find((entry) => entry.scopeTarget === "workspace" && entry.scopeKey === workspaceKey)
      : undefined;
    const documentOverride = documentKey
      ? overrides.find((entry) => entry.scopeTarget === "document" && entry.scopeKey === documentKey)
      : undefined;
    const activeTarget: ConnectorScopeTarget = documentOverride ? "document" : workspaceOverride ? "workspace" : "global";

    return [
      {
        target: "global",
        enabled: definition.defaultEnabled,
        applies: activeTarget === "global",
        label: "Everywhere",
      },
      {
        target: "workspace",
        enabled: workspaceOverride?.enabled ?? definition.defaultEnabled,
        applies: activeTarget === "workspace",
        label: "This folder",
        scopeKey: workspaceKey,
        reason: workspaceKey ? undefined : "Open a saved document in a folder to save a folder override.",
      },
      {
        target: "document",
        enabled: documentOverride?.enabled ?? workspaceOverride?.enabled ?? definition.defaultEnabled,
        applies: activeTarget === "document",
        label: "This document",
        scopeKey: documentKey,
        reason: documentKey ? undefined : "Only saved documents can keep a document-specific override.",
      },
    ];
  }

  appendLog(entry: Omit<StoredConnectorLogEntry, "id" | "timestamp"> & { id?: string | undefined; timestamp?: string | undefined }): void {
    const store = this.read();
    store.logs.push({
      ...entry,
      id: entry.id ?? randomUUID(),
      timestamp: entry.timestamp ?? nowIso(),
    });
    if (store.logs.length > MAX_LOG_ENTRIES) {
      store.logs = store.logs.slice(-MAX_LOG_ENTRIES);
    }
    this.write(store);
  }

  getLogSummary(connectorId: string): { total: number; lastAt?: string | undefined; lastLevel?: ConnectorDiagnosticLevel | undefined } {
    const entries = this.listLogs(connectorId);
    const last = entries.at(-1);
    return {
      total: entries.length,
      lastAt: last?.timestamp,
      lastLevel: last?.level,
    };
  }

  readSecret(definition: StoredConnectorDefinition): string | undefined {
    return this.secretStore.read(definition.secretRef);
  }

  persistSecret(secret: string, keyHint: string): StoredSecretReference {
    return this.secretStore.store(secret, keyHint);
  }

  findConflicts(draft: {
    id?: string | undefined;
    connectorId: string;
    source: "library" | "custom";
    name: string;
    transport: ConnectorTransport;
    url?: string | undefined;
    command?: string | undefined;
    args?: string[] | undefined;
    cwd?: string | undefined;
  }): ConnectorConflict[] {
    const fingerprint = computeConnectorFingerprint(draft);
    return this.listDefinitions()
      .filter((entry) => entry.id !== draft.id)
      .flatMap<ConnectorConflict>((entry) => {
        if (draft.source === "library" && entry.source === "library" && entry.connectorId === draft.connectorId) {
          return [{
            key: `duplicate:${entry.id}`,
            kind: "duplicate_connector" as const,
            existingConnectorId: entry.id,
            existingName: entry.name,
            incomingConnectorId: draft.connectorId,
            title: "This connector already exists",
            message: `${entry.name} is already saved. Open the existing connector or replace it instead of creating a duplicate.`,
            resolution: "open_existing" as const,
          }];
        }

        if (entry.fingerprint !== fingerprint) {
          return [];
        }

        return [{
          key: `duplicate:${entry.id}`,
          kind: draft.transport === "remote_http" ? "duplicate_remote_url" : "duplicate_local_runtime",
          existingConnectorId: entry.id,
          existingName: entry.name,
          incomingConnectorId: draft.connectorId,
          title: "This connector is already configured",
          message:
            draft.transport === "remote_http"
              ? `${entry.name} already uses this MCP URL.`
              : `${entry.name} already uses this local command and working folder.`,
          resolution: "replace" as const,
        }];
      });
  }

  buildExportBundle(): ConnectorExportBundle {
    const store = this.read();
    const connectors: ConnectorExportItem[] = store.definitions.map((entry) => ({
      storedId: entry.id,
      connectorId: entry.connectorId,
      name: entry.name,
      source: entry.source,
      category: entry.category,
      maturity: entry.maturity,
      setupKind: entry.setupKind,
      authMethod: entry.authMethod,
      transport: entry.transport,
      credentialSource: entry.credentialSource,
      secretEnvKey: entry.secretEnvKey,
      useDetectedEnvKey: entry.useDetectedEnvKey,
      url: entry.url,
      command: entry.command,
      args: entry.args,
      cwd: entry.cwd,
      env: entry.env,
      defaultEnabled: entry.defaultEnabled,
    }));

    const scopeOverrides: ConnectorExportScopeOverride[] = store.scopeOverrides.map((entry) => {
      const definition = store.definitions.find((candidate) => candidate.id === entry.connectorId);
      return {
        storedId: definition?.id ?? entry.connectorId,
        connectorId: definition?.connectorId ?? entry.connectorId,
        scopeTarget: entry.scopeTarget,
        scopeKey: entry.scopeKey,
        enabled: entry.enabled,
      };
    });

    return {
      version: CONNECTOR_STORE_VERSION,
      exportedAt: nowIso(),
      connectors,
      scopeOverrides,
      favorites: [...store.preferences.favorites],
      auditPreference: store.preferences.auditPreference,
    };
  }

  previewImport(bundle: ConnectorExportBundle): ConnectorImportPreviewResponse {
    const conflicts: ConnectorConflict[] = [];
    for (const connector of bundle.connectors) {
      const incomingId = connector.storedId ?? connector.connectorId;
      conflicts.push(
        ...this.findConflicts({
          connectorId: connector.connectorId,
          source: connector.source,
          name: connector.name,
          transport: connector.transport,
          url: connector.url,
          command: connector.command,
          args: connector.args,
          cwd: connector.cwd,
        }).map((conflict) => ({
          ...conflict,
          kind: "import_collision" as const,
          incomingConnectorId: incomingId,
          key: `import:${incomingId}:${conflict.existingConnectorId}`,
        })),
      );
    }

    return {
      ok: true,
      bundle,
      conflicts,
    };
  }

  applyImport(request: ConnectorImportApplyRequest): ConnectorImportApplyResponse {
    const preview = this.previewImport(request.bundle);
    const resolutions = request.resolutions ?? {};
    const importedConnectorIds: string[] = [];
    const skippedConflictKeys: string[] = [];
    const importedIdMap = new Map<string, string>();

    for (const connector of request.bundle.connectors) {
      const incomingId = connector.storedId ?? connector.connectorId;
      const connectorConflicts = preview.conflicts.filter((entry) => entry.incomingConnectorId === incomingId);
      const replaceTargets = connectorConflicts
        .filter((entry) => resolutions[entry.key] === "replace")
        .map((entry) => entry.existingConnectorId);
      const shouldSkip = connectorConflicts.some((entry) => resolutions[entry.key] === "skip" || (!resolutions[entry.key] && entry.resolution !== "replace"));

      if (shouldSkip) {
        skippedConflictKeys.push(...connectorConflicts.map((entry) => entry.key));
        continue;
      }

      for (const targetId of replaceTargets) {
        this.removeDefinition(targetId);
      }

      const existingBuiltin = connector.source === "library" ? this.getByLibraryConnectorId(connector.connectorId) : undefined;
      const saved = this.upsertDefinition({
        id: existingBuiltin?.id,
        connectorId: connector.connectorId,
        name: connector.name,
        source: connector.source,
        category: connector.category,
        maturity: connector.maturity,
        setupKind: connector.setupKind,
        authMethod: connector.authMethod,
        transport: connector.transport,
        credentialSource: connector.credentialSource,
        secretEnvKey: connector.secretEnvKey,
        useDetectedEnvKey: connector.useDetectedEnvKey,
        url: connector.url,
        command: connector.command,
        args: connector.args,
        cwd: connector.cwd,
        env: connector.env,
        defaultEnabled: connector.defaultEnabled,
      });

      this.clearVerificationSnapshot(saved.id);
      importedConnectorIds.push(saved.id);
      if (connector.storedId) {
        importedIdMap.set(connector.storedId, saved.id);
      }
    }

    const store = this.read();
    for (const override of request.bundle.scopeOverrides) {
      const targetId = (override.storedId && importedIdMap.get(override.storedId))
        || (override.connectorId !== "custom" ? this.getByLibraryConnectorId(override.connectorId)?.id : undefined);
      if (!targetId) {
        continue;
      }
      const existingIndex = store.scopeOverrides.findIndex(
        (entry) => entry.connectorId === targetId && entry.scopeTarget === override.scopeTarget && entry.scopeKey === override.scopeKey,
      );
      const next: StoredConnectorScopeOverride = {
        id: existingIndex >= 0 ? store.scopeOverrides[existingIndex]!.id : randomUUID(),
        connectorId: targetId,
        scopeTarget: override.scopeTarget,
        scopeKey: override.scopeKey,
        enabled: override.enabled,
        updatedAt: nowIso(),
      };
      if (existingIndex >= 0) {
        store.scopeOverrides[existingIndex] = next;
      } else {
        store.scopeOverrides.push(next);
      }
    }
    store.preferences.auditPreference = request.bundle.auditPreference;
    const favoriteIds = request.bundle.favorites
      .map((entry) => importedIdMap.get(entry) ?? entry)
      .filter((entry): entry is string => Boolean(entry));
    store.preferences.favorites = Array.from(new Set([
      ...store.preferences.favorites,
      ...favoriteIds,
    ])).sort((left, right) => left.localeCompare(right));
    this.write(store);

    return {
      ok: true,
      importedConnectorIds,
      skippedConflictKeys,
    };
  }

  private migrateLegacy(legacy: LegacyConnectorStoreFile): ConnectorStoreFile {
    const connectors = Array.isArray(legacy.connectors) ? legacy.connectors : [];
    const migrated: ConnectorStoreFile = {
      version: CONNECTOR_STORE_VERSION,
      updatedAt: nowIso(),
      definitions: [],
      scopeOverrides: [],
      verificationSnapshots: [],
      preferences: {
        favorites: [],
        auditPreference: { enabled: false },
      },
      logs: [],
    };

    for (const connector of connectors) {
      if (!connector?.id || !connector.connectorId || !connector.name) {
        continue;
      }

      const secretRef = trimString(connector.secret)
        ? this.secretStore.store(trimString(connector.secret)!, connector.id)
        : undefined;
      const definition: StoredConnectorDefinition = {
        id: connector.id,
        connectorId: connector.connectorId,
        name: connector.name,
        source: connector.source === "custom" ? "custom" : "library",
        category: connector.category,
        maturity: connector.maturity,
        setupKind: connector.setupKind,
        authMethod: connector.authMethod,
        transport: connector.transport,
        credentialSource: connector.credentialSource,
        secretRef,
        secretEnvKey: trimString(connector.secretEnvKey),
        useDetectedEnvKey: trimString(connector.useDetectedEnvKey),
        url: normalizeUrl(connector.url),
        command: trimString(connector.command),
        args: trimStringArray(connector.args),
        cwd: trimString(connector.cwd),
        env: trimStringMap(connector.env),
        createdAt: trimString(connector.createdAt) ?? nowIso(),
        updatedAt: trimString(connector.updatedAt) ?? nowIso(),
        defaultEnabled: connector.enabled !== false,
        fingerprint: computeConnectorFingerprint({
          source: connector.source === "custom" ? "custom" : "library",
          connectorId: connector.connectorId,
          transport: connector.transport,
          url: connector.url,
          command: connector.command,
          args: connector.args,
          cwd: connector.cwd,
        }),
        lastHealthyAt: connector.lastTest?.ok ? trimString(connector.lastTest.checkedAt) : undefined,
        lastError: trimString(connector.lastTest?.error),
      };

      migrated.definitions.push(definition);
      if (connector.lastTest) {
        migrated.verificationSnapshots.push({
          connectorId: connector.id,
          verifiedAt: trimString(connector.lastTest.checkedAt) ?? nowIso(),
          catalogRevision: "legacy-migrated",
          inventoryHash: sha256(JSON.stringify({
            tools: connector.lastTest.tools,
            allowedTools: connector.lastTest.allowedTools,
            blockedTools: connector.lastTest.blockedTools,
            resourceCount: connector.lastTest.resourceCount,
            promptCount: connector.lastTest.promptCount,
          })),
          toolNames: trimStringArray(connector.lastTest.tools),
          allowedTools: trimStringArray(connector.lastTest.allowedTools),
          blockedTools: trimStringArray(connector.lastTest.blockedTools),
          resourceToolNames: [],
          promptNames: [],
          allowedPrompts: [],
          blockedPrompts: [],
          resourceCount: typeof connector.lastTest.resourceCount === "number" ? connector.lastTest.resourceCount : 0,
          promptCount: typeof connector.lastTest.promptCount === "number" ? connector.lastTest.promptCount : 0,
          staleReason: "Migrated from the previous connector store. Re-verify before using it again.",
          lastFailure: trimString(connector.lastTest.error),
        });
      }
    }

    return migrated;
  }

  private read(): ConnectorStoreFile {
    if (!existsSync(this.filePath)) {
      return this.emptyStore();
    }

    try {
      const raw = JSON.parse(readFileSync(this.filePath, "utf8")) as ConnectorStoreFile | LegacyConnectorStoreFile;
      if (typeof raw === "object" && raw && "version" in raw && raw.version === CONNECTOR_STORE_VERSION) {
        return this.normalizeStore(raw as ConnectorStoreFile);
      }

      if (typeof raw === "object" && raw && "connectors" in raw) {
        const migrated = this.migrateLegacy(raw as LegacyConnectorStoreFile);
        this.write(migrated);
        return migrated;
      }
    } catch {
      // Fall through.
    }

    return this.emptyStore();
  }

  private normalizeStore(store: ConnectorStoreFile): ConnectorStoreFile {
    return {
      version: CONNECTOR_STORE_VERSION,
      updatedAt: trimString(store.updatedAt) ?? nowIso(),
      definitions: Array.isArray(store.definitions)
        ? store.definitions
            .map((entry) => this.normalizeDefinition(entry))
            .filter((entry): entry is StoredConnectorDefinition => Boolean(entry))
        : [],
      scopeOverrides: Array.isArray(store.scopeOverrides)
        ? store.scopeOverrides
            .map((entry) => this.normalizeScopeOverride(entry))
            .filter((entry): entry is StoredConnectorScopeOverride => Boolean(entry))
        : [],
      verificationSnapshots: Array.isArray(store.verificationSnapshots)
        ? store.verificationSnapshots
            .map((entry) => this.normalizeVerification(entry))
            .filter((entry): entry is StoredConnectorVerification => Boolean(entry))
        : [],
      preferences: {
        favorites: Array.isArray(store.preferences?.favorites)
          ? store.preferences.favorites
              .map((entry) => trimString(entry))
              .filter((entry): entry is string => Boolean(entry))
          : [],
        auditPreference: {
          enabled: store.preferences?.auditPreference?.enabled === true,
        },
      },
      logs: Array.isArray(store.logs)
        ? store.logs
            .map((entry) => this.normalizeLog(entry))
            .filter((entry): entry is StoredConnectorLogEntry => Boolean(entry))
        : [],
    };
  }

  private normalizeDefinition(value: unknown): StoredConnectorDefinition | undefined {
    if (!value || typeof value !== "object" || Array.isArray(value)) {
      return undefined;
    }

    const entry = value as Record<string, unknown>;
    const id = trimString(entry.id);
    const connectorId = trimString(entry.connectorId);
    const name = trimString(entry.name);
    const source = entry.source === "custom" ? "custom" : entry.source === "library" ? "library" : undefined;
    const category = trimString(entry.category) as ConnectorCategory | undefined;
    const maturity = trimString(entry.maturity) as ConnectorMaturity | undefined;
    const setupKind = trimString(entry.setupKind) as ConnectorSetupKind | undefined;
    const authMethod = trimString(entry.authMethod) as ConnectorAuthMethod | undefined;
    const transport = trimString(entry.transport) as ConnectorTransport | undefined;
    const credentialSource = trimString(entry.credentialSource) as ConnectorCredentialSource | undefined;
    if (!id || !connectorId || !name || !source || !category || !maturity || !setupKind || !authMethod || !transport || !credentialSource) {
      return undefined;
    }

    const secretRefRecord = entry.secretRef;
    const secretRef =
      secretRefRecord && typeof secretRefRecord === "object" && !Array.isArray(secretRefRecord)
        ? {
            backend: trimString((secretRefRecord as Record<string, unknown>).backend) as StoredSecretReference["backend"],
            key: trimString((secretRefRecord as Record<string, unknown>).key) ?? "",
          }
        : undefined;

    return {
      id,
      connectorId,
      name,
      source,
      category,
      maturity,
      setupKind,
      authMethod,
      transport,
      credentialSource,
      secretRef: secretRef?.key ? secretRef : undefined,
      secretEnvKey: trimString(entry.secretEnvKey),
      useDetectedEnvKey: trimString(entry.useDetectedEnvKey),
      url: normalizeUrl(trimString(entry.url)),
      command: trimString(entry.command),
      args: trimStringArray(entry.args),
      cwd: trimString(entry.cwd),
      env: trimStringMap(entry.env),
      createdAt: trimString(entry.createdAt) ?? nowIso(),
      updatedAt: trimString(entry.updatedAt) ?? nowIso(),
      defaultEnabled: entry.defaultEnabled !== false,
      fingerprint:
        trimString(entry.fingerprint) ??
        computeConnectorFingerprint({
          source,
          connectorId,
          transport,
          url: trimString(entry.url),
          command: trimString(entry.command),
          args: trimStringArray(entry.args),
          cwd: trimString(entry.cwd),
        }),
      lastHealthyAt: trimString(entry.lastHealthyAt),
      lastError: trimString(entry.lastError),
      oauthExpiresAt: trimString(entry.oauthExpiresAt),
    };
  }

  private normalizeScopeOverride(value: unknown): StoredConnectorScopeOverride | undefined {
    if (!value || typeof value !== "object" || Array.isArray(value)) {
      return undefined;
    }

    const entry = value as Record<string, unknown>;
    const id = trimString(entry.id);
    const connectorId = trimString(entry.connectorId);
    const scopeTarget = trimString(entry.scopeTarget) as Exclude<ConnectorScopeTarget, "global"> | undefined;
    const scopeKey = trimString(entry.scopeKey);
    if (!id || !connectorId || !scopeTarget || !scopeKey) {
      return undefined;
    }

    return {
      id,
      connectorId,
      scopeTarget,
      scopeKey,
      enabled: entry.enabled !== false,
      updatedAt: trimString(entry.updatedAt) ?? nowIso(),
    };
  }

  private normalizeVerification(value: unknown): StoredConnectorVerification | undefined {
    if (!value || typeof value !== "object" || Array.isArray(value)) {
      return undefined;
    }

    const entry = value as Record<string, unknown>;
    const connectorId = trimString(entry.connectorId);
    const verifiedAt = trimString(entry.verifiedAt);
    const catalogRevision = trimString(entry.catalogRevision);
    const inventoryHash = trimString(entry.inventoryHash);
    if (!connectorId || !verifiedAt || !catalogRevision || !inventoryHash) {
      return undefined;
    }

    return {
      connectorId,
      verifiedAt,
      catalogRevision,
      inventoryHash,
      toolNames: trimStringArray(entry.toolNames),
      allowedTools: trimStringArray(entry.allowedTools),
      blockedTools: trimStringArray(entry.blockedTools),
      resourceToolNames: trimStringArray(entry.resourceToolNames),
      promptNames: trimStringArray(entry.promptNames),
      allowedPrompts: trimStringArray(entry.allowedPrompts),
      blockedPrompts: trimStringArray(entry.blockedPrompts),
      resourceCount: typeof entry.resourceCount === "number" ? entry.resourceCount : 0,
      promptCount: typeof entry.promptCount === "number" ? entry.promptCount : 0,
      staleReason: trimString(entry.staleReason),
      lastFailure: trimString(entry.lastFailure),
    };
  }

  private normalizeLog(value: unknown): StoredConnectorLogEntry | undefined {
    if (!value || typeof value !== "object" || Array.isArray(value)) {
      return undefined;
    }

    const entry = value as Record<string, unknown>;
    const id = trimString(entry.id);
    const kind = trimString(entry.kind) as StoredConnectorLogEntry["kind"] | undefined;
    const channel = trimString(entry.channel) as StoredConnectorLogEntry["channel"] | undefined;
    const level = trimString(entry.level) as ConnectorDiagnosticLevel | undefined;
    const message = trimString(entry.message);
    const timestamp = trimString(entry.timestamp);
    if (!id || !kind || !channel || !level || !message || !timestamp) {
      return undefined;
    }

    return {
      id,
      connectorId: trimString(entry.connectorId),
      connectorName: trimString(entry.connectorName),
      kind,
      channel,
      level,
      message,
      timestamp,
      scopeTarget: trimString(entry.scopeTarget) as ConnectorScopeTarget | undefined,
      healthState: trimString(entry.healthState) as ConnectorHealthState | undefined,
      durationMs: typeof entry.durationMs === "number" ? entry.durationMs : undefined,
      errorCode: trimString(entry.errorCode),
      redacted: entry.redacted !== false,
    };
  }

  private write(store: ConnectorStoreFile): void {
    const next = this.normalizeStore({
      ...store,
      version: CONNECTOR_STORE_VERSION,
      updatedAt: nowIso(),
    });
    ensureDir(this.filePath);
    writeFileSync(this.filePath, `${JSON.stringify(next, null, 2)}\n`, "utf8");
  }

  private emptyStore(): ConnectorStoreFile {
    return {
      version: CONNECTOR_STORE_VERSION,
      updatedAt: nowIso(),
      definitions: [],
      scopeOverrides: [],
      verificationSnapshots: [],
      preferences: {
        favorites: [],
        auditPreference: { enabled: false },
      },
      logs: [],
    };
  }
}
