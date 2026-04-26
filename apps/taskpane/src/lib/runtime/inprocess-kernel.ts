import { Agent, type AgentMessage, type AgentTool } from "@mariozechner/pi-agent-core";
import {
  Type,
  completeSimple,
  getModel,
  getModels,
  getProviders,
  supportsXhigh,
  type Context as PiContext,
  type ImageContent,
  type Model,
  type ThinkingLevel as PiThinkingLevel,
} from "@mariozechner/pi-ai";
import {
  HOST_LABELS,
  OFFICE_APPEND_SYSTEM_PROMPT,
  composeAutonomyPrompt,
  composeOfficeAwarePrompt,
  getOfficeDocumentState,
} from "@pi-office/pi-office-pack/defaults";
import {
  AUTONOMY_LEVEL_AUTO_APPROVE,
  DEFAULT_USER_PREFERENCES,
  EDIT_REJECT_REASON_LABELS,
  OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH,
  TOOL_CATEGORY_MAP,
  type AskUserQuestion,
  type AskUserRequest,
  type AskUserResponse,
  type AuthStatusResponse,
  type BridgeClientMessage,
  type BridgeServerMessage,
  type CheckpointMetadata,
  type CompanionConnectorDefinition,
  type CompanionState,
  type ConnectorDiagnostic,
  type ConnectorAuditPreference,
  type ConnectorAuditPreferenceResponse,
  type ConnectorCatalogResponse,
  type ConnectorDiagnosticsResponse,
  type ConnectorExportBundle,
  type ConnectorFavoriteRequest,
  type ConnectorImportApplyRequest,
  type ConnectorImportApplyResponse,
  type ConnectorImportPreviewRequest,
  type ConnectorImportPreviewResponse,
  type ConnectorLogResponse,
  type ConnectorOAuthCallbackRequest,
  type ConnectorOAuthCallbackResponse,
  type ConnectorOAuthStartResponse,
  type ConnectorPrepareResponse,
  type ConnectorScopeContext,
  type ConnectorScopeUpdateRequest,
  type ConnectorSetupRequest,
  type ConnectorSetupResponse,
  type ConnectorStatus,
  type ConnectorStatusResponse,
  type ConnectorTestResponse,
  type OfficeDocumentState,
  type ContextBreakdownEntry,
  type DeriveSubjectRequest,
  type DeriveSubjectResponse,
  type DocumentCheckpointPayload,
  type ImageModelCatalogResponse,
  type ImageModelDescriptor,
  type OfficeContextPayload,
  type OfficeEditProposal,
  type OfficeEditProposalDecision,
  type OfficeHost,
  type OfficeSessionOpenRequest,
  type OfficeSessionOpenResponse,
  type OfficeSessionStateResponse,
  type OfficeStateUpdate,
  type OfficeToolName,
  type OfficeToolResult,
  type PromptImagePayload,
  type PromptMode,
  type PromptSuggestionMessage,
  type PromptSuggestionRequest,
  type PromptSuggestionResponse,
  type ProviderCatalogResponse,
  type ProviderDescriptor,
  type ProviderModelDescriptor,
  type SessionStatsResponse,
  type SetModelRequest,
  type ThinkingCapabilities,
  type ThinkingLevel,
  type ToolCategory,
  type ToolPermissionDecision,
  type ToolPermissionRequest,
  type UserPreferences,
} from "@pi-office/pi-office-pack/protocol";
import { parsePromptSuggestions } from "@pi-office/pi-office-pack/prompt-suggestions";
import { executeOfficeTool } from "../office-tools";
import { BrowserConnectorRuntime } from "./browser-connectors";
import { CompanionClient, type CompanionSessionBinding } from "./companion-client";

type JsonRecord = Record<string, unknown>;

interface StoredAuthRecord {
  provider: string;
  apiKey: string;
}

interface PendingAskUser {
  resolve: (value: AskUserResponse) => void;
  reject: (error: Error) => void;
  timeout: ReturnType<typeof setTimeout>;
}

interface PendingToolPermission {
  resolve: (value: ToolPermissionDecision) => void;
  reject: (error: Error) => void;
  timeout: ReturnType<typeof setTimeout>;
}

interface PendingEditProposal {
  resolve: (value: OfficeEditProposalDecision) => void;
  reject: (error: Error) => void;
  timeout: ReturnType<typeof setTimeout>;
}

const AUTH_STORAGE_KEY = "pi-office-auth";
const AUTH_STORAGE_KEY_VERSION = 2;
const AUTH_CRYPTO_KEY_STORAGE_KEY = "pi-office-auth-key-v1";
const CHECKPOINT_STORAGE_KEY_PREFIX = "pi-office-checkpoints:";
const MAX_CHECKPOINT_AGE_MS = 3 * 24 * 60 * 60 * 1000;
const MAX_CHECKPOINTS_PER_DOC = 50;
const BROWSER_UNSUPPORTED_PROVIDERS = new Set<string>(["amazon-bedrock"]);
interface EncryptedAuthEnvelope {
  version: number;
  algorithm: "AES-GCM";
  iv: string;
  ciphertext: string;
}

interface StoredCheckpointDocument {
  documentId: string;
  host: OfficeHost;
  checkpoints: CheckpointMetadata[];
  payloads: Record<string, DocumentCheckpointPayload>;
}

function isBrowserProviderSupported(provider: string): boolean {
  return !BROWSER_UNSUPPORTED_PROVIDERS.has(provider);
}

const IMAGE_MODEL_CATALOG: Array<
  Omit<ImageModelDescriptor, "configured"> & { key: string }
> = [
  {
    key: "openai::gpt-image-1",
    provider: "openai",
    modelId: "gpt-image-1",
    modelName: "GPT Image 1",
    apiType: "openai-images",
    supportsReasoningEffort: true,
    supportedAspectRatios: ["1:1", "16:9", "9:16", "4:3", "3:4", "3:2", "2:3"],
    supportedSizes: ["1K", "2K", "4K"],
  },
  {
    key: "openai::dall-e-3",
    provider: "openai",
    modelId: "dall-e-3",
    modelName: "DALL-E 3",
    apiType: "openai-images",
    supportsReasoningEffort: false,
    supportedAspectRatios: ["1:1", "16:9", "9:16"],
    supportedSizes: ["1K", "2K"],
  },
];

function nowIso(): string {
  return new Date().toISOString();
}

function titleCase(input: string): string {
  return input
    .split(/[-_\s]+/g)
    .filter(Boolean)
    .map((chunk) => chunk[0]!.toUpperCase() + chunk.slice(1))
    .join(" ");
}

function parseRequestBody(init?: RequestInit): unknown {
  const body = init?.body;
  if (body == null) return undefined;
  if (typeof body === "string") {
    try {
      return JSON.parse(body);
    } catch {
      return undefined;
    }
  }
  return undefined;
}

function disconnectedCompanionState(lastKnown?: Partial<CompanionState>): CompanionState {
  return {
    status: lastKnown?.status === "error" ? "error" : "unavailable",
    endpoint: lastKnown?.endpoint,
    identity: lastKnown?.identity,
    lastError: lastKnown?.lastError,
    manualEndpoint: lastKnown?.manualEndpoint,
    lastSuccessfulEndpoint: lastKnown?.lastSuccessfulEndpoint,
    sessionId: undefined,
    connectorToolNames: [],
    capabilities: {
      fileRead: false,
      localMcp: false,
      endpoint: lastKnown?.capabilities?.endpoint,
    },
  };
}

function summarizeCompanionForPrompt(companion: CompanionState, documentSaved: boolean): string {
  const lines = [
    `Companion status: ${companion.status}`,
  ];

  if (companion.endpoint) {
    lines.push(`Companion endpoint: ${companion.endpoint}`);
  }

  if (documentSaved) {
    lines.push(
      companion.status === "connected" && companion.capabilities.fileRead
        ? "Read-only local file tools are available for the saved document folder."
        : "Read-only local file tools are unavailable for this session.",
    );
  } else {
    lines.push("Local file tools stay unavailable until the document is saved.");
  }

  if (companion.status === "connected" && companion.connectorToolNames?.length) {
    lines.push(
      "Verified companion MCP tools: " + companion.connectorToolNames.join(", "),
    );
  } else if (companion.status === "connected" && companion.capabilities.localMcp) {
    lines.push("Local MCP execution is available through the companion when configured read-only connectors are ready.");
  } else {
    lines.push("Local MCP execution is unavailable in this session.");
  }

  if (companion.lastError) {
    lines.push(`Companion note: ${companion.lastError}`);
  }

  return lines.join("\n");
}

function normalizeOpenState(request: OfficeSessionOpenRequest): OfficeStateUpdate {
  return {
    host: request.host,
    document: {
      id: request.documentId,
      title: request.title,
      saved: request.saved,
      documentPath: request.documentPath,
      documentUrl: request.documentUrl,
      workspaceDir: request.documentPath ? request.documentPath.split(/[/\\]/).slice(0, -1).join("/") : undefined,
    },
    selection: request.selectionSummary ?? { label: "Selection unavailable" },
    capabilities: [],
    timestamp: nowIso(),
  };
}

function buildBrowserSessionKey(host: OfficeHost, documentId: string, windowId?: string): string {
  return `${host}:${documentId}:${windowId ?? "default"}`;
}

function errorCompanionState(current: CompanionState, error: unknown): CompanionState {
  const lastError = error instanceof Error ? error.message : String(error);
  return {
    ...disconnectedCompanionState({
      ...current,
      status: "error",
      lastError,
    }),
    status: "error",
    endpoint: current.endpoint,
    identity: current.identity,
    lastError,
    manualEndpoint: current.manualEndpoint,
    lastSuccessfulEndpoint: current.lastSuccessfulEndpoint,
    capabilities: {
      fileRead: false,
      localMcp: false,
      endpoint: current.capabilities.endpoint ?? current.endpoint,
    },
  };
}

function toPiImages(images: PromptImagePayload[] | undefined): ImageContent[] | undefined {
  if (!images?.length) return undefined;
  return images.map((image) => ({
    type: "image",
    data: image.data,
    mimeType: image.mimeType,
  }));
}

function describeAttachedImages(images: PromptImagePayload[] | undefined): string | undefined {
  if (!images?.length) return undefined;
  return [
    "Attached visuals:",
    ...images.map((image, index) => {
      const label = image.label?.trim() || `Visual ${index + 1}`;
      const dimensions =
        image.width && image.height ? ` (${Math.round(image.width)}x${Math.round(image.height)})` : "";
      return `- ${label}${dimensions}`;
    }),
  ].join("\n");
}

function clipForSuggestionPrompt(text: string, maxLength: number): string {
  const normalized = text.replace(/\s+/g, " ").trim();
  return normalized.length > maxLength ? `${normalized.slice(0, maxLength - 3)}...` : normalized;
}

function extractAgentMessageText(message: AgentMessage): string {
  if (!message || typeof message !== "object") return "";
  const role = (message as { role?: unknown }).role;
  if (role !== "user" && role !== "assistant") return "";
  const content = (message as { content?: unknown }).content;
  if (typeof content === "string") return content.trim();
  if (!Array.isArray(content)) return "";
  return content
    .filter((part): part is { type: "text"; text: string } =>
      Boolean(part && typeof part === "object" && (part as { type?: unknown }).type === "text"),
    )
    .map((part) => part.text)
    .join("\n")
    .trim();
}

function normalizeSuggestionMessages(
  requestMessages: PromptSuggestionRequest["recentMessages"] | undefined,
  fallbackMessages: AgentMessage[],
): PromptSuggestionMessage[] {
  const source = requestMessages?.length
    ? requestMessages
    : fallbackMessages
        .map((message) => {
          const role = (message as { role?: unknown }).role;
          if (role !== "user" && role !== "assistant") return undefined;
          const text = extractAgentMessageText(message);
          return text ? { role, text } : undefined;
        })
        .filter((entry): entry is PromptSuggestionMessage => Boolean(entry));

  return source
    .filter((entry) => (entry.role === "user" || entry.role === "assistant") && Boolean(entry.text.trim()))
    .slice(-8)
    .map((entry) => ({
      role: entry.role,
      text: clipForSuggestionPrompt(entry.text, 900),
    }));
}

function formatSuggestionMessages(messages: PromptSuggestionMessage[]): string {
  if (!messages.length) return "No recent visible chat was provided.";
  return messages
    .map((message, index) => `${index + 1}. ${message.role}: ${message.text}`)
    .join("\n");
}

function normalizeToolParams(params: unknown): Record<string, unknown> {
  if (!params || typeof params !== "object" || Array.isArray(params)) return {};
  return params as Record<string, unknown>;
}

function hasVisuals(value: unknown): value is { visuals: PromptImagePayload[] } {
  if (!value || typeof value !== "object") return false;
  const visuals = (value as { visuals?: unknown }).visuals;
  return (
    Array.isArray(visuals) &&
    visuals.every(
      (entry) =>
        entry &&
        typeof entry === "object" &&
        typeof (entry as PromptImagePayload).data === "string" &&
        typeof (entry as PromptImagePayload).mimeType === "string",
    )
  );
}

function isOfficeContextPayload(value: unknown): value is OfficeContextPayload {
  return Boolean(value && typeof value === "object" && "summary" in value && "state" in value);
}

function stripBinaryData(value: unknown): unknown {
  if (!value || typeof value !== "object") return value;
  if (Array.isArray(value)) {
    return value.map((entry) => stripBinaryData(entry));
  }

  const next: Record<string, unknown> = { ...(value as JsonRecord) };
  if (Array.isArray(next.visuals)) {
    next.visuals = next.visuals.map((visual) => {
      if (!visual || typeof visual !== "object") return visual;
      const copy: JsonRecord = { ...(visual as JsonRecord) };
      if (typeof copy.data === "string") {
        copy.data = `[base64 ${copy.data.length} chars]`;
      }
      return copy;
    });
  }

  return Object.fromEntries(Object.entries(next).map(([key, entry]) => [key, stripBinaryData(entry)]));
}

function toToolText(value: unknown): string {
  if (typeof value === "string") return value;
  if (isOfficeContextPayload(value)) {
    const structured = JSON.stringify(stripBinaryData(value), null, 2);
    return `${value.summary}\n\nStructured data:\n${structured}`;
  }
  return JSON.stringify(stripBinaryData(value), null, 2);
}

function toToolContent(value: unknown): Array<{ type: "text"; text: string } | { type: "image"; data: string; mimeType: string }> {
  const content: Array<{ type: "text"; text: string } | { type: "image"; data: string; mimeType: string }> = [
    { type: "text", text: toToolText(value) },
  ];
  if (hasVisuals(value)) {
    for (const visual of value.visuals) {
      content.push({ type: "image", data: visual.data, mimeType: visual.mimeType });
    }
  }
  return content;
}

function normalizeExternalToolResult(
  value: unknown,
): {
  content: Array<{ type: "text"; text: string } | { type: "image"; data: string; mimeType: string }>;
  details: unknown;
} {
  if (value && typeof value === "object" && Array.isArray((value as { content?: unknown }).content)) {
    return {
      content: (value as {
        content: Array<{ type: "text"; text: string } | { type: "image"; data: string; mimeType: string }>;
      }).content,
      details: (value as { details?: unknown }).details ?? value,
    };
  }

  return {
    content: toToolContent(value),
    details: value,
  };
}

function toSize(size: string | undefined, aspectRatio: string | undefined): string | undefined {
  if (!size) return undefined;
  if (size === "1K") {
    if (aspectRatio === "16:9") return "1536x1024";
    if (aspectRatio === "9:16") return "1024x1536";
    return "1024x1024";
  }
  if (size === "2K") {
    if (aspectRatio === "16:9") return "1792x1024";
    if (aspectRatio === "9:16") return "1024x1792";
    return "1024x1024";
  }
  if (size === "4K") {
    if (aspectRatio === "16:9") return "1792x1024";
    if (aspectRatio === "9:16") return "1024x1792";
    return "1024x1024";
  }
  return undefined;
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

function isEncryptedAuthEnvelope(value: unknown): value is EncryptedAuthEnvelope {
  if (!value || typeof value !== "object") return false;
  const envelope = value as Partial<EncryptedAuthEnvelope>;
  return (
    envelope.version === AUTH_STORAGE_KEY_VERSION &&
    envelope.algorithm === "AES-GCM" &&
    typeof envelope.iv === "string" &&
    typeof envelope.ciphertext === "string"
  );
}

function sanitizeCheckpointDocumentId(documentId: string): string {
  return documentId.replace(/[^a-zA-Z0-9_-]/g, "_").slice(0, 120);
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(new ArrayBuffer(bytes.byteLength));
  copy.set(bytes);
  return copy.buffer;
}

class BrowserAuthStore {
  private store = new Map<string, StoredAuthRecord>();
  readonly ready: Promise<void>;

  constructor() {
    this.ready = this.load();
  }

  list(): string[] {
    return Array.from(this.store.keys()).sort((left, right) => left.localeCompare(right));
  }

  hasAuth(provider: string): boolean {
    return this.store.has(provider);
  }

  getApiKey(provider: string): string | undefined {
    return this.store.get(provider)?.apiKey;
  }

  async setApiKey(provider: string, apiKey: string): Promise<void> {
    this.store.set(provider, { provider, apiKey });
    await this.persist();
  }

  async remove(provider: string): Promise<void> {
    this.store.delete(provider);
    await this.persist();
  }

  private async load(): Promise<void> {
    try {
      const raw = localStorage.getItem(AUTH_STORAGE_KEY);
      if (!raw) return;

      const parsed = JSON.parse(raw) as unknown;
      if (Array.isArray(parsed)) {
        for (const entry of parsed as StoredAuthRecord[]) {
          if (!entry?.provider || !entry.apiKey) continue;
          this.store.set(entry.provider, { provider: entry.provider, apiKey: entry.apiKey });
        }
        await this.persist();
        return;
      }

      if (!isEncryptedAuthEnvelope(parsed)) {
        this.store.clear();
        return;
      }

      const entries = await this.decryptRecords(parsed);
      for (const entry of entries) {
        if (!entry?.provider || !entry.apiKey) continue;
        this.store.set(entry.provider, { provider: entry.provider, apiKey: entry.apiKey });
      }
    } catch {
      this.store.clear();
    }
  }

  private async persist(): Promise<void> {
    try {
      const envelope = await this.encryptRecords(Array.from(this.store.values()));
      localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify(envelope));
    } catch {
      // Ignore quota/storage errors in browser sandbox.
    }
  }

  private async getOrCreateCryptoKey(): Promise<CryptoKey> {
    let keyBase64 = localStorage.getItem(AUTH_CRYPTO_KEY_STORAGE_KEY);
    if (!keyBase64) {
      const keyBytes = new Uint8Array(32);
      crypto.getRandomValues(keyBytes);
      keyBase64 = arrayBufferToBase64(keyBytes.buffer);
      localStorage.setItem(AUTH_CRYPTO_KEY_STORAGE_KEY, keyBase64);
    }

    const rawKey = base64ToUint8Array(keyBase64);
    return crypto.subtle.importKey("raw", toArrayBuffer(rawKey), { name: "AES-GCM" }, false, ["encrypt", "decrypt"]);
  }

  private async encryptRecords(entries: StoredAuthRecord[]): Promise<EncryptedAuthEnvelope> {
    const key = await this.getOrCreateCryptoKey();
    const iv = new Uint8Array(12);
    crypto.getRandomValues(iv);
    const plaintext = new TextEncoder().encode(JSON.stringify(entries));
    const ciphertext = await crypto.subtle.encrypt({ name: "AES-GCM", iv: toArrayBuffer(iv) }, key, toArrayBuffer(plaintext));
    return {
      version: AUTH_STORAGE_KEY_VERSION,
      algorithm: "AES-GCM",
      iv: arrayBufferToBase64(iv.buffer),
      ciphertext: arrayBufferToBase64(ciphertext),
    };
  }

  private async decryptRecords(envelope: EncryptedAuthEnvelope): Promise<StoredAuthRecord[]> {
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
    return Array.isArray(parsed) ? (parsed as StoredAuthRecord[]) : [];
  }
}

class BrowserCheckpointStore {
  persistCheckpoint(documentId: string, checkpoint: DocumentCheckpointPayload): void {
    const next = this.readDocument(documentId) ?? {
      documentId,
      host: checkpoint.host,
      checkpoints: [],
      payloads: {},
    };

    next.host = checkpoint.host;
    next.payloads[checkpoint.id] = checkpoint;
    next.checkpoints = next.checkpoints.filter((entry) => entry.id !== checkpoint.id);
    next.checkpoints.push({
      id: checkpoint.id,
      timestamp: checkpoint.timestamp,
      userPrompt: checkpoint.userPrompt,
      file: `cp-${checkpoint.timestamp}.json`,
    });

    this.prune(next);
    this.writeDocument(documentId, next);
  }

  loadCheckpoints(documentId: string): CheckpointMetadata[] {
    const stored = this.readDocument(documentId);
    if (!stored?.checkpoints?.length) return [];
    return [...stored.checkpoints].sort((left, right) => right.timestamp - left.timestamp);
  }

  loadCheckpointData(documentId: string, checkpointId: string): DocumentCheckpointPayload | null {
    const stored = this.readDocument(documentId);
    if (!stored) return null;
    return stored.payloads[checkpointId] ?? null;
  }

  private prune(document: StoredCheckpointDocument): void {
    const now = Date.now();
    const filtered = document.checkpoints
      .filter((entry) => now - entry.timestamp < MAX_CHECKPOINT_AGE_MS)
      .sort((left, right) => right.timestamp - left.timestamp)
      .slice(0, MAX_CHECKPOINTS_PER_DOC);

    document.checkpoints = filtered;
    const allowed = new Set(filtered.map((entry) => entry.id));
    for (const checkpointId of Object.keys(document.payloads)) {
      if (!allowed.has(checkpointId)) {
        delete document.payloads[checkpointId];
      }
    }
  }

  private storageKey(documentId: string): string {
    return `${CHECKPOINT_STORAGE_KEY_PREFIX}${sanitizeCheckpointDocumentId(documentId)}`;
  }

  private readDocument(documentId: string): StoredCheckpointDocument | null {
    try {
      const raw = localStorage.getItem(this.storageKey(documentId));
      if (!raw) return null;
      const parsed = JSON.parse(raw) as StoredCheckpointDocument;
      if (!parsed || typeof parsed !== "object") return null;
      if (!Array.isArray(parsed.checkpoints) || !parsed.payloads || typeof parsed.payloads !== "object") {
        return null;
      }
      return parsed;
    } catch {
      return null;
    }
  }

  private writeDocument(documentId: string, document: StoredCheckpointDocument): void {
    try {
      localStorage.setItem(this.storageKey(documentId), JSON.stringify(document));
    } catch {
      // Ignore quota/storage errors in browser sandbox.
    }
  }
}

class BrowserModelRegistry {
  constructor(private readonly authStore: BrowserAuthStore) {}

  find(provider: string, modelId: string): Model<any> | undefined {
    const models = this.getModelsForProvider(provider);
    return models.find((model) => model.id === modelId);
  }

  getProviderCatalog(): ProviderCatalogResponse {
    const providers: ProviderDescriptor[] = [];
    for (const provider of getProviders()) {
      const providerId = String(provider);
      if (!isBrowserProviderSupported(providerId)) continue;
      const models = this.getModelsForProvider(provider);
      if (!models.length) continue;
      const configured = this.authStore.hasAuth(providerId);
      const descriptors: ProviderModelDescriptor[] = models
        .map((model) => {
          const totalCostPer1k = (model.cost.input + model.cost.output) / 2;
          const costTier: "$" | "$$" | "$$$" = totalCostPer1k <= 1 ? "$" : totalCostPer1k <= 10 ? "$$" : "$$$";
          return {
            provider: providerId,
            providerLabel: titleCase(providerId),
            modelId: model.id,
            modelName: model.name,
            configured,
            oauthSupported: false,
            usesApiKey: true,
            contextWindow: model.contextWindow,
            costTier,
            supportsThinking: model.reasoning,
            supportsReasoningEffort: model.reasoning,
          };
        })
        .sort((left, right) => left.modelName.localeCompare(right.modelName));

      providers.push({
        provider: providerId,
        label: titleCase(providerId),
        configured,
        oauthSupported: false,
        models: descriptors,
      });
    }

    providers.sort((left, right) => left.label.localeCompare(right.label));
    return { providers };
  }

  getAuthStatus(): AuthStatusResponse {
    const storedProviders = this.authStore.list().filter(isBrowserProviderSupported);
    return {
      storedProviders,
      oauthProviders: [],
      configuredProviders: storedProviders,
    };
  }

  getApiKey(provider: string): string | undefined {
    return this.authStore.getApiKey(provider);
  }

  getApiKeyAndHeaders(model: Model<any>): { ok: true; apiKey?: string; headers?: Record<string, string> } | { ok: false; error: string } {
    const apiKey = this.authStore.getApiKey(String(model.provider));
    if (!apiKey) {
      return { ok: false, error: `Provider ${model.provider} is not configured.` };
    }
    return {
      ok: true,
      apiKey,
      ...(model.headers ? { headers: model.headers } : {}),
    };
  }

  getPreferredModel(): Model<any> | undefined {
    for (const providerId of this.authStore.list()) {
      if (!isBrowserProviderSupported(providerId)) continue;
      const models = this.getModelsForProvider(providerId);
      if (models.length) return models[0];
    }

    const fallbackProvider = getProviders().find((provider) => isBrowserProviderSupported(String(provider)));
    if (!fallbackProvider) return undefined;
    return this.getModelsForProvider(String(fallbackProvider))[0];
  }

  getThinkingCapabilities(model: Model<any>, currentLevel: ThinkingLevel): ThinkingCapabilities {
    const supportsThinking = Boolean(model.reasoning);
    const baseLevels: ThinkingLevel[] = supportsThinking ? ["off", "minimal", "low", "medium", "high"] : ["off"];
    const supportsExtraHigh = supportsThinking && supportsXhigh(model);
    const availableLevels: ThinkingLevel[] = supportsExtraHigh ? [...baseLevels, "xhigh"] : baseLevels;
    const normalizedLevel = availableLevels.includes(currentLevel) ? currentLevel : "off";
    return {
      supportsThinking,
      availableLevels,
      supportsXhigh: supportsExtraHigh,
      currentLevel: normalizedLevel,
    };
  }

  getImageCatalog(): ImageModelCatalogResponse {
    const models: ImageModelDescriptor[] = IMAGE_MODEL_CATALOG.map((entry) => ({
      provider: entry.provider,
      modelId: entry.modelId,
      modelName: entry.modelName,
      apiType: entry.apiType,
      supportsReasoningEffort: entry.supportsReasoningEffort,
      supportedAspectRatios: entry.supportedAspectRatios,
      supportedSizes: entry.supportedSizes,
      configured: this.authStore.hasAuth(entry.provider),
    }));

    return {
      models,
      defaultModelKey: IMAGE_MODEL_CATALOG[0]?.key ?? "",
    };
  }

  private getModelsForProvider(provider: string): Model<any>[] {
    if (!isBrowserProviderSupported(provider)) return [];
    try {
      return getModels(provider as never) as Model<any>[];
    } catch {
      return [];
    }
  }
}

class BrowserOfficeSession {
  readonly sessionId = crypto.randomUUID();
  readonly documentKey: string;
  readonly windowId: string | undefined;
  private officeState: OfficeStateUpdate;
  private companionState: CompanionState = disconnectedCompanionState();
  private companionConnectors: ConnectorStatusResponse["connectors"] = [];
  private bridgeSockets = new Set<LocalBridgeSocket>();
  private readonly pendingAskUser = new Map<string, PendingAskUser>();
  private readonly pendingPermissions = new Map<string, PendingToolPermission>();
  private readonly pendingEditProposals = new Map<string, PendingEditProposal>();
  private readonly sessionApprovedTools = new Set<string>();
  private readonly agent: Agent;
  private unsubscribeAgent: (() => void) | undefined;

  constructor(
    private readonly authStore: BrowserAuthStore,
    private readonly modelRegistry: BrowserModelRegistry,
    private readonly checkpointStore: BrowserCheckpointStore,
    private readonly getPreferences: () => UserPreferences,
    private readonly executeCompanionFileTool: (sessionId: string, toolName: "read" | "grep" | "find" | "ls", params: Record<string, unknown>) => Promise<unknown>,
    private readonly executeCompanionMcpTool: (sessionId: string, toolName: string, params: Record<string, unknown>) => Promise<unknown>,
    request: OfficeSessionOpenRequest,
  ) {
    this.windowId = request.windowId;
    this.documentKey = buildBrowserSessionKey(request.host, request.documentId, request.windowId);
    this.officeState = normalizeOpenState(request);

    this.agent = new Agent({
      getApiKey: (provider) => this.modelRegistry.getApiKey(provider),
      beforeToolCall: async (ctx) => {
        const toolName = ctx.toolCall.name;
        if (this.shouldAutoApproveTool(toolName)) return undefined;
        const params = normalizeToolParams(ctx.args);
        const decision = await this.requestToolPermission(toolName, params);
        if (!decision.allowed) {
          return { block: true, reason: `User denied ${toolName}.` };
        }
        return undefined;
      },
    });

    const tools = this.buildTools();
    this.agent.setTools(tools);
    this.agent.setSystemPrompt(this.buildSystemPrompt(tools.map((tool) => tool.name)));

    const preferredModel = this.modelRegistry.getPreferredModel();
    if (preferredModel) {
      this.agent.setModel(preferredModel);
    }

    const defaultThinkingLevel = this.getPreferences().defaultThinkingLevel;
    this.agent.setThinkingLevel(this.normalizeThinkingLevel(defaultThinkingLevel, this.agent.state.model) as PiThinkingLevel);

    this.unsubscribeAgent = this.agent.subscribe((event) => {
      this.send({ type: "session_event", event });
    });
  }

  get documentState(): OfficeDocumentState {
    return getOfficeDocumentState(this.officeState.document.saved);
  }

  get companion(): CompanionState {
    return { ...this.companionState };
  }

  setCompanion(binding: CompanionSessionBinding | undefined): void {
    this.setCompanionState(
      binding?.companion ? { ...binding.companion } : disconnectedCompanionState(this.companionState),
      binding?.connectors,
    );
  }

  setCompanionState(companion: CompanionState, connectors?: ConnectorStatusResponse["connectors"]): void {
    this.companionState = { ...companion };
    this.companionConnectors = connectors ? [...connectors] : [];
    const tools = this.buildTools();
    this.agent.setTools(tools);
    this.agent.setSystemPrompt(this.buildSystemPrompt(tools.map((tool) => tool.name)));
  }

  getCompanionConnectors(): ConnectorStatus[] {
    return [...this.companionConnectors];
  }

  attachBridge(socket: LocalBridgeSocket): void {
    this.bridgeSockets.add(socket);
  }

  detachBridge(socket: LocalBridgeSocket): void {
    this.bridgeSockets.delete(socket);
    if (this.bridgeSockets.size === 0) {
      this.rejectAllPending(new Error("Bridge disconnected before interactive request completed."));
    }
  }

  toOpenResponse(): OfficeSessionOpenResponse {
    return {
      sessionId: this.sessionId,
      documentState: this.documentState,
      companion: this.companion,
      origin: window.location.origin,
      eventsPath: `/v1/sessions/${this.sessionId}/events`,
    };
  }

  async updateOfficeState(next: OfficeStateUpdate): Promise<void> {
    this.officeState = next;
    const tools = this.buildTools();
    this.agent.setTools(tools);
    this.agent.setSystemPrompt(this.buildSystemPrompt(tools.map((tool) => tool.name)));
  }

  async prompt(text: string, mode: PromptMode = "prompt", images?: PromptImagePayload[]): Promise<void> {
    const visualNote = describeAttachedImages(images);
    const input = composeOfficeAwarePrompt(
      visualNote ? `${text.trim()}\n\n${visualNote}` : text,
      this.officeState,
    );
    const piImages = toPiImages(images);
    const message: AgentMessage = {
      role: "user",
      content: [{ type: "text", text: input }, ...(piImages ?? [])],
      timestamp: Date.now(),
    };

    if (mode === "steer") {
      if (this.agent.state.isStreaming) {
        this.agent.steer(message);
        return;
      }
      await this.agent.prompt([message]);
      return;
    }

    if (mode === "followUp") {
      if (this.agent.state.isStreaming) {
        this.agent.followUp(message);
        return;
      }
      await this.agent.prompt([message]);
      return;
    }

    await this.agent.prompt(input, piImages);
  }

  async abort(): Promise<void> {
    this.agent.abort();
  }

  async setModel(provider: string, modelId: string): Promise<void> {
    const model = this.modelRegistry.find(provider, modelId);
    if (!model) {
      throw new Error(`Model ${provider}/${modelId} is not available.`);
    }
    this.agent.setModel(model);
    this.agent.setThinkingLevel(
      this.normalizeThinkingLevel(this.getPreferences().defaultThinkingLevel, model) as PiThinkingLevel,
    );
  }

  setThinkingLevel(level: ThinkingLevel): void {
    const normalized = this.normalizeThinkingLevel(level, this.agent.state.model);
    this.agent.setThinkingLevel(normalized as PiThinkingLevel);
  }

  getThinkingCapabilities(): ThinkingCapabilities {
    return this.modelRegistry.getThinkingCapabilities(this.agent.state.model, this.agent.state.thinkingLevel as ThinkingLevel);
  }

  getStats(): SessionStatsResponse {
    const messages = this.agent.state.messages;
    const userMessages = messages.filter((message) => message.role === "user").length;
    const assistantMessages = messages.filter((message) => message.role === "assistant").length;
    const toolResults = messages.filter((message) => message.role === "toolResult").length;

    let toolCalls = 0;
    const totals = {
      input: 0,
      output: 0,
      cacheRead: 0,
      cacheWrite: 0,
      total: 0,
    };
    let totalCost = 0;

    for (const message of messages) {
      if (message.role !== "assistant") continue;
      toolCalls += message.content.filter((entry) => entry.type === "toolCall").length;
      totals.input += message.usage.input;
      totals.output += message.usage.output;
      totals.cacheRead += message.usage.cacheRead;
      totals.cacheWrite += message.usage.cacheWrite;
      totals.total += message.usage.totalTokens;
      totalCost += message.usage.cost.total;
    }

    const contextWindow = this.agent.state.model.contextWindow || 0;
    const contextTokens = totals.total || null;
    const percent = contextTokens && contextWindow ? (contextTokens / contextWindow) * 100 : null;
    let breakdown: ContextBreakdownEntry[] | undefined;

    if (contextTokens != null) {
      const CORE_TOOL_NAMES = new Set([
        "read",
        "bash",
        "edit",
        "write",
        "grep",
        "find",
        "ls",
        "mcp",
        "office_get_context",
        "office_apply_edit",
        "edit_doc_text",
        "edit_doc_list",
        "get_cell_ranges",
        "set_cell_range",
        "clear_cell_range",
        "resize_range",
        "copy_to",
        "modify_sheet_structure",
        "modify_object",
        "get_all_objects",
        "search_data",
        "get_range_as_csv",
        "read_range_image",
        "extract_chart_xml",
        "office_navigate",
        "office_capture_snapshot",
        "office_capture_viewport",
        "office_read_section",
        "verify_doc",
        "verify_doc_visual",
        "get_presentation_structure",
        "get_slide",
        "list_slide_shapes",
        "modify_presentation_structure",
        "duplicate_slide",
        "insert_slide_element",
        "remove_slide_element",
        "edit_slide_text",
        "edit_slide_xml",
        "edit_slide_master",
        "edit_slide_chart",
        "copy_image_between_slides",
        "search_icons",
        "insert_icon",
        "verify_slides",
        "verify_slide_visual",
        "office_execute_js",
        "office_propose_edits",
        "ask_user",
        "generate_image",
      ]);

      const systemPromptChars = this.agent.state.systemPrompt.length;
      let coreToolDefChars = 0;
      let integrationToolDefChars = 0;
      for (const tool of this.agent.state.tools) {
        const chars = tool.name.length
          + (tool.description ?? "").length
          + JSON.stringify(tool.parameters ?? {}).length;
        if (CORE_TOOL_NAMES.has(tool.name)) {
          coreToolDefChars += chars;
        } else {
          integrationToolDefChars += chars;
        }
      }

      let messageChars = 0;
      let toolCallChars = 0;
      for (const message of messages) {
        if (message.role === "user") {
          if (typeof message.content === "string") {
            messageChars += message.content.length;
          } else if (Array.isArray(message.content)) {
            for (const part of message.content) {
              if (part.type === "text") {
                messageChars += part.text.length;
              }
            }
          }
        } else if (message.role === "assistant") {
          for (const part of message.content) {
            if (part.type === "text") {
              messageChars += part.text.length;
            } else if (part.type === "toolCall") {
              toolCallChars += (part.name?.length ?? 0) + JSON.stringify(part.arguments ?? {}).length;
            }
          }
        } else if (message.role === "toolResult") {
          for (const part of message.content) {
            if (part.type === "text") {
              toolCallChars += part.text.length;
            }
          }
        }
      }

      const totalRaw = Math.ceil((systemPromptChars + coreToolDefChars + integrationToolDefChars + messageChars + toolCallChars) / 4);
      const ratio = totalRaw > 0 ? contextTokens / totalRaw : 1;
      const sysTokens = Math.round(Math.ceil(systemPromptChars / 4) * ratio);
      const coreToolTokens = Math.round(Math.ceil(coreToolDefChars / 4) * ratio);
      const integrationTokens = Math.round(Math.ceil(integrationToolDefChars / 4) * ratio);
      const msgTokens = Math.round(Math.ceil(messageChars / 4) * ratio);
      const tcTokens = Math.round(Math.ceil(toolCallChars / 4) * ratio);
      const calibratedSum = sysTokens + coreToolTokens + integrationTokens + msgTokens + tcTokens;
      const other = Math.max(0, contextTokens - calibratedSum);
      const available = Math.max(0, contextWindow - contextTokens);

      breakdown = [
        { label: "System Prompt", tokens: sysTokens, color: "#3b82f6" },
        { label: "Tool Definitions", tokens: coreToolTokens, color: "#a855f7" },
        ...(integrationTokens > 10 ? [{ label: "Integrations", tokens: integrationTokens, color: "#ec4899" }] : []),
        { label: "Messages", tokens: msgTokens, color: "#10b981" },
        { label: "Tool Results", tokens: tcTokens, color: "#f59e0b" },
        ...(other > 10 ? [{ label: "Other", tokens: other, color: "#6b7280" }] : []),
        { label: "Available", tokens: available, color: "#e5e7eb" },
      ];
    }

    return {
      sessionId: this.sessionId,
      userMessages,
      assistantMessages,
      toolCalls,
      toolResults,
      totalMessages: messages.length,
      tokens: totals,
      cost: totalCost,
      contextUsage: {
        tokens: contextTokens,
        contextWindow,
        percent,
        ...(breakdown ? { breakdown } : {}),
      },
    };
  }

  async deriveSubject(messages: DeriveSubjectRequest["messages"]): Promise<string> {
    const model = this.agent.state.model;
    const auth = this.modelRegistry.getApiKeyAndHeaders(model);
    if (!auth.ok) throw new Error(auth.error);

    const context: PiContext = {
      systemPrompt: "Generate a concise 3-6 word title for this conversation. Return only the title, nothing else.",
      messages: messages.map((entry) => ({
        role: "user" as const,
        content: entry.role === "assistant" ? `[assistant]: ${entry.text}` : entry.text,
        timestamp: Date.now(),
      })),
    };

    const result = await completeSimple(model, context, {
      ...(auth.apiKey ? { apiKey: auth.apiKey } : {}),
      ...(auth.headers ? { headers: auth.headers } : {}),
    });
    const text = result.content.find((part) => part.type === "text");
    if (!text || text.type !== "text") throw new Error("No text in subject response.");
    return text.text.trim();
  }

  async suggestPrompts(request: PromptSuggestionRequest): Promise<PromptSuggestionResponse> {
    const generationId = String(request?.generationId ?? "");
    const emptyResponse: PromptSuggestionResponse = { generationId, suggestions: [] };
    const latestAssistantText = String(request?.latestAssistantText ?? "").trim();
    if (latestAssistantText.length < 12) return emptyResponse;

    try {
      const model = this.agent.state.model;
      const auth = this.modelRegistry.getApiKeyAndHeaders(model);
      if (!auth.ok) return emptyResponse;

      const context = this.buildPromptSuggestionContext(
        latestAssistantText,
        normalizeSuggestionMessages(request.recentMessages, this.agent.state.messages),
      );
      const result = await completeSimple(model, context, {
        ...(auth.apiKey ? { apiKey: auth.apiKey } : {}),
        ...(auth.headers ? { headers: auth.headers } : {}),
        maxTokens: 320,
        temperature: 0.2,
        reasoning: "minimal" as PiThinkingLevel,
      });
      const rawText = result.content
        .filter((part): part is { type: "text"; text: string } => part.type === "text")
        .map((part) => part.text)
        .join("\n")
        .trim();

      return {
        generationId,
        suggestions: parsePromptSuggestions(rawText, { documentState: this.documentState }),
      };
    } catch {
      return emptyResponse;
    }
  }

  handleClientMessage(message: BridgeClientMessage): void {
    if (message.type === "client_ready") {
      this.send({ type: "connection_state", state: "ready" });
      this.sendAvailableCheckpoints(this.officeState.document.id);
      return;
    }

    if (message.type === "ask_user_response") {
      const pending = this.pendingAskUser.get(message.response.requestId);
      if (!pending) return;
      clearTimeout(pending.timeout);
      this.pendingAskUser.delete(message.response.requestId);
      pending.resolve(message.response);
      return;
    }

    if (message.type === "tool_permission_response") {
      const pending = this.pendingPermissions.get(message.requestId);
      if (!pending) return;
      clearTimeout(pending.timeout);
      this.pendingPermissions.delete(message.requestId);
      const category = (TOOL_CATEGORY_MAP[message.decision.toolName] ?? "connector") as ToolCategory;
      const decision: ToolPermissionDecision = category === "escape-hatch"
        ? { ...message.decision, scope: "once" }
        : message.decision;
      if (decision.allowed && decision.scope === "session") {
        this.sessionApprovedTools.add(decision.toolName);
      }
      pending.resolve(decision);
      return;
    }

    if (message.type === "edit_proposal_decision") {
      const pending = this.pendingEditProposals.get(message.decision.requestId);
      if (!pending) return;
      clearTimeout(pending.timeout);
      this.pendingEditProposals.delete(message.decision.requestId);
      pending.resolve(message.decision);
      return;
    }

    if (message.type === "rewind_session") {
      this.agent.replaceMessages(this.agent.state.messages.slice(0, Math.max(0, message.targetMessageCount)));
      return;
    }

    if (message.type === "persist_checkpoint") {
      this.handlePersistCheckpoint(message.documentId, message.checkpoint);
      return;
    }

    if (message.type === "load_checkpoint") {
      this.handleLoadCheckpoint(message.documentId, message.checkpointId);
      return;
    }
  }

  async dispose(): Promise<void> {
    this.rejectAllPending(new Error("Session disposed."));
    this.agent.abort();
    this.unsubscribeAgent?.();
    this.unsubscribeAgent = undefined;
    for (const socket of this.bridgeSockets) {
      socket.close();
    }
    this.bridgeSockets.clear();
  }

  private canUseCompanionFileTools(): boolean {
    return this.officeState.document.saved && this.companionState.status === "connected" && this.companionState.capabilities.fileRead;
  }

  private getCompanionConnectorToolNames(): string[] {
    return [...(this.companionState.connectorToolNames ?? [])];
  }

  private hasCompanionConnectorTools(): boolean {
    return this.getCompanionConnectorToolNames().length > 0;
  }

  private buildSystemPrompt(availableToolNames: readonly string[]): string {
    return `${OFFICE_APPEND_SYSTEM_PROMPT}\n\n${composeAutonomyPrompt(
      this.getPreferences(),
      this.canUseCompanionFileTools(),
      availableToolNames,
    )}\n\n${composeOfficeAwarePrompt(
      "Prefer Office tools as the source of truth for the active document.",
      this.officeState,
    )}\n\n${summarizeCompanionForPrompt(this.companionState, this.officeState.document.saved)}`;
  }

  private buildPromptSuggestionContext(
    latestAssistantText: string,
    recentMessages: PromptSuggestionMessage[],
  ): PiContext {
    const documentMode = this.documentState === "saved" ? "saved-document mode" : "unsaved-document mode";
    const selectionPreview =
      this.officeState.selection.structuredPreview?.trim() ||
      this.officeState.selection.textPreview?.trim() ||
      "none";
    const systemPrompt = [
      "You generate next-prompt suggestions for Pi-Office after a successful assistant response.",
      "Return strict JSON only, with this shape: {\"suggestions\":[{\"text\":\"...\"}]}",
      "Return 0 to 3 suggestions. Use 0 when the response is complete or any suggestion would be filler.",
      "Each suggestion must be a concise user prompt the user can review before sending.",
      "Prefer actionable Office-aware prompts: inspect selection, apply a native edit, summarize, continue a concrete section, compare alternatives, or verify document state.",
      "Do not call tools, request tools, mutate Office content, or scan local files.",
      this.documentState === "unsaved"
        ? "The document is unsaved; never suggest local filesystem, folder, repo, AGENTS.md, or SKILL.md actions."
        : "The document is saved; local file actions are still not useful unless the visible chat specifically makes them relevant.",
    ].join("\n");
    const userContent = [
      "Office context:",
      `Host: ${HOST_LABELS[this.officeState.host]}`,
      `Document title: ${this.officeState.document.title}`,
      `Document state: ${this.documentState}`,
      `Current document mode: ${documentMode}`,
      `Selection label: ${this.officeState.selection.label}`,
      `Selection preview: ${clipForSuggestionPrompt(selectionPreview, 900)}`,
      "",
      "Recent visible chat:",
      formatSuggestionMessages(recentMessages),
      "",
      "Latest assistant response:",
      clipForSuggestionPrompt(latestAssistantText, 1800),
    ].join("\n");

    return {
      systemPrompt,
      messages: [
        {
          role: "user",
          content: userContent,
          timestamp: Date.now(),
        },
      ],
    };
  }

  private handlePersistCheckpoint(documentId: string, checkpoint: DocumentCheckpointPayload): void {
    try {
      this.checkpointStore.persistCheckpoint(documentId, checkpoint);
      this.sendAvailableCheckpoints(documentId);
    } catch (error) {
      console.error("[inprocess-kernel] Persist checkpoint failed", error);
    }
  }

  private handleLoadCheckpoint(documentId: string, checkpointId: string): void {
    try {
      const data = this.checkpointStore.loadCheckpointData(documentId, checkpointId);
      if (data) {
        this.send({ type: "session_event", event: { type: "checkpoint_data", checkpoint: data } });
      }
    } catch (error) {
      console.error("[inprocess-kernel] Load checkpoint failed", error);
    }
  }

  private sendAvailableCheckpoints(documentId: string): void {
    try {
      const checkpoints = this.checkpointStore.loadCheckpoints(documentId);
      if (checkpoints.length) {
        this.send({ type: "available_checkpoints", checkpoints });
      }
    } catch {
      // Ignore checkpoint listing failures.
    }
  }

  private normalizeThinkingLevel(level: ThinkingLevel, model: Model<any>): ThinkingLevel {
    const caps = this.modelRegistry.getThinkingCapabilities(model, level);
    if (!caps.supportsThinking) return "off";
    if (!caps.availableLevels.includes(level)) return "high";
    return level;
  }

  private buildTools(): AgentTool[] {
    const getContextParams = Type.Object({
      scope: Type.Optional(
        Type.String({
          description:
            "Optional context hint (selection, document, worksheet, workbook, slide, presentation). Scope filtering is currently strongest for Excel and may be treated as a hint for Word/PowerPoint.",
        }),
      ),
    });

    const applyEditParams = Type.Object({
      mode: Type.Optional(Type.String({ description: "Legacy edit mode such as replaceSelection, insertAfterSelection, or setRangeValues." })),
      content: Type.Optional(Type.String({ description: "Legacy text, HTML, or JSON matrix payload to insert into Office." })),
      text: Type.Optional(Type.String({ description: "Alias for legacy text content. Use when operation/type is insertText." })),
      html: Type.Optional(Type.String({ description: "Alias for legacy HTML content. Use when operation/type is insertHtml." })),
      format: Type.Optional(Type.String({ description: "Legacy content format such as text, html, or matrix." })),
      operation: Type.Optional(Type.String({ description: "Top-level action type alias, e.g., insertText, insertHtml, setRangeValues." })),
      type: Type.Optional(Type.String({ description: "Top-level action type alias when not wrapping with action.type." })),
      values: Type.Optional(Type.Any({ description: "2D array of values for setRangeValues actions." })),
      action: Type.Optional(
        Type.Any({
          description:
            "Structured host action. Prefer this over legacy mode/content for workbook, slide, comment, shape image, chart, table, and navigation-aware edits. Destructive actions should set action.options.confirmDestructive=true.",
        }),
      ),
    });

    const navigateParams = Type.Object({
      target: Type.Optional(Type.String({ description: "Legacy anchor label or text to navigate to." })),
      kind: Type.Optional(
        Type.String({
          description:
            "heading, comment, revision, footnote, endnote, paragraph, field, contentControl, cell, range, sheet, workbook, slide, notesRegion, layout, slideMaster, shape, chart, or pivotTable.",
        }),
      ),
      anchor: Type.Optional(
        Type.Any({
          description: "Structured navigation anchor. Prefer this when sheet names, range addresses, slide IDs, or paragraph IDs are known.",
        }),
      ),
    });

    const captureSnapshotParams = Type.Object({
      scope: Type.Optional(
        Type.String({
          description:
            "Optional context hint (selection, document, worksheet, workbook, slide, shape). Scope is currently most effective for Excel and may be treated as a hint in Word/PowerPoint.",
        }),
      ),
      includeFormatting: Type.Optional(Type.Boolean({ description: "Include formatting and layout metadata alongside the visuals." })),
      maxImages: Type.Optional(Type.Number({ minimum: 0, maximum: 4, description: "Maximum number of visual snapshots to include." })),
    });

    const captureViewportParams = Type.Object({
      includeFormatting: Type.Optional(
        Type.Boolean({ description: "Include Word viewport metadata such as visible pages, scroll position, and view mode." }),
      ),
      includeWindowFrame: Type.Optional(
        Type.Boolean({
          description:
            "Reserved for future native capture support. In browser-only runtime this is acknowledged but cannot capture the full OS window frame.",
        }),
      ),
    });

    const readSectionParams = Type.Object({
      startIndex: Type.Optional(Type.Number({ description: "Zero-based paragraph start index (inclusive)." })),
      endIndex: Type.Optional(Type.Number({ description: "Zero-based paragraph end index (exclusive). Defaults to startIndex + 20." })),
      start: Type.Optional(Type.Number({ description: "Alias for startIndex." })),
      end: Type.Optional(Type.Number({ description: "Alias for endIndex." })),
      includeStyles: Type.Optional(Type.Boolean({ description: "Include paragraph styles and heading levels. Defaults to true." })),
    });

    const verifyDocParams = Type.Object({
      scope: Type.Optional(
        Type.String({
          description: "Verification target for Word context capture (selection or document). Defaults to document-level verification context.",
        }),
      ),
      includeFormatting: Type.Optional(
        Type.Boolean({
          description: "Include formatting and review metadata in the structured verification details. Defaults to true.",
        }),
      ),
    });

    const verifyDocVisualParams = Type.Object({
      includeFormatting: Type.Optional(
        Type.Boolean({
          description: "Include viewport formatting metadata in the structured visual verification details. Defaults to true.",
        }),
      ),
      includeWindowFrame: Type.Optional(
        Type.Boolean({
          description:
            "Reserved for future native capture support. Browser-only runtime acknowledges this flag but cannot capture the full OS window frame.",
        }),
      ),
    });

    const getCellRangesParams = Type.Object({
      sheetName: Type.Optional(Type.String({ description: "Worksheet name that contains the range. Defaults to the active worksheet." })),
      address: Type.Optional(Type.String({ description: "A1-style cell/range address. Defaults to the current selection." })),
      includeValues: Type.Optional(Type.Boolean({ description: "Include range values in the response. Defaults to true." })),
      includeText: Type.Optional(Type.Boolean({ description: "Include rendered text values in the response. Defaults to true." })),
      includeFormulas: Type.Optional(Type.Boolean({ description: "Include range formulas in the response. Defaults to true." })),
      includeNumberFormat: Type.Optional(Type.Boolean({ description: "Include number formats in the response. Defaults to true." })),
    }, { additionalProperties: true });

    const setCellRangeParams = Type.Object({
      sheetName: Type.Optional(Type.String({ description: "Worksheet name that contains the destination range. Defaults to active worksheet." })),
      address: Type.Optional(Type.String({ description: "A1-style destination address. Defaults to current selection." })),
      values: Type.Optional(Type.Any({ description: "2D matrix values to write into the target range." })),
      content: Type.Optional(Type.String({ description: "JSON matrix alias when values is omitted." })),
      options: Type.Optional(Type.Any({ description: "Additional write options forwarded to the host adapter." })),
    }, { additionalProperties: true });

    const clearCellRangeParams = Type.Object({
      sheetName: Type.Optional(Type.String({ description: "Worksheet name that contains the range to clear." })),
      address: Type.Optional(Type.String({ description: "A1-style address to clear. Defaults to current selection." })),
      applyTo: Type.Optional(
        Type.String({
          description: "Clear mode: all, contents, formats, hyperlinks, removeHyperlinks. Defaults to all.",
        }),
      ),
      confirmDestructive: Type.Optional(
        Type.Boolean({
          description: "Set true to acknowledge this destructive clear operation.",
        }),
      ),
      options: Type.Optional(Type.Any({ description: "Additional clear options forwarded to the host adapter." })),
    }, { additionalProperties: true });

    const resizeRangeParams = Type.Object({
      sheetName: Type.Optional(Type.String({ description: "Worksheet name that contains the source range." })),
      address: Type.Optional(Type.String({ description: "A1-style source range address. Defaults to current selection." })),
      rowCount: Type.Optional(Type.Number({ minimum: 1, description: "Final row count for the resized range." })),
      columnCount: Type.Optional(Type.Number({ minimum: 1, description: "Final column count for the resized range." })),
      rowDelta: Type.Optional(Type.Number({ description: "Relative row delta when rowCount is not provided." })),
      columnDelta: Type.Optional(Type.Number({ description: "Relative column delta when columnCount is not provided." })),
      activate: Type.Optional(Type.Boolean({ description: "Activate/select the resized range after resolving it." })),
      options: Type.Optional(Type.Any({ description: "Additional resize options forwarded to the host adapter." })),
    }, { additionalProperties: true });

    const copyToParams = Type.Object({
      sourceSheetName: Type.Optional(Type.String({ description: "Worksheet name for the source range. Defaults to active worksheet." })),
      sourceAddress: Type.Optional(Type.String({ description: "A1-style source range address. Defaults to current selection." })),
      destinationSheetName: Type.Optional(Type.String({ description: "Worksheet name for the destination range." })),
      destinationAddress: Type.String({ description: "A1-style destination range address." }),
      copyType: Type.Optional(
        Type.String({
          description: "Excel copy type: All, Formats, Formulas, Values, or Link.",
        }),
      ),
      skipBlanks: Type.Optional(Type.Boolean({ description: "Skip blank cells while copying. Defaults to false." })),
      transpose: Type.Optional(Type.Boolean({ description: "Transpose copied rows/columns. Defaults to false." })),
      options: Type.Optional(Type.Any({ description: "Additional copy options forwarded to the host adapter." })),
    }, { additionalProperties: true });

    const modifySheetStructureParams = Type.Object({
      operation: Type.String({
        description: "Worksheet structure operation: create_worksheet, rename_worksheet, duplicate_worksheet, or delete_worksheet.",
      }),
      sheetName: Type.Optional(Type.String({ description: "Worksheet name targeted by rename/duplicate/delete operations." })),
      name: Type.Optional(Type.String({ description: "Worksheet name for create/rename/duplicate operations." })),
      relativeTo: Type.Optional(Type.String({ description: "Worksheet name used as placement anchor for duplication." })),
      positionType: Type.Optional(Type.String({ description: "Worksheet copy placement type when duplicating (before/after)." })),
      confirmDestructive: Type.Optional(
        Type.Boolean({
          description: "Set true when performing destructive operations such as delete_worksheet.",
        }),
      ),
      options: Type.Optional(Type.Any({ description: "Additional structure options forwarded to the host adapter." })),
    }, { additionalProperties: true });

    const modifyObjectParams = Type.Object({
      operation: Type.String({
        description:
          "Excel object mutation operation (format_range, create_table, format_table, apply_table_filter, clear_table_filter, clear_table_filters, reapply_table_filters, create_chart, update_chart, create_pivot_table, update_pivot_table, sort_pivot_field, sort_pivot_by_labels, sort_pivot_by_values, refresh_pivot_table, set_worksheet_gridlines, set_worksheet_headings, set_print_area, set_data_validation, clear_data_validation, add_conditional_format, clear_conditional_formats, insert_inline_picture).",
      }),
      sheetName: Type.Optional(Type.String({ description: "Worksheet name for range/table/chart/pivot operations." })),
      address: Type.Optional(Type.String({ description: "A1-style range address when an operation targets a worksheet range." })),
      tableName: Type.Optional(Type.String({ description: "Target table name for table-oriented operations." })),
      chartName: Type.Optional(Type.String({ description: "Target chart name for chart-oriented operations." })),
      pivotTableName: Type.Optional(Type.String({ description: "Target PivotTable name for pivot-oriented operations." })),
      confirmDestructive: Type.Optional(Type.Boolean({ description: "Set true for destructive operations when required." })),
      options: Type.Optional(Type.Any({ description: "Additional object-operation options forwarded to the host adapter." })),
    }, { additionalProperties: true });

    const getAllObjectsParams = Type.Object({
      scope: Type.Optional(
        Type.String({
          description: "Inventory scope hint: selection, worksheet, or workbook. Defaults to workbook.",
        }),
      ),
      includeFormatting: Type.Optional(
        Type.Boolean({
          description: "Include worksheet formatting metadata in addition to object inventory. Defaults to true.",
        }),
      ),
      objectTypes: Type.Optional(
        Type.Array(Type.String(), {
          description: "Optional object kinds to include (table, chart, pivotTable, namedItem, worksheet, cell).",
        }),
      ),
    }, { additionalProperties: true });

    const searchDataParams = Type.Object({
      query: Type.String({ description: "Case-insensitive query used to search workbook/worksheet objects and cited cells." }),
      scope: Type.Optional(
        Type.String({
          description: "Search scope hint: selection, worksheet, or workbook. Defaults to workbook.",
        }),
      ),
      objectTypes: Type.Optional(
        Type.Array(Type.String(), {
          description: "Object kinds to search: table, chart, pivotTable, namedItem, worksheet, or cell.",
        }),
      ),
      limit: Type.Optional(Type.Number({ minimum: 1, maximum: 200, description: "Maximum number of matches to return." })),
    }, { additionalProperties: true });

    const getRangeAsCsvParams = Type.Object({
      sheetName: Type.Optional(Type.String({ description: "Worksheet name that contains the source range. Defaults to active worksheet." })),
      address: Type.Optional(Type.String({ description: "A1-style source range address. Defaults to current selection." })),
      delimiter: Type.Optional(Type.String({ description: "CSV delimiter. Defaults to comma." })),
      quoteValues: Type.Optional(Type.Boolean({ description: "Wrap and escape all CSV cells in quotes. Defaults to false." })),
      includeHeaders: Type.Optional(Type.Boolean({ description: "Include header row. Defaults to true." })),
      includeFormulas: Type.Optional(
        Type.Boolean({
          description: "Export formulas instead of displayed values for auditable formula-first reviews. Defaults to false.",
        }),
      ),
    }, { additionalProperties: true });

    const readRangeImageParams = Type.Object({
      sheetName: Type.Optional(Type.String({ description: "Worksheet name hint for the range image capture context." })),
      address: Type.Optional(Type.String({ description: "A1-style range address hint for the range image capture context." })),
      scope: Type.Optional(
        Type.String({
          description: "Visual capture scope hint. Defaults to selection.",
        }),
      ),
      includeFormatting: Type.Optional(Type.Boolean({ description: "Include formatting metadata in the visual payload. Defaults to true." })),
      maxImages: Type.Optional(Type.Number({ minimum: 1, maximum: 4, description: "Maximum number of range images to include." })),
    }, { additionalProperties: true });

    const extractChartXmlParams = Type.Object({
      sheetName: Type.Optional(Type.String({ description: "Worksheet name that contains the chart. Defaults to active worksheet." })),
      chartName: Type.Optional(Type.String({ description: "Chart name to extract. Required when chartId/chartIndex are not provided." })),
      chartId: Type.Optional(Type.String({ description: "Optional chart id alias when chartName is unknown." })),
      chartIndex: Type.Optional(Type.Number({ minimum: 1, description: "Optional one-based chart index when chartName is unknown." })),
    }, { additionalProperties: true });

    const getPresentationStructureParams = Type.Object({
      maxSlides: Type.Optional(
        Type.Number({
          minimum: 1,
          description: "Optional maximum number of slide previews to include. Defaults to a bounded preview size.",
        }),
      ),
      includeSlideText: Type.Optional(
        Type.Boolean({
          description: "Include per-slide title/body text previews when available. Defaults to true.",
        }),
      ),
    }, { additionalProperties: true });

    const getSlideParams = Type.Object({
      slideId: Type.Optional(Type.String({ description: "PowerPoint slide ID to read." })),
      slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based slide index to read." })),
      includeShapes: Type.Optional(Type.Boolean({ description: "Include shape summaries for the resolved slide. Defaults to true." })),
      includeSlideText: Type.Optional(Type.Boolean({ description: "Include title/body text previews for the resolved slide. Defaults to true." })),
    }, { additionalProperties: true });

    const listSlideShapesParams = Type.Object({
      slideId: Type.Optional(Type.String({ description: "PowerPoint slide ID whose shapes should be listed." })),
      slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based slide index whose shapes should be listed." })),
      maxShapes: Type.Optional(Type.Number({ minimum: 1, description: "Optional maximum number of shapes to return." })),
    }, { additionalProperties: true });

    const modifyPresentationStructureParams = Type.Object({
      operation: Type.String({
        description:
          "Presentation structure operation (add_slide, move_slide, reorder_slides, delete_slide, apply_layout, select_slides, add_agenda_slide, add_transition_slide, combine_slides, import_slides_from_base64).",
      }),
      slideId: Type.Optional(Type.String({ description: "Target slide ID for the operation." })),
      slideIds: Type.Optional(Type.Array(Type.String(), { description: "Ordered list of slide IDs for multi-slide operations." })),
      slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based slide index target." })),
      targetSlideId: Type.Optional(Type.String({ description: "Insertion target slide ID where applicable." })),
      formatting: Type.Optional(Type.String({ description: "PowerPoint insert formatting mode when supported." })),
      confirmDestructive: Type.Optional(
        Type.Boolean({
          description: "Required for destructive operations such as delete_slide/delete_slides.",
        }),
      ),
      content: Type.Optional(Type.String({ description: "Optional text payload used by supported slide-creation helpers." })),
      options: Type.Optional(Type.Any({ description: "Additional operation-specific options forwarded to the host adapter." })),
    }, { additionalProperties: true });

    const duplicateSlideParams = Type.Object({
      slideId: Type.Optional(Type.String({ description: "Single source slide ID to duplicate." })),
      slideIds: Type.Optional(Type.Array(Type.String(), { description: "One or more source slide IDs to duplicate in order." })),
      slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based source slide index when slideId is not known." })),
      targetSlideId: Type.Optional(Type.String({ description: "Slide ID to insert duplicates after." })),
      formatting: Type.Optional(Type.String({ description: "PowerPoint insert formatting mode when supported." })),
      options: Type.Optional(Type.Any({ description: "Additional duplication options forwarded to the host adapter." })),
    }, { additionalProperties: true });

    const insertSlideElementParams = Type.Object({
      operation: Type.String({
        description:
          "PowerPoint element insertion operation (add_text_box, add_geometric_shape, add_table, add_line, add_process_flow, add_simple_diagram, insert_inline_picture).",
      }),
      slideId: Type.Optional(Type.String({ description: "Target slide ID for inserting the new element." })),
      slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based slide index target when slideId is not known." })),
      shapeId: Type.Optional(Type.String({ description: "Optional shape target for grouped operations." })),
      content: Type.Optional(Type.String({ description: "Primary text payload (or base64 image payload for insert_inline_picture)." })),
      text: Type.Optional(Type.String({ description: "Alias for content when inserting text." })),
      values: Type.Optional(Type.Any({ description: "Matrix payload for table insertion when applicable." })),
      options: Type.Optional(Type.Any({ description: "Additional insertion options forwarded to the host adapter." })),
    }, { additionalProperties: true });

    const removeSlideElementParams = Type.Object({
      operation: Type.String({
        description:
          "PowerPoint element removal operation (remove_shape, remove_shapes, clear_shape_text).",
      }),
      slideId: Type.Optional(Type.String({ description: "Target slide ID for the removal operation." })),
      slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based slide index target when slideId is not known." })),
      shapeId: Type.Optional(Type.String({ description: "Primary shape ID to remove or clear." })),
      shapeIds: Type.Optional(Type.Array(Type.String(), { description: "One or more shape IDs for multi-shape removal." })),
      confirmDestructive: Type.Optional(
        Type.Boolean({
          description: "Required for destructive operations such as remove_shape/remove_shapes.",
        }),
      ),
      options: Type.Optional(Type.Any({ description: "Additional removal options forwarded to the host adapter." })),
    }, { additionalProperties: true });

    const editSlideTextParams = Type.Object({
      operation: Type.Optional(
        Type.String({
          description:
            "PowerPoint text-edit operation (set_shape_text, append_shape_text, clear_shape_text, insert_text). Defaults to set_shape_text when shapeId is provided, otherwise insert_text.",
        }),
      ),
      slideId: Type.Optional(Type.String({ description: "Target slide ID for text updates." })),
      slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based slide index target when slideId is not known." })),
      shapeId: Type.Optional(Type.String({ description: "Shape ID whose text should be updated." })),
      content: Type.Optional(Type.String({ description: "Text payload to apply." })),
      text: Type.Optional(Type.String({ description: "Alias for content." })),
      placement: Type.Optional(Type.String({ description: "Optional placement hint (replace or after)." })),
      options: Type.Optional(Type.Any({ description: "Additional text-edit options forwarded to the host adapter." })),
    }, { additionalProperties: true });

    const editSlideXmlParams = Type.Object({
      operation: Type.String({
        description:
          "PowerPoint XML/serialized operation (inspect_presentation_package, get_presentation_theme, get_slide_notes, set_slide_notes, replace_slide_notes, import_slides_from_base64, merge_presentation_from_base64, export_slides_as_base64).",
      }),
      slideId: Type.Optional(Type.String({ description: "Target slide ID for slide-scoped XML operations." })),
      slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based slide index target when slideId is not known." })),
      content: Type.Optional(Type.String({ description: "Text or base64 payload used by mutating XML operations." })),
      base64: Type.Optional(Type.String({ description: "Alias for content when providing serialized PPTX payloads." })),
      formatting: Type.Optional(Type.String({ description: "Insert formatting mode for base64 import operations when supported." })),
      options: Type.Optional(Type.Any({ description: "Additional XML operation options forwarded to the host adapter." })),
    }, { additionalProperties: true });

    const editSlideMasterParams = Type.Object({
      operation: Type.Optional(
        Type.String({
          description:
            "PowerPoint layout/master operation. Currently supports apply_layout (default), with layout/master selectors routed through native layout resolution.",
        }),
      ),
      slideId: Type.Optional(Type.String({ description: "Target slide ID whose layout/master mapping should be updated." })),
      slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based slide index target when slideId is not known." })),
      layoutId: Type.Optional(Type.String({ description: "Layout ID to apply." })),
      layoutName: Type.Optional(Type.String({ description: "Layout name to apply." })),
      slideMasterId: Type.Optional(Type.String({ description: "Optional slide master ID used for layout resolution." })),
      slideMasterName: Type.Optional(Type.String({ description: "Optional slide master name used for layout resolution." })),
      options: Type.Optional(Type.Any({ description: "Additional layout/master options forwarded to the host adapter." })),
    }, { additionalProperties: true });

    const editSlideChartParams = Type.Object({
      operation: Type.Optional(
        Type.String({
          description:
            "PowerPoint chart operation (get_slide_charts, add_slide_chart, update_slide_chart). Defaults to update_slide_chart when omitted.",
        }),
      ),
      slideId: Type.Optional(Type.String({ description: "Target slide ID for chart inspection or mutation." })),
      slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based slide index target when slideId is not known." })),
      shapeId: Type.Optional(Type.String({ description: "Optional chart shape ID for precise chart targeting." })),
      chartIndex: Type.Optional(Type.Number({ minimum: 1, description: "Optional one-based chart index within the target slide package." })),
      shapeName: Type.Optional(Type.String({ description: "Optional chart shape name used to resolve chart edits." })),
      title: Type.Optional(Type.String({ description: "Chart title for create/update operations." })),
      categories: Type.Optional(Type.Any({ description: "Ordered category labels for chart create/update operations." })),
      series: Type.Optional(Type.Any({ description: "Series payload for chart create/update operations." })),
      replaceOriginal: Type.Optional(Type.Boolean({ description: "When true (default), replace the source slide after serialized chart updates." })),
      formatting: Type.Optional(Type.String({ description: "Insert formatting mode when serialized chart updates insert replacement slides." })),
      options: Type.Optional(Type.Any({ description: "Additional chart operation options forwarded to the host adapter." })),
    }, { additionalProperties: true });

    const copyImageBetweenSlidesParams = Type.Object({
      sourceSlideId: Type.Optional(Type.String({ description: "Source slide ID that contains the image shape to copy." })),
      sourceSlideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based source slide index when sourceSlideId is unknown." })),
      sourceShapeId: Type.Optional(Type.String({ description: "Source image shape ID to copy from." })),
      sourceImageBase64: Type.Optional(Type.String({ description: "Optional image base64 override when source shape export is unavailable." })),
      targetSlideId: Type.Optional(Type.String({ description: "Destination slide ID for image placement/replacement." })),
      targetSlideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based destination slide index when targetSlideId is unknown." })),
      targetShapeId: Type.Optional(Type.String({ description: "Destination shape ID to update. If omitted, inserts a new image shape." })),
      left: Type.Optional(Type.Number({ description: "Optional destination left position in points for inserted images." })),
      top: Type.Optional(Type.Number({ description: "Optional destination top position in points for inserted images." })),
      width: Type.Optional(Type.Number({ description: "Optional destination width in points for inserted images." })),
      height: Type.Optional(Type.Number({ description: "Optional destination height in points for inserted images." })),
      options: Type.Optional(Type.Any({ description: "Additional media-copy options forwarded to the host adapter." })),
    }, { additionalProperties: true });

    const searchIconsParams = Type.Object({
      query: Type.String({ description: "Icon search query text." }),
      maxResults: Type.Optional(Type.Number({ minimum: 1, maximum: 50, description: "Maximum number of icon matches to return." })),
      style: Type.Optional(Type.String({ description: "Optional style/category hint used by the icon catalog search." })),
      options: Type.Optional(Type.Any({ description: "Additional icon-search options forwarded to the host adapter." })),
    }, { additionalProperties: true });

    const insertIconParams = Type.Object({
      iconId: Type.Optional(Type.String({ description: "Icon ID returned by search_icons." })),
      iconName: Type.Optional(Type.String({ description: "Icon name alias when iconId is unknown." })),
      query: Type.Optional(Type.String({ description: "Fallback query used when selecting an icon by search text." })),
      slideId: Type.Optional(Type.String({ description: "Target slide ID for icon insertion." })),
      slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based target slide index when slideId is unknown." })),
      shapeId: Type.Optional(Type.String({ description: "Optional target shape ID for icon replacement workflows." })),
      left: Type.Optional(Type.Number({ description: "Optional icon left position in points." })),
      top: Type.Optional(Type.Number({ description: "Optional icon top position in points." })),
      width: Type.Optional(Type.Number({ description: "Optional icon width in points." })),
      height: Type.Optional(Type.Number({ description: "Optional icon height in points." })),
      fillColor: Type.Optional(Type.String({ description: "Optional icon fill/text color (hex/rgb)." })),
      lineColor: Type.Optional(Type.String({ description: "Optional icon outline color (hex/rgb)." })),
      options: Type.Optional(Type.Any({ description: "Additional icon insertion options forwarded to the host adapter." })),
    }, { additionalProperties: true });

    const verifySlidesParams = Type.Object({
      scope: Type.Optional(
        Type.String({
          description: "Structural verification scope for PowerPoint slides (presentation or selection). Defaults to presentation structure.",
        }),
      ),
      maxSlides: Type.Optional(Type.Number({ minimum: 1, description: "Maximum number of slide previews to include in verification details." })),
      includeSlideText: Type.Optional(Type.Boolean({ description: "Include bounded slide text previews in structural verification output." })),
      includeFormatting: Type.Optional(Type.Boolean({ description: "Include layout/master metadata where supported. Defaults to true." })),
    }, { additionalProperties: true });

    const verifySlideVisualParams = Type.Object({
      scope: Type.Optional(
        Type.String({
          description: "Visual verification scope. Defaults to slide selection snapshots.",
        }),
      ),
      includeFormatting: Type.Optional(Type.Boolean({ description: "Include slide/shape formatting metadata in visual verification details. Defaults to true." })),
      maxImages: Type.Optional(Type.Number({ minimum: 1, maximum: 4, description: "Maximum number of slide/shape snapshot images to include." })),
    }, { additionalProperties: true });

    const executeJsParams = Type.Object({
      code: Type.Optional(
        Type.String({
          description:
            "Office.js code to execute. Must use the active host run function (Word.run, Excel.run, or PowerPoint.run). Return a JSON-serializable value. The runtime enforces a best-effort restricted subset (regex checks only, not an isolated sandbox) and blocks network, storage, eval, and system-access patterns.",
        }),
      ),
      script: Type.Optional(Type.String({ description: "Alias for code." })),
    });

    const proposeEditsParams = Type.Object({
      edits: Type.Array(
        Type.Object({
          kind: Type.String({ description: "insert, replace, or delete." }),
          searchText: Type.Optional(Type.String({
            maxLength: OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH,
            description:
              `Text to locate in the document for replace/delete. Must be under ${OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH} characters.`,
          })),
          oldText: Type.Optional(Type.String({ description: "Expected existing text (for replace/delete verification)." })),
          newText: Type.Optional(Type.String({ description: "Replacement text (for insert/replace)." })),
          anchor: Type.Optional(Type.String({ description: "Heading or paragraph label to scope the search." })),
          paragraphId: Type.Optional(Type.String({ description: "Paragraph unique ID for precise targeting." })),
          explanation: Type.Optional(Type.String({ description: "Brief rationale for this edit." })),
        }),
        { description: "Ordered list of proposed edits." },
      ),
      summary: Type.String({ description: "One-sentence summary of all proposed changes." }),
    });

    const simpleOfficeTool = (
      toolName: OfficeToolName,
      label: string,
      description: string,
      parameters: any = Type.Any(),
    ): AgentTool => ({
      name: toolName,
      label,
      description,
      parameters,
      execute: async (_toolCallId, params) => {
        const result = await this.invokeOfficeTool(toolName, normalizeToolParams(params));
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    const reviewableWordEditTool = (
      toolName: "office_propose_edits" | "edit_doc_list",
      label: string,
      description: string,
    ): AgentTool => ({
      name: toolName,
      label,
      description,
      parameters: proposeEditsParams,
      execute: async (_toolCallId, params) => {
        const result = await this.invokeOfficeTool(toolName, normalizeToolParams(params));
        const resultObj = result as JsonRecord | undefined;
        if (!resultObj || !Array.isArray(resultObj.edits) || typeof resultObj.summary !== "string") {
          return {
            content: toToolContent(result),
            details: result,
          };
        }

        const proposal: OfficeEditProposal = {
          requestId: crypto.randomUUID(),
          edits: resultObj.edits as OfficeEditProposal["edits"],
          summary: resultObj.summary,
        };

        const decision = await this.invokeEditProposal(proposal);
        const accepted = decision.decisions.filter((entry) => entry.accepted);
        const rejected = decision.decisions.filter((entry) => !entry.accepted);

        const lines: string[] = [
          `Edit proposal reviewed: ${accepted.length} accepted, ${rejected.length} rejected out of ${proposal.edits.length} edits.`,
        ];

        if (decision.applicationResult) {
          const resultSummary = decision.applicationResult;
          if (resultSummary.failed > 0) {
            lines.push(
              `Application result: ${resultSummary.applied} applied, ${resultSummary.failed} failed. Errors: ${resultSummary.errors.join("; ")}`,
            );
          } else if (resultSummary.applied > 0) {
            lines.push(`All ${resultSummary.applied} accepted edit${resultSummary.applied === 1 ? "" : "s"} applied successfully.`);
          } else {
            lines.push("No edits were applied.");
          }
        } else {
          lines.push(accepted.length ? "Accepted edits were applied to the document." : "No edits were applied.");
        }

        if (rejected.length) {
          lines.push("", "Rejected edits:");
          for (const rejectedEdit of rejected) {
            const reason = rejectedEdit.rejectReason
              ? EDIT_REJECT_REASON_LABELS[rejectedEdit.rejectReason]
              : "No reason given";
            lines.push(`- ${rejectedEdit.editId}: ${reason}${rejectedEdit.rejectNote ? ` (note: ${rejectedEdit.rejectNote})` : ""}`);
          }
        }

        if (decision.globalFeedback) {
          lines.push("", `User feedback: "${decision.globalFeedback}"`);
        }

        return {
          content: [{ type: "text", text: lines.join("\n") }],
          details: {
            proposalId: proposal.requestId,
            decisions: decision.decisions,
            applicationResult: decision.applicationResult,
          },
        };
      },
    });

    const tools: AgentTool[] = [
      simpleOfficeTool(
        "office_get_context",
        "Office Context",
        "Read the current Office document or selection context from the active host.",
        getContextParams,
      ),
      simpleOfficeTool(
        "office_apply_edit",
        "Office Edit",
        "Apply native edits to the active Office document, worksheet, or slide. Prefer action.type/action.content; legacy operation/text params are also supported.",
        applyEditParams,
      ),
      simpleOfficeTool(
        "edit_doc_text",
        "Edit Word Text",
        "Word-only first-class text editing. Use for direct clause/sentence updates through native Word actions when no per-edit review card is required.",
        applyEditParams,
      ),
      simpleOfficeTool(
        "office_navigate",
        "Office Navigate",
        "Move to an Office anchor such as heading, range, or slide.",
        navigateParams,
      ),
      simpleOfficeTool(
        "office_capture_snapshot",
        "Office Snapshot",
        "Capture visual snapshots and formatting metadata for the active Office surface.",
        captureSnapshotParams,
      ),
      simpleOfficeTool(
        "office_capture_viewport",
        "Office Viewport",
        "Capture Word viewport metadata for layout-sensitive tasks (not a pixel-perfect OS/window screenshot).",
        captureViewportParams,
      ),
      simpleOfficeTool(
        "office_read_section",
        "Read Document Section",
        "Read a paginated range of Word paragraphs by paragraph index.",
        readSectionParams,
      ),
      simpleOfficeTool(
        "verify_doc",
        "Verify Word Document",
        "Collect a non-mutating, structured Word verification context with summary text and detailed anchors/snippets for document checks.",
        verifyDocParams,
      ),
      simpleOfficeTool(
        "verify_doc_visual",
        "Verify Word Visual",
        "Capture non-mutating Word visual verification context through the supported viewport path. Word-only; returns structured visual/details payloads.",
        verifyDocVisualParams,
      ),
      simpleOfficeTool(
        "get_cell_ranges",
        "Get Cell Ranges",
        "Excel-only first-class range read tool for cell/range values, text, formulas, and number formats.",
        getCellRangesParams,
      ),
      simpleOfficeTool(
        "set_cell_range",
        "Set Cell Range",
        "Excel-only first-class range write tool for setting values in a target cell/range.",
        setCellRangeParams,
      ),
      simpleOfficeTool(
        "clear_cell_range",
        "Clear Cell Range",
        "Excel-only first-class range clear tool. Destructive clears should set confirmDestructive=true.",
        clearCellRangeParams,
      ),
      simpleOfficeTool(
        "resize_range",
        "Resize Range",
        "Excel-only first-class range layout tool for computing/activating resized ranges by count or delta.",
        resizeRangeParams,
      ),
      simpleOfficeTool(
        "copy_to",
        "Copy To Range",
        "Excel-only first-class range copy tool that copies a source range into a destination range.",
        copyToParams,
      ),
      simpleOfficeTool(
        "modify_sheet_structure",
        "Modify Sheet Structure",
        "Excel-only first-class worksheet structure tool for create, rename, duplicate, and delete operations.",
        modifySheetStructureParams,
      ),
      simpleOfficeTool(
        "modify_object",
        "Modify Excel Object",
        "Excel-only first-class object mutation tool for table/chart/pivot/worksheet object operations through native workbook actions.",
        modifyObjectParams,
      ),
      simpleOfficeTool(
        "get_all_objects",
        "Get Excel Objects",
        "Excel-only first-class object inventory read for workbook/worksheet tables, charts, PivotTables, and named items.",
        getAllObjectsParams,
      ),
      simpleOfficeTool(
        "search_data",
        "Search Excel Data",
        "Excel-only first-class workbook/worksheet data search across tables, charts, PivotTables, named items, and cited cells.",
        searchDataParams,
      ),
      simpleOfficeTool(
        "get_range_as_csv",
        "Export Range as CSV",
        "Excel-only first-class CSV export for auditable range snapshots. Use includeFormulas=true when formula-first verification is required.",
        getRangeAsCsvParams,
      ),
      simpleOfficeTool(
        "read_range_image",
        "Read Range Image",
        "Excel-only first-class range imagery read for visual verification workflows on the active worksheet selection/range.",
        readRangeImageParams,
      ),
      simpleOfficeTool(
        "extract_chart_xml",
        "Extract Chart XML",
        "Excel-only first-class chart XML extraction that returns a runtime-generated chart metadata XML snapshot (not full package OOXML).",
        extractChartXmlParams,
      ),
      simpleOfficeTool(
        "get_presentation_structure",
        "Read Presentation Structure",
        "PowerPoint-only first-class presentation structure read. Returns slide order plus layout/master structure metadata and bounded slide previews.",
        getPresentationStructureParams,
      ),
      simpleOfficeTool(
        "get_slide",
        "Read Slide",
        "PowerPoint-only first-class per-slide read. Resolve a slide by slideId/slideIndex (or selection) and return structured slide details.",
        getSlideParams,
      ),
      simpleOfficeTool(
        "list_slide_shapes",
        "List Slide Shapes",
        "PowerPoint-only first-class shape inventory read. Returns structured shape summaries for the resolved slide.",
        listSlideShapesParams,
      ),
      simpleOfficeTool(
        "modify_presentation_structure",
        "Modify Presentation Structure",
        "PowerPoint-only first-class structure mutation tool for slide create/move/reorder/delete/layout operations through native host actions.",
        modifyPresentationStructureParams,
      ),
      simpleOfficeTool(
        "duplicate_slide",
        "Duplicate Slide",
        "PowerPoint-only first-class slide duplication tool supporting one or multiple source slides and optional insertion target/formatting controls.",
        duplicateSlideParams,
      ),
      simpleOfficeTool(
        "insert_slide_element",
        "Insert Slide Element",
        "PowerPoint-only first-class element insertion tool for explicit shape/table/diagram/picture authoring through native actions.",
        insertSlideElementParams,
      ),
      simpleOfficeTool(
        "remove_slide_element",
        "Remove Slide Element",
        "PowerPoint-only first-class element removal tool for explicit shape/text removal operations with destructive-action policy support.",
        removeSlideElementParams,
      ),
      simpleOfficeTool(
        "edit_slide_text",
        "Edit Slide Text",
        "PowerPoint-only first-class text editing for slide shapes/selection. Use this instead of generic office_apply_edit when the intent is text-focused slide authoring.",
        editSlideTextParams,
      ),
      simpleOfficeTool(
        "edit_slide_xml",
        "Edit Slide XML",
        "PowerPoint-only first-class serialized/XML editing tool for slide notes and package-level OOXML workflows, including base64 import/export paths.",
        editSlideXmlParams,
      ),
      simpleOfficeTool(
        "edit_slide_master",
        "Edit Slide Layout/Master",
        "PowerPoint-only first-class layout/master editing tool that applies slide layouts via native layout/master resolution.",
        editSlideMasterParams,
      ),
      simpleOfficeTool(
        "edit_slide_chart",
        "Edit Slide Chart",
        "PowerPoint-only first-class chart workflow tool for chart inspection and serialized chart create/update paths.",
        editSlideChartParams,
      ),
      simpleOfficeTool(
        "copy_image_between_slides",
        "Copy Image Between Slides",
        "PowerPoint-only first-class media workflow tool that copies an image from a source slide/shape to a destination slide or shape.",
        copyImageBetweenSlidesParams,
      ),
      simpleOfficeTool(
        "search_icons",
        "Search Slide Icons",
        "PowerPoint-only first-class icon search. Returns icon matches from the supported runtime icon catalog without mutating slides.",
        searchIconsParams,
      ),
      simpleOfficeTool(
        "insert_icon",
        "Insert Slide Icon",
        "PowerPoint-only first-class icon insertion tool that inserts or updates an icon-like visual on the target slide.",
        insertIconParams,
      ),
      simpleOfficeTool(
        "verify_slides",
        "Verify Slides",
        "PowerPoint-only first-class structural verification. Returns a non-mutating summary/details payload for slide/layout/master checks.",
        verifySlidesParams,
      ),
      simpleOfficeTool(
        "verify_slide_visual",
        "Verify Slide Visual",
        "PowerPoint-only first-class visual verification using supported Office.js slide/shape snapshot paths (not slideshow-frame capture).",
        verifySlideVisualParams,
      ),
      simpleOfficeTool(
        "office_execute_js",
        "Execute Office.js",
        "Execute Office.js code as an escape hatch when structured tools are insufficient. This tool is a best-effort restricted subset enforced by regex checks (not an isolated sandbox) and blocks network, storage, eval, and system-access patterns.",
        executeJsParams,
      ),
      reviewableWordEditTool(
        "edit_doc_list",
        "Edit Word List",
        `Word-only first-class reviewable list editing. Keep each searchText under ${OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH} characters and include paragraphId/anchor locators when available.`,
      ),
      reviewableWordEditTool(
        "office_propose_edits",
        "Propose Document Edits",
        `Propose a batch of text edits for user review before applying changes. ` +
          `CRITICAL: each edit's searchText MUST be under ${OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH} characters. ` +
          "Split large paragraph rewrites into multiple small, targeted edits.",
      ),
      {
        name: "ask_user",
        label: "Ask User",
        description:
          "Ask the user one or more clarifying questions and continue execution once answers are received.",
        parameters: Type.Any(),
        execute: async (_toolCallId, params) => {
          const typed = params as {
            question?: string;
            context?: string;
            options?: Array<string | { title: string; description?: string }>;
            questions?: Array<{
              question: string;
              context?: string;
              options?: Array<string | { title: string; description?: string }>;
            }>;
          };

          const normalizeOptions = (raw?: Array<string | { title: string; description?: string }>) =>
            (raw ?? []).map((option) => (typeof option === "string" ? { title: option } : option));

          const questions: AskUserQuestion[] = [];
          if (typed.questions?.length) {
            for (const question of typed.questions) {
              questions.push({
                id: crypto.randomUUID(),
                question: question.question,
                context: question.context,
                options: normalizeOptions(question.options),
              });
            }
          } else if (typed.question) {
            questions.push({
              id: crypto.randomUUID(),
              question: typed.question,
              context: typed.context,
              options: normalizeOptions(typed.options),
            });
          } else {
            return {
              content: [{ type: "text", text: "No question provided." }],
              details: { error: "missing_question" },
            };
          }

          const request: AskUserRequest = { requestId: crypto.randomUUID(), questions };
          const response = await this.invokeAskUser(request);
          const lines = response.answers.map((answer) => {
            const question = questions.find((entry) => entry.id === answer.questionId);
            const notes = answer.notes ? ` (notes: ${answer.notes})` : "";
            return `Q: ${question?.question ?? "?"}\nA: ${answer.selectedOption ?? "None of the above"}${notes}`;
          });

          return {
            content: [{ type: "text", text: lines.join("\n\n") }],
            details: { questions, answers: response.answers },
          };
        },
      },
      {
        name: "generate_image",
        label: "Generate Image",
        description: "Generate an image from a text prompt and optionally insert it into the active Office document.",
        parameters: Type.Any(),
        execute: async (_toolCallId, params) => {
          const typed = params as {
            prompt?: string;
            aspectRatio?: string;
            size?: string;
            quality?: string;
            insert?: boolean;
          };
          const prompt = String(typed.prompt ?? "").trim();
          if (!prompt) {
            throw new Error("prompt is required.");
          }

          const prefs = this.getPreferences();
          if (!prefs.imageGenerationEnabled) {
            return {
              content: [
                {
                  type: "text",
                  text: "Image generation is disabled. Enable it in Settings > Preferences > Image Generation.",
                },
              ],
              details: { enabled: false },
            };
          }

          const configuredModelKey = prefs.defaultImageModel || IMAGE_MODEL_CATALOG[0]?.key;
          const [provider = "openai", modelId = "gpt-image-1"] = (configuredModelKey ?? "openai::gpt-image-1").split("::");
          if (provider !== "openai") {
            throw new Error(`Image generation for provider ${provider} is not available in browser-only mode yet.`);
          }

          const apiKey = this.authStore.getApiKey(provider);
          if (!apiKey) {
            throw new Error(`Provider ${provider} is not configured.`);
          }

          const size = toSize(typed.size, typed.aspectRatio) ?? "1024x1024";
          const response = await fetch("https://api.openai.com/v1/images/generations", {
            method: "POST",
            headers: {
              Authorization: `Bearer ${apiKey}`,
              "Content-Type": "application/json",
            },
            body: JSON.stringify({
              model: modelId,
              prompt,
              size,
              quality: typed.quality ?? "auto",
              response_format: "b64_json",
            }),
          });

          if (!response.ok) {
            const text = await response.text();
            throw new Error(text || `${response.status} ${response.statusText}`);
          }

          const payload = (await response.json()) as {
            data?: Array<{ b64_json?: string; revised_prompt?: string; url?: string }>;
          };
          const image = payload.data?.[0];
          let base64 = image?.b64_json;
          if (!base64 && image?.url) {
            const binary = await fetch(image.url);
            const buffer = await binary.arrayBuffer();
            base64 = arrayBufferToBase64(buffer);
          }

          if (!base64) {
            throw new Error("Image API did not return image data.");
          }

          const shouldInsert = typed.insert !== false;
          if (shouldInsert) {
            await this.invokeOfficeTool("office_apply_edit", {
              action: {
                type: "insertInlinePicture",
                content: base64,
                placement: "after",
                options: { altText: prompt.slice(0, 120) },
              },
            });
          }

          return {
            content: [
              {
                type: "text",
                text: `Image generated with ${modelId}.${shouldInsert ? " Inserted into document." : ""}`,
              },
              { type: "image", data: base64, mimeType: "image/png" },
            ],
            details: { provider, modelId, inserted: shouldInsert },
          };
        },
      },
    ];

    if (this.canUseCompanionFileTools()) {
      for (const toolName of ["read", "grep", "find", "ls"] as const) {
        tools.push({
          name: toolName,
          label: toolName === "ls" ? "List Local Files" : `Local ${toolName[0]!.toUpperCase()}${toolName.slice(1)}`,
          description:
            toolName === "read"
              ? "Read file contents from the saved document folder through the optional local companion."
              : toolName === "grep"
                ? "Search file contents in the saved document folder through the optional local companion."
                : toolName === "find"
                  ? "Find files by name inside the saved document folder through the optional local companion."
                  : "List directory contents inside the saved document folder through the optional local companion.",
          parameters: Type.Any(),
          execute: async (_toolCallId, params) => normalizeExternalToolResult(
            await this.executeCompanionFileTool(this.sessionId, toolName, normalizeToolParams(params)),
          ),
        });
      }
    }

    if (this.hasCompanionConnectorTools()) {
      tools.push({
        name: "mcp",
        label: "Local MCP Connector",
        description:
          "Execute a verified read-only local MCP tool through the optional companion. Use one of the exact tool names listed in the companion inventory.",
        parameters: Type.Object({
          toolName: Type.String({
            description: "Exact verified companion connector tool name to execute.",
          }),
          arguments: Type.Optional(
            Type.Any({
              description: "JSON arguments to pass to the selected companion connector tool.",
            }),
          ),
        }),
        execute: async (_toolCallId, params) => {
          const typed = normalizeToolParams(params);
          const toolName = String(typed.toolName ?? "").trim();
          if (!toolName) {
            throw new Error("toolName is required.");
          }

          return normalizeExternalToolResult(
            await this.executeCompanionMcpTool(
              this.sessionId,
              toolName,
              normalizeToolParams(typed.arguments),
            ),
          );
        },
      });
    }

    return tools;
  }

  private async invokeOfficeTool(toolName: OfficeToolName, params: Record<string, unknown>): Promise<unknown> {
    const result: OfficeToolResult = await executeOfficeTool({
      requestId: crypto.randomUUID(),
      toolName,
      host: this.officeState.host,
      params,
    });

    if (!result.success) {
      throw new Error(result.error ?? `${toolName} failed.`);
    }

    return result.content;
  }

  private async invokeAskUser(request: AskUserRequest): Promise<AskUserResponse> {
    return new Promise<AskUserResponse>((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.pendingAskUser.delete(request.requestId);
        reject(new Error("ask_user timed out waiting for user response."));
      }, 300_000);

      this.pendingAskUser.set(request.requestId, { resolve, reject, timeout });
      this.send({ type: "ask_user_request", request });
    });
  }

  private async invokeEditProposal(proposal: OfficeEditProposal): Promise<OfficeEditProposalDecision> {
    return new Promise<OfficeEditProposalDecision>((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.pendingEditProposals.delete(proposal.requestId);
        reject(new Error("Edit proposal timed out waiting for user review."));
      }, 300_000);

      this.pendingEditProposals.set(proposal.requestId, { resolve, reject, timeout });
      this.send({ type: "edit_proposal_request", proposal });
    });
  }

  private shouldAutoApproveTool(toolName: string): boolean {
    const preferences = this.getPreferences();
    const override = preferences.toolPermissionOverrides?.find((entry) => entry.toolName === toolName);
    if (override?.autoApproveAtLevel === "disabled") return false;

    const category = (TOOL_CATEGORY_MAP[toolName] ?? "connector") as ToolCategory;
    if (category === "interaction") return true;
    if (this.sessionApprovedTools.has(toolName)) return true;

    const effectiveLevel = override?.autoApproveAtLevel ?? preferences.autonomyLevel;
    return AUTONOMY_LEVEL_AUTO_APPROVE[effectiveLevel].has(category);
  }

  private async requestToolPermission(
    toolName: string,
    params: Record<string, unknown>,
  ): Promise<ToolPermissionDecision> {
    const category = (TOOL_CATEGORY_MAP[toolName] ?? "connector") as ToolCategory;
    const requestId = crypto.randomUUID();
    const expiresAt = new Date(Date.now() + 120_000).toISOString();
    const request: ToolPermissionRequest = {
      requestId,
      toolName,
      toolCategory: category,
      params,
      expiresAt,
    };

    return new Promise<ToolPermissionDecision>((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.pendingPermissions.delete(requestId);
        this.send({ type: "tool_permission_expired", requestId, toolName });
        resolve({ toolName, allowed: false, scope: "once" });
      }, 120_000);

      this.pendingPermissions.set(requestId, { resolve, reject, timeout });
      this.send({ type: "tool_permission_request", request });
    });
  }

  private rejectAllPending(error: Error): void {
    for (const [requestId, pending] of this.pendingAskUser.entries()) {
      clearTimeout(pending.timeout);
      pending.reject(error);
      this.pendingAskUser.delete(requestId);
    }
    for (const [requestId, pending] of this.pendingPermissions.entries()) {
      clearTimeout(pending.timeout);
      pending.reject(error);
      this.pendingPermissions.delete(requestId);
    }
    for (const [requestId, pending] of this.pendingEditProposals.entries()) {
      clearTimeout(pending.timeout);
      pending.reject(error);
      this.pendingEditProposals.delete(requestId);
    }
  }

  private send(message: BridgeServerMessage): void {
    for (const socket of this.bridgeSockets) {
      socket.emitServerMessage(message);
    }
  }
}

class InProcessKernel {
  private readonly authStore = new BrowserAuthStore();
  private readonly modelRegistry = new BrowserModelRegistry(this.authStore);
  private readonly checkpointStore = new BrowserCheckpointStore();
  private readonly connectorRuntime = new BrowserConnectorRuntime();
  private readonly companionClient = new CompanionClient();
  private readonly sessionsById = new Map<string, BrowserOfficeSession>();
  private readonly sessionsByDocument = new Map<string, BrowserOfficeSession>();
  private userPreferences: UserPreferences = { ...DEFAULT_USER_PREFERENCES };

  setPreferences(patch: Partial<UserPreferences>): { ok: true; preferences: UserPreferences } {
    this.userPreferences = { ...this.userPreferences, ...patch };
    return { ok: true, preferences: { ...this.userPreferences } };
  }

  connectBridge(sessionId: string): LocalBridgeSocket {
    const session = this.sessionsById.get(sessionId);
    if (!session) {
      throw new Error("Unknown session.");
    }
    return new LocalBridgeSocket(session);
  }

  private async getCompanionState(): Promise<CompanionState> {
    await this.companionClient.ensureInitialized();
    return this.companionClient.getState();
  }

  private getSessionCompanionConnectors(sessionId: string | undefined): ConnectorStatus[] | undefined {
    if (!sessionId) {
      return undefined;
    }
    return this.sessionsById.get(sessionId)?.getCompanionConnectors();
  }

  private addCompanionDiagnostics(
    diagnostics: ConnectorDiagnostic[],
    transport: ConnectorSetupRequest["transport"],
    companion: CompanionState,
  ): ConnectorDiagnostic[] {
    if (transport !== "local_stdio") {
      return [
        ...diagnostics,
        {
          level: "info",
          code: "remote_http_setup_only",
          title: "Remote connector setup-only",
          message:
            "Remote HTTP connectors can be configured in the taskpane, but agent execution is disabled until the browser remote-MCP execution path is implemented.",
        },
      ];
    }

    if (companion.status === "connected") {
      return [
        ...diagnostics,
        {
          level: "info",
          code: "local_stdio_companion_connected",
          title: "Optional companion connected",
          message: "Read-only local connector execution is available through the optional companion.",
        },
      ];
    }

    return [
      ...diagnostics,
      {
        level: "warning",
        code: "local_stdio_companion_unavailable",
        title: "Optional companion unavailable",
        message:
          "Local stdio connectors need the optional companion to verify and execute. Remote HTTP connectors are setup-only until browser remote-MCP execution is implemented.",
      },
    ];
  }

  private applyCompanionExecutionMetadata(
    status: ConnectorStatusResponse["connectors"][number],
    overlay?: ConnectorStatus | undefined,
  ): ConnectorStatus {
    const companion = this.companionClient.getState();
    const isLocalConnector = status.transport === "local_stdio";
    return {
      ...status,
      ...(overlay ?? {}),
      executionEnvironment: isLocalConnector
        ? "companion"
        : "browser",
      executionAvailable: overlay?.executionAvailable
        ?? (isLocalConnector
          ? companion.status === "connected"
          : false),
    };
  }

  private mergeConnectorStatuses(
    statuses: ConnectorStatusResponse["connectors"],
    overlayStatuses: ConnectorStatus[] | undefined,
  ): ConnectorStatusResponse {
    const overlayById = new Map((overlayStatuses ?? []).map((status) => [status.id, status]));
    return {
      connectors: statuses.map((status) => this.applyCompanionExecutionMetadata(status, overlayById.get(status.id))),
    };
  }

  private async prepareCompanionBinding(
    session: BrowserOfficeSession,
    officeState: OfficeStateUpdate,
    windowId?: string,
  ): Promise<CompanionSessionBinding | undefined> {
    const connectors = this.connectorRuntime.buildCompanionSessionConnectors({
      host: officeState.host,
      documentId: officeState.document.id,
      documentTitle: officeState.document.title,
      documentSaved: officeState.document.saved,
      documentUrl: officeState.document.documentUrl,
      workspaceId: officeState.document.workspaceDir,
    });
    return this.companionClient.openSession(session.sessionId, officeState, connectors, windowId);
  }

  private async syncSessionCompanion(
    session: BrowserOfficeSession,
    officeState: OfficeStateUpdate,
  ): Promise<CompanionState> {
    try {
      const binding = await this.prepareCompanionBinding(session, officeState, session.windowId);
      if (binding) {
        session.setCompanion(binding);
        return session.companion;
      }

      const companion = await this.getCompanionState();
      session.setCompanionState(disconnectedCompanionState(companion));
      return session.companion;
    } catch (error) {
      const companion = errorCompanionState(this.companionClient.getState(), error);
      session.setCompanionState(companion);
      return session.companion;
    }
  }

  private async probeLocalConnector(request: ConnectorSetupRequest): Promise<{
    ok: boolean;
    status: ConnectorStatus;
    diagnostics: ConnectorDiagnostic[];
  } | undefined> {
    const definition = this.connectorRuntime.buildCompanionConnectorDefinitionFromSetup(request);
    if (!definition) {
      return undefined;
    }
    return this.companionClient.probeConnector(definition);
  }

  async request<T>(path: string, init?: RequestInit): Promise<T> {
    await Promise.all([this.authStore.ready, this.connectorRuntime.ready]);
    const method = (init?.method ?? "GET").toUpperCase();
    const body = parseRequestBody(init);

    if (method === "GET" && path === "/v1/providers") {
      return this.modelRegistry.getProviderCatalog() as T;
    }
    if (method === "GET" && path === "/v1/auth/status") {
      return this.modelRegistry.getAuthStatus() as T;
    }
    if (method === "POST" && path === "/v1/auth/api-key") {
      const record = body as { provider?: string; apiKey?: string } | undefined;
      if (!record?.provider || !record.apiKey) {
        throw new Error("provider and apiKey are required.");
      }
      await this.authStore.setApiKey(record.provider, record.apiKey);
      return { ok: true } as T;
    }
    if (method === "POST" && path === "/v1/auth/start") {
      throw new Error("OAuth sign-in is unavailable in browser-only mode. Use API keys in Settings.");
    }

    const authDeleteMatch = method === "DELETE" ? path.match(/^\/v1\/auth\/([^/]+)$/) : null;
    if (authDeleteMatch) {
      await this.authStore.remove(decodeURIComponent(authDeleteMatch[1] ?? ""));
      return { ok: true } as T;
    }

    if (method === "GET" && path === "/v1/companion/state") {
      return (await this.getCompanionState()) as T;
    }
    if (method === "POST" && path === "/v1/companion/discover") {
      return (await this.companionClient.discover()) as T;
    }
    if (method === "POST" && path === "/v1/companion/manual-endpoint") {
      const request = body as { endpoint?: string } | undefined;
      return (await this.companionClient.setManualEndpoint(
        typeof request?.endpoint === "string" ? request.endpoint : undefined,
      )) as T;
    }

    if (method === "GET" && path === "/v1/connectors/catalog") {
      return this.connectorRuntime.getCatalogResponse() as T;
    }
    if (method === "POST" && path === "/v1/connectors/status") {
      const request = body as { scopeContext?: ConnectorScopeContext; sessionId?: string } | undefined;
      const companion = await this.getCompanionState();
      const overlayStatuses = companion.status === "connected"
        ? this.getSessionCompanionConnectors(request?.sessionId)
        : undefined;
      return this.mergeConnectorStatuses(
        this.connectorRuntime.getStatusResponse(request?.scopeContext).connectors,
        overlayStatuses,
      ) as T;
    }
    if (method === "GET" && path === "/v1/connectors/diagnostics") {
      const companion = await this.getCompanionState();
      const diagnostics = this.connectorRuntime.getDiagnostics();
      return {
        ...diagnostics,
        diagnostics: [
          ...diagnostics.diagnostics,
          companion.status === "connected"
            ? {
                level: "info",
                code: "optional_companion_connected",
                title: "Optional companion connected",
                message: `Read-only local file tools and local stdio MCP connectors are available through ${companion.endpoint}.`,
              }
            : {
                level: companion.status === "error" ? "warning" : "info",
                code: "optional_companion_unavailable",
                title: "Optional companion not connected",
                message:
                  "Local stdio connectors need the optional companion to verify and execute. Remote HTTP connectors are setup-only until browser remote-MCP execution is implemented.",
              },
        ],
      } as T;
    }
    if (method === "GET" && path === "/v1/connectors/audit") {
      return this.connectorRuntime.getAuditPreferenceResponse() as T;
    }
    if (method === "POST" && path === "/v1/connectors/audit") {
      return (await this.connectorRuntime.setAuditPreference(body as ConnectorAuditPreference)) as T;
    }
    if (method === "GET" && path === "/v1/connectors/export") {
      return this.connectorRuntime.getExportBundle() as T;
    }
    if (method === "POST" && path === "/v1/connectors/import/preview") {
      const request = body as ConnectorImportPreviewRequest;
      return this.connectorRuntime.previewImport(request) as T;
    }
    if (method === "POST" && path === "/v1/connectors/import/apply") {
      return (await this.connectorRuntime.applyImport(body as ConnectorImportApplyRequest)) as T;
    }
    if (method === "POST" && path === "/v1/connectors/setup/prepare") {
      const request = body as { connectorId?: string; scopeContext?: ConnectorScopeContext } | undefined;
      const connectorId = String(request?.connectorId ?? "");
      if (!connectorId) {
        throw new Error("connectorId is required.");
      }
      const companion = await this.getCompanionState();
      const response = this.connectorRuntime.prepareConnector(connectorId, request?.scopeContext);
      return {
        ...response,
        diagnostics: this.addCompanionDiagnostics(response.diagnostics, response.connector.transport, companion),
        executionEnvironment: response.connector.transport === "local_stdio" ? "companion" : "browser",
        executionAvailable: response.connector.transport === "local_stdio" ? companion.status === "connected" : false,
      } as T;
    }
    if (method === "POST" && path === "/v1/connectors/setup/connect") {
      const request = body as ConnectorSetupRequest;
      const companion = await this.getCompanionState();
      const response = await this.connectorRuntime.connectConnector(request);
      let probeDiagnostics: ConnectorDiagnostic[] = [];
      let probe = undefined as Awaited<ReturnType<InProcessKernel["probeLocalConnector"]>>;
      try {
        probe = await this.probeLocalConnector({ ...request, existingId: response.status.id });
      } catch (error) {
        probeDiagnostics = [
          {
            level: "warning",
            code: "companion_probe_failed",
            title: "Companion verification failed",
            message: error instanceof Error ? error.message : String(error),
            connectorId: response.status.connectorId,
          },
        ];
      }
      return {
        ...response,
        status: this.applyCompanionExecutionMetadata(response.status, probe?.status),
        diagnostics: this.addCompanionDiagnostics(
          [...response.diagnostics, ...probeDiagnostics, ...(probe?.diagnostics ?? [])],
          response.status.transport,
          companion,
        ),
      } as T;
    }
    if (method === "POST" && path === "/v1/connectors/setup/test") {
      const request = body as ConnectorSetupRequest;
      const companion = await this.getCompanionState();
      const response = this.connectorRuntime.testConnector(request);
      let probeDiagnostics: ConnectorDiagnostic[] = [];
      let probe = undefined as Awaited<ReturnType<InProcessKernel["probeLocalConnector"]>>;
      try {
        probe = await this.probeLocalConnector(request);
      } catch (error) {
        probeDiagnostics = [
          {
            level: "warning",
            code: "companion_probe_failed",
            title: "Companion verification failed",
            message: error instanceof Error ? error.message : String(error),
            connectorId: response.status.connectorId,
          },
        ];
      }
      return {
        ...response,
        status: this.applyCompanionExecutionMetadata(response.status, probe?.status),
        diagnostics: this.addCompanionDiagnostics(
          [...response.diagnostics, ...probeDiagnostics, ...(probe?.diagnostics ?? [])],
          response.status.transport,
          companion,
        ),
      } as T;
    }
    if (method === "POST" && path === "/v1/connectors/reverify") {
      const request = body as { connectorId?: string; scopeContext?: ConnectorScopeContext } | undefined;
      const connectorId = String(request?.connectorId ?? "");
      if (!connectorId) {
        throw new Error("connectorId is required.");
      }
      const companion = await this.getCompanionState();
      const response = await this.connectorRuntime.reverifyConnector(connectorId, request?.scopeContext);
      let probeDiagnostics: ConnectorDiagnostic[] = [];
      let probe: {
        ok: boolean;
        status: ConnectorStatus;
        diagnostics: ConnectorDiagnostic[];
      } | undefined;
      const definition = this.connectorRuntime.buildCompanionConnectorDefinition(connectorId);
      try {
        probe = definition ? await this.companionClient.probeConnector(definition) : undefined;
      } catch (error) {
        probeDiagnostics = [
          {
            level: "warning",
            code: "companion_probe_failed",
            title: "Companion verification failed",
            message: error instanceof Error ? error.message : String(error),
            connectorId: response.status.connectorId,
          },
        ];
      }
      return {
        ...response,
        status: this.applyCompanionExecutionMetadata(response.status, probe?.status),
        diagnostics: this.addCompanionDiagnostics(
          [...response.diagnostics, ...probeDiagnostics, ...(probe?.diagnostics ?? [])],
          response.status.transport,
          companion,
        ),
      } as T;
    }
    if (method === "POST" && path === "/v1/connectors/oauth/start") {
      const request = body as { connectorId?: string } | undefined;
      const connectorId = String(request?.connectorId ?? "");
      if (!connectorId) {
        throw new Error("connectorId is required.");
      }
      return (await this.connectorRuntime.startOAuth(connectorId) as ConnectorOAuthStartResponse) as T;
    }
    if (method === "POST" && path === "/v1/connectors/oauth/callback") {
      return (await this.connectorRuntime.completeOAuth(body as ConnectorOAuthCallbackRequest) as ConnectorOAuthCallbackResponse) as T;
    }
    if (method === "POST" && path === "/v1/connectors/favorite") {
      return (await this.connectorRuntime.setFavorite(body as ConnectorFavoriteRequest)) as T;
    }
    if (method === "POST" && path === "/v1/connectors/scope") {
      return (await this.connectorRuntime.updateScope(body as ConnectorScopeUpdateRequest)) as T;
    }
    const connectorDeleteMatch = method === "DELETE" ? path.match(/^\/v1\/connectors\/([^/]+)$/) : null;
    if (connectorDeleteMatch) {
      return (await this.connectorRuntime.removeConnector(decodeURIComponent(connectorDeleteMatch[1] ?? ""))) as T;
    }
    const connectorLogsMatch = method === "GET" ? path.match(/^\/v1\/connectors\/([^/]+)\/logs$/) : null;
    if (connectorLogsMatch) {
      return this.connectorRuntime.getLogs(decodeURIComponent(connectorLogsMatch[1] ?? "")) as T;
    }
    if (path.startsWith("/v1/connectors/")) {
      throw new Error(`Unknown connector route: ${method} ${path}`);
    }

    if (method === "POST" && path === "/v1/sessions/open") {
      const request = body as OfficeSessionOpenRequest;
      if (!request?.host || !request.documentId) {
        throw new Error("host and documentId are required.");
      }

      const documentKey = buildBrowserSessionKey(request.host, request.documentId, request.windowId);
      let session = this.sessionsByDocument.get(documentKey);
      const normalizedState = normalizeOpenState(request);

      if (request.forceNew && session) {
        this.sessionsByDocument.delete(documentKey);
        this.sessionsById.delete(session.sessionId);
        await session.dispose();
        session = undefined;
      }

      if (!session) {
        session = new BrowserOfficeSession(
          this.authStore,
          this.modelRegistry,
          this.checkpointStore,
          () => this.userPreferences,
          (sessionId, toolName, params) => this.companionClient.executeFileTool(sessionId, toolName, params),
          (sessionId, toolName, params) => this.companionClient.executeMcpTool(sessionId, toolName, params),
          request,
        );
        this.sessionsByDocument.set(documentKey, session);
        this.sessionsById.set(session.sessionId, session);
      } else {
        await session.updateOfficeState(normalizedState);
      }

      await this.syncSessionCompanion(session, normalizedState);
      return session.toOpenResponse() as T;
    }

    const officeStateMatch = method === "POST" ? path.match(/^\/v1\/sessions\/([^/]+)\/office-state$/) : null;
    if (officeStateMatch) {
      const session = this.getSession(officeStateMatch[1] ?? "");
      const officeState = body as OfficeStateUpdate;
      await session.updateOfficeState(officeState);
      await this.syncSessionCompanion(session, officeState);
      const response: OfficeSessionStateResponse = {
        ok: true,
        documentState: session.documentState,
        companion: session.companion,
      };
      return response as T;
    }

    const promptMatch = method === "POST" ? path.match(/^\/v1\/sessions\/([^/]+)\/(prompt|steer|follow-up)$/) : null;
    if (promptMatch) {
      const session = this.getSession(promptMatch[1] ?? "");
      const payload = body as { text?: string; images?: PromptImagePayload[] } | undefined;
      const text = String(payload?.text ?? "");
      if (!text.trim()) throw new Error("Prompt text is required.");
      const mode: PromptMode = promptMatch[2] === "follow-up" ? "followUp" : (promptMatch[2] as PromptMode);
      await session.prompt(text, mode, payload?.images);
      return { ok: true } as T;
    }

    const abortMatch = method === "POST" ? path.match(/^\/v1\/sessions\/([^/]+)\/abort$/) : null;
    if (abortMatch) {
      const session = this.getSession(abortMatch[1] ?? "");
      await session.abort();
      return { ok: true } as T;
    }

    const modelMatch = method === "POST" ? path.match(/^\/v1\/sessions\/([^/]+)\/model$/) : null;
    if (modelMatch) {
      const session = this.getSession(modelMatch[1] ?? "");
      const payload = body as SetModelRequest;
      await session.setModel(payload.provider, payload.modelId);
      return { ok: true } as T;
    }

    const statsMatch = method === "GET" ? path.match(/^\/v1\/sessions\/([^/]+)\/stats$/) : null;
    if (statsMatch) {
      const session = this.getSession(statsMatch[1] ?? "");
      return session.getStats() as T;
    }

    const promptSuggestionsMatch = method === "POST" ? path.match(/^\/v1\/sessions\/([^/]+)\/prompt-suggestions$/) : null;
    if (promptSuggestionsMatch) {
      const session = this.getSession(promptSuggestionsMatch[1] ?? "");
      const payload = body as PromptSuggestionRequest;
      return (await session.suggestPrompts(payload)) as T;
    }

    const thinkingMatch = method === "GET" ? path.match(/^\/v1\/sessions\/([^/]+)\/thinking$/) : null;
    if (thinkingMatch) {
      const session = this.getSession(thinkingMatch[1] ?? "");
      return session.getThinkingCapabilities() as T;
    }

    const thinkingLevelMatch = method === "POST" ? path.match(/^\/v1\/sessions\/([^/]+)\/thinking-level$/) : null;
    if (thinkingLevelMatch) {
      const session = this.getSession(thinkingLevelMatch[1] ?? "");
      const payload = body as { level?: ThinkingLevel };
      if (!payload?.level) throw new Error("level is required.");
      session.setThinkingLevel(payload.level);
      return { ok: true, level: payload.level } as T;
    }

    const deriveSubjectMatch = method === "POST" ? path.match(/^\/v1\/sessions\/([^/]+)\/derive-subject$/) : null;
    if (deriveSubjectMatch) {
      const session = this.getSession(deriveSubjectMatch[1] ?? "");
      const payload = body as DeriveSubjectRequest;
      if (!payload?.messages?.length) throw new Error("messages array is required.");
      const response: DeriveSubjectResponse = { subject: await session.deriveSubject(payload.messages) };
      return response as T;
    }

    if (method === "GET" && path === "/v1/preferences") {
      return { ...this.userPreferences } as T;
    }
    if (method === "POST" && path === "/v1/preferences") {
      const patch = body as Partial<UserPreferences>;
      return this.setPreferences(patch) as T;
    }
    if (method === "GET" && path === "/v1/image-models") {
      return this.modelRegistry.getImageCatalog() as T;
    }

    throw new Error(`Unknown route: ${method} ${path}`);
  }

  private getSession(sessionId: string): BrowserOfficeSession {
    const session = this.sessionsById.get(sessionId);
    if (!session) throw new Error("Unknown session.");
    return session;
  }
}

export class LocalBridgeSocket extends EventTarget {
  readyState: number = WebSocket.CONNECTING;

  constructor(private readonly session: BrowserOfficeSession) {
    super();
    this.session.attachBridge(this);
    queueMicrotask(() => {
      if (this.readyState !== WebSocket.CONNECTING) return;
      this.readyState = WebSocket.OPEN;
      this.dispatchEvent(new Event("open"));
    });
  }

  send(data: string): void {
    if (this.readyState !== WebSocket.OPEN && this.readyState !== WebSocket.CONNECTING) return;
    try {
      const payload = JSON.parse(data) as BridgeClientMessage;
      this.session.handleClientMessage(payload);
    } catch {
      this.dispatchEvent(new Event("error"));
    }
  }

  close(): void {
    if (this.readyState === WebSocket.CLOSED) return;
    this.readyState = WebSocket.CLOSED;
    this.session.detachBridge(this);
    this.dispatchEvent(new CloseEvent("close"));
  }

  emitServerMessage(payload: BridgeServerMessage): void {
    if (this.readyState === WebSocket.CLOSED) return;
    this.dispatchEvent(new MessageEvent("message", { data: JSON.stringify(payload) }));
  }
}

const kernel = new InProcessKernel();

export async function dispatchKernelRequest<T>(path: string, init?: RequestInit): Promise<T> {
  return kernel.request<T>(path, init);
}

export function createLocalBridgeSocket(sessionId: string): LocalBridgeSocket {
  return kernel.connectBridge(sessionId);
}

export function syncKernelPreferences(preferences: UserPreferences): void {
  kernel.setPreferences(preferences);
}
