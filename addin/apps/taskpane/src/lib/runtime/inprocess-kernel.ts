import { Agent, type AgentMessage, type AgentTool } from "@mariozechner/pi-agent-core";
import {
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
import { Type } from "@sinclair/typebox";
import {
  HOST_LABELS,
  OFFICE_APPEND_SYSTEM_PROMPT,
  composeAutonomyPrompt,
  composeOfficeAwarePrompt,
  getOfficeDocumentState,
} from "@pi-office/pi-office-pack/defaults";
import {
  getAvailableToolNames,
  getResolvedCapability,
  resolvePiOfficeCapabilities,
  type CapabilityResolution,
} from "@pi-office/pi-office-pack/capabilities";
import {
  AUTONOMY_LEVEL_AUTO_APPROVE,
  DEFAULT_USER_PREFERENCES,
  ARTIFACT_CLARIFICATION_MODES,
  COMPANION_RUNTIME_MODES,
  EDIT_REJECT_REASON_LABELS,
  OFFICE_TOOL_NAMES,
  TASKPANE_COMPANION_PROTOCOL,
  TOOL_CATEGORY_MAP,
  type AskUserQuestion,
  type AskUserRequest,
  type AskUserResponse,
  type AuthStatusResponse,
  type BridgeClientMessage,
  type BridgeServerMessage,
  type CheckpointMetadata,
  type CompanionConnectorDefinition,
  type CompanionNativeCaptureRequest,
  type CompanionNativeCaptureResponse,
  type CompanionProviderAuthStatusResponse,
  type CompanionRuntimeMode,
  type CompanionSettingsSyncRequest,
  type CompanionShellExecuteRequest,
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
  type ConnectorSetupProfile,
  type ConnectorSetupResponse,
  type ConnectorStatus,
  type ConnectorStatusResponse,
  type ConnectorTestResponse,
  type ConnectorToolPolicyUpdateRequest,
  type McpResultClearRequest,
  type McpResultPageRequest,
  type McpResultSummarizeRequest,
  type McpToolSearchRequest,
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
  type ProviderAuthDescriptor,
  type ProviderAuthMethod,
  type ProviderAuthState,
  type ProviderCatalogResponse,
  type ProviderDescriptor,
  type ProviderModelDescriptor,
  type ProviderRuntimeSurface,
  type ProviderSupportStatus,
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
import {
  SIMPLE_RECOMMENDED_MODELS_BY_PROVIDER,
  SIMPLE_VISIBLE_PROVIDERS,
  getProviderDefaultModel,
  getProviderModelPreference,
  getProviderSettingsPreference,
} from "@pi-office/pi-office-pack/provider-model-preferences";
import { executeOfficeTool } from "../office-tools";
import {
  getOfficeToolDefinition,
  getCoreOfficeToolDefinitionsForHost,
  getOfficeToolCapabilityDetail,
  officeToolSupportsHost,
  searchOfficeToolDefinitions,
  type OfficeToolDefinition,
} from "../office/tools/index.js";
import { isBrowserDebugOfficeState } from "../office/shared";
import { executeMcpBatchPlan, executeOfficeBatchPlan } from "./batch-executor";
import { CompanionClient, type CompanionSessionBinding } from "./companion-client";
import type { BrowserConnectorRuntime } from "./browser-connectors";

type JsonRecord = Record<string, unknown>;
type ToolContentPart = { type: "text"; text: string } | ImageContent;

const TOOL_RESULT_TEXT_MAX_CHARS = 24_000;
const OLD_TOOL_RESULT_TEXT_MAX_CHARS = 4_000;
const RECENT_TOOL_RESULTS_WITH_FULL_TEXT = 20;
const BINARY_TEXT_FIELD_PATTERN = /(^|[-_])(b64[-_]?json|base64|image[-_]?data|binary[-_]?data|screenshot[-_]?data)([-_]|$)/i;

const OFFICE_TOOL_NAME_SET = new Set<string>(OFFICE_TOOL_NAMES);

interface StoredAuthRecord {
  provider: string;
  apiKey: string;
  authState?: ProviderAuthState | undefined;
  savedAt?: string | undefined;
  verifiedAt?: string | undefined;
  lastVerificationAttemptAt?: string | undefined;
  lastVerificationError?: string | undefined;
}

interface PendingAskUser {
  resolve: (value: AskUserResponse) => void;
  reject: (error: Error) => void;
  timeout: ReturnType<typeof setTimeout>;
}

export interface ProviderModelSelectionState {
  enabledProviders: string[];
  enabledModels: string[];
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

type DebugLogDirection = "client" | "server" | "runtime";

interface ConversationDebugLogEvent {
  sequence: number;
  timestamp: string;
  direction: DebugLogDirection;
  type: string;
  payload: unknown;
}

interface ConversationDebugLogExport {
  version: 1;
  exportedAt: string;
  sessionId: string;
  documentKey: string;
  windowId?: string | undefined;
  documentState: OfficeDocumentState;
  officeState: unknown;
  companion: unknown;
  pendingRequests: {
    askUser: number;
    toolPermissions: number;
    editProposals: number;
  };
  sessionApprovedTools: string[];
  agent: unknown;
  stats: SessionStatsResponse;
  events: ConversationDebugLogEvent[];
  redaction: {
    secretFields: string;
    note: string;
  };
}

const AUTH_STORAGE_KEY = "pi-office-auth";
const AUTH_STORAGE_KEY_VERSION = 2;
const AUTH_CRYPTO_KEY_STORAGE_KEY = "pi-office-auth-key-v1";
const CONNECTOR_STORAGE_KEY = "pi-office-connectors";
const CHECKPOINT_STORAGE_KEY_PREFIX = "pi-office-checkpoints:";
const MAX_CHECKPOINT_AGE_MS = 3 * 24 * 60 * 60 * 1000;
const MAX_CHECKPOINTS_PER_DOC = 50;
const PROVIDER_AUTH_STATE_VALUES = new Set<ProviderAuthState>([
  "not_configured",
  "credential_stored",
  "verified_usable",
  "verification_failed",
]);
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

interface BrowserProviderCapability {
  supportStatus: ProviderSupportStatus;
  runtimeSurface: ProviderRuntimeSurface;
  authMethods: ProviderAuthMethod[];
  apiKeySupported: boolean;
  oauthSupported: boolean;
  browserCallable: boolean;
  companionRequired: boolean;
  subscriptionBacked: boolean;
  imageGenerationSupported: boolean;
  capabilityNote?: string | undefined;
}

const DEFAULT_PROVIDER_CAPABILITY: BrowserProviderCapability = {
  supportStatus: "research_only",
  runtimeSurface: "not_implemented",
  authMethods: [],
  apiKeySupported: false,
  oauthSupported: false,
  browserCallable: false,
  companionRequired: false,
  subscriptionBacked: false,
  imageGenerationSupported: false,
  capabilityNote: "Pi-Office has not classified this provider for browser taskpane execution yet.",
};

const BROWSER_API_KEY_CAPABILITY: BrowserProviderCapability = {
  supportStatus: "supported",
  runtimeSurface: "browser_taskpane",
  authMethods: ["api_key"],
  apiKeySupported: true,
  oauthSupported: false,
  browserCallable: true,
  companionRequired: false,
  subscriptionBacked: false,
  imageGenerationSupported: false,
};

const PROVIDER_CAPABILITIES: Record<string, BrowserProviderCapability> = {
  "amazon-bedrock": {
    supportStatus: "planned",
    runtimeSurface: "companion",
    authMethods: ["aws_credentials", "manual_token"],
    apiKeySupported: false,
    oauthSupported: false,
    browserCallable: false,
    companionRequired: true,
    subscriptionBacked: false,
    imageGenerationSupported: false,
    capabilityNote: "Amazon Bedrock needs AWS credential discovery or bearer-token handling in the companion, not browser localStorage.",
  },
  anthropic: {
    ...BROWSER_API_KEY_CAPABILITY,
    authMethods: ["api_key", "oauth"],
    capabilityNote: "Anthropic API keys work in the taskpane. Claude subscription OAuth is planned for companion-owned auth.",
  },
  "azure-openai-responses": {
    supportStatus: "planned",
    runtimeSurface: "companion",
    authMethods: ["api_key"],
    apiKeySupported: false,
    oauthSupported: false,
    browserCallable: false,
    companionRequired: true,
    subscriptionBacked: false,
    imageGenerationSupported: false,
    capabilityNote: "Azure OpenAI requires endpoint, deployment, and tenant-specific configuration before Pi-Office can call it honestly.",
  },
  cerebras: BROWSER_API_KEY_CAPABILITY,
  deepseek: {
    ...BROWSER_API_KEY_CAPABILITY,
    capabilityNote: "Direct DeepSeek API-key setup is available, but Pi-Office keeps it in Advanced settings so regional-provider use is intentional.",
  },
  fireworks: BROWSER_API_KEY_CAPABILITY,
  "github-copilot": {
    supportStatus: "planned",
    runtimeSurface: "companion",
    authMethods: ["oauth"],
    apiKeySupported: false,
    oauthSupported: false,
    browserCallable: false,
    companionRequired: true,
    subscriptionBacked: true,
    imageGenerationSupported: false,
    capabilityNote: "GitHub Copilot requires OAuth/subscription token brokerage; Pi-Office does not start that flow from the browser taskpane yet.",
  },
  google: BROWSER_API_KEY_CAPABILITY,
  "google-antigravity": {
    supportStatus: "planned",
    runtimeSurface: "companion",
    authMethods: ["oauth"],
    apiKeySupported: false,
    oauthSupported: false,
    browserCallable: false,
    companionRequired: true,
    subscriptionBacked: true,
    imageGenerationSupported: false,
    capabilityNote: "Antigravity uses Google OAuth and should be handled by the companion before it is offered as executable.",
  },
  "google-gemini-cli": {
    supportStatus: "planned",
    runtimeSurface: "companion",
    authMethods: ["oauth"],
    apiKeySupported: false,
    oauthSupported: false,
    browserCallable: false,
    companionRequired: true,
    subscriptionBacked: true,
    imageGenerationSupported: false,
    capabilityNote: "Gemini CLI / Cloud Code Assist uses Google OAuth and project state that the browser taskpane does not own.",
  },
  "google-vertex": {
    supportStatus: "planned",
    runtimeSurface: "companion",
    authMethods: ["api_key", "cloud_identity"],
    apiKeySupported: false,
    oauthSupported: false,
    browserCallable: false,
    companionRequired: true,
    subscriptionBacked: false,
    imageGenerationSupported: false,
    capabilityNote: "Vertex AI needs project, location, and ADC/API-key handling outside the current browser-only provider setup.",
  },
  groq: BROWSER_API_KEY_CAPABILITY,
  huggingface: BROWSER_API_KEY_CAPABILITY,
  "kimi-coding": BROWSER_API_KEY_CAPABILITY,
  minimax: BROWSER_API_KEY_CAPABILITY,
  "minimax-cn": BROWSER_API_KEY_CAPABILITY,
  mistral: BROWSER_API_KEY_CAPABILITY,
  openai: {
    ...BROWSER_API_KEY_CAPABILITY,
    imageGenerationSupported: true,
    capabilityNote: "OpenAI API keys work in the taskpane for chat and the current OpenAI-only image generation tools.",
  },
  "openai-codex": {
    supportStatus: "planned",
    runtimeSurface: "companion",
    authMethods: ["oauth"],
    apiKeySupported: false,
    oauthSupported: false,
    browserCallable: false,
    companionRequired: true,
    subscriptionBacked: true,
    imageGenerationSupported: false,
    capabilityNote: "OpenAI Codex models require ChatGPT subscription OAuth; Pi-Office needs companion token brokerage before exposing them.",
  },
  opencode: BROWSER_API_KEY_CAPABILITY,
  "opencode-go": BROWSER_API_KEY_CAPABILITY,
  openrouter: BROWSER_API_KEY_CAPABILITY,
  "vercel-ai-gateway": BROWSER_API_KEY_CAPABILITY,
  xai: BROWSER_API_KEY_CAPABILITY,
  zai: BROWSER_API_KEY_CAPABILITY,
};

function getProviderCapability(provider: string): BrowserProviderCapability {
  if (provider.startsWith("faux")) return BROWSER_API_KEY_CAPABILITY;
  return PROVIDER_CAPABILITIES[provider] ?? DEFAULT_PROVIDER_CAPABILITY;
}

function isBrowserProviderSupported(provider: string): boolean {
  return getProviderCapability(provider).browserCallable;
}

function isProviderConfigured(auth: ProviderAuthDescriptor, capability: BrowserProviderCapability): boolean {
  return capability.browserCallable && capability.apiKeySupported && auth.credentialStored;
}

function providerCapabilityFields(capability: BrowserProviderCapability) {
  return {
    supportStatus: capability.supportStatus,
    runtimeSurface: capability.runtimeSurface,
    authMethods: capability.authMethods,
    apiKeySupported: capability.apiKeySupported,
    browserCallable: capability.browserCallable,
    companionRequired: capability.companionRequired,
    subscriptionBacked: capability.subscriptionBacked,
    imageGenerationSupported: capability.imageGenerationSupported,
    ...(capability.capabilityNote ? { capabilityNote: capability.capabilityNote } : {}),
  };
}

function uniqueStrings(values: Iterable<string | undefined>): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const value of values) {
    if (!value || seen.has(value)) continue;
    seen.add(value);
    result.push(value);
  }
  return result;
}

const SIMPLE_RECOMMENDED_MODEL_LOOKUP = SIMPLE_RECOMMENDED_MODELS_BY_PROVIDER as Record<string, readonly string[]>;

function providerSettingsFields(providerId: string) {
  const settings = getProviderSettingsPreference(providerId);
  return {
    label: settings.label || titleCase(providerId),
    settingsVisibility: settings.settingsVisibility,
    lab: settings.lab || settings.label || titleCase(providerId),
    defaultModelId: settings.defaultModel,
  };
}

function providerModelPreferenceFields(providerId: string, modelId: string) {
  const providerSettings = getProviderSettingsPreference(providerId);
  const preference = getProviderModelPreference(providerId, modelId);
  const recommended = preference?.recommended ?? false;
  return {
    settingsVisibility: preference?.settingsVisibility ?? providerSettings.settingsVisibility,
    lab: preference?.lab ?? providerSettings.lab ?? titleCase(providerId),
    family: preference?.family,
    recommended,
    recommendationReason: preference?.recommendationReason,
    defaultForProvider: preference?.defaultForProvider ?? providerSettings.defaultModel === modelId,
    requiresUnrecommendedWarning: preference?.requiresUnrecommendedWarning ?? !recommended,
  };
}

type BrowserImageModelCatalogEntry = Omit<
  ImageModelDescriptor,
  "authState" | "credentialStored" | "verifiedUsable" | "verificationError" | "verifiedAt" | "configured"
> & { key: string };

const IMAGE_MODEL_CATALOG: BrowserImageModelCatalogEntry[] = [
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

function disconnectedCompanionCapabilities(endpoint?: string): CompanionState["capabilities"] {
  return {
    fileRead: false,
    localMcp: false,
    endpoint,
    version: "companion-capabilities-v1",
    protocol: TASKPANE_COMPANION_PROTOCOL,
    agent: {
      state: "unavailable",
      available: false,
      officeToolProxy: true,
      providerAuth: false,
      smartAuto: true,
      reason: "Optional companion is not connected.",
    },
    providerAuth: {
      state: "unavailable",
      available: false,
      explicitMigrationRequired: true,
      reason: "Optional companion is not connected.",
    },
    settingsSync: {
      state: "unavailable",
      available: false,
      secretsIncluded: false,
      reason: "Optional companion is not connected.",
    },
    nativeCapture: {
      state: "unavailable",
      available: false,
      hosts: [],
      trueViewportScreenshot: false,
      includeWindowFrame: false,
      reason: "Optional companion is not connected.",
    },
    mcp: {
      state: "unavailable",
      available: false,
      readOnly: true,
      toolCount: 0,
      reason: "Optional companion is not connected.",
    },
    memory: {
      state: "unavailable",
      available: false,
      reason: "Optional companion is not connected.",
    },
  };
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
    capabilities: disconnectedCompanionCapabilities(lastKnown?.capabilities?.endpoint ?? lastKnown?.endpoint),
  };
}

function unavailableCompanionProviderAuthStatus(providerIds: string[] = []): CompanionProviderAuthStatusResponse {
  const providerStates = providerIds
    .map((provider) => provider.trim())
    .filter(Boolean)
    .sort()
    .map((provider): ProviderAuthDescriptor => ({
      provider,
      state: "not_configured",
      credentialStored: false,
      verifiedUsable: false,
    }));
  return {
    ok: true,
    storageKind: "unavailable",
    secureStorage: false,
    explicitMigrationRequired: true,
    supportedAuthMethods: ["api_key"],
    storedProviders: [],
    configuredProviders: [],
    verifiedProviders: [],
    unverifiedProviders: [],
    verificationFailedProviders: [],
    providerStates,
  };
}

function companionRuntimeModeLabel(mode: CompanionRuntimeMode): string {
  if (mode === "basic") return "Basic";
  if (mode === "advanced") return "Advanced";
  return "Smart Auto";
}

function summarizeCompanionForPrompt(
  companion: CompanionState,
  documentSaved: boolean,
  runtimeMode: CompanionRuntimeMode,
): string {
  const lines = [
    `Companion status: ${companion.status}`,
    `Routing mode: ${companionRuntimeModeLabel(runtimeMode)}. Office.js document execution always remains in this taskpane.`,
  ];

  if (runtimeMode === "basic") {
    lines.push("Basic mode keeps companion-only tools hidden from model turns even if a companion is connected.");
  } else if (runtimeMode === "advanced") {
    lines.push("Advanced mode prefers companion-owned provider/auth/agent capability, with taskpane fallback until the companion advertises those surfaces.");
  } else {
    lines.push("Smart Auto prefers eligible non-Office companion capabilities only when the companion advertises them.");
  }

  if (companion.endpoint) {
    lines.push(`Companion endpoint: ${companion.endpoint}`);
  }
  if (companion.capabilities.protocol?.version) {
    lines.push(`Companion protocol: ${companion.capabilities.protocol.version}; Office.js execution remains ${companion.capabilities.protocol.advancedMode.officeJsExecutor}-owned.`);
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
    lines.push("MCP execution is available through the companion when configured read-only connectors are verified.");
  } else {
    lines.push("MCP execution is unavailable in this session.");
  }

  const agentCapability = companion.capabilities.agent;
  if (companion.status === "connected" && agentCapability?.state === "available") {
    lines.push("Companion-owned inference is available; Office tool calls must still be proxied back to the taskpane.");
  } else {
    lines.push("Inference is currently taskpane-owned unless companion provider auth is explicitly configured.");
  }

  const nativeCapture = companion.capabilities.nativeCapture;
  if (companion.status === "connected" && nativeCapture?.state === "available" && nativeCapture.hosts.length) {
    lines.push(`True viewport/window screenshot is available through companion native capture for: ${nativeCapture.hosts.join(", ")}.`);
  } else {
    lines.push("True viewport/window screenshot is unavailable; use Office.js snapshot/verification tools instead.");
  }

  const shellCapability = companion.capabilities.shell;
  if (companion.status === "connected" && shellCapability?.state === "available") {
    lines.push("Sandboxed shell is available through the companion for saved-document context; writes are limited to companion scratch.");
  } else if (companion.status === "connected" && shellCapability) {
    lines.push(`Sandboxed shell is ${shellCapability.state}: ${shellCapability.reason ?? "not enabled for this session"}`);
  } else {
    lines.push("Sandboxed shell is unavailable; do not ask for or attempt bash, PowerShell, cmd, edit, or write tools.");
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
    capabilities: disconnectedCompanionCapabilities(current.capabilities.endpoint ?? current.endpoint),
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

export function normalizeToolParams(params: unknown): Record<string, unknown> {
  if (typeof params === "string" && params.trim()) {
    try {
      const parsed = JSON.parse(params) as unknown;
      if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
        return parsed as Record<string, unknown>;
      }
    } catch {
      return {};
    }
    return {};
  }
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

function binaryPlaceholder(value: string): string {
  return `[base64 ${value.length} chars]`;
}

function shouldStripBinaryField(key: string, value: string, container: Record<string, unknown>): boolean {
  if (!value) return false;
  if (key === "data" && typeof container.mimeType === "string" && container.mimeType.startsWith("image/")) {
    return true;
  }
  return value.length > 512 && BINARY_TEXT_FIELD_PATTERN.test(key);
}

function stripBinaryData(value: unknown): unknown {
  if (!value || typeof value !== "object") return value;
  if (Array.isArray(value)) {
    return value.map((entry) => stripBinaryData(entry));
  }

  const next: Record<string, unknown> = { ...(value as JsonRecord) };
  for (const [key, entry] of Object.entries(next)) {
    if (typeof entry === "string" && shouldStripBinaryField(key, entry, next)) {
      next[key] = binaryPlaceholder(entry);
    }
  }
  if (Array.isArray(next.visuals)) {
    next.visuals = next.visuals.map((visual) => {
      if (!visual || typeof visual !== "object") return visual;
      const copy: JsonRecord = { ...(visual as JsonRecord) };
      if (typeof copy.data === "string") {
        copy.data = binaryPlaceholder(copy.data);
      }
      return copy;
    });
  }

  return Object.fromEntries(Object.entries(next).map(([key, entry]) => [key, stripBinaryData(entry)]));
}

const DEBUG_SECRET_KEY_PATTERN =
  /(^|[-_])(api[-_]?key|authorization|access[-_]?token|refresh[-_]?token|id[-_]?token|secret|password|client[-_]?secret|private[-_]?key|bearer|cookie|session[-_]?token)([-_]|$)/i;
const DEBUG_BEARER_PATTERN = /\bBearer\s+[A-Za-z0-9._~+/=-]+/g;
const DEBUG_API_KEY_PATTERN = /\b(sk|pk|ghp|github_pat|glpat|xox[baprs])-[-A-Za-z0-9_]{12,}\b/g;

function redactDebugString(value: string): string {
  return value
    .replace(DEBUG_BEARER_PATTERN, "Bearer [redacted]")
    .replace(DEBUG_API_KEY_PATTERN, "$1-[redacted]");
}

function sanitizeDebugPayload(value: unknown, seen = new WeakSet<object>()): unknown {
  if (value == null) return value;
  if (typeof value === "string") return redactDebugString(value);
  if (typeof value === "number" || typeof value === "boolean") return value;
  if (typeof value === "bigint") return value.toString();
  if (typeof value === "function") return `[function ${(value as { name?: string }).name || "anonymous"}]`;
  if (typeof value !== "object") return String(value);

  if (seen.has(value)) return "[circular]";
  seen.add(value);

  if (value instanceof Error) {
    return {
      name: value.name,
      message: redactDebugString(value.message),
      stack: value.stack ? redactDebugString(value.stack) : undefined,
    };
  }

  if (Array.isArray(value)) {
    return value.map((entry) => sanitizeDebugPayload(entry, seen));
  }

  const record = value as Record<string, unknown>;
  const sanitized: Record<string, unknown> = {};
  for (const [key, entry] of Object.entries(record)) {
    if (DEBUG_SECRET_KEY_PATTERN.test(key)) {
      sanitized[key] = entry == null || typeof entry === "boolean" || typeof entry === "number"
        ? entry
        : "[redacted]";
      continue;
    }
    sanitized[key] = sanitizeDebugPayload(entry, seen);
  }
  return sanitized;
}

function toToolText(value: unknown): string {
  if (typeof value === "string") return value;
  if (isOfficeContextPayload(value)) {
    const structured = JSON.stringify(stripBinaryData(value), null, 2);
    return `${value.summary}\n\nStructured data:\n${structured}`;
  }
  return JSON.stringify(stripBinaryData(value), null, 2);
}

function toToolContent(value: unknown): ToolContentPart[] {
  const content: ToolContentPart[] = [
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
  content: ToolContentPart[];
  details: unknown;
} {
  if (value && typeof value === "object" && Array.isArray((value as { content?: unknown }).content)) {
    return {
      content: (value as {
        content: ToolContentPart[];
      }).content,
      details: (value as { details?: unknown }).details ?? value,
    };
  }

  return {
    content: toToolContent(value),
    details: value,
  };
}

function truncateToolText(text: string, maxChars: number): string {
  if (text.length <= maxChars) return text;
  return `${text.slice(0, maxChars)}\n\n[truncated ${text.length - maxChars} chars from tool result for context safety]`;
}

function compactToolContentForContext(
  toolName: string,
  content: ToolContentPart[],
  options: { maxTextChars?: number; keepImages?: boolean } = {},
): ToolContentPart[] {
  const maxTextChars = options.maxTextChars ?? TOOL_RESULT_TEXT_MAX_CHARS;
  const keepImages = options.keepImages ?? toolName === "generate_image";
  const next: ToolContentPart[] = [];
  const omittedImages: string[] = [];

  for (const part of content) {
    if (part.type === "text") {
      next.push({ ...part, text: truncateToolText(part.text, maxTextChars) });
      continue;
    }

    if (keepImages) {
      next.push(part);
    } else {
      omittedImages.push(`${part.mimeType || "image"} ${part.data.length} base64 chars`);
    }
  }

  if (omittedImages.length) {
    next.push({
      type: "text",
      text:
        `[${omittedImages.length} image payload${omittedImages.length === 1 ? "" : "s"} omitted from model context: ` +
        `${omittedImages.join(", ")}. Use verify_doc_visual, native page metadata, or a fresh viewport capture when visual evidence is needed.]`,
    });
  }

  return next.length ? next : [{ type: "text", text: "[Tool result omitted from model context.]" }];
}

function compactAgentMessagesForContext(messages: AgentMessage[]): AgentMessage[] {
  const toolResultIndexes = messages
    .map((message, index) => (message.role === "toolResult" ? index : -1))
    .filter((index) => index >= 0);
  const fullTextStart = toolResultIndexes[Math.max(0, toolResultIndexes.length - RECENT_TOOL_RESULTS_WITH_FULL_TEXT)] ?? 0;

  return messages.map((message, index) => {
    if (message.role !== "toolResult") {
      return message;
    }

    const maxTextChars = index >= fullTextStart ? TOOL_RESULT_TEXT_MAX_CHARS : OLD_TOOL_RESULT_TEXT_MAX_CHARS;
    return {
      ...message,
      content: compactToolContentForContext(message.toolName, message.content as ToolContentPart[], { maxTextChars }),
      details: stripBinaryData(message.details),
    };
  });
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

function normalizeProviderAuthState(record: StoredAuthRecord | undefined): ProviderAuthState {
  if (!record?.apiKey) return "not_configured";
  return record.authState && PROVIDER_AUTH_STATE_VALUES.has(record.authState)
    ? record.authState
    : "credential_stored";
}

function toProviderAuthDescriptor(provider: string, record: StoredAuthRecord | undefined): ProviderAuthDescriptor {
  const state = normalizeProviderAuthState(record);
  const credentialStored = state !== "not_configured";
  const verifiedUsable = state === "verified_usable";
  return {
    provider,
    state,
    credentialStored,
    verifiedUsable,
    ...(record?.verifiedAt ? { verifiedAt: record.verifiedAt } : {}),
    ...(record?.lastVerificationAttemptAt ? { lastVerificationAttemptAt: record.lastVerificationAttemptAt } : {}),
    ...(state === "verification_failed" && record?.lastVerificationError
      ? { lastVerificationError: record.lastVerificationError }
      : {}),
  };
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function isProviderAuthFailure(error: unknown): boolean {
  const message = getErrorMessage(error).toLowerCase();
  const status = typeof error === "object" && error
    ? (error as { status?: unknown; statusCode?: unknown; code?: unknown }).status
      ?? (error as { statusCode?: unknown }).statusCode
      ?? (error as { code?: unknown }).code
    : undefined;
  if (status === 401 || status === 403 || status === "401" || status === "403") return true;
  return /\b(401|403)\b/.test(message)
    || message.includes("unauthorized")
    || message.includes("forbidden")
    || message.includes("invalid api key")
    || message.includes("incorrect api key")
    || message.includes("invalid_api_key")
    || message.includes("authentication");
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

  getAuthState(provider: string): ProviderAuthDescriptor {
    return toProviderAuthDescriptor(provider, this.store.get(provider));
  }

  listAuthStates(providerIds?: string[]): ProviderAuthDescriptor[] {
    const ids = providerIds?.length ? providerIds : this.list();
    return ids
      .map((provider) => this.getAuthState(provider))
      .sort((left, right) => left.provider.localeCompare(right.provider));
  }

  async setApiKey(provider: string, apiKey: string): Promise<void> {
    const now = new Date().toISOString();
    this.store.set(provider, {
      provider,
      apiKey,
      authState: "credential_stored",
      savedAt: now,
    });
    await this.persist();
  }

  async markVerificationSuccess(provider: string): Promise<void> {
    const record = this.store.get(provider);
    if (!record?.apiKey) return;
    const now = new Date().toISOString();
    this.store.set(provider, {
      ...record,
      authState: "verified_usable",
      verifiedAt: now,
      lastVerificationAttemptAt: now,
      lastVerificationError: undefined,
    });
    await this.persist();
  }

  async markVerificationFailure(provider: string, error: unknown): Promise<void> {
    const record = this.store.get(provider);
    if (!record?.apiKey) return;
    this.store.set(provider, {
      ...record,
      authState: "verification_failed",
      lastVerificationAttemptAt: new Date().toISOString(),
      lastVerificationError: getErrorMessage(error).slice(0, 300),
    });
    await this.persist();
  }

  async remove(provider: string): Promise<void> {
    this.store.delete(provider);
    await this.persist();
  }

  async clearAll(): Promise<void> {
    this.store.clear();
    try {
      localStorage.removeItem(AUTH_STORAGE_KEY);
      localStorage.removeItem(AUTH_CRYPTO_KEY_STORAGE_KEY);
    } catch {
      // Ignore storage errors in browser sandbox.
    }
  }

  private async load(): Promise<void> {
    try {
      const raw = localStorage.getItem(AUTH_STORAGE_KEY);
      if (!raw) return;

      const parsed = JSON.parse(raw) as unknown;
      if (Array.isArray(parsed)) {
        for (const entry of parsed as StoredAuthRecord[]) {
          if (!entry?.provider || !entry.apiKey) continue;
          this.store.set(entry.provider, this.normalizeRecord(entry));
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
        this.store.set(entry.provider, this.normalizeRecord(entry));
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

  private normalizeRecord(entry: StoredAuthRecord): StoredAuthRecord {
    return {
      provider: String(entry.provider),
      apiKey: String(entry.apiKey),
      authState: normalizeProviderAuthState(entry),
      ...(entry.savedAt ? { savedAt: entry.savedAt } : {}),
      ...(entry.verifiedAt ? { verifiedAt: entry.verifiedAt } : {}),
      ...(entry.lastVerificationAttemptAt ? { lastVerificationAttemptAt: entry.lastVerificationAttemptAt } : {}),
      ...(entry.lastVerificationError ? { lastVerificationError: entry.lastVerificationError } : {}),
    };
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
    const models = this.getExecutableModelsForProvider(provider);
    return models.find((model) => model.id === modelId);
  }

  getProviderCatalog(): ProviderCatalogResponse {
    const providers: ProviderDescriptor[] = [];
    for (const provider of getProviders()) {
      const providerId = String(provider);
      const capability = getProviderCapability(providerId);
      const providerSettings = providerSettingsFields(providerId);
      const models = this.getCatalogModelsForProvider(provider);
      if (!models.length) continue;
      const auth = this.authStore.getAuthState(providerId);
      const configured = isProviderConfigured(auth, capability);
      const descriptors: ProviderModelDescriptor[] = models
        .map((model) => {
          const totalCostPer1k = (model.cost.input + model.cost.output) / 2;
          const costTier: "$" | "$$" | "$$$" = totalCostPer1k <= 1 ? "$" : totalCostPer1k <= 10 ? "$$" : "$$$";
          const preference = providerModelPreferenceFields(providerId, model.id);
          return {
            provider: providerId,
            providerLabel: providerSettings.label,
            modelId: model.id,
            modelName: model.name,
            settingsVisibility: preference.settingsVisibility,
            lab: preference.lab,
            family: preference.family,
            recommended: preference.recommended,
            recommendationReason: preference.recommendationReason,
            defaultForProvider: preference.defaultForProvider,
            requiresUnrecommendedWarning: preference.requiresUnrecommendedWarning,
            ...providerCapabilityFields(capability),
            authState: auth.state,
            credentialStored: auth.credentialStored,
            verifiedUsable: auth.verifiedUsable,
            verificationError: auth.lastVerificationError,
            verifiedAt: auth.verifiedAt,
            configured,
            oauthSupported: capability.oauthSupported,
            usesApiKey: capability.apiKeySupported,
            contextWindow: model.contextWindow,
            costTier,
            supportsThinking: model.reasoning,
            supportsReasoningEffort: model.reasoning,
          };
        })
        .sort((left, right) => left.modelName.localeCompare(right.modelName));

      providers.push({
        provider: providerId,
        label: providerSettings.label,
        settingsVisibility: providerSettings.settingsVisibility,
        lab: providerSettings.lab,
        defaultModelId: providerSettings.defaultModelId,
        ...providerCapabilityFields(capability),
        authState: auth.state,
        credentialStored: auth.credentialStored,
        verifiedUsable: auth.verifiedUsable,
        verificationError: auth.lastVerificationError,
        verifiedAt: auth.verifiedAt,
        configured,
        oauthSupported: capability.oauthSupported,
        models: descriptors,
      });
    }

    providers.sort((left, right) => left.label.localeCompare(right.label));
    return { providers };
  }

  getAuthStatus(): AuthStatusResponse {
    const providerIds = getProviders().map((provider) => String(provider));
    const providerStates = this.authStore.listAuthStates(providerIds);
    const storedProviders = providerStates.filter((entry) => entry.credentialStored).map((entry) => entry.provider);
    const verifiedProviders = providerStates.filter((entry) => entry.verifiedUsable).map((entry) => entry.provider);
    const configuredProviders = providerStates
      .filter((entry) => entry.verifiedUsable && isBrowserProviderSupported(entry.provider))
      .map((entry) => entry.provider);
    return {
      storedProviders,
      oauthProviders: [],
      configuredProviders,
      verifiedProviders,
      unverifiedProviders: providerStates
        .filter((entry) => entry.state === "credential_stored")
        .map((entry) => entry.provider),
      verificationFailedProviders: providerStates
        .filter((entry) => entry.state === "verification_failed")
        .map((entry) => entry.provider),
      providerStates,
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

  getPreferredModel(preferences?: Pick<UserPreferences, "defaultModelByProvider">): Model<any> | undefined {
    const providerOrder = uniqueStrings([
      ...Object.keys(preferences?.defaultModelByProvider ?? {}),
      ...this.authStore.list(),
      ...SIMPLE_VISIBLE_PROVIDERS,
      ...getProviders().map((provider) => String(provider)),
    ]);

    for (const providerId of providerOrder) {
      if (!isBrowserProviderSupported(providerId)) continue;
      const models = this.getExecutableModelsForProvider(providerId);
      if (!models.length) continue;

      const modelOrder = uniqueStrings([
        preferences?.defaultModelByProvider?.[providerId],
        getProviderDefaultModel(providerId),
        ...(SIMPLE_RECOMMENDED_MODEL_LOOKUP[providerId] ?? []),
        ...models.map((model) => model.id),
      ]);
      const preferred = modelOrder
        .map((modelId) => models.find((model) => model.id === modelId))
        .find(Boolean);
      if (preferred) return preferred;
    }

    const fallbackProvider = getProviders().find((provider) => isBrowserProviderSupported(String(provider)));
    if (!fallbackProvider) return undefined;
    return this.getExecutableModelsForProvider(String(fallbackProvider))[0];
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
      ...(() => {
        const auth = this.authStore.getAuthState(entry.provider);
        return {
          authState: auth.state,
          credentialStored: auth.credentialStored,
          verifiedUsable: auth.verifiedUsable,
          verificationError: auth.lastVerificationError,
          verifiedAt: auth.verifiedAt,
          configured: auth.credentialStored,
        };
      })(),
      supportsReasoningEffort: entry.supportsReasoningEffort,
      supportedAspectRatios: entry.supportedAspectRatios,
      supportedSizes: entry.supportedSizes,
    }));

    return {
      models,
      defaultModelKey: IMAGE_MODEL_CATALOG[0]?.key ?? "",
    };
  }

  hasImageModelKey(modelKey: string): boolean {
    return IMAGE_MODEL_CATALOG.some((entry) => entry.key === modelKey);
  }

  hasCatalogModel(provider: string, modelId: string): boolean {
    return this.getCatalogModelsForProvider(provider).some((model) => model.id === modelId);
  }

  private getCatalogModelsForProvider(provider: string): Model<any>[] {
    try {
      return getModels(provider as never) as Model<any>[];
    } catch {
      return [];
    }
  }

  private getExecutableModelsForProvider(provider: string): Model<any>[] {
    if (!isBrowserProviderSupported(provider)) return [];
    return this.getCatalogModelsForProvider(provider);
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
  private readonly debugEvents: ConversationDebugLogEvent[] = [];
  private debugEventSequence = 0;
  private readonly agent: Agent;
  private unsubscribeAgent: (() => void) | undefined;

  constructor(
    private readonly authStore: BrowserAuthStore,
    private readonly modelRegistry: BrowserModelRegistry,
    private readonly checkpointStore: BrowserCheckpointStore,
    private readonly getPreferences: () => UserPreferences,
    private readonly executeCompanionFileTool: (sessionId: string, toolName: "read" | "grep" | "find" | "ls", params: Record<string, unknown>) => Promise<unknown>,
    private readonly executeCompanionMcpTool: (sessionId: string, toolName: string, params: Record<string, unknown>) => Promise<unknown>,
    private readonly getBrowserMcpToolNames: (scopeContext: ConnectorScopeContext) => string[],
    private readonly searchBrowserMcpTools: (request: McpToolSearchRequest, scopeContext: ConnectorScopeContext) => unknown,
    private readonly searchCompanionMcpTools: (sessionId: string, request: McpToolSearchRequest) => Promise<unknown>,
    private readonly executeBrowserMcpTool: (toolName: string, params: Record<string, unknown>, scopeContext: ConnectorScopeContext) => Promise<unknown>,
    private readonly getBrowserMcpResultPage: (request: { handleId: string; page?: number | undefined }) => unknown,
    private readonly summarizeBrowserMcpResult: (request: { handleId: string; query?: string | undefined; maxChars?: number | undefined }) => unknown,
    private readonly clearBrowserMcpResults: (request: { handleId?: string | undefined }) => unknown,
    private readonly getCompanionMcpResultPage: (sessionId: string, request: McpResultPageRequest) => Promise<unknown>,
    private readonly summarizeCompanionMcpResult: (sessionId: string, request: McpResultSummarizeRequest) => Promise<unknown>,
    private readonly clearCompanionMcpResults: (sessionId: string, request: McpResultClearRequest) => Promise<unknown>,
    private readonly executeCompanionShellCommand: (sessionId: string, request: CompanionShellExecuteRequest) => Promise<unknown>,
    private readonly executeCompanionNativeCapture: (sessionId: string, request: CompanionNativeCaptureRequest) => Promise<CompanionNativeCaptureResponse>,
    request: OfficeSessionOpenRequest,
  ) {
    this.windowId = request.windowId;
    this.documentKey = buildBrowserSessionKey(request.host, request.documentId, request.windowId);
    this.officeState = normalizeOpenState(request);

    this.agent = new Agent({
      getApiKey: (provider) => this.modelRegistry.getApiKey(provider),
      toolExecution: "sequential",
      transformContext: async (messages) => compactAgentMessagesForContext(messages),
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
      afterToolCall: async (ctx) => ({
        content: compactToolContentForContext(ctx.toolCall.name, ctx.result.content as ToolContentPart[]),
        details: stripBinaryData(ctx.result.details),
      }),
    });

    const tools = this.buildTools();
    this.agent.setTools(tools);
    this.agent.setSystemPrompt(this.buildSystemPrompt(tools.map((tool) => tool.name)));

    const preferredModel = this.modelRegistry.getPreferredModel(this.getPreferences());
    if (preferredModel) {
      this.agent.setModel(preferredModel);
    }

    const defaultThinkingLevel = this.getPreferences().defaultThinkingLevel;
    this.agent.setThinkingLevel(this.normalizeThinkingLevel(defaultThinkingLevel, this.agent.state.model) as PiThinkingLevel);

    this.unsubscribeAgent = this.agent.subscribe((event) => {
      this.send({ type: "session_event", event });
    });

    this.recordDebugEvent("runtime", {
      type: "session_initialized",
      host: request.host,
      documentId: request.documentId,
      saved: request.saved,
      title: request.title,
      model: this.agent.state.model
        ? {
            provider: String(this.agent.state.model.provider),
            id: this.agent.state.model.id,
            name: this.agent.state.model.name,
          }
        : undefined,
      thinkingLevel: this.agent.state.thinkingLevel,
      toolNames: this.agent.state.tools.map((tool) => tool.name),
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
    this.recordDebugEvent("runtime", {
      type: "companion_state_updated",
      companion: this.companionState,
      connectorCount: this.companionConnectors.length,
      connectorToolNames: this.companionState.connectorToolNames ?? [],
      toolNames: tools.map((tool) => tool.name),
    });
  }

  getCompanionConnectors(): ConnectorStatus[] {
    return [...this.companionConnectors];
  }

  getCapabilities(): CapabilityResolution[] {
    return this.resolveCapabilities();
  }

  exportDebugLog(): ConversationDebugLogExport {
    return {
      version: 1,
      exportedAt: nowIso(),
      sessionId: this.sessionId,
      documentKey: this.documentKey,
      windowId: this.windowId,
      documentState: this.documentState,
      officeState: sanitizeDebugPayload(this.officeState),
      companion: sanitizeDebugPayload(this.companionState),
      pendingRequests: {
        askUser: this.pendingAskUser.size,
        toolPermissions: this.pendingPermissions.size,
        editProposals: this.pendingEditProposals.size,
      },
      sessionApprovedTools: [...this.sessionApprovedTools].sort(),
      agent: sanitizeDebugPayload({
        model: this.agent.state.model
          ? {
              provider: String(this.agent.state.model.provider),
              id: this.agent.state.model.id,
              name: this.agent.state.model.name,
              contextWindow: this.agent.state.model.contextWindow,
              reasoning: this.agent.state.model.reasoning,
            }
          : undefined,
        thinkingLevel: this.agent.state.thinkingLevel,
        isStreaming: this.agent.state.isStreaming,
        systemPrompt: this.agent.state.systemPrompt,
        tools: this.agent.state.tools.map((tool) => ({
          name: tool.name,
          description: tool.description,
          parameters: tool.parameters,
        })),
        messages: this.agent.state.messages,
      }),
      stats: this.getStats(),
      events: this.debugEvents.map((entry) => ({ ...entry })),
      redaction: {
        secretFields: DEBUG_SECRET_KEY_PATTERN.source,
        note: "Known secret-bearing fields and bearer/API-key shaped strings are redacted. Prompts, document snippets, reasoning deltas, tool arguments, and tool results are intentionally included for sideload debugging.",
      },
    };
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
    this.recordDebugEvent("runtime", {
      type: "office_state_updated",
      host: next.host,
      document: next.document,
      selection: next.selection,
      toolNames: tools.map((tool) => tool.name),
    });
  }

  refreshPreferences(): void {
    const tools = this.buildTools();
    this.agent.setTools(tools);
    this.agent.setSystemPrompt(this.buildSystemPrompt(tools.map((tool) => tool.name)));
    this.recordDebugEvent("runtime", {
      type: "preferences_updated",
      companionRuntimeMode: this.getPreferences().companionRuntimeMode,
      autonomyLevel: this.getPreferences().autonomyLevel,
      artifactClarificationMode: this.getPreferences().artifactClarificationMode,
      toolNames: tools.map((tool) => tool.name),
    });
  }

  async prompt(text: string, mode: PromptMode = "prompt", images?: PromptImagePayload[]): Promise<void> {
    this.recordDebugEvent("runtime", {
      type: "prompt_request",
      mode,
      text,
      images,
    });
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
      await this.runWithCurrentModelVerification(() => this.agent.prompt([message]));
      return;
    }

    if (mode === "followUp") {
      if (this.agent.state.isStreaming) {
        this.agent.followUp(message);
        return;
      }
      await this.runWithCurrentModelVerification(() => this.agent.prompt([message]));
      return;
    }

    await this.runWithCurrentModelVerification(() => this.agent.prompt(input, piImages));
  }

  async abort(): Promise<void> {
    this.recordDebugEvent("runtime", { type: "abort_requested" });
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
    this.recordDebugEvent("runtime", {
      type: "model_changed",
      provider,
      modelId,
      modelName: model.name,
      thinkingLevel: this.agent.state.thinkingLevel,
    });
  }

  setThinkingLevel(level: ThinkingLevel): void {
    const normalized = this.normalizeThinkingLevel(level, this.agent.state.model);
    this.agent.setThinkingLevel(normalized as PiThinkingLevel);
    this.recordDebugEvent("runtime", {
      type: "thinking_level_changed",
      requestedLevel: level,
      activeLevel: normalized,
    });
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

    const toTokenEstimate = (chars: number) => Math.ceil(chars / 4);
    const systemPromptTokens = toTokenEstimate(this.agent.state.systemPrompt.length);
    let coreToolDefTokens = 0;
    let integrationToolDefTokens = 0;
    for (const tool of this.agent.state.tools) {
      const chars = tool.name.length
        + (tool.description ?? "").length
        + JSON.stringify(tool.parameters ?? {}).length;
      if (CORE_TOOL_NAMES.has(tool.name)) {
        coreToolDefTokens += toTokenEstimate(chars);
      } else {
        integrationToolDefTokens += toTokenEstimate(chars);
      }
    }

    let messageTokens = 0;
    let toolCallTokens = 0;
    let toolResultTokens = 0;
    let imageTokens = 0;
    const contextMessages = compactAgentMessagesForContext(messages);
    for (const message of contextMessages) {
      if (message.role === "user") {
        if (typeof message.content === "string") {
          messageTokens += toTokenEstimate(message.content.length);
        } else if (Array.isArray(message.content)) {
          for (const part of message.content) {
            if (part.type === "text") {
              messageTokens += toTokenEstimate(part.text.length);
            } else if (part.type === "image") {
              imageTokens += toTokenEstimate(part.data.length);
            }
          }
        }
      } else if (message.role === "assistant") {
        for (const part of message.content) {
          if (part.type === "text") {
            messageTokens += toTokenEstimate(part.text.length);
          } else if (part.type === "thinking") {
            messageTokens += toTokenEstimate(part.thinking.length + (part.thinkingSignature?.length ?? 0));
          } else if (part.type === "toolCall") {
            toolCallTokens += toTokenEstimate((part.name?.length ?? 0) + JSON.stringify(part.arguments ?? {}).length);
          }
        }
      } else if (message.role === "toolResult") {
        for (const part of message.content) {
          if (part.type === "text") {
            toolResultTokens += toTokenEstimate(part.text.length);
          } else if (part.type === "image") {
            imageTokens += toTokenEstimate(part.data.length);
          }
        }
      }
    }

    const contextTokens =
      systemPromptTokens +
      coreToolDefTokens +
      integrationToolDefTokens +
      messageTokens +
      toolCallTokens +
      toolResultTokens +
      imageTokens;
    const percent = contextTokens && contextWindow ? (contextTokens / contextWindow) * 100 : null;
    const available = Math.max(0, contextWindow - contextTokens);
    const breakdown: ContextBreakdownEntry[] = [
      { label: "System Prompt", tokens: systemPromptTokens, color: "#3b82f6" },
      { label: "Tool Definitions", tokens: coreToolDefTokens, color: "#a855f7" },
      ...(integrationToolDefTokens > 10 ? [{ label: "Integrations", tokens: integrationToolDefTokens, color: "#ec4899" }] : []),
      { label: "Messages", tokens: messageTokens, color: "#10b981" },
      { label: "Tool Calls", tokens: toolCallTokens, color: "#6366f1" },
      { label: "Tool Results", tokens: toolResultTokens, color: "#f59e0b" },
      ...(imageTokens > 0 ? [{ label: "Images", tokens: imageTokens, color: "#06b6d4" }] : []),
      { label: "Available", tokens: available, color: "#e5e7eb" },
    ];

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

    try {
      const result = await completeSimple(model, context, {
        ...(auth.apiKey ? { apiKey: auth.apiKey } : {}),
        ...(auth.headers ? { headers: auth.headers } : {}),
      });
      await this.authStore.markVerificationSuccess(String(model.provider));
      const text = result.content.find((part) => part.type === "text");
      if (!text || text.type !== "text") throw new Error("No text in subject response.");
      return text.text.trim();
    } catch (error) {
      await this.markProviderVerificationFailureIfAuthError(String(model.provider), error);
      throw error;
    }
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
      await this.authStore.markVerificationSuccess(String(model.provider));
      const rawText = result.content
        .filter((part): part is { type: "text"; text: string } => part.type === "text")
        .map((part) => part.text)
        .join("\n")
        .trim();

      return {
        generationId,
        suggestions: parsePromptSuggestions(rawText, { documentState: this.documentState }),
      };
    } catch (error) {
      await this.markProviderVerificationFailureIfAuthError(String(this.agent.state.model.provider), error);
      return emptyResponse;
    }
  }

  handleClientMessage(message: BridgeClientMessage): void {
    this.recordDebugEvent("client", message);

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
    this.recordDebugEvent("runtime", { type: "session_disposed" });
    this.rejectAllPending(new Error("Session disposed."));
    this.agent.abort();
    this.unsubscribeAgent?.();
    this.unsubscribeAgent = undefined;
    for (const socket of this.bridgeSockets) {
      socket.close();
    }
    this.bridgeSockets.clear();
  }

  private resolveCapabilities(): CapabilityResolution[] {
    return resolvePiOfficeCapabilities({
      host: this.officeState.host,
      documentSaved: this.officeState.document.saved,
      companion: this.companionState,
      companionRuntimeMode: this.getPreferences().companionRuntimeMode,
    });
  }

  private isCapabilityAvailable(id: Parameters<typeof getResolvedCapability>[1]): boolean {
    return getResolvedCapability(this.resolveCapabilities(), id).available;
  }

  private isToolAvailableByName(toolName: string): boolean {
    return getAvailableToolNames(this.resolveCapabilities()).has(toolName);
  }

  private isCompanionMcpResultHandle(handleId: string): boolean {
    return handleId.startsWith("mcp-result-companion-");
  }

  private async getMcpResultPage(request: McpResultPageRequest): Promise<unknown> {
    if (this.isCompanionMcpResultHandle(request.handleId)) {
      return this.getCompanionMcpResultPage(this.sessionId, request);
    }
    return this.getBrowserMcpResultPage(request);
  }

  private async summarizeMcpResult(request: McpResultSummarizeRequest): Promise<unknown> {
    if (this.isCompanionMcpResultHandle(request.handleId)) {
      return this.summarizeCompanionMcpResult(this.sessionId, request);
    }
    return this.summarizeBrowserMcpResult(request);
  }

  private async clearMcpResultHandles(request: McpResultClearRequest = {}): Promise<unknown> {
    if (request.handleId) {
      if (this.isCompanionMcpResultHandle(request.handleId)) {
        return this.clearCompanionMcpResults(this.sessionId, request);
      }
      return this.clearBrowserMcpResults(request);
    }

    const [browser, companion] = await Promise.all([
      Promise.resolve(this.clearBrowserMcpResults(request)),
      this.companionState.status === "connected"
        ? this.clearCompanionMcpResults(this.sessionId, request).catch((error) => ({
            ok: false,
            error: error instanceof Error ? error.message : String(error),
          }))
        : Promise.resolve({ ok: true, cleared: 0 }),
    ]);
    return { browser, companion };
  }

  private canUseCompanionFileTools(): boolean {
    return this.isCapabilityAvailable("local_files");
  }

  private getCompanionConnectorToolNames(): string[] {
    return [...(this.companionState.connectorToolNames ?? [])];
  }

  private connectorScopeContext(): ConnectorScopeContext {
    return {
      host: this.officeState.host,
      documentId: this.officeState.document.id,
      documentTitle: this.officeState.document.title,
      documentSaved: this.officeState.document.saved,
      documentUrl: this.officeState.document.documentUrl ?? this.officeState.document.documentPath,
      workspaceId: this.officeState.document.workspaceDir,
    };
  }

  private getBrowserConnectorToolNames(): string[] {
    return this.getBrowserMcpToolNames(this.connectorScopeContext());
  }

  private hasMcpConnectorTools(): boolean {
    return (
      (this.isCapabilityAvailable("mcp_connectors") && this.getCompanionConnectorToolNames().length > 0) ||
      this.getBrowserConnectorToolNames().length > 0
    );
  }

  private canUseCompanionShellTools(): boolean {
    return this.isCapabilityAvailable("shell_sandbox");
  }

  private canUseCompanionNativeCapture(): boolean {
    return this.isCapabilityAvailable("native_viewport_capture");
  }

  private buildSystemPrompt(availableToolNames: readonly string[]): string {
    const browserDebugGuidance = this.isBrowserDebugMode()
      ? "\n\nBrowser preview mode: no real Office host is attached. Do not call Office tools or claim document state. Use this session only for taskpane UI, settings, provider, and non-document debugging."
      : "";
    return `${OFFICE_APPEND_SYSTEM_PROMPT}\n\n${composeAutonomyPrompt(
      this.getPreferences(),
      this.canUseCompanionFileTools(),
      availableToolNames,
    )}\n\n${composeOfficeAwarePrompt(
      "Prefer Office tools as the source of truth for the active document.",
      this.officeState,
    )}\n\n${summarizeCompanionForPrompt(
      this.companionState,
      this.officeState.document.saved,
      this.getPreferences().companionRuntimeMode,
    )}${browserDebugGuidance}`;
  }

  private isBrowserDebugMode(): boolean {
    return isBrowserDebugOfficeState(this.officeState);
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

  private async runWithCurrentModelVerification<T>(operation: () => Promise<T>): Promise<T> {
    const provider = String(this.agent.state.model.provider);
    try {
      const result = await operation();
      await this.authStore.markVerificationSuccess(provider);
      return result;
    } catch (error) {
      await this.markProviderVerificationFailureIfAuthError(provider, error);
      throw error;
    }
  }

  private async markProviderVerificationFailureIfAuthError(provider: string, error: unknown): Promise<void> {
    if (isProviderAuthFailure(error)) {
      await this.authStore.markVerificationFailure(provider, error);
    }
  }

  private createOfficeAgentTool(definition: OfficeToolDefinition): AgentTool {
    const base = {
      name: definition.name,
      label: definition.label,
      description: definition.description,
      parameters: definition.parameters,
    };

    if (definition.executor === "reviewable-word-edit") {
      return {
        ...base,
        execute: async (_toolCallId, params) => {
          const result = await this.invokeOfficeTool(definition.name, normalizeToolParams(params));
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
      };
    }

    if (definition.executor === "companion-native-capture") {
      throw new Error(`${definition.name} must be registered through the companion native capture path.`);
    }

    return {
      ...base,
      execute: async (_toolCallId, params) => {
        const result = await this.invokeOfficeTool(definition.name, normalizeToolParams(params));
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    };
  }

  private createRuntimeAgentTools(): AgentTool[] {
    return [
      {
        name: "office_tool_call",
        label: "Call Discovered Office Tool",
        description:
          "Execute a structured Office tool discovered through office_tool_search / office_tool_get. This dispatcher cannot run office_execute_js and is permissioned as document-write capable.",
        parameters: Type.Object({
          toolName: Type.String({
            description: "Exact Office tool name returned by office_tool_search / office_tool_get.",
          }),
          arguments: Type.Optional(Type.Any({
            description: "JSON arguments matching the discovered Office tool schema.",
          })),
        }),
        execute: async (_toolCallId, params) => {
          const typed = normalizeToolParams(params);
          const toolName = String(typed.toolName ?? "").trim() as OfficeToolName;
          if (!toolName) {
            throw new Error("office_tool_call requires toolName.");
          }
          if (toolName === "office_tool_call") {
            throw new Error("office_tool_call cannot call itself.");
          }
          if (toolName === "office_execute_js") {
            throw new Error("office_execute_js must be called directly as a manual one-time escape hatch.");
          }
          const definition = getOfficeToolDefinition(toolName);
          if (definition.executor === "runtime-registry" || definition.executor === "companion-native-capture") {
            throw new Error(`${toolName} is not executable through office_tool_call.`);
          }
          if (!officeToolSupportsHost(definition, this.officeState.host)) {
            throw new Error(`${toolName} is not available for ${this.officeState.host}.`);
          }
          if (!this.isToolAvailableByName(toolName)) {
            throw new Error(`${toolName} is not available in the current Pi-Office capability state.`);
          }
          const tool = this.createOfficeAgentTool(definition);
          return tool.execute(_toolCallId, normalizeToolParams(typed.arguments));
        },
      },
      {
        name: "office_batch_execute",
        label: "Execute Office Batch Plan",
        description:
          "Execute a constrained, typed batch plan over approved Office tools. This is not arbitrary code execution and cannot run office_execute_js.",
        parameters: Type.Any(),
        execute: async (_toolCallId, params) => {
          const result = await executeOfficeBatchPlan({
            host: this.officeState.host,
            request: normalizeToolParams(params) as never,
            invokeOfficeTool: (toolName, toolParams) => this.invokeOfficeTool(toolName, toolParams),
          });
          return {
            content: toToolContent(result),
            details: result,
          };
        },
      },
      {
        name: "mcp_batch_execute",
        label: "Execute MCP Batch Plan",
        description:
          "Execute a constrained, typed batch plan over enabled MCP connector tools. Connector policy is checked per call.",
        parameters: Type.Any(),
        execute: async (_toolCallId, params) => {
          const result = await executeMcpBatchPlan({
            request: normalizeToolParams(params) as never,
            searchMcpTools: (request) => this.searchMcpTools(request),
            invokeMcpTool: (toolName, toolParams) => this.invokeMcpTool(toolName, toolParams),
          });
          return {
            content: toToolContent(result),
            details: result,
          };
        },
      },
      {
        name: "mcp_result_get",
        label: "Get MCP Result Page",
        description: "Retrieve a page from a cached browser-direct or companion-routed MCP result handle.",
        parameters: Type.Any(),
        execute: async (_toolCallId, params) => {
          const typed = normalizeToolParams(params);
          const result = await this.getMcpResultPage({
            handleId: String(typed.handleId ?? ""),
            page: typeof typed.page === "number" ? typed.page : undefined,
          });
          return normalizeExternalToolResult(result);
        },
      },
      {
        name: "mcp_result_summarize",
        label: "Summarize MCP Result",
        description: "Return a compact extract from a cached MCP result handle.",
        parameters: Type.Any(),
        execute: async (_toolCallId, params) => {
          const typed = normalizeToolParams(params);
          const result = await this.summarizeMcpResult({
            handleId: String(typed.handleId ?? ""),
            query: typeof typed.query === "string" ? typed.query : undefined,
            maxChars: typeof typed.maxChars === "number" ? typed.maxChars : undefined,
          });
          return normalizeExternalToolResult(result);
        },
      },
      {
        name: "mcp_result_clear",
        label: "Clear MCP Result Cache",
        description: "Clear one cached MCP result handle or all browser-direct and companion-routed MCP result handles.",
        parameters: Type.Any(),
        execute: async (_toolCallId, params) => {
          const typed = normalizeToolParams(params);
          const result = await this.clearMcpResultHandles({
            handleId: typeof typed.handleId === "string" ? typed.handleId : undefined,
          });
          return normalizeExternalToolResult(result);
        },
      },
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
            const error = new Error(text || `${response.status} ${response.statusText}`);
            if (response.status === 401 || response.status === 403) {
              await this.authStore.markVerificationFailure(provider, error);
            }
            throw error;
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
          await this.authStore.markVerificationSuccess(provider);

          const shouldInsert = typed.insert !== false;
          if (shouldInsert) {
            await this.invokeOfficeTool("office_apply_edit", {
              action: {
                type: "insertInlinePicture",
                content: base64,
                placement: "after",
                options: {
                  altText: prompt.slice(0, 120),
                  assetSourceId: "generated-image-base64",
                  assetMimeType: "image/png",
                  generatedImagePrompt: prompt,
                  generatedImageModel: modelId,
                },
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
  }

  private buildTools(): AgentTool[] {
    const availableToolNames = getAvailableToolNames(this.resolveCapabilities());
    const isToolAvailable = (toolName: string) => availableToolNames.has(toolName);
    const browserDebugMode = this.isBrowserDebugMode();
    const tools: AgentTool[] = getCoreOfficeToolDefinitionsForHost(this.officeState.host)
      .filter((definition) => definition.executor !== "companion-native-capture")
      .filter((definition) => definition.executor !== "runtime-registry")
      .filter((definition) => isToolAvailable(definition.name))
      .map((definition) => this.createOfficeAgentTool(definition))
      .filter((tool) => !browserDebugMode || !OFFICE_TOOL_NAME_SET.has(tool.name));

    tools.push(...this.createRuntimeAgentTools().filter((tool) => isToolAvailable(tool.name)));

    if (this.canUseCompanionNativeCapture()) {
      const definition = getOfficeToolDefinition("office_capture_viewport");
      tools.push({
        name: definition.name,
        label: definition.label,
        description: definition.description,
        parameters: definition.parameters,
        execute: async (_toolCallId, params) => {
          const typed = normalizeToolParams(params);
          const response = await this.executeCompanionNativeCapture(this.sessionId, {
            host: this.officeState.host,
            includeWindowFrame: typed.includeWindowFrame !== false,
          });
          if (!response.ok || !response.visual) {
            throw new Error(response.error ?? "Companion native capture unavailable.");
          }
          return normalizeExternalToolResult({
            summary: `Captured ${this.officeState.host} ${response.visual.kind} screenshot through the companion native capture backend.`,
            visuals: [response.visual],
            details: response.details,
          });
        },
      });
    }

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

    if (this.hasMcpConnectorTools()) {
      tools.push({
        name: "mcp",
        label: "MCP Connector",
        description:
          "Execute a verified read-only MCP tool through a browser-direct or companion connector. Use one of the exact tool names listed in the connector inventory.",
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
          if (this.getBrowserConnectorToolNames().includes(toolName)) {
            return normalizeExternalToolResult(
              await this.executeBrowserMcpTool(
                toolName,
                normalizeToolParams(typed.arguments),
                this.connectorScopeContext(),
              ),
            );
          }
          if (!this.getCompanionConnectorToolNames().includes(toolName)) {
            throw new Error(`Connector tool "${toolName}" is not enabled.`);
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

    if (this.canUseCompanionShellTools()) {
      tools.push({
        name: "bash",
        label: "Sandboxed Shell",
        description:
          "Execute a command through the optional companion shell sandbox. It is unavailable unless the companion reports an isolation backend and destructive probes have passed. Read-only commands can inspect the saved document folder; any writes must use the companion scratch directory only. Network, package installs, git push/commit, secret reads, and paths outside the approved roots are denied.",
        parameters: Type.Object({
          command: Type.String({
            description: "Shell command to execute through the companion sandbox.",
          }),
          cwd: Type.Optional(Type.String({
            description: "Working directory inside the saved document folder or companion scratch directory.",
          })),
          category: Type.Optional(Type.Union([
            Type.Literal("read-only"),
            Type.Literal("scratch-write"),
          ])),
          timeoutMs: Type.Optional(Type.Number({
            description: "Requested timeout in milliseconds, capped by the companion policy.",
          })),
        }),
        execute: async (_toolCallId, params) => {
          const typed = normalizeToolParams(params);
          const command = String(typed.command ?? "").trim();
          if (!command) {
            throw new Error("command is required.");
          }

          return normalizeExternalToolResult(
            await this.executeCompanionShellCommand(this.sessionId, {
              command,
              cwd: typeof typed.cwd === "string" ? typed.cwd : undefined,
              category: typed.category === "scratch-write" ? "scratch-write" : "read-only",
              timeoutMs: typeof typed.timeoutMs === "number" ? typed.timeoutMs : undefined,
            }),
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

  private async searchMcpTools(request: { query?: string; connectorId?: string; limit?: number }): Promise<unknown> {
    const rawBrowserResults = this.searchBrowserMcpTools(request, this.connectorScopeContext()) as unknown;
    const browserResults = Array.isArray(rawBrowserResults) ? rawBrowserResults : [];
    const rawCompanionResults = this.companionState.status === "connected"
      ? await this.searchCompanionMcpTools(this.sessionId, request).catch(() => ({ results: [] }))
      : { results: [] };
    const companionResults = Array.isArray((rawCompanionResults as { results?: unknown }).results)
      ? (rawCompanionResults as { results: unknown[] }).results
      : [];
    return {
      query: request.query,
      results: [...browserResults, ...companionResults],
    };
  }

  private async invokeMcpTool(toolName: string, args: Record<string, unknown>): Promise<unknown> {
    if (this.getBrowserConnectorToolNames().includes(toolName)) {
      return this.executeBrowserMcpTool(toolName, args, this.connectorScopeContext());
    }
    if (!this.getCompanionConnectorToolNames().includes(toolName)) {
      throw new Error(`Connector tool "${toolName}" is not enabled.`);
    }
    return this.executeCompanionMcpTool(this.sessionId, toolName, args);
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

  private recordDebugEvent(direction: DebugLogDirection, payload: unknown): void {
    const record = payload && typeof payload === "object" ? payload as { type?: unknown } : undefined;
    const type = typeof record?.type === "string" ? record.type : direction;
    this.debugEventSequence += 1;
    this.debugEvents.push({
      sequence: this.debugEventSequence,
      timestamp: nowIso(),
      direction,
      type,
      payload: sanitizeDebugPayload(payload),
    });
  }

  private send(message: BridgeServerMessage): void {
    this.recordDebugEvent("server", message);
    for (const socket of this.bridgeSockets) {
      socket.emitServerMessage(message);
    }
  }
}

class InProcessKernel {
  private readonly authStore = new BrowserAuthStore();
  private readonly modelRegistry = new BrowserModelRegistry(this.authStore);
  private readonly checkpointStore = new BrowserCheckpointStore();
  private connectorRuntime: BrowserConnectorRuntime | undefined;
  private connectorRuntimePromise: Promise<BrowserConnectorRuntime> | undefined;
  private readonly companionClient = new CompanionClient();
  private readonly sessionsById = new Map<string, BrowserOfficeSession>();
  private readonly sessionsByDocument = new Map<string, BrowserOfficeSession>();
  private userPreferences: UserPreferences = { ...DEFAULT_USER_PREFERENCES };
  private providerModelSelection: ProviderModelSelectionState = {
    enabledProviders: [],
    enabledModels: [],
  };

  private async getConnectorRuntime(): Promise<BrowserConnectorRuntime> {
    if (this.connectorRuntime) {
      await this.connectorRuntime.ready;
      return this.connectorRuntime;
    }
    this.connectorRuntimePromise ??= import("./browser-connectors").then((module) => {
      const runtime = new module.BrowserConnectorRuntime();
      this.connectorRuntime = runtime;
      return runtime;
    });
    const runtime = await this.connectorRuntimePromise;
    await runtime.ready;
    return runtime;
  }

  private hasStoredConnectorState(): boolean {
    try {
      return Boolean(localStorage.getItem(CONNECTOR_STORAGE_KEY));
    } catch {
      return false;
    }
  }

  setPreferences(patch: Partial<UserPreferences>): { ok: true; preferences: UserPreferences } {
    if (
      typeof patch.defaultImageModel === "string" &&
      patch.defaultImageModel &&
      !this.modelRegistry.hasImageModelKey(patch.defaultImageModel)
    ) {
      throw new Error(
        `Image model ${patch.defaultImageModel} is not available in the browser taskpane image catalog.`,
      );
    }
    if (patch.defaultModelByProvider !== undefined) {
      if (
        !patch.defaultModelByProvider ||
        typeof patch.defaultModelByProvider !== "object" ||
        Array.isArray(patch.defaultModelByProvider)
      ) {
        throw new Error("defaultModelByProvider must be an object keyed by provider.");
      }
      for (const [provider, modelId] of Object.entries(patch.defaultModelByProvider)) {
        if (!modelId) continue;
        if (!this.modelRegistry.hasCatalogModel(provider, modelId)) {
          throw new Error(`Model ${provider}/${modelId} is not available in the Pi provider catalog.`);
        }
      }
    }
    if (
      patch.artifactClarificationMode !== undefined &&
      !ARTIFACT_CLARIFICATION_MODES.includes(patch.artifactClarificationMode)
    ) {
      throw new Error("artifactClarificationMode must be balanced, draft_now, or ask_first.");
    }
    if (
      patch.companionRuntimeMode !== undefined &&
      !COMPANION_RUNTIME_MODES.includes(patch.companionRuntimeMode)
    ) {
      throw new Error("companionRuntimeMode must be basic, smart_auto, or advanced.");
    }
    this.userPreferences = { ...this.userPreferences, ...patch };
    for (const session of this.sessionsById.values()) {
      session.refreshPreferences();
      void this.syncCompanionSettingsForSession(session);
    }
    return { ok: true, preferences: { ...this.userPreferences } };
  }

  setProviderModelSelection(selection: Partial<ProviderModelSelectionState>): { ok: true; selection: ProviderModelSelectionState } {
    const enabledProviders = Array.isArray(selection.enabledProviders)
      ? selection.enabledProviders.filter((entry): entry is string => typeof entry === "string" && entry.trim().length > 0)
      : this.providerModelSelection.enabledProviders;
    const enabledModels = Array.isArray(selection.enabledModels)
      ? selection.enabledModels.filter((entry): entry is string => typeof entry === "string" && entry.trim().length > 0)
      : this.providerModelSelection.enabledModels;
    this.providerModelSelection = {
      enabledProviders: Array.from(new Set(enabledProviders)).sort(),
      enabledModels: Array.from(new Set(enabledModels)).sort(),
    };
    for (const session of this.sessionsById.values()) {
      void this.syncCompanionSettingsForSession(session);
    }
    return { ok: true, selection: { ...this.providerModelSelection } };
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
    requiresCompanion = transport === "local_stdio",
    oauthBrokerRequired = false,
  ): ConnectorDiagnostic[] {
    if (!requiresCompanion) {
      return diagnostics;
    }
    if (companion.status === "connected") {
      const isLocal = transport === "local_stdio";
      return [
        ...diagnostics,
        {
          level: "info",
          code: isLocal ? "local_stdio_companion_connected" : "remote_http_companion_connected",
          title: oauthBrokerRequired ? "Companion connected" : "Optional companion connected",
          message: oauthBrokerRequired
            ? "System-browser sign-in and connector execution will run through the local companion."
            : isLocal
            ? "Read-only local connector execution is available through the optional companion."
            : "Read-only remote MCP connector execution is available through the optional companion after verification.",
        },
      ];
    }

    return [
      ...diagnostics,
      {
        level: "warning",
        code: transport === "local_stdio" ? "local_stdio_companion_unavailable" : "remote_http_companion_unavailable",
        title: oauthBrokerRequired ? "Companion required" : "Optional companion unavailable",
        message: oauthBrokerRequired
          ? "This connector uses the local companion for secure system-browser sign-in, callback handling, and token exchange."
          : transport === "local_stdio"
          ? "Local stdio connectors need the optional companion to verify and execute."
          : "Remote HTTP MCP connectors need the optional companion to verify and execute.",
      },
    ];
  }

  private setupProfileUsesCompanionOAuthBroker(profile: ConnectorSetupProfile | undefined): boolean {
    return profile?.authMethod === "oauth" && profile.oauth?.broker === "companion";
  }

  private setupProfileIsHostedHttp(profile: ConnectorSetupProfile | undefined): boolean {
    if (profile?.transport !== "remote_http" || profile.availability === "needs_companion") {
      return false;
    }
    const endpoint = profile.endpoint ?? "";
    try {
      const host = new URL(endpoint).hostname.toLowerCase();
      return host !== "localhost" && host !== "127.0.0.1" && host !== "0.0.0.0" && host !== "::1" && host !== "[::1]";
    } catch {
      return true;
    }
  }

  private setupProfileNeedsCompanion(profile: ConnectorSetupProfile | undefined, transport: ConnectorSetupRequest["transport"]): boolean {
    if (this.setupProfileUsesCompanionOAuthBroker(profile)) {
      return true;
    }
    if (this.setupProfileIsHostedHttp(profile) && profile?.browserDirect !== "unsupported" && profile?.setupDisabled !== true) {
      return false;
    }
    return profile?.requiresCompanion === true || profile?.availability === "needs_companion" || profile?.transport === "local_stdio" || transport === "local_stdio";
  }

  private applyCompanionExecutionMetadata(
    connectorRuntime: BrowserConnectorRuntime,
    status: ConnectorStatusResponse["connectors"][number],
    overlay?: ConnectorStatus | undefined,
    forceCompanion = false,
  ): ConnectorStatus {
    const catalog = connectorRuntime.getCatalogItem(status.connectorId);
    const profile = catalog?.setupProfiles?.find((entry) => entry.id === status.setupProfileId)
      ?? catalog?.setupProfiles?.find((entry) => entry.defaultWhenCompanionAbsent)
      ?? catalog?.setupProfiles?.[0];
    const browserDirect = !forceCompanion && !this.setupProfileUsesCompanionOAuthBroker(profile) && this.setupProfileIsHostedHttp(profile) && profile?.browserDirect !== "unsupported" && profile?.setupDisabled !== true;
    const usesCompanion = forceCompanion || (!browserDirect && this.setupProfileNeedsCompanion(profile, status.transport));
    return {
      ...status,
      ...(overlay ?? {}),
      executionEnvironment: browserDirect ? "browser" : (usesCompanion ? "companion" : "browser"),
      executionAvailable: overlay?.executionAvailable ?? (!usesCompanion && browserDirect && status.healthState === "ready" && Boolean(status.capabilities?.allowedTools.length)),
    };
  }

  private async mergeConnectorStatuses(
    statuses: ConnectorStatusResponse["connectors"],
    overlayStatuses: ConnectorStatus[] | undefined,
  ): Promise<ConnectorStatusResponse> {
    const connectorRuntime = await this.getConnectorRuntime();
    const overlayById = new Map((overlayStatuses ?? []).map((status) => [status.id, status]));
    return {
      connectors: statuses.map((status) => this.applyCompanionExecutionMetadata(
        connectorRuntime,
        status,
        overlayById.get(status.id),
        Boolean(connectorRuntime.buildCompanionConnectorDefinition(status.id)),
      )),
    };
  }

  private async prepareCompanionBinding(
    session: BrowserOfficeSession,
    officeState: OfficeStateUpdate,
    windowId?: string,
  ): Promise<CompanionSessionBinding | undefined> {
    const connectors = this.hasStoredConnectorState()
      ? (await this.getConnectorRuntime()).buildCompanionSessionConnectors({
          host: officeState.host,
          documentId: officeState.document.id,
          documentTitle: officeState.document.title,
          documentSaved: officeState.document.saved,
          documentUrl: officeState.document.documentUrl,
          workspaceId: officeState.document.workspaceDir,
        })
      : [];
    return this.companionClient.openSession(
      session.sessionId,
      officeState,
      connectors,
      this.buildCompanionSettingsSyncRequest(session.sessionId, connectors.length),
      windowId,
    );
  }

  private buildCompanionSettingsSyncRequest(
    browserSessionId: string,
    connectorCount = 0,
  ): CompanionSettingsSyncRequest {
    return {
      browserSessionId,
      preferences: { ...this.userPreferences },
      providerSelection: {
        enabledProviders: [...this.providerModelSelection.enabledProviders],
        enabledModels: [...this.providerModelSelection.enabledModels],
        defaultModelByProvider: { ...this.userPreferences.defaultModelByProvider },
      },
      connectorCount,
      secretsIncluded: false,
    };
  }

  private async syncCompanionSettingsForSession(session: BrowserOfficeSession): Promise<void> {
    try {
      await this.companionClient.syncSettings(
        session.sessionId,
        this.buildCompanionSettingsSyncRequest(
          session.sessionId,
          session.companion.connectorToolNames?.length ?? 0,
        ),
      );
      const binding = this.companionClient.getBinding(session.sessionId);
      if (binding) {
        session.setCompanion(binding);
      }
    } catch {
      // Best-effort non-secret sync; disconnected companions stay covered by taskpane fallback.
    }
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

  private async probeConnectorThroughCompanion(request: ConnectorSetupRequest): Promise<{
    ok: boolean;
    status: ConnectorStatus;
    diagnostics: ConnectorDiagnostic[];
  } | undefined> {
    const connectorRuntime = await this.getConnectorRuntime();
    const definition = connectorRuntime.buildCompanionConnectorDefinitionFromSetup(request);
    if (!definition) {
      return undefined;
    }
    return this.companionClient.probeConnector(definition);
  }

  async request<T>(path: string, init?: RequestInit): Promise<T> {
    await this.authStore.ready;
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
      const capability = getProviderCapability(record.provider);
      if (!capability.browserCallable || !capability.apiKeySupported) {
        throw new Error(
          `${titleCase(record.provider)} does not accept browser-stored API keys in Pi-Office yet. ${
            capability.capabilityNote ?? "Use a supported API-key provider or wait for companion-owned auth."
          }`,
        );
      }
      await this.authStore.setApiKey(record.provider, record.apiKey);
      return { ok: true } as T;
    }
    if (method === "POST" && path === "/v1/auth/start") {
      const request = body as { provider?: string; providerId?: string } | undefined;
      const provider = request?.providerId ?? request?.provider ?? "";
      const capability = getProviderCapability(provider);
      if (provider && capability.authMethods.includes("oauth")) {
        throw new Error(
          `${titleCase(provider)} requires companion-owned OAuth/token brokerage before Pi-Office can start sign-in.`,
        );
      }
      throw new Error("OAuth sign-in is unavailable in browser-only mode. Use a supported API-key provider in Settings.");
    }

    if (method === "DELETE" && path === "/v1/auth") {
      await this.authStore.clearAll();
      return { ok: true } as T;
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
    if (method === "GET" && path === "/v1/companion/provider-auth/status") {
      const companion = await this.getCompanionState();
      if (companion.status !== "connected") {
        return unavailableCompanionProviderAuthStatus() as T;
      }
      return ((await this.companionClient.getProviderAuthStatus()) ?? unavailableCompanionProviderAuthStatus()) as T;
    }
    if (method === "POST" && path === "/v1/companion/provider-auth/copy-api-key") {
      const request = body as { provider?: string; explicitUserAction?: boolean } | undefined;
      const provider = typeof request?.provider === "string" ? request.provider.trim() : "";
      if (!provider) {
        throw new Error("provider is required.");
      }
      if (request?.explicitUserAction !== true) {
        throw new Error("explicitUserAction=true is required before copying taskpane provider credentials to the companion.");
      }
      const capability = getProviderCapability(provider);
      if (!capability.authMethods.includes("api_key")) {
        throw new Error(`${titleCase(provider)} does not support API-key companion storage yet.`);
      }
      const apiKey = this.modelRegistry.getApiKey(provider);
      if (!apiKey) {
        throw new Error(`No taskpane API key is stored for ${titleCase(provider)}.`);
      }
      return (await this.companionClient.setProviderApiKey({
        provider,
        apiKey,
        explicitUserAction: true,
      })) as T;
    }
    if (method === "DELETE" && path === "/v1/companion/provider-auth") {
      const request = body as { provider?: string } | undefined;
      return ((await this.companionClient.clearProviderAuth({
        provider: typeof request?.provider === "string" && request.provider.trim()
          ? request.provider.trim()
          : undefined,
      })) ?? unavailableCompanionProviderAuthStatus(request?.provider ? [request.provider] : [])) as T;
    }

    const connectorRuntime = path.startsWith("/v1/connectors")
      ? await this.getConnectorRuntime()
      : undefined;

    if (connectorRuntime) {
    if (method === "GET" && path === "/v1/connectors/catalog") {
      return connectorRuntime.getCatalogResponse() as T;
    }
    if (method === "POST" && path === "/v1/connectors/status") {
      const request = body as { scopeContext?: ConnectorScopeContext; sessionId?: string } | undefined;
      const companion = await this.getCompanionState();
      const overlayStatuses = companion.status === "connected"
        ? this.getSessionCompanionConnectors(request?.sessionId)
        : undefined;
      return this.mergeConnectorStatuses(
        connectorRuntime.getStatusResponse(request?.scopeContext).connectors,
        overlayStatuses,
      ) as T;
    }
    if (method === "GET" && path === "/v1/connectors/diagnostics") {
      const companion = await this.getCompanionState();
      const diagnostics = connectorRuntime.getDiagnostics();
      const companionDiagnostics = companion.status === "connected"
        ? await this.companionClient.getConnectorDiagnostics().catch((error): ConnectorDiagnosticsResponse => ({
            generatedAt: new Date().toISOString(),
            runtimes: diagnostics.runtimes,
            envSuggestions: [],
            diagnostics: [{
              level: "warning",
              code: "companion_diagnostics_unavailable",
              title: "Companion diagnostics unavailable",
              message: error instanceof Error ? error.message : String(error),
            }],
          }))
        : undefined;
      return {
        ...diagnostics,
        generatedAt: companionDiagnostics?.generatedAt ?? diagnostics.generatedAt,
        runtimes: companionDiagnostics?.runtimes ?? diagnostics.runtimes,
        diagnostics: [
          ...(companionDiagnostics?.diagnostics ?? diagnostics.diagnostics),
          companion.status === "connected"
            ? {
                level: "info",
                code: "optional_companion_connected",
                title: "Optional companion connected",
                message: `Read-only local file tools plus companion-required MCP connectors are available through ${companion.endpoint}.`,
              }
            : {
                level: companion.status === "error" ? "warning" : "info",
                code: "optional_companion_unavailable",
                title: "Optional companion not connected",
                message:
                  "Local STDIO, local HTTP, and unverified hosted MCP profiles need the optional companion to verify and execute.",
              },
        ],
      } as T;
    }
    if (method === "GET" && path === "/v1/connectors/audit") {
      return connectorRuntime.getAuditPreferenceResponse() as T;
    }
    if (method === "POST" && path === "/v1/connectors/audit") {
      return (await connectorRuntime.setAuditPreference(body as ConnectorAuditPreference)) as T;
    }
    if (method === "DELETE" && path === "/v1/connectors") {
      await this.companionClient.clearConnectorOAuthTokens();
      return (await connectorRuntime.clearAll()) as T;
    }
    if (method === "GET" && path === "/v1/connectors/export") {
      return connectorRuntime.getExportBundle() as T;
    }
    if (method === "POST" && path === "/v1/connectors/import/preview") {
      const request = body as ConnectorImportPreviewRequest;
      return connectorRuntime.previewImport(request) as T;
    }
    if (method === "POST" && path === "/v1/connectors/import/apply") {
      return (await connectorRuntime.applyImport(body as ConnectorImportApplyRequest)) as T;
    }
    if (method === "POST" && path === "/v1/connectors/setup/prepare") {
      const request = body as { connectorId?: string; scopeContext?: ConnectorScopeContext } | undefined;
      const connectorId = String(request?.connectorId ?? "");
      if (!connectorId) {
        throw new Error("connectorId is required.");
      }
      const companion = await this.getCompanionState();
      const response = connectorRuntime.prepareConnector(connectorId, request?.scopeContext);
      const selectedProfile = response.connector.setupProfiles?.find((profile) => profile.id === response.draft?.setupProfileId)
        ?? response.connector.setupProfiles?.[0];
      const requiresCompanion = this.setupProfileNeedsCompanion(selectedProfile, selectedProfile?.transport ?? response.connector.transport);
      const oauthBrokerRequired = this.setupProfileUsesCompanionOAuthBroker(selectedProfile);
      return {
        ...response,
        diagnostics: this.addCompanionDiagnostics(
          response.diagnostics,
          selectedProfile?.transport ?? response.connector.transport,
          companion,
          requiresCompanion,
          oauthBrokerRequired,
        ),
        executionEnvironment: requiresCompanion
          ? "companion"
          : "browser",
        executionAvailable: false,
      } as T;
    }
    if (method === "POST" && path === "/v1/connectors/setup/connect") {
      const request = body as ConnectorSetupRequest;
      const companion = await this.getCompanionState();
      const response = await connectorRuntime.connectConnector(request);
      let probeDiagnostics: ConnectorDiagnostic[] = [];
      let probe = undefined as Awaited<ReturnType<InProcessKernel["probeConnectorThroughCompanion"]>>;
      const definition = connectorRuntime.buildCompanionConnectorDefinitionFromSetup({ ...request, existingId: response.status.id });
      try {
        probe = definition ? await this.probeConnectorThroughCompanion({ ...request, existingId: response.status.id }) : undefined;
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
        status: this.applyCompanionExecutionMetadata(connectorRuntime, response.status, probe?.status, Boolean(definition)),
        diagnostics: this.addCompanionDiagnostics(
          [...response.diagnostics, ...probeDiagnostics, ...(probe?.diagnostics ?? [])],
          response.status.transport,
          companion,
          Boolean(definition),
          definition?.oauth?.broker === "companion",
        ),
      } as T;
    }
    if (method === "POST" && path === "/v1/connectors/setup/test") {
      const request = body as ConnectorSetupRequest;
      const companion = await this.getCompanionState();
      const response = await connectorRuntime.testConnector(request);
      let probeDiagnostics: ConnectorDiagnostic[] = [];
      let probe = undefined as Awaited<ReturnType<InProcessKernel["probeConnectorThroughCompanion"]>>;
      const definition = connectorRuntime.buildCompanionConnectorDefinitionFromSetup(request);
      try {
        probe = definition ? await this.probeConnectorThroughCompanion(request) : undefined;
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
        status: this.applyCompanionExecutionMetadata(connectorRuntime, response.status, probe?.status, Boolean(definition)),
        diagnostics: this.addCompanionDiagnostics(
          [...response.diagnostics, ...probeDiagnostics, ...(probe?.diagnostics ?? [])],
          response.status.transport,
          companion,
          Boolean(definition),
          definition?.oauth?.broker === "companion",
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
      const definition = connectorRuntime.buildCompanionConnectorDefinition(connectorId);
      if (definition?.oauth?.broker === "companion" && companion.status === "connected") {
        const companionStatus = await this.companionClient.getConnectorOAuthStatus({ connectorId });
        await connectorRuntime.syncCompanionOAuthStatus({ connectorId }, companionStatus, request?.scopeContext);
      }
      const response = await connectorRuntime.reverifyConnector(connectorId, request?.scopeContext);
      let probeDiagnostics: ConnectorDiagnostic[] = [];
      let probe: {
        ok: boolean;
        status: ConnectorStatus;
        diagnostics: ConnectorDiagnostic[];
      } | undefined;
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
        status: this.applyCompanionExecutionMetadata(connectorRuntime, response.status, probe?.status, Boolean(definition)),
        diagnostics: this.addCompanionDiagnostics(
          [...response.diagnostics, ...probeDiagnostics, ...(probe?.diagnostics ?? [])],
          response.status.transport,
          companion,
          Boolean(definition),
          definition?.oauth?.broker === "companion",
        ),
      } as T;
    }
    if (method === "POST" && path === "/v1/connectors/oauth/start") {
      const request = body as { connectorId?: string } | undefined;
      const connectorId = String(request?.connectorId ?? "");
      if (!connectorId) {
        throw new Error("connectorId is required.");
      }
      const definition = connectorRuntime.buildCompanionConnectorDefinition(connectorId);
      if (definition?.oauth?.broker === "companion") {
        const companion = await this.getCompanionState();
        if (companion.status !== "connected") {
          throw new Error("Start the local companion before signing in to this connector.");
        }
        const started = await this.companionClient.startConnectorOAuth(definition);
        return (await connectorRuntime.markCompanionOAuthStarted(connectorId, started) as ConnectorOAuthStartResponse) as T;
      }
      return (await connectorRuntime.startOAuth(connectorId) as ConnectorOAuthStartResponse) as T;
    }
    if (method === "POST" && path === "/v1/connectors/oauth/status") {
      const request = body as { connectorId?: string; state?: string; scopeContext?: ConnectorScopeContext } | undefined;
      const connectorId = String(request?.connectorId ?? "");
      const state = String(request?.state ?? "");
      const definition = connectorId ? connectorRuntime.buildCompanionConnectorDefinition(connectorId) : undefined;
      if (definition?.oauth?.broker === "companion") {
        const companion = await this.getCompanionState();
        if (companion.status !== "connected") {
          return connectorRuntime.getOAuthStatus({ connectorId, state }, request?.scopeContext) as T;
        }
        const companionStatus = await this.companionClient.getConnectorOAuthStatus({ connectorId, state });
        return (await connectorRuntime.syncCompanionOAuthStatus(
          { connectorId, state },
          companionStatus,
          request?.scopeContext,
        )) as T;
      }
      return connectorRuntime.getOAuthStatus({ connectorId, state }, request?.scopeContext) as T;
    }
    if (method === "POST" && path === "/v1/connectors/oauth/callback") {
      return (await connectorRuntime.completeOAuth(body as ConnectorOAuthCallbackRequest) as ConnectorOAuthCallbackResponse) as T;
    }
    if (method === "POST" && path === "/v1/connectors/favorite") {
      return (await connectorRuntime.setFavorite(body as ConnectorFavoriteRequest)) as T;
    }
    if (method === "POST" && path === "/v1/connectors/scope") {
      return (await connectorRuntime.updateScope(body as ConnectorScopeUpdateRequest)) as T;
    }
    if (method === "POST" && path === "/v1/connectors/tools") {
      return (await connectorRuntime.updateToolPolicy(body as ConnectorToolPolicyUpdateRequest)) as T;
    }
    const connectorDeleteMatch = method === "DELETE" ? path.match(/^\/v1\/connectors\/([^/]+)$/) : null;
    if (connectorDeleteMatch) {
      const connectorId = decodeURIComponent(connectorDeleteMatch[1] ?? "");
      await this.companionClient.clearConnectorOAuthTokens({ connectorId });
      return (await connectorRuntime.removeConnector(connectorId)) as T;
    }
    const connectorLogsMatch = method === "GET" ? path.match(/^\/v1\/connectors\/([^/]+)\/logs$/) : null;
    if (connectorLogsMatch) {
      return connectorRuntime.getLogs(decodeURIComponent(connectorLogsMatch[1] ?? "")) as T;
    }
    if (path.startsWith("/v1/connectors/")) {
      throw new Error(`Unknown connector route: ${method} ${path}`);
    }
    }

    if (method === "POST" && path === "/v1/sessions/open") {
      const request = body as OfficeSessionOpenRequest;
      if (!request?.host || !request.documentId) {
        throw new Error("host and documentId are required.");
      }

      const documentKey = buildBrowserSessionKey(request.host, request.documentId, request.windowId);
      let session = this.sessionsByDocument.get(documentKey);
      const normalizedState = normalizeOpenState(request);
      const sessionConnectorRuntime = connectorRuntime ?? (this.hasStoredConnectorState() ? await this.getConnectorRuntime() : undefined);

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
          (scopeContext) => sessionConnectorRuntime?.getBrowserConnectorToolNames(scopeContext) ?? [],
          (searchRequest, scopeContext) => sessionConnectorRuntime?.searchBrowserMcpTools(searchRequest, scopeContext) ?? [],
          (sessionId, searchRequest) => this.companionClient.searchMcpTools(sessionId, searchRequest),
          (toolName, params, scopeContext) => {
            if (!sessionConnectorRuntime) {
              throw new Error(`Connector tool "${toolName}" is not enabled for browser-direct execution.`);
            }
            return sessionConnectorRuntime.executeBrowserMcpTool(toolName, params, scopeContext);
          },
          (request) => sessionConnectorRuntime?.getMcpResult(request) ?? { ok: false, error: "No browser MCP result handles are loaded." },
          (request) => sessionConnectorRuntime?.summarizeMcpResult(request) ?? { ok: false, error: "No browser MCP result handles are loaded." },
          (request) => sessionConnectorRuntime?.clearMcpResults(request) ?? { ok: true, cleared: 0 },
          (sessionId, request) => this.companionClient.getMcpResult(sessionId, request),
          (sessionId, request) => this.companionClient.summarizeMcpResult(sessionId, request),
          (sessionId, request) => this.companionClient.clearMcpResults(sessionId, request),
          (sessionId, request) => this.companionClient.executeShellCommand(sessionId, request),
          (sessionId, request) => this.companionClient.captureNativeViewport(sessionId, request),
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

    const capabilitiesMatch = method === "GET" ? path.match(/^\/v1\/sessions\/([^/]+)\/capabilities$/) : null;
    if (capabilitiesMatch) {
      return this.getSession(capabilitiesMatch[1] ?? "").getCapabilities() as T;
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

    const debugLogMatch = method === "GET" ? path.match(/^\/v1\/sessions\/([^/]+)\/debug-log$/) : null;
    if (debugLogMatch) {
      const session = this.getSession(debugLogMatch[1] ?? "");
      return session.exportDebugLog() as T;
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

export function syncKernelProviderSelection(selection: ProviderModelSelectionState): void {
  kernel.setProviderModelSelection(selection);
}
