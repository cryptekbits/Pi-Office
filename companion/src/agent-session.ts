import type {
  CompanionAgentOfficeToolResultRequest,
  CompanionAgentOfficeToolResultResponse,
  CompanionAgentPromptRequest,
  CompanionAgentPromptResponse,
  CompanionRuntimeMode,
  CompanionSettingsSyncRequest,
  OfficeHost,
} from "@pi-office/pi-office-pack/protocol";
import { Agent, type AgentMessage, type StreamFn } from "@mariozechner/pi-agent-core";
import {
  getModels,
  supportsXhigh,
  type Model,
} from "@mariozechner/pi-ai";
import type { CompanionProviderAuthStore } from "./provider-auth-store.js";

export interface CompanionAgentSessionContext {
  id: string;
  browserSessionId: string;
  host?: OfficeHost | undefined;
  documentId?: string | undefined;
  title?: string | undefined;
  saved?: boolean | undefined;
  workspaceDir?: string | undefined;
  settings?: CompanionSettingsSyncRequest | undefined;
}

interface CompanionAgentRuntime {
  provider: string;
  modelId: string;
  agent: Agent;
  unsubscribe: () => void;
  eventCount: number;
}

export interface CompanionAgentSessionManagerOptions {
  resolveModel?: (provider: string, modelId: string | undefined) => Model<any> | undefined;
  streamFn?: StreamFn | undefined;
}

function requireNoSecrets(value: { secretsIncluded?: unknown } | undefined, action: string): void {
  if (!value || typeof value !== "object") {
    throw new Error(`${action} request is required.`);
  }
  if (value.secretsIncluded !== false) {
    throw new Error(`${action} must not include taskpane provider secrets.`);
  }
}

function firstEnabledProvider(settings: CompanionSettingsSyncRequest | undefined): string | undefined {
  return settings?.providerSelection?.enabledProviders
    ?.map((provider) => provider.trim())
    .find(Boolean);
}

function defaultModelForProvider(
  settings: CompanionSettingsSyncRequest | undefined,
  provider: string | undefined,
): string | undefined {
  if (!provider) return undefined;
  const model = settings?.providerSelection?.defaultModelByProvider?.[provider];
  return typeof model === "string" && model.trim() ? model.trim() : undefined;
}

function resolveCatalogModel(provider: string, modelId: string | undefined): Model<any> | undefined {
  let models: Model<any>[] = [];
  try {
    models = getModels(provider as never) as Model<any>[];
  } catch {
    models = [];
  }
  if (!modelId) return models[0];
  return models.find((model) => model.id === modelId);
}

function normalizeThinkingLevel(level: CompanionSettingsSyncRequest["preferences"]["defaultThinkingLevel"], model: Model<any>) {
  if (!model.reasoning) return "off";
  if (level === "off") return "off";
  if (level === "xhigh" && !supportsXhigh(model)) return "high";
  return level;
}

function extractAssistantText(messages: AgentMessage[]): string {
  const assistant = [...messages].reverse().find((message) => message.role === "assistant");
  if (!assistant || assistant.role !== "assistant") return "";
  return assistant.content
    .map((part) => (part.type === "text" ? part.text : ""))
    .filter(Boolean)
    .join("\n")
    .trim();
}

function latestAssistantError(messages: AgentMessage[]): string | undefined {
  const assistant = [...messages].reverse().find((message) => message.role === "assistant");
  if (!assistant || assistant.role !== "assistant") return undefined;
  return assistant.stopReason === "error" || assistant.stopReason === "aborted"
    ? assistant.errorMessage || `Companion agent stopped with ${assistant.stopReason}.`
    : undefined;
}

function isProviderAuthFailure(error: unknown): boolean {
  const message = (error instanceof Error ? error.message : String(error)).toLowerCase();
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

function buildCompanionSystemPrompt(session: CompanionAgentSessionContext): string {
  const mode = session.settings?.preferences.companionRuntimeMode ?? "smart_auto";
  const title = session.title?.trim() || "Untitled Office document";
  const host = session.host ?? "word";
  const saved = session.saved === true ? "saved" : "unsaved";
  return [
    "You are the Pi-Office Advanced companion agent.",
    "Own model inference and non-Office reasoning from the local companion process.",
    "Do not claim to execute Office.js directly. Office.js document reads and writes remain taskpane-owned and must be proxied through the Office taskpane in later protocol steps.",
    "This first companion runtime slice has no Office tool schemas attached, so answer from the user's prompt and synced taskpane context only.",
    `Companion runtime mode: ${mode}.`,
    `Office host: ${host}.`,
    `Document title: ${title}.`,
    `Document state: ${saved}.`,
    session.workspaceDir ? `Saved document workspace folder: ${session.workspaceDir}.` : "No saved document workspace folder is available.",
    "Never request or reveal provider secrets. Taskpane provider secrets are not silently migrated.",
  ].join("\n");
}

function buildCompanionUserPrompt(text: string, session: CompanionAgentSessionContext): string {
  const lines = [
    "User request:",
    text.trim(),
    "",
    "Synced Office context:",
    `Host: ${session.host ?? "unknown"}`,
    `Document title: ${session.title?.trim() || "Untitled Office document"}`,
    `Document state: ${session.saved === true ? "saved" : "unsaved"}`,
  ];
  if (session.documentId) {
    lines.push(`Document id: ${session.documentId}`);
  }
  return lines.join("\n");
}

export class CompanionAgentSessionManager {
  private readonly runtimes = new Map<string, CompanionAgentRuntime>();

  constructor(
    private readonly providerAuthStore: CompanionProviderAuthStore,
    private readonly options: CompanionAgentSessionManagerOptions = {},
  ) {}

  async handlePrompt(
    session: CompanionAgentSessionContext,
    request: CompanionAgentPromptRequest,
  ): Promise<CompanionAgentPromptResponse> {
    requireNoSecrets(request, "agent prompt");
    if (request.browserSessionId && request.browserSessionId !== session.browserSessionId) {
      throw new Error("browserSessionId does not match this companion session.");
    }
    if (typeof request.text !== "string" || !request.text.trim()) {
      throw new Error("text is required.");
    }

    const mode: CompanionRuntimeMode = session.settings?.preferences.companionRuntimeMode ?? "smart_auto";
    const requestedProvider = typeof request.modelProvider === "string" && request.modelProvider.trim()
      ? request.modelProvider.trim()
      : firstEnabledProvider(session.settings);
    const requestedModel = typeof request.modelId === "string" && request.modelId.trim()
      ? request.modelId.trim()
      : defaultModelForProvider(session.settings, requestedProvider);

    if (mode !== "advanced") {
      return {
        ok: false,
        status: "taskpane_fallback",
        sessionId: session.id,
        promptId: request.promptId,
        companionRuntimeMode: mode,
        modelProvider: requestedProvider,
        modelId: requestedModel,
        reason: "Companion agent preflight is session-owned, but Basic/Smart Auto keeps model inference in the taskpane until Advanced mode is selected.",
        taskpaneFallback: true,
        officeToolProxy: true,
        secretsIncluded: false,
      };
    }

    if (!requestedProvider) {
      return {
        ok: false,
        status: "provider_auth_required",
        sessionId: session.id,
        promptId: request.promptId,
        companionRuntimeMode: mode,
        reason: "Advanced mode requires an enabled provider/model selection synced from the taskpane before companion-owned inference can start.",
        taskpaneFallback: true,
        officeToolProxy: true,
        secretsIncluded: false,
      };
    }

    const providerAuth = this.providerAuthStore.status([requestedProvider]);
    if (!providerAuth.storedProviders.includes(requestedProvider)) {
      return {
        ok: false,
        status: "provider_auth_required",
        sessionId: session.id,
        promptId: request.promptId,
        companionRuntimeMode: mode,
        modelProvider: requestedProvider,
        modelId: requestedModel,
        reason: `Advanced mode requires explicit companion-held auth for ${requestedProvider}; taskpane provider secrets are not silently migrated.`,
        taskpaneFallback: true,
        officeToolProxy: true,
        secretsIncluded: false,
      };
    }

    const model = (this.options.resolveModel ?? resolveCatalogModel)(requestedProvider, requestedModel);
    if (!model) {
      return {
        ok: false,
        status: "agent_runtime_unavailable",
        sessionId: session.id,
        promptId: request.promptId,
        companionRuntimeMode: mode,
        modelProvider: requestedProvider,
        modelId: requestedModel,
        reason: requestedModel
          ? `Companion runtime could not resolve model ${requestedProvider}/${requestedModel}.`
          : `Companion runtime could not resolve a default model for ${requestedProvider}.`,
        taskpaneFallback: true,
        officeToolProxy: true,
        secretsIncluded: false,
      };
    }

    const runtime = this.getOrCreateRuntime(session, model);
    const message: AgentMessage = {
      role: "user",
      content: buildCompanionUserPrompt(request.text, session),
      timestamp: Date.now(),
    };

    try {
      if (request.mode === "steer" && runtime.agent.state.isStreaming) {
        runtime.agent.steer(message);
      } else if (request.mode === "followUp" && runtime.agent.state.isStreaming) {
        runtime.agent.followUp(message);
      } else {
        await runtime.agent.prompt(message);
      }

      const error = latestAssistantError(runtime.agent.state.messages);
      if (error) {
        throw new Error(error);
      }
      this.providerAuthStore.markVerificationSuccess(requestedProvider);
      return {
        ok: true,
        status: "accepted",
        sessionId: session.id,
        promptId: request.promptId,
        companionRuntimeMode: mode,
        modelProvider: requestedProvider,
        modelId: model.id,
        reason: "Companion-owned Pi agent runtime accepted and completed the prompt. Office.js execution remains taskpane-owned.",
        taskpaneFallback: false,
        officeToolProxy: true,
        secretsIncluded: false,
        completed: !runtime.agent.state.isStreaming,
        assistantText: extractAssistantText(runtime.agent.state.messages),
        messageCount: runtime.agent.state.messages.length,
        toolCallCount: runtime.agent.state.messages
          .filter((message) => message.role === "assistant")
          .flatMap((message) => message.role === "assistant" ? message.content : [])
          .filter((part) => part.type === "toolCall").length,
      };
    } catch (error) {
      if (isProviderAuthFailure(error)) {
        this.providerAuthStore.markVerificationFailure(requestedProvider, error);
      }
      return {
        ok: false,
        status: "agent_runtime_unavailable",
        sessionId: session.id,
        promptId: request.promptId,
        companionRuntimeMode: mode,
        modelProvider: requestedProvider,
        modelId: model.id,
        reason: `Companion agent runtime failed: ${error instanceof Error ? error.message : String(error)}`,
        taskpaneFallback: true,
        officeToolProxy: true,
        secretsIncluded: false,
      };
    }
  }

  private getOrCreateRuntime(session: CompanionAgentSessionContext, model: Model<any>): CompanionAgentRuntime {
    const existing = this.runtimes.get(session.id);
    if (existing && existing.provider === String(model.provider) && existing.modelId === model.id) {
      existing.agent.setSystemPrompt(buildCompanionSystemPrompt(session));
      existing.agent.setThinkingLevel(normalizeThinkingLevel(
        session.settings?.preferences.defaultThinkingLevel ?? "medium",
        model,
      ));
      return existing;
    }

    existing?.unsubscribe();
    const agent = new Agent({
      getApiKey: (provider) => this.providerAuthStore.getApiKey(provider),
      ...(this.options.streamFn ? { streamFn: this.options.streamFn } : {}),
      toolExecution: "sequential",
      sessionId: `pi-office-companion:${session.id}`,
    });
    agent.setModel(model);
    agent.setThinkingLevel(normalizeThinkingLevel(
      session.settings?.preferences.defaultThinkingLevel ?? "medium",
      model,
    ));
    agent.setTools([]);
    agent.setSystemPrompt(buildCompanionSystemPrompt(session));
    const runtime: CompanionAgentRuntime = {
      provider: String(model.provider),
      modelId: model.id,
      agent,
      eventCount: 0,
      unsubscribe: agent.subscribe(() => {
        runtime.eventCount += 1;
      }),
    };
    this.runtimes.set(session.id, runtime);
    return runtime;
  }

  handleOfficeToolResult(
    session: CompanionAgentSessionContext,
    request: CompanionAgentOfficeToolResultRequest,
  ): CompanionAgentOfficeToolResultResponse {
    requireNoSecrets(request, "Office tool result");
    if (request.browserSessionId && request.browserSessionId !== session.browserSessionId) {
      throw new Error("browserSessionId does not match this companion session.");
    }
    return {
      ok: false,
      status: "no_pending_office_tool_request",
      sessionId: session.id,
      toolCallId: typeof request.toolCallId === "string" && request.toolCallId.trim()
        ? request.toolCallId.trim()
        : undefined,
      reason: "No companion-owned agent turn is waiting for an Office tool result. Office.js execution remains taskpane-owned.",
      secretsIncluded: false,
    };
  }
}
