import type {
  CompanionAgentOfficeToolResultRequest,
  CompanionAgentOfficeToolResultResponse,
  CompanionAgentPromptRequest,
  CompanionAgentPromptResponse,
  CompanionRuntimeMode,
  CompanionSettingsSyncRequest,
} from "@pi-office/pi-office-pack/protocol";
import type { CompanionProviderAuthStore } from "./provider-auth-store.js";

export interface CompanionAgentSessionContext {
  id: string;
  browserSessionId: string;
  settings?: CompanionSettingsSyncRequest | undefined;
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

export class CompanionAgentSessionManager {
  constructor(private readonly providerAuthStore: CompanionProviderAuthStore) {}

  handlePrompt(
    session: CompanionAgentSessionContext,
    request: CompanionAgentPromptRequest,
  ): CompanionAgentPromptResponse {
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

    return {
      ok: false,
      status: "agent_runtime_unavailable",
      sessionId: session.id,
      promptId: request.promptId,
      companionRuntimeMode: mode,
      modelProvider: requestedProvider,
      modelId: requestedModel,
      reason: "Companion-held provider auth is present, but the companion Pi agent runtime is not enabled yet. Office.js execution remains taskpane-owned.",
      taskpaneFallback: true,
      officeToolProxy: true,
      secretsIncluded: false,
    };
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
