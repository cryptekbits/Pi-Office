import { mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { randomUUID } from "node:crypto";
import WebSocket from "ws";
import { completeSimple, type Context, type Model, type ThinkingLevel as PiThinkingLevel } from "@mariozechner/pi-ai";
import {
  AuthStorage,
  createAgentSession,
  createCodingTools,
  DefaultResourceLoader,
  ModelRegistry,
  SessionManager,
  type AgentSession,
} from "@mariozechner/pi-coding-agent";
import {
  composeAutonomyPrompt,
  composeOfficeAwarePrompt,
  createOfficeExtension,
  getOfficeMode,
  getOfficeSkillPaths,
  HOST_LABELS,
  OFFICE_APPEND_SYSTEM_PROMPT,
  type BridgeClientMessage,
  type BridgeServerMessage,
  type OfficeContextPayload,
  type OfficeHost,
  type OfficeSessionOpenRequest,
  type OfficeSessionOpenResponse,
  type OfficeStateUpdate,
  type OfficeToolName,
  type PromptImagePayload,
  type PromptMode,
  type SessionStatsResponse,
  type ConnectorScopeContext,
} from "@pi-office/pi-office-pack";
import type { CompanionConfig } from "./config.js";
import { saveCheckpointToDisk, pruneOldCheckpoints, loadCheckpointData, loadCheckpointsFromDisk } from "./rewind-store.js";
import { createConnectorGuardExtension } from "./connectors/guard-extension.js";
import type { ConnectorService } from "./connectors/runtime.js";
import { headlessUiContext } from "./headless-ui.js";
import { captureWordViewport } from "./viewport-capture.js";
import type { ImageGenService } from "./image-gen.js";
import type {
  AskUserRequest,
  AskUserResponse,
  ImageReasoningEffort,
  ToolPermissionDecision,
  ToolPermissionRequest,
  UserPreferences,
} from "@pi-office/pi-office-pack/protocol";
import {
  AUTONOMY_LEVEL_AUTO_APPROVE,
  TOOL_CATEGORY_MAP,
  type AutonomyLevel,
  type ToolCategory,
} from "@pi-office/pi-office-pack/protocol";

function nowIso(): string {
  return new Date().toISOString();
}

function normalizeDocumentState(request: OfficeSessionOpenRequest): OfficeStateUpdate {
  return {
    host: request.host,
    document: {
      id: request.documentId,
      title: request.title,
      saved: request.saved,
      documentPath: request.documentPath,
      documentUrl: request.documentUrl,
      workspaceDir: request.documentPath ? dirname(request.documentPath) : undefined,
    },
    selection: request.selectionSummary ?? { label: "Selection unavailable" },
    capabilities: [],
    timestamp: nowIso(),
  };
}

interface PendingToolCall {
  toolName: OfficeToolName;
  resolve: (value: unknown) => void;
  reject: (error: Error) => void;
  timeout: NodeJS.Timeout;
}

interface PendingAskUser {
  resolve: (value: AskUserResponse) => void;
  reject: (error: Error) => void;
  timeout: NodeJS.Timeout;
}

interface PendingPermission {
  resolve: (value: ToolPermissionDecision) => void;
  reject: (error: Error) => void;
  timeout: NodeJS.Timeout;
}

interface PendingEditProposal {
  resolve: (value: import("@pi-office/pi-office-pack/protocol").OfficeEditProposalDecision) => void;
  reject: (error: Error) => void;
  timeout: NodeJS.Timeout;
}

interface PiImageContent {
  type: "image";
  data: string;
  mimeType: string;
}

function describeAttachedImages(images: PromptImagePayload[] | undefined): string | undefined {
  if (!images?.length) {
    return undefined;
  }

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

function toPiImages(images: PromptImagePayload[] | undefined): PiImageContent[] | undefined {
  if (!images?.length) {
    return undefined;
  }

  return images.map((image) => ({
    type: "image",
    data: image.data,
    mimeType: image.mimeType,
  }));
}

function isOfficeContextPayload(value: unknown): value is OfficeContextPayload {
  return Boolean(value && typeof value === "object" && "summary" in value && "state" in value);
}

function getNestedString(record: Record<string, unknown> | undefined, path: string[]): string | undefined {
  let current: unknown = record;
  for (const segment of path) {
    if (!current || typeof current !== "object") {
      return undefined;
    }
    current = (current as Record<string, unknown>)[segment];
  }

  return typeof current === "string" && current.trim() ? current : undefined;
}

export class OfficeDocumentSession {
  readonly sessionId = randomUUID();
  readonly documentKey: string;

  private officeState: OfficeStateUpdate;
  private bridgeSocket: WebSocket | undefined;
  private pendingToolCalls = new Map<string, PendingToolCall>();
  private pendingAskUser = new Map<string, PendingAskUser>();
  private pendingPermissions = new Map<string, PendingPermission>();
  private pendingEditProposals = new Map<string, PendingEditProposal>();
  private sessionApprovedTools = new Set<string>();
  private agentSession?: AgentSession;
  private agentUnsubscribe?: () => void;

  constructor(
    private readonly config: CompanionConfig,
    private readonly authStorage: AuthStorage,
    private readonly modelRegistry: ModelRegistry,
    private readonly connectorService: ConnectorService,
    private readonly imageGenService: ImageGenService,
    private readonly getPreferences: () => UserPreferences,
    request: OfficeSessionOpenRequest,
  ) {
    this.documentKey = `${request.host}:${request.documentId}`;
    this.officeState = normalizeDocumentState(request);
  }

  async initialize(): Promise<void> {
    await this.rebuildAgent(false);
  }

  get mode() {
    return getOfficeMode(this.officeState.document.saved);
  }

  get host(): OfficeHost {
    return this.officeState.host;
  }

  toOpenResponse(origin: string): OfficeSessionOpenResponse {
    return {
      sessionId: this.sessionId,
      mode: this.mode,
      origin,
      eventsPath: `/v1/sessions/${this.sessionId}/events`,
    };
  }

  attachBridge(socket: WebSocket): void {
    this.bridgeSocket = socket;
    this.send({ type: "connection_state", state: "ready" });
  }

  detachBridge(socket: WebSocket): void {
    if (this.bridgeSocket === socket) {
      this.bridgeSocket = undefined;
      this.rejectAllPendingToolCalls(new Error("Office bridge disconnected."));
    }
  }

  async updateOfficeState(next: OfficeStateUpdate): Promise<void> {
    const previousDir = this.officeState.document.workspaceDir;
    const wasSaved = this.officeState.document.saved;
    this.officeState = next;
    const nextDir = this.officeState.document.workspaceDir;
    const workspaceChanged = this.officeState.document.saved && (!wasSaved || previousDir !== nextDir);

    if (workspaceChanged) {
      await this.rebuildAgent(true);
      this.send({
        type: "error",
        message: `Workspace mode enabled for ${HOST_LABELS[this.host]}. Pi is now bound to ${nextDir}.`,
      });
    }
  }

  async setModel(provider: string, modelId: string): Promise<void> {
    const session = this.requireSession();
    const model = this.modelRegistry.find(provider, modelId) as Model<any> | undefined;
    if (!model) {
      throw new Error(`Model ${provider}/${modelId} is not available.`);
    }
    await session.setModel(model);
  }

  setThinkingLevel(level: string): void {
    const session = this.requireSession();
    session.setThinkingLevel(level as PiThinkingLevel);
  }

  getThinkingCapabilities(): {
    supportsThinking: boolean;
    availableLevels: string[];
    supportsXhigh: boolean;
    currentLevel: string;
  } {
    const session = this.requireSession();
    return {
      supportsThinking: session.supportsThinking(),
      availableLevels: session.getAvailableThinkingLevels(),
      supportsXhigh: session.supportsXhighThinking(),
      currentLevel: session.thinkingLevel,
    };
  }

  async prompt(text: string, mode: PromptMode = "prompt", images?: PromptImagePayload[]): Promise<void> {
    const session = this.requireSession();
    const attachmentNotes = describeAttachedImages(images);
    const input = composeOfficeAwarePrompt(
      attachmentNotes ? `${text.trim()}\n\n${attachmentNotes}` : text,
      this.officeState,
    );
    const piImages = toPiImages(images);

    if (mode === "steer") {
      await session.steer(input, piImages);
      return;
    }

    if (mode === "followUp") {
      await session.followUp(input, piImages);
      return;
    }

    await session.prompt(input, piImages ? { images: piImages } : undefined);
  }

  async abort(): Promise<void> {
    const session = this.requireSession();
    await session.abort();
  }

  getStats(): SessionStatsResponse {
    const session = this.requireSession();
    const stats = session.getSessionStats();
    const result: SessionStatsResponse = { ...stats };

    if (stats.contextUsage && stats.contextUsage.tokens != null) {
      const contextWindow = stats.contextUsage.contextWindow;
      const totalActual = stats.contextUsage.tokens;

      const CORE_TOOL_NAMES = new Set([
        "read", "bash", "edit", "write", "grep", "find", "ls", "mcp",
        "office_get_context", "office_apply_edit", "office_navigate",
        "office_capture_snapshot", "office_capture_viewport", "office_read_section",
        "office_execute_js", "office_propose_edits", "ask_user", "generate_image",
      ]);

      const systemPromptChars = session.systemPrompt.length;
      let coreToolDefChars = 0;
      let integrationToolDefChars = 0;
      for (const tool of session.state.tools) {
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
      for (const msg of session.messages) {
        if (msg.role === "user") {
          const content = (msg as { content: string | unknown[] }).content;
          if (typeof content === "string") messageChars += content.length;
          else if (Array.isArray(content)) {
            for (const p of content) {
              if ((p as { type: string }).type === "text" && (p as { text?: string }).text) {
                messageChars += ((p as { text: string }).text).length;
              }
            }
          }
        } else if (msg.role === "assistant") {
          const content = (msg as { content: unknown[] }).content;
          if (Array.isArray(content)) {
            for (const p of content) {
              const pt = p as { type: string; text?: string; name?: string; arguments?: unknown };
              if (pt.type === "text" && pt.text) messageChars += pt.text.length;
              else if (pt.type === "toolCall") toolCallChars += (pt.name?.length ?? 0) + JSON.stringify(pt.arguments ?? {}).length;
            }
          }
        } else if (msg.role === "toolResult") {
          const content = (msg as { content: string | unknown[] }).content;
          if (typeof content === "string") toolCallChars += content.length;
          else if (Array.isArray(content)) {
            for (const p of content) {
              if ((p as { type: string }).type === "text" && (p as { text?: string }).text) {
                toolCallChars += ((p as { text: string }).text).length;
              }
            }
          }
        }
      }

      const totalRaw = Math.ceil((systemPromptChars + coreToolDefChars + integrationToolDefChars + messageChars + toolCallChars) / 4);
      const ratio = totalRaw > 0 ? totalActual / totalRaw : 1;
      const sysTokens = Math.round(Math.ceil(systemPromptChars / 4) * ratio);
      const coreToolTokens = Math.round(Math.ceil(coreToolDefChars / 4) * ratio);
      const integrationTokens = Math.round(Math.ceil(integrationToolDefChars / 4) * ratio);
      const msgTokens = Math.round(Math.ceil(messageChars / 4) * ratio);
      const tcTokens = Math.round(Math.ceil(toolCallChars / 4) * ratio);
      const calibratedSum = sysTokens + coreToolTokens + integrationTokens + msgTokens + tcTokens;
      const other = Math.max(0, totalActual - calibratedSum);
      const available = Math.max(0, contextWindow - totalActual);

      result.contextUsage = {
        ...stats.contextUsage,
        breakdown: [
          { label: "System Prompt", tokens: sysTokens, color: "#3b82f6" },
          { label: "Tool Definitions", tokens: coreToolTokens, color: "#a855f7" },
          ...(integrationTokens > 10 ? [{ label: "Integrations", tokens: integrationTokens, color: "#ec4899" }] : []),
          { label: "Messages", tokens: msgTokens, color: "#10b981" },
          { label: "Tool Results", tokens: tcTokens, color: "#f59e0b" },
          ...(other > 10 ? [{ label: "Other", tokens: other, color: "#6b7280" }] : []),
          { label: "Available", tokens: available, color: "#e5e7eb" },
        ],
      };
    }

    return result;
  }

  async deriveSubject(messages: { role: string; text: string }[]): Promise<string> {
    const session = this.requireSession();
    const model = session.model;
    if (!model) throw new Error("No model available to derive subject.");
    const auth = await this.modelRegistry.getApiKeyAndHeaders(model);
    if (!auth.ok) throw new Error(auth.error);

    const context: Context = {
      systemPrompt:
        "Generate a concise 3-6 word title for this conversation. Return only the title, nothing else.",
      messages: messages.map((m) => ({
        role: "user" as const,
        content: m.role === "assistant" ? `[assistant]: ${m.text}` : m.text,
        timestamp: Date.now(),
      })),
    };

    const opts: Record<string, unknown> = {};
    if (auth.apiKey) opts.apiKey = auth.apiKey;
    if (auth.headers) opts.headers = auth.headers;
    const result = await completeSimple(model, context, opts);
    const text = result.content.find((c) => c.type === "text");
    if (!text || text.type !== "text") throw new Error("No text in subject response.");
    return text.text.trim();
  }

  handleBridgeMessage(message: BridgeClientMessage): void {
    if (message.type === "ask_user_response") {
      const pending = this.pendingAskUser.get(message.response.requestId);
      if (!pending) return;
      clearTimeout(pending.timeout);
      this.pendingAskUser.delete(message.response.requestId);
      pending.resolve(message.response);
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

    if (message.type === "tool_permission_response") {
      const pending = this.pendingPermissions.get(message.requestId);
      if (!pending) return;
      clearTimeout(pending.timeout);
      this.pendingPermissions.delete(message.requestId);
      if (message.decision.allowed && message.decision.scope === "session") {
        this.sessionApprovedTools.add(message.decision.toolName);
      }
      pending.resolve(message.decision);
      return;
    }

    if (message.type === "rewind_session") {
      this.handleRewindSession(message.targetMessageCount);
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

    if (message.type !== "office_tool_result") {
      return;
    }

    const pending = this.pendingToolCalls.get(message.result.requestId);
    if (!pending) {
      return;
    }

    clearTimeout(pending.timeout);
    this.pendingToolCalls.delete(message.result.requestId);

    if (message.result.success) {
      pending.resolve(message.result.content);
    } else {
      console.error(`[office-session] ${pending.toolName} failed`, {
        requestId: message.result.requestId,
        error: message.result.error ?? "Office tool call failed.",
      });
      pending.reject(new Error(message.result.error ?? "Office tool call failed."));
    }
  }

  private handleRewindSession(targetMessageCount: number): void {
    try {
      const session = this.requireSession();
      session.agent.state.messages = session.agent.state.messages.slice(0, targetMessageCount);
      console.log(`[office-session] Rewound conversation to ${targetMessageCount} messages`);
    } catch (error) {
      console.error("[office-session] Rewind failed", error);
      this.send({ type: "error", message: `Rewind failed: ${error instanceof Error ? error.message : String(error)}` });
    }
  }

  private handlePersistCheckpoint(documentId: string, checkpoint: import("@pi-office/pi-office-pack/protocol").DocumentCheckpointPayload): void {
    try {
      saveCheckpointToDisk(this.config, documentId, checkpoint);
      pruneOldCheckpoints(this.config, documentId);
    } catch (error) {
      console.error("[office-session] Persist checkpoint failed", error);
    }
  }

  private handleLoadCheckpoint(documentId: string, checkpointId: string): void {
    try {
      const data = loadCheckpointData(this.config, documentId, checkpointId);
      if (data) {
        this.send({ type: "session_event", event: { type: "checkpoint_data", checkpoint: data } });
      }
    } catch (error) {
      console.error("[office-session] Load checkpoint failed", error);
    }
  }

  sendAvailableCheckpoints(documentId: string): void {
    try {
      const checkpoints = loadCheckpointsFromDisk(this.config, documentId);
      if (checkpoints.length) {
        this.send({ type: "available_checkpoints", checkpoints });
      }
    } catch { /* ignore */ }
  }

  async dispose(): Promise<void> {
    this.rejectAllPendingToolCalls(new Error("Session disposed."));
    this.agentUnsubscribe?.();
    this.agentSession?.dispose();
    this.connectorService.clearPreparedSession(this.sessionId);
  }

  async reloadConnectorRuntime(): Promise<void> {
    await this.rebuildAgent(true);
  }

  private async rebuildAgent(preserveMessages: boolean): Promise<void> {
    const previousMessages = preserveMessages ? this.agentSession?.messages : undefined;

    this.agentUnsubscribe?.();
    this.agentSession?.dispose();

    const cwd = this.resolveWorkingDirectory();
    mkdirSync(cwd, { recursive: true });
    mkdirSync(this.config.agentDir, { recursive: true });

    const connectorConfigPath = await this.connectorService.prepareSessionRuntime(this.sessionId, this.getConnectorScopeContext());

    const resourceLoader = new DefaultResourceLoader({
      cwd,
      agentDir: this.config.agentDir,
      additionalExtensionPaths: [this.connectorService.getAdapterExtensionPath()],
      additionalSkillPaths: getOfficeSkillPaths(),
      appendSystemPrompt: `${OFFICE_APPEND_SYSTEM_PROMPT}\n\n${composeAutonomyPrompt(
        this.getPreferences(),
        this.officeState.document.saved,
      )}\n\n${composeOfficeAwarePrompt(
        "Prefer Office tools as the source of truth for the active document.",
        this.officeState,
      )}`,
      extensionFactories: [
        createOfficeExtension({
          getHost: () => this.host,
          getState: () => this.officeState,
          isToolDisabled: (toolName) => {
            const overrides = this.getPreferences().toolPermissionOverrides ?? [];
            return overrides.some((o) => o.toolName === toolName && o.autoApproveAtLevel === "disabled");
          },
          invokeTool: (toolName, params) => this.invokeOfficeTool(toolName, params),
          invokeAskUser: (request) => this.invokeAskUser(request),
          invokeEditProposal: (proposal) => this.invokeEditProposal(proposal),
          isImageGenerationEnabled: () => this.getPreferences().imageGenerationEnabled,
          getImageReasoningEffort: () => this.getPreferences().imageReasoningEffort,
          getDefaultImageModel: () => this.getPreferences().defaultImageModel,
          generateImage: async (params) => {
            const prefs = this.getPreferences();
            console.log("[generate_image] invoked", {
              host: this.host,
              modelKey: prefs.defaultImageModel || "(auto)",
              imageGenerationEnabled: prefs.imageGenerationEnabled,
              aspectRatio: params.aspectRatio,
              size: params.size,
            });
            try {
              const result = await this.imageGenService.generate({
                prompt: params.prompt,
                modelKey: prefs.defaultImageModel || undefined,
                aspectRatio: params.aspectRatio,
                size: params.size,
                quality: params.quality,
                reasoningEffort: prefs.imageReasoningEffort,
                host: this.host,
              });
              console.log("[generate_image] success", {
                modelKey: result.modelKey,
                width: result.width,
                height: result.height,
              });
              return result;
            } catch (error) {
              console.error("[generate_image] failed", error);
              throw error;
            }
          },
        }),
        createConnectorGuardExtension(
          () => this.connectorService.getGuardSnapshotForSession(this.sessionId),
          (toolName, blocked) => this.connectorService.recordSessionToolUse(this.sessionId, toolName, blocked),
        ),
        (pi) => {
          pi.on("tool_call", async (event) => {
            const toolName = event.toolName;
            if (this.shouldAutoApproveTool(toolName)) return undefined;
            const inputObj = (event.input && typeof event.input === "object" ? event.input : {}) as Record<string, unknown>;
            const decision = await this.requestToolPermission(toolName, inputObj);
            if (!decision.allowed) {
              return { block: true, reason: `User denied ${toolName}.` };
            }
            return undefined;
          });
        },
      ],
    });
    const configFlagPresent = process.argv.includes("--mcp-config");
    if (!configFlagPresent) {
      process.argv.push("--mcp-config", connectorConfigPath);
    }

    try {
      await resourceLoader.reload();
    } finally {
      if (!configFlagPresent) {
        process.argv.splice(process.argv.length - 2, 2);
      }
    }

    const { session, modelFallbackMessage } = await createAgentSession({
      cwd,
      agentDir: this.config.agentDir,
      authStorage: this.authStorage,
      modelRegistry: this.modelRegistry,
      tools: this.officeState.document.saved ? createCodingTools(cwd) : [],
      resourceLoader,
      sessionManager: SessionManager.inMemory(cwd),
    });

    if (previousMessages?.length) {
      session.agent.state.messages = [...previousMessages];
    }

    await session.bindExtensions({
      uiContext: headlessUiContext,
      onError: (error) => {
        this.send({
          type: "error",
          message: `Extension error (${error.extensionPath}): ${error.error}`,
        });
      },
    });

    this.agentUnsubscribe = session.subscribe((event) => {
      this.send({ type: "session_event", event });
    });

    this.agentSession = session;

    if (modelFallbackMessage) {
      this.send({
        type: "error",
        message: modelFallbackMessage,
      });
    }
  }

  private resolveWorkingDirectory(): string {
    if (this.officeState.document.workspaceDir) {
      return this.officeState.document.workspaceDir;
    }

    return join(this.config.scratchDir, this.sessionId);
  }

  private getConnectorScopeContext(): ConnectorScopeContext {
    return {
      host: this.host,
      documentId: this.officeState.document.id,
      documentTitle: this.officeState.document.title,
      documentSaved: this.officeState.document.saved,
      documentUrl: this.officeState.document.documentUrl,
      workspaceId: this.officeState.document.workspaceDir,
    };
  }

  private requireSession(): AgentSession {
    if (!this.agentSession) {
      throw new Error("Pi session is not initialized yet.");
    }

    return this.agentSession;
  }

  private async invokeOfficeTool(toolName: OfficeToolName, params: Record<string, unknown>): Promise<unknown> {
    if (toolName === "office_capture_viewport") {
      return this.captureViewportTool(params);
    }

    return this.invokeBridgeOfficeTool(toolName, params);
  }

  private async captureViewportTool(params: Record<string, unknown>): Promise<OfficeContextPayload> {
    if (this.host !== "word") {
      throw new Error("office_capture_viewport is only available for Word.");
    }

    const includeFormatting = params.includeFormatting !== false;
    const includeWindowFrame = params.includeWindowFrame === true;
    const contextResult = await this.invokeBridgeOfficeTool("office_get_context", {});
    if (!isOfficeContextPayload(contextResult)) {
      throw new Error("office_get_context did not return a valid Office context payload.");
    }

    const windowCaption = getNestedString(contextResult.formatting, ["viewport", "window", "caption"]);
    const capture = await captureWordViewport(this.config, {
      documentTitle: this.officeState.document.title,
      windowCaption,
      includeWindowFrame,
    });

    const payload: OfficeContextPayload = {
      ...contextResult,
      formatting: includeFormatting
        ? {
            ...(contextResult.formatting ?? {}),
            viewportCapture: capture.metadata,
          }
        : undefined,
      visuals: capture.visuals,
    };

    payload.summary = [
      contextResult.summary,
      "Visible Word viewport capture attached.",
      includeWindowFrame ? "Full Word window capture also attached." : undefined,
    ]
      .filter(Boolean)
      .join("\n");

    return payload;
  }

  private async invokeBridgeOfficeTool(toolName: OfficeToolName, params: Record<string, unknown>): Promise<unknown> {
    if (!this.bridgeSocket || this.bridgeSocket.readyState !== WebSocket.OPEN) {
      throw new Error("Office bridge is not connected.");
    }

    const requestId = randomUUID();

    return new Promise<unknown>((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.pendingToolCalls.delete(requestId);
        reject(new Error(`${toolName} timed out waiting for the Office host.`));
      }, 45_000);

      this.pendingToolCalls.set(requestId, { toolName, resolve, reject, timeout });
      this.send({
        type: "office_tool_call",
        request: {
          requestId,
          toolName,
          host: this.host,
          params,
        },
      });
    });
  }

  private async invokeAskUser(request: AskUserRequest): Promise<AskUserResponse> {
    if (!this.bridgeSocket || this.bridgeSocket.readyState !== WebSocket.OPEN) {
      throw new Error("Office bridge is not connected.");
    }

    return new Promise<AskUserResponse>((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.pendingAskUser.delete(request.requestId);
        reject(new Error("ask_user timed out waiting for user response."));
      }, 300_000);

      this.pendingAskUser.set(request.requestId, { resolve, reject, timeout });
      this.send({ type: "ask_user_request", request });
    });
  }

  private async invokeEditProposal(
    proposal: import("@pi-office/pi-office-pack/protocol").OfficeEditProposal,
  ): Promise<import("@pi-office/pi-office-pack/protocol").OfficeEditProposalDecision> {
    if (!this.bridgeSocket || this.bridgeSocket.readyState !== WebSocket.OPEN) {
      throw new Error("Office bridge is not connected.");
    }

    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.pendingEditProposals.delete(proposal.requestId);
        reject(new Error("Edit proposal timed out waiting for user review."));
      }, 300_000);

      this.pendingEditProposals.set(proposal.requestId, { resolve, reject, timeout });
      this.send({ type: "edit_proposal_request", proposal });
    });
  }

  private shouldAutoApproveTool(toolName: string): boolean {
    const prefs = this.getPreferences();
    const override = prefs.toolPermissionOverrides?.find((o) => o.toolName === toolName);
    if (override?.autoApproveAtLevel === "disabled") return false;

    const category = (TOOL_CATEGORY_MAP[toolName] ?? "connector") as ToolCategory;
    if (category === "interaction") return true;
    if (this.sessionApprovedTools.has(toolName)) return true;

    const effectiveLevel: AutonomyLevel = override ? override.autoApproveAtLevel as AutonomyLevel : prefs.autonomyLevel;
    return AUTONOMY_LEVEL_AUTO_APPROVE[effectiveLevel].has(category);
  }

  private async requestToolPermission(
    toolName: string,
    params: Record<string, unknown>,
  ): Promise<ToolPermissionDecision> {
    if (!this.bridgeSocket || this.bridgeSocket.readyState !== WebSocket.OPEN) {
      return { toolName, allowed: true, scope: "once" };
    }

    const category = (TOOL_CATEGORY_MAP[toolName] ?? "connector") as ToolCategory;
    const requestId = randomUUID();
    const request: ToolPermissionRequest = { requestId, toolName, toolCategory: category, params };

    return new Promise<ToolPermissionDecision>((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.pendingPermissions.delete(requestId);
        resolve({ toolName, allowed: true, scope: "once" });
      }, 120_000);

      this.pendingPermissions.set(requestId, { resolve, reject, timeout });
      this.send({ type: "tool_permission_request", request });
    });
  }

  private rejectAllPendingToolCalls(error: Error): void {
    for (const [requestId, pending] of this.pendingToolCalls.entries()) {
      clearTimeout(pending.timeout);
      pending.reject(error);
      this.pendingToolCalls.delete(requestId);
    }
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
    if (!this.bridgeSocket || this.bridgeSocket.readyState !== WebSocket.OPEN) {
      return;
    }

    this.bridgeSocket.send(JSON.stringify(message));
  }
}
