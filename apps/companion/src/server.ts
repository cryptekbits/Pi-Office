import { createServer as createHttpsServer } from "node:https";
import type { IncomingMessage } from "node:http";
import type { AddressInfo } from "node:net";
import express from "express";
import { AuthStorage, ModelRegistry } from "@mariozechner/pi-coding-agent";
import {
  type AuthStatusResponse,
  type BridgeClientMessage,
  type CompanionHealthResponse,
  type ConnectorAuditPreference,
  type ConnectorAuditPreferenceResponse,
  type ConnectorCatalogResponse,
  type ConnectorDiagnosticsResponse,
  type ConnectorFavoriteRequest,
  type ConnectorImportApplyRequest,
  type ConnectorImportApplyResponse,
  type ConnectorImportPreviewRequest,
  type ConnectorImportPreviewResponse,
  type ConnectorLogResponse,
  type ConnectorPrepareResponse,
  type ConnectorScopeContext,
  type ConnectorScopeUpdateRequest,
  type ConnectorSetupRequest,
  type ConnectorSetupResponse,
  type ConnectorStatusResponse,
  type ConnectorTestResponse,
  type DeriveSubjectRequest,
  type DeriveSubjectResponse,
  type OfficeSessionOpenRequest,
  type OfficeStateUpdate,
  type PromptRequest,
  type ProviderCatalogResponse,
  type SaveApiKeyRequest,
  type SessionStatsResponse,
  type SetModelRequest,
  type UserPreferences,
  DEFAULT_USER_PREFERENCES,
} from "@pi-office/pi-office-pack";
import { createServer as createViteServer, type ViteDevServer } from "vite";
import WebSocket, { WebSocketServer, type RawData } from "ws";
import { loadConfig, type CompanionConfig } from "./config.js";
import { ConnectorService } from "./connectors/runtime.js";
import { ImageGenService } from "./image-gen.js";
import { OfficeDocumentSession } from "./office-session.js";
import { openExternal } from "./open-browser.js";

function titleCase(input: string): string {
  return input
    .split(/[-_\s]+/g)
    .filter(Boolean)
    .map((chunk) => chunk[0]!.toUpperCase() + chunk.slice(1))
    .join(" ");
}

function readScopeContext(body: unknown): ConnectorScopeContext | undefined {
  if (!body || typeof body !== "object" || Array.isArray(body)) {
    return undefined;
  }
  const record = body as Record<string, unknown>;
  const raw = record.scopeContext;
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
    return undefined;
  }
  return raw as ConnectorScopeContext;
}

export class CompanionServer {
  private readonly config: CompanionConfig;
  private readonly authStorage: AuthStorage;
  private readonly modelRegistry: ModelRegistry;
  private readonly connectorService: ConnectorService;
  private readonly imageGenService: ImageGenService;
  private readonly sessionsById = new Map<string, OfficeDocumentSession>();
  private readonly sessionsByDocument = new Map<string, OfficeDocumentSession>();
  private userPreferences: UserPreferences = { ...DEFAULT_USER_PREFERENCES };
  private viteServer?: ViteDevServer;

  constructor() {
    this.config = loadConfig();
    process.env.PI_CODING_AGENT_DIR = this.config.agentDir;
    this.authStorage = AuthStorage.create(`${this.config.agentDir}/auth.json`);
    this.modelRegistry = ModelRegistry.create(this.authStorage, `${this.config.agentDir}/models.json`);
    this.connectorService = new ConnectorService(this.config);
    this.imageGenService = new ImageGenService(this.authStorage, this.modelRegistry);
  }

  async start(): Promise<void> {
    const app = express();
    app.use(express.json({ limit: "25mb" }));

    const server = createHttpsServer(
      { pfx: this.config.tls.pfx, passphrase: this.config.tls.passphrase },
      app,
    );

    await this.mountRoutes(app);
    await this.mountTaskpane(app, server);

    const socketServer = new WebSocketServer({ noServer: true });
    server.on("upgrade", (request, socket, head) => {
      const url = new URL(request.url ?? "/", this.config.origin);
      const match = url.pathname.match(/^\/v1\/sessions\/([^/]+)\/events$/);
      if (!match) {
        return;
      }

      const session = this.sessionsById.get(match[1] ?? "");
      if (!session) {
        socket.destroy();
        return;
      }

      socketServer.handleUpgrade(request, socket, head, (ws) => {
        socketServer.emit("connection", ws, request, session);
      });
    });

    socketServer.on("connection", (ws: WebSocket, _request: IncomingMessage, session: OfficeDocumentSession) => {
      session.attachBridge(ws);

      ws.on("message", (buffer: RawData) => {
        try {
          const payload = JSON.parse(buffer.toString()) as BridgeClientMessage;
          session.handleBridgeMessage(payload);
        } catch (error) {
          ws.send(
            JSON.stringify({
              type: "error",
              message: error instanceof Error ? error.message : String(error),
            }),
          );
        }
      });

      ws.on("close", () => {
        session.detachBridge(ws);
      });
    });

    await new Promise<void>((resolve) => server.listen(this.config.port, this.config.host, resolve));
    const address = server.address() as AddressInfo;
    console.log(`Pi-Office companion running at https://${address.address}:${address.port}`);
  }

  private async mountRoutes(app: express.Express): Promise<void> {
    app.get("/v1/health", (_request, response) => {
      const body: CompanionHealthResponse = {
        ok: true,
        origin: this.config.origin,
        mode: this.config.mode,
        port: this.config.port,
      };
      response.json(body);
    });

    app.get("/v1/providers", (_request, response) => {
      this.modelRegistry.refresh();
      response.json(this.buildProviderCatalog());
    });

    app.get("/v1/auth/status", (_request, response) => {
      this.modelRegistry.refresh();
      const body: AuthStatusResponse = {
        storedProviders: this.authStorage.list(),
        oauthProviders: this.authStorage.getOAuthProviders().map((provider) => provider.id),
        configuredProviders: Array.from(
          new Set(
            this.modelRegistry
              .getAvailable()
              .map((model) => model.provider)
              .filter((provider): provider is string => Boolean(provider)),
          ),
        ),
      };
      response.json(body);
    });

    app.post("/v1/auth/start", async (request, response) => {
      const providerId = String(request.body?.providerId ?? "");
      if (!providerId) {
        response.status(400).json({ error: "providerId is required." });
        return;
      }

      try {
        await this.authStorage.login(providerId, {
          onAuth: (info) => {
            void openExternal(info.url);
          },
          onPrompt: async (prompt) => {
            throw new Error(`OAuth prompt is not wired into the taskpane flow yet: ${prompt.message}`);
          },
          onManualCodeInput: async () => {
            throw new Error("Manual OAuth code input is not wired into the taskpane flow yet.");
          },
        });
        this.modelRegistry.refresh();
        response.json({ ok: true });
      } catch (error) {
        response.status(500).json({
          error: error instanceof Error ? error.message : String(error),
        });
      }
    });

    app.post("/v1/auth/api-key", (request, response) => {
      const body = request.body as SaveApiKeyRequest;
      if (!body?.provider || !body.apiKey) {
        response.status(400).json({ error: "provider and apiKey are required." });
        return;
      }

      this.authStorage.set(body.provider, { type: "api_key", key: body.apiKey });
      this.modelRegistry.refresh();
      response.json({ ok: true });
    });

    app.delete("/v1/auth/:providerId", (request, response) => {
      this.authStorage.remove(request.params.providerId);
      this.modelRegistry.refresh();
      response.json({ ok: true });
    });

    app.get("/v1/connectors/catalog", (_request, response) => {
      const body: ConnectorCatalogResponse = {
        connectors: this.connectorService.listCatalog(),
      };
      response.json(body);
    });

    app.get("/v1/connectors/status", (_request, response) => {
      const body: ConnectorStatusResponse = this.connectorService.getStatusResponse();
      response.json(body);
    });

    app.post("/v1/connectors/status", (request, response) => {
      const body: ConnectorStatusResponse = this.connectorService.getStatusResponse(readScopeContext(request.body));
      response.json(body);
    });

    app.get("/v1/connectors/diagnostics", (_request, response) => {
      const body: ConnectorDiagnosticsResponse = this.connectorService.getDiagnostics();
      response.json(body);
    });

    app.post("/v1/connectors/setup/prepare", (request, response) => {
      const connectorId = String(request.body?.connectorId ?? "");
      if (!connectorId) {
        response.status(400).json({ error: "connectorId is required." });
        return;
      }

      try {
        const body: ConnectorPrepareResponse = this.connectorService.prepare(connectorId, readScopeContext(request.body));
        response.json(body);
      } catch (error) {
        response.status(404).json({ error: error instanceof Error ? error.message : String(error) });
      }
    });

    app.post("/v1/connectors/setup/connect", async (request, response) => {
      try {
        const body: ConnectorSetupResponse = await this.connectorService.connect(request.body as ConnectorSetupRequest);
        await this.reloadConnectorSessions();
        response.json(body);
      } catch (error) {
        response.status(400).json({ error: error instanceof Error ? error.message : String(error) });
      }
    });

    app.post("/v1/connectors/setup/test", async (request, response) => {
      try {
        const body: ConnectorTestResponse = await this.connectorService.test(request.body as ConnectorSetupRequest);
        response.json(body);
      } catch (error) {
        response.status(400).json({ error: error instanceof Error ? error.message : String(error) });
      }
    });

    app.post("/v1/connectors/reverify", async (request, response) => {
      const connectorId = String(request.body?.connectorId ?? "");
      if (!connectorId) {
        response.status(400).json({ error: "connectorId is required." });
        return;
      }

      try {
        const body: ConnectorTestResponse = await this.connectorService.reverify(connectorId, readScopeContext(request.body));
        await this.reloadConnectorSessions();
        response.json(body);
      } catch (error) {
        response.status(400).json({ error: error instanceof Error ? error.message : String(error) });
      }
    });

    app.post("/v1/connectors/oauth/start", async (request, response) => {
      const connectorId = String(request.body?.connectorId ?? "");
      if (!connectorId) {
        response.status(400).json({ error: "connectorId is required." });
        return;
      }

      const connector = this.connectorService.listCatalog().find((item) => item.id === connectorId);
      if (!connector?.authUrl) {
        response.status(404).json({ error: "This connector does not expose a browser sign-in URL." });
        return;
      }

      await openExternal(connector.authUrl);
      response.json({ ok: true, url: connector.authUrl });
    });

    app.post("/v1/connectors/custom/validate", async (request, response) => {
      try {
        const body: ConnectorTestResponse = await this.connectorService.validateCustom(request.body as ConnectorSetupRequest);
        response.json(body);
      } catch (error) {
        response.status(400).json({ error: error instanceof Error ? error.message : String(error) });
      }
    });

    app.post("/v1/connectors/favorite", (request, response) => {
      const body = request.body as ConnectorFavoriteRequest;
      if (!body?.connectorId) {
        response.status(400).json({ error: "connectorId is required." });
        return;
      }
      const status = this.connectorService.setFavorite(body);
      if (!status) {
        response.status(404).json({ error: "Unknown connector." });
        return;
      }
      response.json({ ok: true, status });
    });

    app.post("/v1/connectors/scope", async (request, response) => {
      const body = request.body as ConnectorScopeUpdateRequest;
      if (!body?.connectorId || !body.scopeTarget) {
        response.status(400).json({ error: "connectorId and scopeTarget are required." });
        return;
      }
      const status = this.connectorService.updateScope(body);
      if (!status) {
        response.status(404).json({ error: "Unknown connector." });
        return;
      }
      await this.reloadConnectorSessions();
      response.json({ ok: true, status });
    });

    app.get("/v1/connectors/audit", (_request, response) => {
      const body: ConnectorAuditPreferenceResponse = {
        preference: this.connectorService.getAuditPreference(),
      };
      response.json(body);
    });

    app.post("/v1/connectors/audit", (request, response) => {
      const body: ConnectorAuditPreferenceResponse = {
        preference: this.connectorService.setAuditPreference(request.body as ConnectorAuditPreference),
      };
      response.json(body);
    });

    app.get("/v1/connectors/export", (_request, response) => {
      response.json(this.connectorService.exportBundle());
    });

    app.post("/v1/connectors/import/preview", (request, response) => {
      const body: ConnectorImportPreviewResponse = this.connectorService.previewImport(request.body as ConnectorImportPreviewRequest);
      response.json(body);
    });

    app.post("/v1/connectors/import/apply", async (request, response) => {
      const body: ConnectorImportApplyResponse = this.connectorService.applyImport(request.body as ConnectorImportApplyRequest);
      await this.reloadConnectorSessions();
      response.json(body);
    });

    app.get("/v1/connectors/:connectorId/logs", (request, response) => {
      const body: ConnectorLogResponse = this.connectorService.getLogs(request.params.connectorId);
      response.json(body);
    });

    app.delete("/v1/connectors/:connectorId", async (request, response) => {
      const removed = this.connectorService.remove(request.params.connectorId);
      if (!removed) {
        response.status(404).json({ error: "Unknown connector." });
        return;
      }

      await this.reloadConnectorSessions();
      response.json({ ok: true });
    });

    app.post("/v1/sessions/open", async (request, response) => {
      const body = request.body as OfficeSessionOpenRequest;
      if (!body?.documentId || !body.host) {
        response.status(400).json({ error: "host and documentId are required." });
        return;
      }

      const documentKey = `${body.host}:${body.documentId}`;
      let session = this.sessionsByDocument.get(documentKey);
      if (body.forceNew && session) {
        this.sessionsById.delete(session.sessionId);
        this.sessionsByDocument.delete(documentKey);
        await session.dispose();
        session = undefined;
      }
      if (!session) {
        session = new OfficeDocumentSession(
          this.config,
          this.authStorage,
          this.modelRegistry,
          this.connectorService,
          this.imageGenService,
          () => this.userPreferences,
          body,
        );
        await session.initialize();
        this.sessionsByDocument.set(documentKey, session);
        this.sessionsById.set(session.sessionId, session);
      }

      response.json(session.toOpenResponse(this.config.origin));
    });

    app.post("/v1/sessions/:sessionId/office-state", async (request, response) => {
      const session = this.sessionsById.get(request.params.sessionId);
      if (!session) {
        response.status(404).json({ error: "Unknown session." });
        return;
      }

      await session.updateOfficeState(request.body as OfficeStateUpdate);
      response.json({ ok: true, mode: session.mode });
    });

    app.get("/v1/sessions/:sessionId/stats", (request, response) => {
      const session = this.sessionsById.get(request.params.sessionId);
      if (!session) {
        response.status(404).json({ error: "Unknown session." });
        return;
      }

      const stats: SessionStatsResponse = session.getStats();
      response.json(stats);
    });

    app.post("/v1/sessions/:sessionId/prompt", async (request, response) => {
      await this.handlePromptMode(request.params.sessionId, "prompt", request.body as PromptRequest, response);
    });

    app.post("/v1/sessions/:sessionId/steer", async (request, response) => {
      await this.handlePromptMode(request.params.sessionId, "steer", request.body as PromptRequest, response);
    });

    app.post("/v1/sessions/:sessionId/follow-up", async (request, response) => {
      await this.handlePromptMode(request.params.sessionId, "followUp", request.body as PromptRequest, response);
    });

    app.post("/v1/sessions/:sessionId/model", async (request, response) => {
      const session = this.sessionsById.get(request.params.sessionId);
      if (!session) {
        response.status(404).json({ error: "Unknown session." });
        return;
      }

      const body = request.body as SetModelRequest;
      try {
        await session.setModel(body.provider, body.modelId);
        response.json({ ok: true });
      } catch (error) {
        response.status(400).json({ error: error instanceof Error ? error.message : String(error) });
      }
    });

    app.post("/v1/sessions/:sessionId/abort", async (request, response) => {
      const session = this.sessionsById.get(request.params.sessionId);
      if (!session) {
        response.status(404).json({ error: "Unknown session." });
        return;
      }

      await session.abort();
      response.json({ ok: true });
    });

    app.post("/v1/sessions/:sessionId/derive-subject", async (request, response) => {
      const session = this.sessionsById.get(request.params.sessionId);
      if (!session) {
        response.status(404).json({ error: "Unknown session." });
        return;
      }

      const body = request.body as DeriveSubjectRequest;
      if (!body?.messages?.length) {
        response.status(400).json({ error: "messages array is required." });
        return;
      }

      try {
        const subject = await session.deriveSubject(body.messages);
        const resp: DeriveSubjectResponse = { subject };
        response.json(resp);
      } catch (error) {
        response.status(500).json({ error: error instanceof Error ? error.message : String(error) });
      }
    });

    app.post("/v1/sessions/:sessionId/thinking-level", (request, response) => {
      const session = this.sessionsById.get(request.params.sessionId);
      if (!session) {
        response.status(404).json({ error: "Unknown session." });
        return;
      }
      const body = request.body as { level: string };
      if (!body?.level) {
        response.status(400).json({ error: "level is required." });
        return;
      }
      try {
        session.setThinkingLevel(body.level);
        response.json({ ok: true, level: body.level });
      } catch (error) {
        response.status(400).json({ error: error instanceof Error ? error.message : String(error) });
      }
    });

    app.get("/v1/sessions/:sessionId/thinking", (request, response) => {
      const session = this.sessionsById.get(request.params.sessionId);
      if (!session) {
        response.status(404).json({ error: "Unknown session." });
        return;
      }
      try {
        response.json(session.getThinkingCapabilities());
      } catch (error) {
        response.status(500).json({ error: error instanceof Error ? error.message : String(error) });
      }
    });

    app.get("/v1/preferences", (_request, response) => {
      response.json({ ...this.userPreferences });
    });

    app.post("/v1/preferences", async (request, response) => {
      const body = request.body as Partial<UserPreferences>;
      const prevLevel = this.userPreferences.autonomyLevel;
      const prevOverrides = JSON.stringify(this.userPreferences.toolPermissionOverrides);
      this.userPreferences = { ...this.userPreferences, ...body };

      const toolConfigChanged =
        this.userPreferences.autonomyLevel !== prevLevel ||
        JSON.stringify(this.userPreferences.toolPermissionOverrides) !== prevOverrides;
      if (toolConfigChanged) {
        for (const session of this.sessionsById.values()) {
          session.reloadConnectorRuntime().catch((err) => {
            console.error("[server] Failed to reload session after preference change", err);
          });
        }
      }

      response.json({ ok: true, preferences: { ...this.userPreferences } });
    });

    app.get("/v1/image-models", (_request, response) => {
      this.modelRegistry.refresh();
      response.json(this.imageGenService.getCatalog());
    });
  }

  private async mountTaskpane(
    app: express.Express,
    server: ReturnType<typeof createHttpsServer>,
  ): Promise<void> {
    if (this.config.mode === "development") {
      this.viteServer = await createViteServer({
        root: this.config.taskpaneRoot,
        server: {
          middlewareMode: true,
          hmr: { server },
        },
        appType: "spa",
      });
      app.use(this.viteServer.middlewares);
      return;
    }

    app.use(express.static(this.config.taskpaneDist));
    app.get("*", (_request, response) => {
      response.sendFile(`${this.config.taskpaneDist}/index.html`);
    });
  }

  private buildProviderCatalog(): ProviderCatalogResponse {
    const available = new Set(this.modelRegistry.getAvailable().map((model) => `${model.provider}:${model.id}`));
    const oauthProviders = new Set(this.authStorage.getOAuthProviders().map((provider) => provider.id));
    const grouped = new Map<string, ProviderCatalogResponse["providers"][number]>();

    for (const model of this.modelRegistry.getAll()) {
      const key = model.provider;
      if (!grouped.has(key)) {
        grouped.set(key, {
          provider: key,
          label: titleCase(key),
          configured: this.authStorage.hasAuth(key),
          oauthSupported: oauthProviders.has(key),
          models: [],
        });
      }

      const totalCostPer1k = (model.cost.input + model.cost.output) / 2;
      const costTier: "$" | "$$" | "$$$" =
        totalCostPer1k <= 1 ? "$" : totalCostPer1k <= 10 ? "$$" : "$$$";

      grouped.get(key)!.models.push({
        provider: key,
        providerLabel: titleCase(key),
        modelId: model.id,
        modelName: model.name,
        configured: available.has(`${model.provider}:${model.id}`),
        oauthSupported: oauthProviders.has(key),
        usesApiKey: !oauthProviders.has(key),
        contextWindow: model.contextWindow,
        costTier,
        supportsThinking: model.reasoning,
        supportsReasoningEffort: model.reasoning,
      });
    }

    return {
      providers: Array.from(grouped.values())
        .sort((left, right) => left.label.localeCompare(right.label))
        .map((provider) => ({
          ...provider,
          models: provider.models.sort((left, right) => left.modelName.localeCompare(right.modelName)),
        })),
    };
  }

  private async reloadConnectorSessions(): Promise<void> {
    await Promise.all(
      Array.from(this.sessionsById.values()).map((session) => session.reloadConnectorRuntime()),
    );
  }

  private async handlePromptMode(
    sessionId: string,
    mode: "prompt" | "steer" | "followUp",
    body: PromptRequest,
    response: express.Response,
  ): Promise<void> {
    const session = this.sessionsById.get(sessionId);
    if (!session) {
      response.status(404).json({ error: "Unknown session." });
      return;
    }

    const text = String(body?.text ?? "");
    if (!text.trim()) {
      response.status(400).json({ error: "Prompt text is required." });
      return;
    }

    try {
      await session.prompt(text, mode, body.images);
      response.json({ ok: true });
    } catch (error) {
      response.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  }
}
