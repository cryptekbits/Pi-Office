import { randomUUID } from "node:crypto";
import { mkdirSync } from "node:fs";
import { createServer as createHttpsServer } from "node:https";
import { dirname } from "node:path";
import express from "express";
import { TASKPANE_COMPANION_PROTOCOL } from "@pi-office/pi-office-pack/protocol";
import type {
  CompanionConnectorOAuthClearRequest,
  CompanionConnectorOAuthStartRequest,
  CompanionConnectorOAuthStatusRequest,
  CompanionHealthResponse,
  CompanionNativeCaptureRequest,
  CompanionSettingsSyncCapability,
  CompanionSettingsSyncRequest,
  CompanionSettingsSyncResponse,
  CompanionShellCapability,
  McpToolSearchRequest,
  CompanionShellExecuteRequest,
  CompanionSessionOpenRequest,
  CompanionSessionOpenResponse,
  CompanionState,
  McpResultClearRequest,
  McpResultPageRequest,
  McpResultSummarizeRequest,
} from "@pi-office/pi-office-pack/protocol";
import { loadConfig, type CompanionConfig } from "./config.js";
import { CompanionConnectorBridge } from "./connector-bridge.js";
import { executeFileTool } from "./file-tools.js";
import { companionCorsMiddleware } from "./http.js";
import { captureNativeViewport, createNativeCaptureCapability } from "./native-capture.js";
import { CompanionOAuthBroker } from "./oauth-broker.js";
import { createCompanionRuntimeDiagnostics } from "./runtime-diagnostics.js";
import { CompanionShellSandbox } from "./shell-sandbox.js";

interface SessionRecord {
  id: string;
  browserSessionId: string;
  windowId?: string | undefined;
  host: CompanionSessionOpenRequest["host"];
  documentId: string;
  documentPath?: string | undefined;
  documentUrl?: string | undefined;
  saved: boolean;
  title: string;
  workspaceDir?: string | undefined;
  connectorToolNames: string[];
  settingsSync?: CompanionSettingsSyncCapability | undefined;
}

function createShellSandbox(config: CompanionConfig, session?: SessionRecord | undefined): CompanionShellSandbox {
  return new CompanionShellSandbox({
    sessionId: session?.id ?? "discovery",
    dataDir: config.dataDir,
    workspaceDir: session?.workspaceDir,
  });
}

function createCompanionState(
  config: CompanionConfig,
  sessionId?: string | undefined,
  connectorToolNames?: string[] | undefined,
  shell?: CompanionShellCapability | undefined,
  settingsSync?: CompanionSettingsSyncCapability | undefined,
): CompanionState {
  const connectorToolCount = connectorToolNames?.length ?? 0;
  return {
    status: "connected",
    endpoint: config.endpoint,
    identity: config.identity,
    sessionId,
    connectorToolNames,
    capabilities: {
      fileRead: true,
      localMcp: true,
      endpoint: config.endpoint,
      protocol: TASKPANE_COMPANION_PROTOCOL,
      shell,
      version: "companion-capabilities-v1",
      agent: {
        state: "unavailable",
        available: false,
        version: "companion-agent-v1",
        officeToolProxy: true,
        providerAuth: false,
        smartAuto: true,
        reason: "Companion-owned inference is capability-gated until provider auth is explicitly configured in the companion.",
      },
      providerAuth: {
        state: "unavailable",
        available: false,
        version: "companion-provider-auth-v1",
        explicitMigrationRequired: true,
        supportedAuthMethods: ["oauth", "manual_token", "api_key", "cloud_identity", "aws_credentials"],
        reason: "Taskpane provider secrets are not silently migrated; use an explicit companion auth move/setup flow when implemented.",
      },
      settingsSync: settingsSync ?? {
        state: "available",
        available: true,
        version: "companion-settings-sync-v1",
        secretsIncluded: false,
        reason: "Non-secret taskpane settings sync on session open or preference changes.",
      },
      nativeCapture: createNativeCaptureCapability(),
      mcp: {
        state: "available",
        available: true,
        version: "companion-mcp-v1",
        readOnly: true,
        toolCount: connectorToolCount,
      },
      memory: {
        state: "unavailable",
        available: false,
        version: "companion-memory-v1",
        reason: "Durable companion memory is reserved for advanced mode.",
      },
    },
  };
}

function uniqueStrings(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  return Array.from(new Set(value.filter((entry): entry is string => typeof entry === "string" && entry.trim().length > 0)))
    .map((entry) => entry.trim())
    .sort();
}

function normalizeSettingsSync(request: CompanionSettingsSyncRequest | undefined): CompanionSettingsSyncResponse {
  if (!request || typeof request !== "object") {
    throw new Error("settings sync request is required.");
  }
  if (request.secretsIncluded !== false) {
    throw new Error("settings sync must not include taskpane provider secrets.");
  }
  const mode = request.preferences?.companionRuntimeMode;
  if (mode !== "basic" && mode !== "smart_auto" && mode !== "advanced") {
    throw new Error("preferences.companionRuntimeMode must be basic, smart_auto, or advanced.");
  }
  const enabledProviders = uniqueStrings(request.providerSelection?.enabledProviders);
  const enabledModels = uniqueStrings(request.providerSelection?.enabledModels);
  const connectorCount = Number.isFinite(request.connectorCount) && request.connectorCount >= 0
    ? Math.floor(request.connectorCount)
    : 0;
  return {
    ok: true,
    syncedAt: new Date().toISOString(),
    companionRuntimeMode: mode,
    enabledProviderCount: enabledProviders.length,
    enabledModelCount: enabledModels.length,
    connectorCount,
    secretsIncluded: false,
  };
}

function settingsCapabilityFromSync(sync: CompanionSettingsSyncResponse): CompanionSettingsSyncCapability {
  return {
    state: "available",
    available: true,
    version: "companion-settings-sync-v1",
    lastSyncedAt: sync.syncedAt,
    companionRuntimeMode: sync.companionRuntimeMode,
    enabledProviderCount: sync.enabledProviderCount,
    enabledModelCount: sync.enabledModelCount,
    connectorCount: sync.connectorCount,
    secretsIncluded: false,
  };
}

export class CompanionServer {
  private readonly config = loadConfig();
  private readonly oauthBroker = new CompanionOAuthBroker(this.config);
  private readonly connectorBridge = new CompanionConnectorBridge(this.oauthBroker);
  private readonly sessionsByBrowserId = new Map<string, SessionRecord>();
  private readonly sessionsById = new Map<string, SessionRecord>();

  async start(): Promise<void> {
    mkdirSync(this.config.dataDir, { recursive: true });

    const app = express();
    app.use(companionCorsMiddleware);
    app.use(express.json({ limit: "10mb" }));
    this.mountRoutes(app);

    const server = createHttpsServer(
      {
        pfx: this.config.tls.pfx,
        passphrase: this.config.tls.passphrase,
      },
      app,
    );

    await new Promise<void>((resolve) => server.listen(this.config.port, this.config.host, resolve));
    console.log(`Pi-Office companion running at ${this.config.endpoint}`);
  }

  private mountRoutes(app: express.Express): void {
    app.get("/v1/health", (_request, response) => {
      const body: CompanionHealthResponse = {
        ok: true,
        endpoint: this.config.endpoint,
        identity: this.config.identity,
        capabilities: createCompanionState(this.config, undefined, []).capabilities,
      };
      response.json(body);
    });

    app.get("/v1/diagnostics", (_request, response) => {
      response.json(createCompanionRuntimeDiagnostics());
    });

    app.get("/v1/shell/capability", (_request, response) => {
      response.json(createShellSandbox(this.config).getCapability());
    });

    app.post("/v1/connectors/probe", async (request, response) => {
      try {
        const result = await this.connectorBridge.probeConnector(request.body);
        response.json(result);
      } catch (error) {
        response.status(400).json({
          error: error instanceof Error ? error.message : String(error),
        });
      }
    });

    app.post("/v1/connectors/oauth/start", async (request, response) => {
      try {
        const body = request.body as CompanionConnectorOAuthStartRequest;
        const result = await this.oauthBroker.start(body.definition);
        response.json(result);
      } catch (error) {
        response.status(400).json({
          error: error instanceof Error ? error.message : String(error),
        });
      }
    });

    app.post("/v1/connectors/oauth/status", (request, response) => {
      try {
        response.json(this.oauthBroker.status(request.body as CompanionConnectorOAuthStatusRequest));
      } catch (error) {
        response.status(400).json({
          error: error instanceof Error ? error.message : String(error),
        });
      }
    });

    app.delete("/v1/connectors/oauth/tokens", (request, response) => {
      try {
        response.json(this.oauthBroker.clearTokens(request.body as CompanionConnectorOAuthClearRequest | undefined));
      } catch (error) {
        response.status(400).json({
          error: error instanceof Error ? error.message : String(error),
        });
      }
    });

    app.get("/v1/connectors/oauth/callback", async (request, response) => {
      const result = await this.oauthBroker.completeCallback(request.query as Record<string, unknown>);
      response.status(result.statusCode).type("html").send(result.html);
    });

    app.post("/v1/sessions/open", async (request, response) => {
      const body = request.body as CompanionSessionOpenRequest;
      if (!body?.browserSessionId || !body?.host || !body?.documentId) {
        response.status(400).json({ error: "browserSessionId, host, and documentId are required." });
        return;
      }

      let session = this.sessionsByBrowserId.get(body.browserSessionId);
      if (!session) {
        session = {
          id: randomUUID(),
          browserSessionId: body.browserSessionId,
          host: body.host,
          documentId: body.documentId,
          title: body.title,
          saved: body.saved,
          connectorToolNames: [],
        };
        this.sessionsByBrowserId.set(body.browserSessionId, session);
        this.sessionsById.set(session.id, session);
      }

      session.windowId = body.windowId;
      session.host = body.host;
      session.documentId = body.documentId;
      session.documentPath = body.documentPath;
      session.documentUrl = body.documentUrl;
      session.saved = body.saved;
      session.title = body.title;
      session.workspaceDir = body.saved && body.documentPath ? dirname(body.documentPath) : undefined;

      const prepared = await this.connectorBridge.prepareSession(session.id, body.connectors ?? []);
      session.connectorToolNames = prepared.connectorToolNames;
      if (body.settings) {
        session.settingsSync = settingsCapabilityFromSync(normalizeSettingsSync(body.settings));
      }

      const reply: CompanionSessionOpenResponse = {
        ok: true,
        sessionId: session.id,
        companion: createCompanionState(
          this.config,
          session.id,
          prepared.connectorToolNames,
          createShellSandbox(this.config, session).getCapability(),
          session.settingsSync,
        ),
        connectors: prepared.connectors,
      };
      response.json(reply);
    });

    app.post("/v1/sessions/:sessionId/settings/sync", (request, response) => {
      const session = this.sessionsById.get(request.params.sessionId);
      if (!session) {
        response.status(404).json({ error: "Unknown companion session." });
        return;
      }

      try {
        const result = normalizeSettingsSync(request.body as CompanionSettingsSyncRequest);
        session.settingsSync = settingsCapabilityFromSync(result);
        response.json(result);
      } catch (error) {
        response.status(400).json({
          error: error instanceof Error ? error.message : String(error),
        });
      }
    });

    app.get("/v1/sessions/:sessionId/shell/capability", (request, response) => {
      const session = this.sessionsById.get(request.params.sessionId);
      if (!session) {
        response.status(404).json({ error: "Unknown companion session." });
        return;
      }

      response.json(createShellSandbox(this.config, session).getCapability());
    });

    app.post("/v1/sessions/:sessionId/shell/execute", async (request, response) => {
      const session = this.sessionsById.get(request.params.sessionId);
      if (!session) {
        response.status(404).json({ error: "Unknown companion session." });
        return;
      }

      try {
        const body = request.body as CompanionShellExecuteRequest;
        if (!body || typeof body.command !== "string" || !body.command.trim()) {
          response.status(400).json({ error: "command is required." });
          return;
        }
        const result = await createShellSandbox(this.config, session).execute(body);
        response.json(result);
      } catch (error) {
        response.status(400).json({
          error: error instanceof Error ? error.message : String(error),
        });
      }
    });

    app.post("/v1/sessions/:sessionId/native-capture/viewport", (request, response) => {
      const session = this.sessionsById.get(request.params.sessionId);
      if (!session) {
        response.status(404).json({ error: "Unknown companion session." });
        return;
      }

      const result = captureNativeViewport(session, request.body as CompanionNativeCaptureRequest);
      response.json(result);
    });

    app.post("/v1/sessions/:sessionId/agent/prompt", (_request, response) => {
      response.status(501).json({
        ok: false,
        error:
          "Companion-owned Pi agent sessions require explicit companion provider auth setup. Smart Auto will keep using the taskpane runtime until that capability is available.",
      });
    });

    app.post("/v1/sessions/:sessionId/agent/office-tool-result", (_request, response) => {
      response.status(501).json({
        ok: false,
        error:
          "Companion-owned Office tool proxying is reserved for companion agent mode. Office.js execution remains taskpane-owned.",
      });
    });

    app.post("/v1/sessions/:sessionId/files/:toolName", async (request, response) => {
      const session = this.sessionsById.get(request.params.sessionId);
      if (!session) {
        response.status(404).json({ error: "Unknown companion session." });
        return;
      }

      const toolName = request.params.toolName;
      if (toolName !== "read" && toolName !== "grep" && toolName !== "find" && toolName !== "ls") {
        response.status(404).json({ error: "Unknown file tool." });
        return;
      }

      try {
        const result = await executeFileTool(session.workspaceDir, toolName, request.body ?? {});
        response.json(result);
      } catch (error) {
        response.status(400).json({
          error: error instanceof Error ? error.message : String(error),
        });
      }
    });

    app.post("/v1/sessions/:sessionId/mcp/execute", async (request, response) => {
      const session = this.sessionsById.get(request.params.sessionId);
      if (!session) {
        response.status(404).json({ error: "Unknown companion session." });
        return;
      }

      const toolName = typeof request.body?.toolName === "string" ? request.body.toolName : "";
      if (!toolName) {
        response.status(400).json({ error: "toolName is required." });
        return;
      }

      try {
        const result = await this.connectorBridge.executePreparedTool(
          session.id,
          toolName,
          typeof request.body?.arguments === "object" && request.body.arguments && !Array.isArray(request.body.arguments)
            ? request.body.arguments as Record<string, unknown>
            : {},
        );
        response.json(result);
      } catch (error) {
        response.status(400).json({
          error: error instanceof Error ? error.message : String(error),
        });
      }
    });

    app.post("/v1/sessions/:sessionId/mcp/search", (request, response) => {
      const session = this.sessionsById.get(request.params.sessionId);
      if (!session) {
        response.status(404).json({ error: "Unknown companion session." });
        return;
      }

      try {
        const body = request.body as McpToolSearchRequest;
        response.json({
          query: body?.query,
          results: this.connectorBridge.searchPreparedTools(session.id, body),
        });
      } catch (error) {
        response.status(400).json({
          error: error instanceof Error ? error.message : String(error),
        });
      }
    });

    app.post("/v1/sessions/:sessionId/mcp/results/get", (request, response) => {
      const session = this.sessionsById.get(request.params.sessionId);
      if (!session) {
        response.status(404).json({ error: "Unknown companion session." });
        return;
      }

      try {
        response.json(this.connectorBridge.getMcpResult(request.body as McpResultPageRequest));
      } catch (error) {
        response.status(400).json({
          error: error instanceof Error ? error.message : String(error),
        });
      }
    });

    app.post("/v1/sessions/:sessionId/mcp/results/summarize", (request, response) => {
      const session = this.sessionsById.get(request.params.sessionId);
      if (!session) {
        response.status(404).json({ error: "Unknown companion session." });
        return;
      }

      try {
        response.json(this.connectorBridge.summarizeMcpResult(request.body as McpResultSummarizeRequest));
      } catch (error) {
        response.status(400).json({
          error: error instanceof Error ? error.message : String(error),
        });
      }
    });

    app.delete("/v1/sessions/:sessionId/mcp/results", (request, response) => {
      const session = this.sessionsById.get(request.params.sessionId);
      if (!session) {
        response.status(404).json({ error: "Unknown companion session." });
        return;
      }

      try {
        response.json(this.connectorBridge.clearMcpResults(request.body as McpResultClearRequest | undefined));
      } catch (error) {
        response.status(400).json({
          error: error instanceof Error ? error.message : String(error),
        });
      }
    });

  }
}
