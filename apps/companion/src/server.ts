import { randomUUID } from "node:crypto";
import { mkdirSync } from "node:fs";
import { createServer as createHttpsServer } from "node:https";
import { dirname } from "node:path";
import express from "express";
import type {
  CompanionHealthResponse,
  CompanionSessionOpenRequest,
  CompanionSessionOpenResponse,
  CompanionState,
} from "@pi-office/pi-office-pack/protocol";
import { loadConfig, type CompanionConfig } from "./config.js";
import { CompanionConnectorBridge } from "./connector-bridge.js";
import { executeFileTool } from "./file-tools.js";

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
}

function createCompanionState(
  config: CompanionConfig,
  sessionId?: string | undefined,
  connectorToolNames?: string[] | undefined,
): CompanionState {
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
    },
  };
}

export class CompanionServer {
  private readonly config = loadConfig();
  private readonly connectorBridge = new CompanionConnectorBridge();
  private readonly sessionsByBrowserId = new Map<string, SessionRecord>();
  private readonly sessionsById = new Map<string, SessionRecord>();

  async start(): Promise<void> {
    mkdirSync(this.config.dataDir, { recursive: true });

    const app = express();
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
        capabilities: {
          fileRead: true,
          localMcp: true,
          endpoint: this.config.endpoint,
        },
      };
      response.json(body);
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

      const reply: CompanionSessionOpenResponse = {
        ok: true,
        sessionId: session.id,
        companion: createCompanionState(this.config, session.id, prepared.connectorToolNames),
        connectors: prepared.connectors,
      };
      response.json(reply);
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
  }
}
