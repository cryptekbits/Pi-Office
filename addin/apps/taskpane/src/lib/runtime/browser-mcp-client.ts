import type {
  ConnectorCapabilitySummary,
  ConnectorReadPolicy,
  ConnectorRemoteHttpHeader,
  ConnectorToolClassification,
  ConnectorToolInventoryItem,
  ConnectorToolPolicyOverride,
  ConnectorVerificationSnapshot,
} from "@pi-office/pi-office-pack/protocol";

const MCP_PROTOCOL_VERSION = "2025-06-18";

interface JsonRpcResponse<T = unknown> {
  jsonrpc?: "2.0";
  id?: string | number;
  result?: T;
  error?: { code?: number; message?: string; data?: unknown };
}

interface McpTool {
  name?: string;
  description?: string;
  inputSchema?: unknown;
  annotations?: ConnectorToolInventoryItem["annotations"];
}

export interface BrowserMcpConnectorConfig {
  connectorId: string;
  name: string;
  url: string;
  accessToken?: string | undefined;
  headers?: ConnectorRemoteHttpHeader[] | undefined;
  readPolicy: ConnectorReadPolicy;
  toolPolicyOverrides?: ConnectorToolPolicyOverride[] | undefined;
  catalogRevision?: string | undefined;
}

export interface BrowserMcpProbeResult {
  ok: boolean;
  diagnostics: string[];
  capabilities?: ConnectorCapabilitySummary | undefined;
  verification?: ConnectorVerificationSnapshot | undefined;
}

function makeInventoryHash(values: string[]): string {
  return values.join("|").slice(0, 240) || "empty";
}

function safePatternTest(pattern: string, value: string): boolean {
  try {
    return new RegExp(pattern, "i").test(value);
  } catch {
    return false;
  }
}

function matchesAny(patterns: string[] | undefined, value: string): boolean {
  return (patterns ?? []).some((pattern) => safePatternTest(pattern, value));
}

function classifyTool(tool: McpTool, readPolicy: ConnectorReadPolicy): { classification: ConnectorToolClassification; reason: string } {
  const name = String(tool.name ?? "");
  const text = `${name} ${tool.description ?? ""}`;
  const lower = text.toLowerCase();
  if (tool.annotations?.destructiveHint === true || /\b(delete|drop|truncate|destroy|remove|refund|charge|revoke)\b/.test(lower)) {
    return { classification: "destructive", reason: "Tool is destructive by MCP annotation or name." };
  }
  if (matchesAny(readPolicy.blockToolPatterns, name)) {
    const destructive = /\b(delete|drop|truncate|destroy|remove|refund|charge|revoke)\b/i.test(name);
    return {
      classification: destructive ? "destructive" : "write",
      reason: "Tool matches the connector read-only block list.",
    };
  }
  if (tool.annotations?.readOnlyHint === true || matchesAny(readPolicy.allowToolPatterns, name)) {
    if (/\b(secret|token|credential|password|private|email|message|transcript|meeting|customer|payment|invoice)\b/.test(lower)) {
      return { classification: "sensitive_read", reason: "Read-only tool may return sensitive workspace or account data." };
    }
    if (tool.annotations?.openWorldHint === true || /\b(web|search|crawl|query|research|fetch)\b/.test(lower)) {
      return { classification: "costly_read", reason: "Read-only tool may call a remote search or data service." };
    }
    return { classification: "read", reason: "Tool is read-only by MCP annotation or curated allow pattern." };
  }
  return { classification: "unknown", reason: "Tool did not provide enough metadata to prove it is read-only." };
}

function enabledByDefault(classification: ConnectorToolClassification): boolean {
  return classification === "read" || classification === "sensitive_read" || classification === "costly_read";
}

function applyOverrides(
  inventory: ConnectorToolInventoryItem[],
  overrides: ConnectorToolPolicyOverride[] | undefined,
): ConnectorToolInventoryItem[] {
  if (!overrides?.length) return inventory;
  const byName = new Map(overrides.map((entry) => [entry.toolName, entry]));
  return inventory.map((tool) => {
    const override = byName.get(tool.name);
    return override ? { ...tool, enabled: override.enabled } : tool;
  });
}

function buildCapabilitySummary(
  tools: McpTool[],
  readPolicy: ConnectorReadPolicy,
  overrides: ConnectorToolPolicyOverride[] | undefined,
): ConnectorCapabilitySummary {
  const inventory = applyOverrides(
    tools
      .filter((tool): tool is Required<Pick<McpTool, "name">> & McpTool => typeof tool.name === "string" && tool.name.trim().length > 0)
      .map((tool): ConnectorToolInventoryItem => {
        const classified = classifyTool(tool, readPolicy);
        const defaultEnabled = enabledByDefault(classified.classification);
        return {
          name: tool.name,
          rawName: tool.name,
          description: tool.description,
          inputSchema: tool.inputSchema,
          annotations: tool.annotations,
          classification: classified.classification,
          defaultEnabled,
          enabled: defaultEnabled,
          reason: classified.reason,
          source: "mcp_tool",
        };
      }),
    overrides,
  );
  const names = inventory.map((tool) => tool.name);
  return {
    tools: names,
    allowedTools: inventory.filter((tool) => tool.enabled).map((tool) => tool.name),
    blockedTools: inventory.filter((tool) => !tool.enabled).map((tool) => tool.name),
    toolInventory: inventory,
    resourceToolNames: [],
    promptNames: [],
    allowedPrompts: [],
    blockedPrompts: [],
    resourceCount: 0,
    promptCount: 0,
    inventoryHash: makeInventoryHash(names),
  };
}

function responseError(response: Response, bodyText: string): Error {
  const trimmed = bodyText.trim();
  return new Error(trimmed ? `HTTP ${response.status}: ${trimmed.slice(0, 240)}` : `HTTP ${response.status}`);
}

async function parseJsonRpcResponse<T>(response: Response): Promise<JsonRpcResponse<T> | undefined> {
  const bodyText = await response.text();
  if (!response.ok) {
    throw responseError(response, bodyText);
  }
  const trimmed = bodyText.trim();
  if (!trimmed) {
    return undefined;
  }
  const contentType = response.headers.get("content-type") ?? "";
  if (contentType.includes("application/json")) {
    return JSON.parse(trimmed) as JsonRpcResponse<T>;
  }
  for (const rawLine of trimmed.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line.startsWith("data:")) continue;
    const data = line.slice("data:".length).trim();
    if (!data || data === "[DONE]") continue;
    try {
      const parsed = JSON.parse(data) as JsonRpcResponse<T>;
      if (parsed.result !== undefined || parsed.error) return parsed;
    } catch {
      // Ignore non-JSON SSE comments or keep-alives.
    }
  }
  try {
    return JSON.parse(trimmed) as JsonRpcResponse<T>;
  } catch {
    throw new Error("MCP server returned a response Pi-Office could not parse.");
  }
}

class BrowserMcpClient {
  private readonly endpoint: string;
  private readonly accessToken?: string | undefined;
  private readonly staticHeaders: ConnectorRemoteHttpHeader[];
  private sessionId?: string | undefined;
  private requestCounter = 0;

  constructor(config: BrowserMcpConnectorConfig) {
    this.endpoint = config.url;
    this.accessToken = config.accessToken;
    this.staticHeaders = config.headers ?? [];
  }

  async initialize(): Promise<void> {
    await this.call("initialize", {
      protocolVersion: MCP_PROTOCOL_VERSION,
      capabilities: {},
      clientInfo: {
        name: "Pi-Office",
        version: "0.1.0",
      },
    });
    await this.notify("notifications/initialized", {});
  }

  async listTools(): Promise<McpTool[]> {
    const result = await this.call<{ tools?: McpTool[] }>("tools/list", {});
    return Array.isArray(result?.tools) ? result.tools : [];
  }

  async callTool(name: string, args: Record<string, unknown>): Promise<unknown> {
    return this.call("tools/call", {
      name,
      arguments: args,
    });
  }

  private headers(): Headers {
    const headers = new Headers();
    headers.set("content-type", "application/json");
    headers.set("accept", "application/json, text/event-stream");
    if (this.accessToken) {
      headers.set("authorization", `Bearer ${this.accessToken}`);
    }
    if (this.sessionId) {
      headers.set("mcp-session-id", this.sessionId);
    }
    for (const header of this.staticHeaders) {
      if (header.name.trim() && header.value.trim()) {
        headers.set(header.name.trim(), header.value);
      }
    }
    return headers;
  }

  private async post(payload: unknown): Promise<Response> {
    if (typeof fetch !== "function") {
      throw new Error("Browser fetch is unavailable for this MCP connector.");
    }
    const response = await fetch(this.endpoint, {
      method: "POST",
      headers: this.headers(),
      body: JSON.stringify(payload),
    });
    const sessionId = response.headers.get("mcp-session-id");
    if (sessionId) this.sessionId = sessionId;
    return response;
  }

  private async call<T>(method: string, params: Record<string, unknown>): Promise<T> {
    const id = `browser-mcp-${++this.requestCounter}`;
    const response = await this.post({ jsonrpc: "2.0", id, method, params });
    const parsed = await parseJsonRpcResponse<T>(response);
    if (!parsed) {
      return undefined as T;
    }
    if (parsed.error) {
      throw new Error(parsed.error.message ?? `MCP ${method} failed.`);
    }
    return parsed.result as T;
  }

  private async notify(method: string, params: Record<string, unknown>): Promise<void> {
    const response = await this.post({ jsonrpc: "2.0", method, params });
    if (response.status === 202 || response.status === 204) {
      return;
    }
    const parsed = await parseJsonRpcResponse(response);
    if (parsed?.error) {
      throw new Error(parsed.error.message ?? `MCP ${method} failed.`);
    }
  }
}

export async function probeBrowserMcpConnector(config: BrowserMcpConnectorConfig): Promise<BrowserMcpProbeResult> {
  try {
    const client = new BrowserMcpClient(config);
    await client.initialize();
    const tools = await client.listTools();
    const capabilities = buildCapabilitySummary(tools, config.readPolicy, config.toolPolicyOverrides);
    const toolNames = capabilities.tools;
    const verifiedAt = new Date().toISOString();
    return {
      ok: true,
      diagnostics: [`Discovered ${toolNames.length} MCP tools through browser-direct HTTP.`],
      capabilities,
      verification: {
        verifiedAt,
        catalogRevision: config.catalogRevision ?? "browser-direct-v1",
        inventoryHash: capabilities.inventoryHash ?? makeInventoryHash(toolNames),
        toolNames,
        allowedTools: capabilities.allowedTools,
        blockedTools: capabilities.blockedTools,
        toolInventory: capabilities.toolInventory,
        resourceToolNames: [],
        promptNames: [],
        allowedPrompts: [],
        blockedPrompts: [],
        resourceCount: 0,
        promptCount: 0,
      },
    };
  } catch (error) {
    return {
      ok: false,
      diagnostics: [error instanceof Error ? error.message : "Browser-direct MCP verification failed."],
    };
  }
}

export async function executeBrowserMcpTool(
  config: BrowserMcpConnectorConfig,
  toolName: string,
  args: Record<string, unknown>,
): Promise<unknown> {
  const probe = await probeBrowserMcpConnector(config);
  if (!probe.ok || !probe.capabilities) {
    throw new Error(probe.diagnostics[0] ?? "Browser-direct MCP connector is not verified.");
  }
  if (!probe.capabilities.allowedTools.includes(toolName)) {
    throw new Error(`Connector tool "${toolName}" is disabled by Pi-Office read-only policy.`);
  }
  const client = new BrowserMcpClient(config);
  await client.initialize();
  return client.callTool(toolName, args);
}
