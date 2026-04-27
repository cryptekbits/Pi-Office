import { createHash, randomUUID } from "node:crypto";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { SSEClientTransport } from "@modelcontextprotocol/sdk/client/sse.js";
import { DEFAULT_INHERITED_ENV_VARS, StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import type {
  CompanionConnectorDefinition,
  McpBatchExecuteRequest,
  McpBatchExecuteResponse,
  ConnectorCredentialSource,
  ConnectorDiagnostic,
  ConnectorRemoteHttpHeader,
  ConnectorRemoteHttpHeaderFromEnv,
  ConnectorStatus,
  ConnectorMcpToolAnnotations,
  ConnectorToolClassification,
  ConnectorToolInventoryItem,
  ConnectorVerificationSnapshot,
  ConnectorHealthState,
  McpToolSearchRequest,
  McpToolSearchResponse,
  ToolCapabilitySearchResult,
} from "@pi-office/pi-office-pack/protocol";

type ProbeTransport = StdioClientTransport | StreamableHTTPClientTransport | SSEClientTransport;

interface ResolvedConnectorRuntime {
  transport: "local_stdio" | "remote_http";
  command?: string | undefined;
  args?: string[] | undefined;
  cwd?: string | undefined;
  env?: Record<string, string> | undefined;
  url?: string | undefined;
  bearerToken?: string | undefined;
  requestHeaders?: Record<string, string> | undefined;
}

interface PreparedExecutionTarget {
  exposedToolName: string;
  connectorId: string;
  kind: "tool" | "resource";
  rawName?: string | undefined;
  resourceUri?: string | undefined;
  definition: CompanionConnectorDefinition;
  runtime: ResolvedConnectorRuntime;
}

interface ProbeInventory {
  verification: ConnectorVerificationSnapshot;
  allowedTools: PreparedExecutionTarget[];
}

export interface McpToolInfo {
  name: string;
  description?: string | undefined;
  inputSchema?: unknown;
  annotations?: ConnectorMcpToolAnnotations | undefined;
}

type EnvironmentSource = Record<string, string | undefined>;
type MissingCredentialReason = "missing_value" | "local_stdio_env_key_required";

interface ResolvedConnectorCredential {
  source: ConnectorCredentialSource;
  value?: string | undefined;
  envKey?: string | undefined;
}

export interface LocalStdioEnvironmentResolution {
  env: Record<string, string>;
  credential: ResolvedConnectorCredential;
  credentialInjected: boolean;
  missingCredentialReason?: MissingCredentialReason | undefined;
}

interface PreparedSession {
  sessionId: string;
  connectors: ConnectorStatus[];
  connectorToolNames: string[];
  executionTargets: Map<string, PreparedExecutionTarget>;
}

export interface CompanionOAuthTokenProvider {
  getAccessToken(definition: CompanionConnectorDefinition): Promise<string | undefined>;
}

function trimString(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed ? trimmed : undefined;
}

function readEnvironmentValue(env: EnvironmentSource, key: string): string | undefined {
  const direct = env[key];
  if (typeof direct === "string") return direct;

  const lowerKey = key.toLowerCase();
  const match = Object.keys(env).find((item) => item.toLowerCase() === lowerKey);
  if (!match) return undefined;
  const value = env[match];
  return typeof value === "string" ? value : undefined;
}

function resolveEnvironmentCredential(
  env: EnvironmentSource,
  envKey: string | undefined,
  source: ConnectorCredentialSource,
): ResolvedConnectorCredential {
  if (!envKey) {
    return { source };
  }
  const value = trimString(readEnvironmentValue(env, envKey));
  return value ? { source, envKey, value } : { source, envKey };
}

function sanitizeServerName(input: string): string {
  return input
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "") || "connector";
}

function formatToolName(serverName: string, toolName: string): string {
  return `${serverName.replace(/-/g, "_")}_${toolName}`;
}

function resourceNameToToolName(name: string): string {
  let result = name
    .replace(/[^a-zA-Z0-9]/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "")
    .toLowerCase();
  if (!result) result = "resource";
  if (/^\d/.test(result)) result = `resource_${result}`;
  return result;
}

function matchesAnyPattern(value: string, patterns: string[]): boolean {
  return patterns.some((pattern) => new RegExp(pattern, "i").test(value));
}

function toInventoryHash(definition: CompanionConnectorDefinition, toolNames: string[], promptNames: string[]): string {
  return createHash("sha1")
    .update(JSON.stringify({
      id: definition.id,
      connectorId: definition.connectorId,
      transport: definition.transport,
      command: definition.command,
      args: definition.args,
      url: definition.url,
      stdioEnvPassthrough: definition.stdioEnvPassthrough,
      remoteHttpHeaders: definition.remoteHttpHeaders,
      remoteHttpHeadersFromEnv: definition.remoteHttpHeadersFromEnv?.map((entry) => ({
        name: entry.name,
        envVarName: entry.envVarName,
      })),
      tools: toolNames,
      prompts: promptNames,
    }))
    .digest("hex");
}

function normalizeRemoteHttpHeaders(entries: ConnectorRemoteHttpHeader[] | undefined): Record<string, string> {
  const headers: Record<string, string> = {};
  for (const entry of entries ?? []) {
    const name = trimString(entry.name);
    const value = trimString(entry.value);
    if (name && value) {
      headers[name] = value;
    }
  }
  return headers;
}

function normalizeRemoteHttpHeadersFromEnv(
  entries: ConnectorRemoteHttpHeaderFromEnv[] | undefined,
  env: EnvironmentSource,
): Record<string, string> {
  const headers: Record<string, string> = {};
  for (const entry of entries ?? []) {
    const name = trimString(entry.name);
    const envVarName = trimString(entry.envVarName);
    if (!name || !envVarName) {
      continue;
    }
    const value = trimString(readEnvironmentValue(env, envVarName));
    if (value) {
      headers[name] = value;
    }
  }
  return headers;
}

/** Build merged HTTP headers for remote MCP (static + env-sourced). Exported for tests. */
export function buildRemoteHttpRequestHeaders(
  definition: Pick<CompanionConnectorDefinition, "remoteHttpHeaders" | "remoteHttpHeadersFromEnv">,
  env: EnvironmentSource = process.env,
): Record<string, string> {
  return {
    ...normalizeRemoteHttpHeadersFromEnv(definition.remoteHttpHeadersFromEnv, env),
    ...normalizeRemoteHttpHeaders(definition.remoteHttpHeaders),
  };
}

function resolveCredential(
  definition: CompanionConnectorDefinition,
  env: EnvironmentSource = process.env,
): ResolvedConnectorCredential {
  if (definition.authMethod === "none" || definition.credentialSource === "none") {
    return { source: "none" };
  }

  const detected = trimString(definition.useDetectedEnvKey);
  const explicitEnv = trimString(definition.secretEnvKey);
  const envKey = explicitEnv ?? detected;

  if (definition.credentialSource === "detected_env") {
    return resolveEnvironmentCredential(env, detected ?? explicitEnv, "detected_env");
  }

  if (definition.credentialSource === "env") {
    return resolveEnvironmentCredential(env, explicitEnv ?? detected, "env");
  }

  if (definition.credentialSource === "manual") {
    const manual = trimString(definition.secret);
    return manual ? { source: "manual", envKey, value: manual } : { source: "manual", envKey };
  }

  if (definition.credentialSource === "oauth") {
    const token = trimString(definition.secret);
    return token ? { source: "oauth", value: token } : { source: "oauth" };
  }

  return { source: definition.credentialSource, envKey };
}

function missingCredentialReason(
  definition: CompanionConnectorDefinition,
  credential = resolveCredential(definition),
): MissingCredentialReason | undefined {
  if (definition.authMethod === "none") return undefined;
  if (!credential.value) return "missing_value";
  if (definition.transport === "local_stdio" && !credential.envKey) return "local_stdio_env_key_required";
  return undefined;
}

function needsCredential(definition: CompanionConnectorDefinition): boolean {
  return Boolean(missingCredentialReason(definition));
}

function isConfigured(definition: CompanionConnectorDefinition): boolean {
  if (definition.transport === "local_stdio") {
    return Boolean(trimString(definition.command));
  }
  return Boolean(trimString(definition.url));
}

function buildDefaultLocalStdioEnvironment(env: EnvironmentSource): Record<string, string> {
  const result: Record<string, string> = {};
  for (const key of [...DEFAULT_INHERITED_ENV_VARS, "USERPROFILE"]) {
    const value = readEnvironmentValue(env, key);
    if (value === undefined || value.startsWith("()")) {
      continue;
    }
    result[key] = value;
  }
  return result;
}

function mergeConnectorEnvironment(target: Record<string, string>, source: Record<string, string> | undefined): void {
  if (!source) return;
  for (const [rawKey, value] of Object.entries(source)) {
    const key = trimString(rawKey);
    if (!key || typeof value !== "string") {
      continue;
    }
    target[key] = value;
  }
}

export function resolveLocalStdioProcessEnvironment(
  definition: CompanionConnectorDefinition,
  env: EnvironmentSource = process.env,
): LocalStdioEnvironmentResolution {
  const credential = resolveCredential(definition, env);
  const processEnv = buildDefaultLocalStdioEnvironment(env);
  mergeConnectorEnvironment(processEnv, definition.env);

  for (const rawKey of definition.stdioEnvPassthrough ?? []) {
    const key = trimString(rawKey);
    if (!key) {
      continue;
    }
    const value = readEnvironmentValue(env, key);
    if (typeof value === "string" && value.length) {
      processEnv[key] = value;
    }
  }

  let credentialInjected = false;
  if (definition.authMethod !== "none" && credential.value && credential.envKey) {
    processEnv[credential.envKey] = credential.value;
    credentialInjected = true;
  }

  const reason = missingCredentialReason(definition, credential);
  return {
    env: processEnv,
    credential,
    credentialInjected,
    ...(reason ? { missingCredentialReason: reason } : {}),
  };
}

async function createRemoteTransport(runtime: ResolvedConnectorRuntime): Promise<ProbeTransport> {
  const headers: Record<string, string> = { ...(runtime.requestHeaders ?? {}) };
  if (runtime.bearerToken && !headers.Authorization) {
    headers.Authorization = `Bearer ${runtime.bearerToken}`;
  }

  const requestInit = Object.keys(headers).length ? ({ headers } satisfies RequestInit) : undefined;
  const url = new URL(runtime.url!);
  const probe = new StreamableHTTPClientTransport(url, requestInit ? { requestInit } : {});
  try {
    const client = new Client({ name: "pi-office-companion-probe", version: "1.0.0" });
    await client.connect(probe as never);
    await client.close().catch(() => {});
    await probe.close().catch(() => {});
    return new StreamableHTTPClientTransport(url, requestInit ? { requestInit } : {});
  } catch {
    await probe.close().catch(() => {});
    return new SSEClientTransport(url, requestInit ? { requestInit } : {});
  }
}

async function withClient<T>(
  runtime: ResolvedConnectorRuntime,
  fn: (client: Client) => Promise<T>,
): Promise<T> {
  const client = new Client({ name: "pi-office-companion", version: "1.0.0" });
  const transport =
    runtime.transport === "local_stdio"
      ? new StdioClientTransport({
          command: runtime.command!,
          args: runtime.args ?? [],
          ...(runtime.cwd ? { cwd: runtime.cwd } : {}),
          ...(runtime.env ? { env: runtime.env } : {}),
          stderr: "ignore",
        })
      : await createRemoteTransport(runtime);

  try {
    await client.connect(transport as never);
    return await fn(client);
  } finally {
    await client.close().catch(() => {});
    await transport.close().catch(() => {});
  }
}

async function fetchAllTools(client: Client): Promise<McpToolInfo[]> {
  const tools: McpToolInfo[] = [];
  let cursor: string | undefined;
  do {
    const result = await client.listTools(cursor ? { cursor } : undefined);
    tools.push(...(result.tools ?? []).map((tool): McpToolInfo => ({
      name: tool.name,
      ...(tool.description ? { description: tool.description } : {}),
      ...(tool.inputSchema ? { inputSchema: tool.inputSchema } : {}),
      ...(tool.annotations ? { annotations: tool.annotations as ConnectorMcpToolAnnotations } : {}),
    })));
    cursor = result.nextCursor;
  } while (cursor);
  return tools;
}

async function fetchAllResources(client: Client): Promise<Array<{ name: string; uri: string }>> {
  try {
    const resources: Array<{ name: string; uri: string }> = [];
    let cursor: string | undefined;
    do {
      const result = await client.listResources(cursor ? { cursor } : undefined);
      resources.push(...(result.resources ?? []).map((resource: { name: string; uri: string }) => ({
        name: resource.name,
        uri: resource.uri,
      })));
      cursor = result.nextCursor;
    } while (cursor);
    return resources;
  } catch {
    return [];
  }
}

async function fetchAllPrompts(client: Client): Promise<Array<{ name: string }>> {
  const listPrompts = (client as unknown as {
    listPrompts?: (input?: { cursor?: string }) => Promise<{ prompts?: Array<{ name: string }>; nextCursor?: string }>;
  }).listPrompts;
  if (!listPrompts) return [];

  try {
    const prompts: Array<{ name: string }> = [];
    let cursor: string | undefined;
    do {
      const result = await listPrompts.call(client, cursor ? { cursor } : undefined);
      prompts.push(...(result.prompts ?? []));
      cursor = result.nextCursor;
    } while (cursor);
    return prompts;
  } catch {
    return [];
  }
}

function resolveRuntime(definition: CompanionConnectorDefinition): ResolvedConnectorRuntime | undefined {
  if (!isConfigured(definition)) {
    return undefined;
  }

  const credential = resolveCredential(definition);
  if (definition.transport === "local_stdio") {
    const resolvedEnvironment = resolveLocalStdioProcessEnvironment(definition);
    return {
      transport: definition.transport,
      command: definition.command,
      args: definition.args,
      cwd: definition.cwd,
      env: resolvedEnvironment.env,
    };
  }

  const requestHeaders = buildRemoteHttpRequestHeaders(definition);
  let bearerToken = trimString(credential.value);
  if (bearerToken && requestHeaders.Authorization) {
    bearerToken = undefined;
  }

  return {
    transport: definition.transport,
    url: definition.url,
    bearerToken: bearerToken || undefined,
    requestHeaders: Object.keys(requestHeaders).length ? requestHeaders : undefined,
  };
}

function destructiveName(name: string): boolean {
  return /(^|_)(delete|remove|drop|truncate|refund|charge|pay|revoke|archive|destroy)(_|$)/i.test(name);
}

export function classifyConnectorToolForPolicy(definition: Pick<CompanionConnectorDefinition, "readPolicy">, tool: McpToolInfo): {
  classification: ConnectorToolClassification;
  defaultEnabled: boolean;
  reason: string;
} {
  const annotations = tool.annotations;
  if (annotations?.destructiveHint === true) {
    return {
      classification: "destructive",
      defaultEnabled: false,
      reason: "MCP annotations mark this tool as destructive.",
    };
  }
  if (annotations?.readOnlyHint === true) {
    return {
      classification: annotations.openWorldHint ? "sensitive_read" : "read",
      defaultEnabled: true,
      reason: annotations.openWorldHint
        ? "MCP annotations mark this as read-only, but it can read outside the local document."
        : "MCP annotations mark this tool as read-only.",
    };
  }
  if (matchesAnyPattern(tool.name, definition.readPolicy.blockToolPatterns)) {
    return {
      classification: destructiveName(tool.name) ? "destructive" : "write",
      defaultEnabled: false,
      reason: "Catalog policy blocks this side-effecting tool by default.",
    };
  }
  if (matchesAnyPattern(tool.name, definition.readPolicy.allowToolPatterns)) {
    return {
      classification: "read",
      defaultEnabled: true,
      reason: "Catalog policy allows this as a read-safe tool.",
    };
  }
  return {
    classification: "unknown",
    defaultEnabled: false,
    reason: "Tool was not matched by read-safe policy and is disabled until reviewed.",
  };
}

function applyPolicyOverride(
  tool: ConnectorToolInventoryItem,
  definition: CompanionConnectorDefinition,
): ConnectorToolInventoryItem {
  const override = definition.toolPolicyOverrides?.find((entry) => entry.toolName === tool.name);
  return override ? { ...tool, enabled: override.enabled } : tool;
}

function buildVerification(
  definition: CompanionConnectorDefinition,
  serverName: string,
  tools: McpToolInfo[],
  resources: Array<{ name: string; uri: string }>,
  prompts: Array<{ name: string }>,
): ProbeInventory {
  const allowedToolTargets: PreparedExecutionTarget[] = [];
  const toolInventory: ConnectorToolInventoryItem[] = [];

  for (const tool of tools) {
    const exposedName = formatToolName(serverName, tool.name);
    const classified = classifyConnectorToolForPolicy(definition, tool);
    const inventoryItem = applyPolicyOverride({
      name: exposedName,
      rawName: tool.name,
      description: tool.description,
      inputSchema: tool.inputSchema,
      annotations: tool.annotations,
      classification: classified.classification,
      defaultEnabled: classified.defaultEnabled,
      enabled: classified.defaultEnabled,
      reason: classified.reason,
      source: "mcp_tool",
    }, definition);
    toolInventory.push(inventoryItem);
    if (!inventoryItem.enabled) continue;

    allowedToolTargets.push({
      exposedToolName: exposedName,
      connectorId: definition.id,
      kind: "tool",
      rawName: tool.name,
      definition,
      runtime: resolveRuntime(definition)!,
    });
  }

  if (definition.readPolicy.allowResources) {
    for (const resource of resources) {
      const exposedToolName = formatToolName(serverName, `get_${resourceNameToToolName(resource.name)}`);
      const inventoryItem = applyPolicyOverride({
        name: exposedToolName,
        rawName: resource.name,
        classification: "read",
        defaultEnabled: true,
        enabled: true,
        reason: "MCP resource helper is read-only.",
        source: "resource",
      }, definition);
      toolInventory.push(inventoryItem);
      if (!inventoryItem.enabled) continue;
      allowedToolTargets.push({
        exposedToolName,
        connectorId: definition.id,
        kind: "resource",
        resourceUri: resource.uri,
        definition,
        runtime: resolveRuntime(definition)!,
      });
    }
  }

  const allowedPromptNames = definition.readPolicy.allowPrompts
    ? prompts
        .map((prompt) => prompt.name)
        .filter((name) => !matchesAnyPattern(name, definition.readPolicy.blockPromptPatterns ?? []))
        .filter((name) => matchesAnyPattern(name, definition.readPolicy.allowPromptPatterns ?? [".*"]))
    : [];

  const promptNames = prompts.map((prompt) => prompt.name);
  const toolNames = toolInventory.map((tool) => tool.name);
  const allowedToolNames = allowedToolTargets.map((target) => target.exposedToolName);
  const blockedToolNames = toolInventory.filter((tool) => !allowedToolNames.includes(tool.name)).map((tool) => tool.name);

  return {
    verification: {
      verifiedAt: new Date().toISOString(),
      catalogRevision: definition.catalogRevision ?? "companion-v1",
      inventoryHash: toInventoryHash(definition, toolNames, promptNames),
      toolNames,
      allowedTools: allowedToolNames,
      blockedTools: blockedToolNames,
      toolInventory,
      resourceToolNames: allowedToolTargets
        .filter((target) => target.kind === "resource")
        .map((target) => target.exposedToolName),
      promptNames,
      allowedPrompts: allowedPromptNames,
      blockedPrompts: promptNames.filter((name) => !allowedPromptNames.includes(name)),
      resourceCount: resources.length,
      promptCount: promptNames.length,
    },
    allowedTools: allowedToolTargets,
  };
}

function buildDiagnostics(
  definition: CompanionConnectorDefinition,
  healthState: ConnectorHealthState,
  message?: string | undefined,
): ConnectorDiagnostic[] {
  const diagnostics: ConnectorDiagnostic[] = [];

  if (!trimString(definition.name)) {
    diagnostics.push({
      level: "error",
      code: "name_required",
      title: "Name required",
      message: "Give this connector a recognizable name before using it.",
      connectorId: definition.id,
    });
  }

  if (!isConfigured(definition)) {
    diagnostics.push({
      level: "error",
      code: definition.transport === "local_stdio" ? "command_required" : "url_required",
      title: definition.transport === "local_stdio" ? "Launch command required" : "Service URL required",
      message: definition.transport === "local_stdio"
        ? "Local MCP connectors need a launch command."
        : "Remote MCP connectors need a full service URL.",
      connectorId: definition.id,
    });
  }

  const credentialIssue = missingCredentialReason(definition);
  if (credentialIssue) {
    diagnostics.push({
      level: "warning",
      code: credentialIssue === "local_stdio_env_key_required" ? "credential_env_key_required" : "credential_required",
      title: credentialIssue === "local_stdio_env_key_required" ? "Credential variable required" : "Credentials still needed",
      message: credentialIssue === "local_stdio_env_key_required"
        ? "Local MCP connectors need an environment variable name so the companion can pass the saved credential to the process."
        : "This connector needs credentials before it can execute.",
      connectorId: definition.id,
    });
  }

  if (message) {
    diagnostics.push({
      level: healthState === "ready" ? "info" : "warning",
      code: healthState === "ready" ? "probe_ok" : "probe_failed",
      title: healthState === "ready" ? "Connector verified" : "Verification failed",
      message,
      connectorId: definition.id,
    });
  }

  return diagnostics;
}

function buildStatus(
  definition: CompanionConnectorDefinition,
  verification: ConnectorVerificationSnapshot | undefined,
  healthState: ConnectorHealthState,
  lastError?: string | undefined,
): ConnectorStatus {
  const enabled = true;
  const configured = isConfigured(definition);
  return {
    id: definition.id,
    connectorId: definition.connectorId,
    name: definition.name,
    enabled,
    configured,
    connected: healthState === "ready",
    source: definition.source,
    category: definition.category,
    maturity: definition.maturity,
    setupKind: definition.setupKind,
    authMethod: definition.authMethod,
    transport: definition.transport,
    setupProfileId: definition.setupProfileId,
    credentialSource: definition.credentialSource,
    detectedEnvKey: definition.useDetectedEnvKey ?? definition.secretEnvKey,
    usesDetectedCredential: Boolean(definition.useDetectedEnvKey),
    lastTestedAt: verification?.verifiedAt,
    lastError,
    lastHealthyAt: healthState === "ready" ? verification?.verifiedAt : undefined,
    healthState,
    staleReason: verification?.staleReason,
    activeScope: "global",
    scopeStates: [
      {
        target: "global",
        enabled,
        applies: true,
        label: "Everywhere",
      },
    ],
    favorite: false,
    recommendedHosts: ["word", "excel", "powerpoint"],
    setupDifficulty: "advanced",
    needsCredential: needsCredential(definition),
    verification,
    capabilities: verification
      ? {
          tools: verification.toolNames,
          allowedTools: verification.allowedTools,
          blockedTools: verification.blockedTools,
          toolInventory: verification.toolInventory,
          resourceToolNames: verification.resourceToolNames,
          promptNames: verification.promptNames,
          allowedPrompts: verification.allowedPrompts,
          blockedPrompts: verification.blockedPrompts,
          resourceCount: verification.resourceCount,
          promptCount: verification.promptCount,
          inventoryHash: verification.inventoryHash,
        }
      : undefined,
    executionEnvironment: "companion",
    executionAvailable: healthState === "ready",
  };
}

async function withBrokeredOAuthToken(
  definition: CompanionConnectorDefinition,
  oauthTokenProvider: CompanionOAuthTokenProvider | undefined,
): Promise<CompanionConnectorDefinition> {
  if (
    definition.authMethod !== "oauth" ||
    definition.credentialSource !== "oauth" ||
    definition.oauth?.broker !== "companion" ||
    trimString(definition.secret)
  ) {
    return definition;
  }
  const token = await oauthTokenProvider?.getAccessToken(definition);
  return token ? { ...definition, secret: token } : definition;
}

async function inspectConnector(
  definition: CompanionConnectorDefinition,
  oauthTokenProvider?: CompanionOAuthTokenProvider | undefined,
): Promise<{
  status: ConnectorStatus;
  diagnostics: ConnectorDiagnostic[];
  executionTargets: PreparedExecutionTarget[];
}> {
  definition = await withBrokeredOAuthToken(definition, oauthTokenProvider);
  if (!isConfigured(definition)) {
    const healthState: ConnectorHealthState = "unverified";
    return {
      status: buildStatus(definition, undefined, healthState, "Connector configuration is incomplete."),
      diagnostics: buildDiagnostics(definition, healthState, "Connector configuration is incomplete."),
      executionTargets: [],
    };
  }

  if (needsCredential(definition)) {
    const healthState: ConnectorHealthState = "auth_required";
    return {
      status: buildStatus(definition, undefined, healthState, "Connector credentials are missing."),
      diagnostics: buildDiagnostics(definition, healthState, "Connector credentials are missing."),
      executionTargets: [],
    };
  }

  const runtime = resolveRuntime(definition)!;
  const serverName = sanitizeServerName(definition.id);

  try {
    const { verification, allowedTools } = await withClient(runtime, async (client) => {
      const [tools, resources, prompts] = await Promise.all([
        fetchAllTools(client),
        fetchAllResources(client),
        fetchAllPrompts(client),
      ]);
      return buildVerification(definition, serverName, tools, resources, prompts);
    });

    const healthState: ConnectorHealthState =
      verification.allowedTools.length > 0 || verification.allowedPrompts.length > 0 ? "ready" : "degraded";
    const message =
      healthState === "ready"
        ? `${verification.allowedTools.length} read-safe tool${verification.allowedTools.length === 1 ? "" : "s"} available through the companion.`
        : "The connector responded, but no read-safe tools or prompts were exposed.";

    return {
      status: buildStatus(definition, verification, healthState, healthState === "ready" ? undefined : message),
      diagnostics: buildDiagnostics(definition, healthState, message),
      executionTargets: healthState === "ready" ? allowedTools : [],
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    const healthState: ConnectorHealthState = "offline";
    return {
      status: buildStatus(definition, undefined, healthState, message),
      diagnostics: buildDiagnostics(definition, healthState, message),
      executionTargets: [],
    };
  }
}

function resourceContentToToolResult(result: {
  contents?: Array<{ text?: string; uri?: string; mimeType?: string }>;
}) {
  const text = (result.contents ?? [])
    .map((entry) => {
      if (entry.text) return entry.text;
      if (entry.uri) return `${entry.mimeType ? `[${entry.mimeType}] ` : ""}${entry.uri}`;
      return "";
    })
    .filter(Boolean)
    .join("\n\n");

  return {
    content: [
      {
        type: "text" as const,
        text: text || JSON.stringify(result, null, 2),
      },
    ],
    details: result,
  };
}

export class CompanionConnectorBridge {
  constructor(private readonly oauthTokenProvider?: CompanionOAuthTokenProvider | undefined) {}

  private readonly preparedSessions = new Map<string, PreparedSession>();

  async probeConnector(definition: CompanionConnectorDefinition): Promise<{
    ok: boolean;
    status: ConnectorStatus;
    diagnostics: ConnectorDiagnostic[];
  }> {
    const inspected = await inspectConnector(definition, this.oauthTokenProvider);
    return {
      ok: inspected.status.healthState === "ready",
      status: inspected.status,
      diagnostics: inspected.diagnostics,
    };
  }

  async prepareSession(
    sessionId: string,
    connectors: CompanionConnectorDefinition[],
  ): Promise<PreparedSession> {
    const prepared: PreparedSession = {
      sessionId,
      connectors: [],
      connectorToolNames: [],
      executionTargets: new Map<string, PreparedExecutionTarget>(),
    };

    for (const definition of connectors) {
      const inspected = await inspectConnector(definition, this.oauthTokenProvider);
      prepared.connectors.push(inspected.status);
      for (const target of inspected.executionTargets) {
        prepared.executionTargets.set(target.exposedToolName, target);
      }
    }

    prepared.connectorToolNames = Array.from(prepared.executionTargets.keys()).sort((left, right) => left.localeCompare(right));
    this.preparedSessions.set(sessionId, prepared);
    return prepared;
  }

  getPreparedSession(sessionId: string): PreparedSession | undefined {
    return this.preparedSessions.get(sessionId);
  }

  clearSession(sessionId: string): void {
    this.preparedSessions.delete(sessionId);
  }

  async executePreparedTool(
    sessionId: string,
    toolName: string,
    args: Record<string, unknown>,
  ): Promise<unknown> {
    const prepared = this.preparedSessions.get(sessionId);
    if (!prepared) {
      throw new Error("Unknown companion connector session.");
    }

    const target = prepared.executionTargets.get(toolName);
    if (!target) {
      throw new Error(`Unknown or unavailable companion connector tool: ${toolName}`);
    }

    return withClient(target.runtime, async (client) => {
      if (target.kind === "resource") {
        const result = await client.readResource({ uri: target.resourceUri! });
        return resourceContentToToolResult(result as { contents?: Array<{ text?: string; uri?: string; mimeType?: string }> });
      }

      const result = await client.callTool({
        name: target.rawName!,
        arguments: args,
      });
      return {
        content: Array.isArray((result as { content?: unknown[] }).content)
          ? (result as { content: unknown[] }).content
          : [
              {
                type: "text",
                text: JSON.stringify(result, null, 2),
              },
            ],
        details: result,
      };
    });
  }

  searchPreparedTools(
    sessionId: string,
    request: McpToolSearchRequest = {},
  ): McpToolSearchResponse {
    const prepared = this.preparedSessions.get(sessionId);
    if (!prepared) return { query: request.query, results: [] };
    const limit = typeof request.limit === "number" && Number.isFinite(request.limit)
      ? Math.max(1, Math.min(25, Math.trunc(request.limit)))
      : 8;
    const connectorFilter = trimString(request.connectorId);
    const terms = String(request.query ?? "")
      .toLowerCase()
      .split(/\s+/g)
      .map((term) => term.trim())
      .filter(Boolean);

    const scored = Array.from(prepared.executionTargets.values()).filter((target) => {
      return !connectorFilter || target.definition.connectorId === connectorFilter || target.definition.id === connectorFilter;
    }).map((target) => {
      const inventory = target.definition.toolPolicyOverrides?.find((entry) => entry.toolName === target.exposedToolName);
      const haystack = [
        target.exposedToolName,
        target.rawName,
        target.definition.name,
        target.definition.connectorId,
        target.definition.category,
        target.definition.transport,
      ].filter(Boolean).join(" ").toLowerCase();
      const query = String(request.query ?? "");
      const terms = query
        .toLowerCase()
        .split(/\s+/g)
        .map((term) => term.trim())
        .filter(Boolean);
      const score = terms.length
        ? terms.reduce((sum, term) => sum + (haystack.includes(term) ? 2 : 0), 0)
        : 1;
      return {
        score,
        result: {
          id: `mcp:${target.definition.id}:${target.exposedToolName}`,
          toolName: target.exposedToolName,
          label: target.rawName ?? target.exposedToolName,
          description: target.kind === "resource"
            ? `Read MCP resource ${target.resourceUri} from ${target.definition.name}.`
            : `Execute MCP tool ${target.rawName ?? target.exposedToolName} from ${target.definition.name}.`,
          host: "all" as const,
          category: "connector" as const,
          runtime: "companion" as const,
          risk: "low" as const,
          support: "supported" as const,
          keywords: [
            target.definition.connectorId,
            target.definition.name,
            target.definition.transport,
            target.kind,
          ],
          score,
          source: "companion-mcp",
          connectorId: target.definition.connectorId,
          connectorName: target.definition.name,
          schemaAvailable: target.kind === "tool",
          fallback: inventory?.enabled === false ? "Connector policy currently disables this tool." : undefined,
        } satisfies ToolCapabilitySearchResult,
      };
    });

    const results = scored
      .filter((entry) => entry.score > 0)
      .sort((left, right) => right.score - left.score || left.result.toolName.localeCompare(right.result.toolName))
      .slice(0, limit)
      .map((entry) => entry.result);
    return { query: request.query, results };
  }
}
