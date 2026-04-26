import { createHash, randomUUID } from "node:crypto";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { SSEClientTransport } from "@modelcontextprotocol/sdk/client/sse.js";
import { DEFAULT_INHERITED_ENV_VARS, StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import type {
  CompanionConnectorDefinition,
  ConnectorCredentialSource,
  ConnectorDiagnostic,
  ConnectorStatus,
  ConnectorVerificationSnapshot,
  ConnectorHealthState,
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
      tools: toolNames,
      prompts: promptNames,
    }))
    .digest("hex");
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
  for (const key of DEFAULT_INHERITED_ENV_VARS) {
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
  const headers: Record<string, string> = {};
  if (runtime.bearerToken) {
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

async function fetchAllTools(client: Client): Promise<Array<{ name: string }>> {
  const tools: Array<{ name: string }> = [];
  let cursor: string | undefined;
  do {
    const result = await client.listTools(cursor ? { cursor } : undefined);
    tools.push(...(result.tools ?? []).map((tool: { name: string }) => ({ name: tool.name })));
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

  return {
    transport: definition.transport,
    url: definition.url,
    bearerToken: credential.value,
  };
}

function buildVerification(
  definition: CompanionConnectorDefinition,
  serverName: string,
  tools: Array<{ name: string }>,
  resources: Array<{ name: string; uri: string }>,
  prompts: Array<{ name: string }>,
): ProbeInventory {
  const allowedToolTargets: PreparedExecutionTarget[] = [];
  const blockedToolNames: string[] = [];

  for (const tool of tools) {
    const exposedName = formatToolName(serverName, tool.name);
    if (matchesAnyPattern(tool.name, definition.readPolicy.blockToolPatterns)) {
      blockedToolNames.push(exposedName);
      continue;
    }
    if (!matchesAnyPattern(tool.name, definition.readPolicy.allowToolPatterns)) {
      blockedToolNames.push(exposedName);
      continue;
    }

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
      allowedToolTargets.push({
        exposedToolName: formatToolName(serverName, `get_${resourceNameToToolName(resource.name)}`),
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
  const toolNames = allowedToolTargets.map((target) => target.exposedToolName).concat(blockedToolNames);

  return {
    verification: {
      verifiedAt: new Date().toISOString(),
      catalogRevision: definition.catalogRevision ?? "companion-v1",
      inventoryHash: toInventoryHash(definition, toolNames, promptNames),
      toolNames,
      allowedTools: allowedToolTargets.map((target) => target.exposedToolName),
      blockedTools: blockedToolNames,
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

async function inspectConnector(definition: CompanionConnectorDefinition): Promise<{
  status: ConnectorStatus;
  diagnostics: ConnectorDiagnostic[];
  executionTargets: PreparedExecutionTarget[];
}> {
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
  private readonly preparedSessions = new Map<string, PreparedSession>();

  async probeConnector(definition: CompanionConnectorDefinition): Promise<{
    ok: boolean;
    status: ConnectorStatus;
    diagnostics: ConnectorDiagnostic[];
  }> {
    const inspected = await inspectConnector(definition);
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
      const inspected = await inspectConnector(definition);
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
}
