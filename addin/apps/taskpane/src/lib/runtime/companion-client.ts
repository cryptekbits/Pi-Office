import type {
  CompanionConnectorDefinition,
  CompanionDiscoveryAttempt,
  CompanionConnectorOAuthStartRequest,
  CompanionConnectorOAuthStatusRequest,
  CompanionConnectorOAuthStatusResponse,
  CompanionHealthResponse,
  CompanionNativeCaptureRequest,
  CompanionNativeCaptureResponse,
  CompanionShellCapability,
  CompanionShellExecuteRequest,
  CompanionShellExecuteResponse,
  CompanionSessionOpenResponse,
  CompanionState,
  ConnectorDiagnosticsResponse,
  ConnectorDiagnostic,
  ConnectorOAuthStartResponse,
  ConnectorStatus,
  McpToolSearchRequest,
  McpToolSearchResponse,
  OfficeStateUpdate,
} from "@pi-office/pi-office-pack/protocol";
import { DEFAULT_COMPANION_HOST, DEFAULT_COMPANION_PORT } from "@pi-office/pi-office-pack/defaults";

const MANUAL_ENDPOINT_KEY = "pi-office-companion-manual-endpoint";
const LAST_SUCCESSFUL_ENDPOINT_KEY = "pi-office-companion-last-endpoint";
const DEFAULT_ENDPOINT = `https://${DEFAULT_COMPANION_HOST}:${DEFAULT_COMPANION_PORT}`;
const LOOPBACK_ENDPOINT = `https://127.0.0.1:${DEFAULT_COMPANION_PORT}`;
export const COMPANION_DISCOVERY_TIMEOUT_MS = 5_000;
const REQUEST_TIMEOUT_MS = 15_000;

export interface CompanionSessionBinding {
  browserSessionId: string;
  companionSessionId: string;
  companion: CompanionState;
  connectors: ConnectorStatus[];
}

function readStorage(key: string): string | undefined {
  try {
    const value = localStorage.getItem(key)?.trim();
    return value || undefined;
  } catch {
    return undefined;
  }
}

function writeStorage(key: string, value: string | undefined): void {
  try {
    if (!value) {
      localStorage.removeItem(key);
      return;
    }
    localStorage.setItem(key, value);
  } catch {
    // Ignore storage failures in browser sandbox.
  }
}

function normalizeEndpoint(value: string | undefined): string | undefined {
  if (!value) return undefined;
  let next = value.trim();
  if (!next) return undefined;
  next = next.replace(/\/v1\/health\/?$/i, "");
  next = next.replace(/\/+$/g, "");
  if (!/^https?:\/\//i.test(next)) {
    next = `https://${next}`;
  }
  return next;
}

export function buildCompanionDiscoveryCandidates(
  manualEndpoint: string | undefined,
  lastSuccessfulEndpoint: string | undefined,
): string[] {
  return [
    normalizeEndpoint(manualEndpoint),
    normalizeEndpoint(lastSuccessfulEndpoint),
    DEFAULT_ENDPOINT,
    LOOPBACK_ENDPOINT,
  ].filter((value, index, values): value is string => Boolean(value) && values.indexOf(value) === index);
}

function errorName(error: unknown): string | undefined {
  return typeof error === "object" && error !== null && "name" in error
    ? String((error as { name?: unknown }).name ?? "")
    : undefined;
}

function errorMessage(error: unknown): string {
  if (error instanceof Error) return error.message;
  return String(error);
}

export function describeCompanionDiscoveryError(error: unknown, endpoint: string, timeoutMs: number): string {
  const name = errorName(error);
  const message = errorMessage(error);
  if (name === "AbortError" || /signal is aborted|aborted without reason|operation was aborted/i.test(message)) {
    return `Timed out contacting the Pi-Office companion at ${endpoint} after ${Math.round(timeoutMs / 1000)}s. Start it with npm run dev:companion and confirm it is listening on port ${DEFAULT_COMPANION_PORT}.`;
  }
  if (/failed to fetch|networkerror|load failed|could not connect|connection refused|err_connection/i.test(message)) {
    return `Could not reach the Pi-Office companion at ${endpoint}. The taskpane dev server uses https://localhost:3443; the optional companion must be started separately with npm run dev:companion on ${DEFAULT_COMPANION_PORT}.`;
  }
  if (/certificate|cert_authority|ssl|tls|self[- ]signed|err_cert/i.test(message)) {
    return `Could not establish a trusted HTTPS connection to the Pi-Office companion at ${endpoint}. Re-run npm run prepare:certs, trust the local certificate, then restart npm run dev:companion.`;
  }
  if (/not pi-office companion|not as pi-office companion/i.test(message)) {
    return message;
  }
  return `Companion discovery failed at ${endpoint}: ${message}`;
}

export function isCompanionHealthResponse(value: unknown): value is CompanionHealthResponse {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Partial<CompanionHealthResponse>;
  return candidate.ok === true &&
    typeof candidate.endpoint === "string" &&
    typeof candidate.identity === "string" &&
    Boolean(candidate.capabilities && typeof candidate.capabilities === "object");
}

function defaultCompanionState(): CompanionState {
  const manualEndpoint = readStorage(MANUAL_ENDPOINT_KEY);
  const lastSuccessfulEndpoint = readStorage(LAST_SUCCESSFUL_ENDPOINT_KEY);
  return {
    status: "unavailable",
    manualEndpoint,
    lastSuccessfulEndpoint,
    capabilities: {
      fileRead: false,
      localMcp: false,
      version: "companion-capabilities-v1",
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
    },
  };
}

async function fetchJsonWithTimeout<T>(url: string, init?: RequestInit, timeoutMs = REQUEST_TIMEOUT_MS): Promise<T> {
  const controller = new AbortController();
  const timeout = window.setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(url, {
      ...init,
      signal: controller.signal,
      headers: {
        "Content-Type": "application/json",
        ...(init?.headers ?? {}),
      },
    });
    if (!response.ok) {
      const text = await response.text();
      throw new Error(text || `${response.status} ${response.statusText}`);
    }
    return await response.json() as T;
  } finally {
    window.clearTimeout(timeout);
  }
}

export class CompanionClient {
  private state: CompanionState = defaultCompanionState();
  private readonly bindings = new Map<string, CompanionSessionBinding>();
  private discoveryPromise: Promise<CompanionState> | undefined;
  private hasInitialized = false;

  getState(): CompanionState {
    return { ...this.state };
  }

  async ensureInitialized(): Promise<CompanionState> {
    if (!this.hasInitialized && !this.discoveryPromise) {
      this.discoveryPromise = this.discover().finally(() => {
        this.discoveryPromise = undefined;
      });
    }
    return this.discoveryPromise ?? this.state;
  }

  async discover(): Promise<CompanionState> {
    this.hasInitialized = true;
    const manualEndpoint = normalizeEndpoint(readStorage(MANUAL_ENDPOINT_KEY));
    const lastSuccessfulEndpoint = normalizeEndpoint(readStorage(LAST_SUCCESSFUL_ENDPOINT_KEY));
    const candidates = buildCompanionDiscoveryCandidates(manualEndpoint, lastSuccessfulEndpoint);

    this.state = {
      ...this.state,
      status: "discovering",
      manualEndpoint,
      lastSuccessfulEndpoint,
      lastDiscoveryAttempts: [],
      lastError: undefined,
    };

    let lastError: string | undefined;
    const attempts: CompanionDiscoveryAttempt[] = [];
    for (const endpoint of candidates) {
      const started = Date.now();
      try {
        const health = await fetchJsonWithTimeout<CompanionHealthResponse>(`${endpoint}/v1/health`, undefined, COMPANION_DISCOVERY_TIMEOUT_MS);
        if (!isCompanionHealthResponse(health)) {
          throw new Error(`Endpoint ${endpoint} responded, but not as Pi-Office companion.`);
        }
        attempts.push({ endpoint, ok: true, durationMs: Date.now() - started });
        this.state = {
          status: "connected",
          endpoint: health.endpoint,
          identity: health.identity,
          manualEndpoint,
          lastSuccessfulEndpoint: health.endpoint,
          lastDiscoveryAttempts: attempts,
          capabilities: {
            ...health.capabilities,
            endpoint: health.endpoint,
          },
        };
        writeStorage(LAST_SUCCESSFUL_ENDPOINT_KEY, health.endpoint);
        this.bindings.clear();
        return this.getState();
      } catch (error) {
        const message = describeCompanionDiscoveryError(error, endpoint, COMPANION_DISCOVERY_TIMEOUT_MS);
        attempts.push({ endpoint, ok: false, message, durationMs: Date.now() - started });
        lastError = message;
      }
    }

    this.bindings.clear();
    this.state = {
      ...defaultCompanionState(),
      manualEndpoint,
      lastSuccessfulEndpoint,
      lastDiscoveryAttempts: attempts,
      lastError,
      status: manualEndpoint ? "error" : "unavailable",
    };
    return this.getState();
  }

  async setManualEndpoint(endpoint: string | undefined): Promise<CompanionState> {
    const normalized = normalizeEndpoint(endpoint);
    writeStorage(MANUAL_ENDPOINT_KEY, normalized);
    this.bindings.clear();
    return this.discover();
  }

  async openSession(
    browserSessionId: string,
    officeState: OfficeStateUpdate,
    connectors: CompanionConnectorDefinition[],
    windowId?: string,
  ): Promise<CompanionSessionBinding | undefined> {
    await this.ensureInitialized();
    if (this.state.status !== "connected" || !this.state.endpoint) {
      return undefined;
    }

    try {
      const response = await fetchJsonWithTimeout<CompanionSessionOpenResponse>(`${this.state.endpoint}/v1/sessions/open`, {
        method: "POST",
        body: JSON.stringify({
          browserSessionId,
          windowId,
          host: officeState.host,
          documentId: officeState.document.id,
          documentPath: officeState.document.documentPath,
          documentUrl: officeState.document.documentUrl,
          saved: officeState.document.saved,
          title: officeState.document.title,
          connectors,
        }),
      });

      const binding: CompanionSessionBinding = {
        browserSessionId,
        companionSessionId: response.sessionId,
        companion: response.companion,
        connectors: response.connectors,
      };
      this.bindings.set(browserSessionId, binding);
      this.state = {
        ...this.state,
        ...response.companion,
      };
      return binding;
    } catch (error) {
      this.bindings.delete(browserSessionId);
      this.state = {
        ...this.state,
        status: "error",
        lastError: error instanceof Error ? error.message : String(error),
        sessionId: undefined,
        connectorToolNames: [],
        capabilities: {
          fileRead: false,
          localMcp: false,
          endpoint: this.state.capabilities.endpoint ?? this.state.endpoint,
          agent: defaultCompanionState().capabilities.agent,
          providerAuth: defaultCompanionState().capabilities.providerAuth,
          nativeCapture: defaultCompanionState().capabilities.nativeCapture,
          mcp: defaultCompanionState().capabilities.mcp,
          memory: defaultCompanionState().capabilities.memory,
        },
      };
      return undefined;
    }
  }

  getBinding(browserSessionId: string): CompanionSessionBinding | undefined {
    return this.bindings.get(browserSessionId);
  }

  async probeConnector(definition: CompanionConnectorDefinition): Promise<{
    ok: boolean;
    status: ConnectorStatus;
    diagnostics: ConnectorDiagnostic[];
  } | undefined> {
    await this.ensureInitialized();
    if (this.state.status !== "connected" || !this.state.endpoint) {
      return undefined;
    }

    return fetchJsonWithTimeout(`${this.state.endpoint}/v1/connectors/probe`, {
      method: "POST",
      body: JSON.stringify(definition),
    });
  }

  async startConnectorOAuth(definition: CompanionConnectorDefinition): Promise<ConnectorOAuthStartResponse> {
    await this.ensureInitialized();
    if (this.state.status !== "connected" || !this.state.endpoint) {
      throw new Error("Optional companion is not connected for connector sign-in.");
    }

    const body: CompanionConnectorOAuthStartRequest = { definition };
    return fetchJsonWithTimeout(`${this.state.endpoint}/v1/connectors/oauth/start`, {
      method: "POST",
      body: JSON.stringify(body),
    });
  }

  async getConnectorOAuthStatus(request: CompanionConnectorOAuthStatusRequest): Promise<CompanionConnectorOAuthStatusResponse> {
    await this.ensureInitialized();
    if (this.state.status !== "connected" || !this.state.endpoint) {
      throw new Error("Optional companion is not connected for connector sign-in.");
    }

    return fetchJsonWithTimeout(`${this.state.endpoint}/v1/connectors/oauth/status`, {
      method: "POST",
      body: JSON.stringify(request),
    });
  }

  async getConnectorDiagnostics(): Promise<ConnectorDiagnosticsResponse | undefined> {
    await this.ensureInitialized();
    if (this.state.status !== "connected" || !this.state.endpoint) {
      return undefined;
    }

    return fetchJsonWithTimeout(`${this.state.endpoint}/v1/diagnostics`);
  }

  async executeFileTool(
    browserSessionId: string,
    toolName: "read" | "grep" | "find" | "ls",
    params: Record<string, unknown>,
  ): Promise<unknown> {
    const binding = this.bindings.get(browserSessionId);
    if (!binding || !this.state.endpoint) {
      throw new Error("Optional companion is not connected for this taskpane session.");
    }

    return fetchJsonWithTimeout(`${this.state.endpoint}/v1/sessions/${binding.companionSessionId}/files/${toolName}`, {
      method: "POST",
      body: JSON.stringify(params),
    });
  }

  async executeMcpTool(
    browserSessionId: string,
    toolName: string,
    args: Record<string, unknown>,
  ): Promise<unknown> {
    const binding = this.bindings.get(browserSessionId);
    if (!binding || !this.state.endpoint) {
      throw new Error("Optional companion is not connected for this taskpane session.");
    }

    return fetchJsonWithTimeout(`${this.state.endpoint}/v1/sessions/${binding.companionSessionId}/mcp/execute`, {
      method: "POST",
      body: JSON.stringify({
        toolName,
        arguments: args,
      }),
    });
  }

  async searchMcpTools(
    browserSessionId: string,
    request: McpToolSearchRequest,
  ): Promise<McpToolSearchResponse> {
    const binding = this.bindings.get(browserSessionId);
    if (!binding || !this.state.endpoint) {
      return { query: request.query, results: [] };
    }

    return fetchJsonWithTimeout(`${this.state.endpoint}/v1/sessions/${binding.companionSessionId}/mcp/search`, {
      method: "POST",
      body: JSON.stringify(request),
    });
  }

  async getShellCapability(browserSessionId?: string | undefined): Promise<CompanionShellCapability | undefined> {
    await this.ensureInitialized();
    if (this.state.status !== "connected" || !this.state.endpoint) {
      return undefined;
    }

    const binding = browserSessionId ? this.bindings.get(browserSessionId) : undefined;
    const path = binding
      ? `/v1/sessions/${binding.companionSessionId}/shell/capability`
      : "/v1/shell/capability";
    return fetchJsonWithTimeout(`${this.state.endpoint}${path}`);
  }

  async executeShellCommand(
    browserSessionId: string,
    request: CompanionShellExecuteRequest,
  ): Promise<CompanionShellExecuteResponse> {
    const binding = this.bindings.get(browserSessionId);
    if (!binding || !this.state.endpoint) {
      throw new Error("Optional companion is not connected for this taskpane session.");
    }

    return fetchJsonWithTimeout(`${this.state.endpoint}/v1/sessions/${binding.companionSessionId}/shell/execute`, {
      method: "POST",
      body: JSON.stringify(request),
    });
  }

  async captureNativeViewport(
    browserSessionId: string,
    request: CompanionNativeCaptureRequest,
  ): Promise<CompanionNativeCaptureResponse> {
    const binding = this.bindings.get(browserSessionId);
    if (!binding || !this.state.endpoint) {
      throw new Error("Optional companion native capture is not connected for this taskpane session.");
    }

    return fetchJsonWithTimeout(`${this.state.endpoint}/v1/sessions/${binding.companionSessionId}/native-capture/viewport`, {
      method: "POST",
      body: JSON.stringify(request),
    });
  }
}
