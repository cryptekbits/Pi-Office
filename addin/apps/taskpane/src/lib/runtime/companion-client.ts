import type {
  CompanionConnectorDefinition,
  CompanionHealthResponse,
  CompanionShellCapability,
  CompanionShellExecuteRequest,
  CompanionShellExecuteResponse,
  CompanionSessionOpenResponse,
  CompanionState,
  ConnectorDiagnosticsResponse,
  ConnectorDiagnostic,
  ConnectorStatus,
  OfficeStateUpdate,
} from "@pi-office/pi-office-pack/protocol";
import { DEFAULT_COMPANION_HOST, DEFAULT_COMPANION_PORT } from "@pi-office/pi-office-pack/defaults";

const MANUAL_ENDPOINT_KEY = "pi-office-companion-manual-endpoint";
const LAST_SUCCESSFUL_ENDPOINT_KEY = "pi-office-companion-last-endpoint";
const DEFAULT_ENDPOINT = `https://${DEFAULT_COMPANION_HOST}:${DEFAULT_COMPANION_PORT}`;
const DISCOVERY_TIMEOUT_MS = 1_200;
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
    const candidates = [manualEndpoint, lastSuccessfulEndpoint, DEFAULT_ENDPOINT]
      .filter((value, index, values): value is string => Boolean(value) && values.indexOf(value) === index);

    this.state = {
      ...this.state,
      status: "discovering",
      manualEndpoint,
      lastSuccessfulEndpoint,
      lastError: undefined,
    };

    let lastError: string | undefined;
    for (const endpoint of candidates) {
      try {
        const health = await fetchJsonWithTimeout<CompanionHealthResponse>(`${endpoint}/v1/health`, undefined, DISCOVERY_TIMEOUT_MS);
        this.state = {
          status: "connected",
          endpoint: health.endpoint,
          identity: health.identity,
          manualEndpoint,
          lastSuccessfulEndpoint: health.endpoint,
          capabilities: {
            ...health.capabilities,
            endpoint: health.endpoint,
          },
        };
        writeStorage(LAST_SUCCESSFUL_ENDPOINT_KEY, health.endpoint);
        this.bindings.clear();
        return this.getState();
      } catch (error) {
        lastError = error instanceof Error ? error.message : String(error);
      }
    }

    this.bindings.clear();
    this.state = {
      ...defaultCompanionState(),
      manualEndpoint,
      lastSuccessfulEndpoint,
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
}
