import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import test from "node:test";

import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
import { getOfficeSkillPaths } from "../../../packages/pi-office-pack/src/extension.js";
import { OFFICE_TOOL_NAMES, TOOL_CATEGORY_MAP } from "../../../packages/pi-office-pack/src/protocol.js";
import { listConnectorCatalog } from "../../../apps/taskpane/src/lib/runtime/connector-catalog.js";

class MemoryStorage {
  private readonly store = new Map<string, string>();

  clear(): void {
    this.store.clear();
  }

  getItem(key: string): string | null {
    return this.store.has(key) ? this.store.get(key)! : null;
  }

  key(index: number): string | null {
    const keys = Array.from(this.store.keys());
    return keys[index] ?? null;
  }

  removeItem(key: string): void {
    this.store.delete(key);
  }

  setItem(key: string, value: string): void {
    this.store.set(key, String(value));
  }

  get length(): number {
    return this.store.size;
  }
}

function installRuntimePolyfills(): void {
  const globalAny = globalThis as unknown as {
    localStorage?: Storage;
    window?: { location?: { origin?: string }; open?: (...args: unknown[]) => unknown };
    WebSocket?: { CONNECTING: number; OPEN: number; CLOSING: number; CLOSED: number };
    CloseEvent?: typeof CloseEvent;
    MessageEvent?: typeof MessageEvent;
    atob?: (value: string) => string;
    btoa?: (value: string) => string;
  };

  globalAny.localStorage = new MemoryStorage() as unknown as Storage;
  globalAny.window = globalAny.window ?? {};
  globalAny.window.location = globalAny.window.location ?? { origin: "https://localhost:3443" };
  globalAny.window.location.origin = globalAny.window.location.origin ?? "https://localhost:3443";
  globalAny.window.open = globalAny.window.open ?? (() => null);

  if (!globalAny.WebSocket) {
    globalAny.WebSocket = {
      CONNECTING: 0,
      OPEN: 1,
      CLOSING: 2,
      CLOSED: 3,
    };
  }

  if (!globalAny.CloseEvent) {
    class CloseEventPolyfill extends Event {
      readonly code: number;
      readonly reason: string;
      readonly wasClean: boolean;

      constructor(type: string, init?: { code?: number; reason?: string; wasClean?: boolean }) {
        super(type);
        this.code = init?.code ?? 0;
        this.reason = init?.reason ?? "";
        this.wasClean = init?.wasClean ?? true;
      }
    }
    globalAny.CloseEvent = CloseEventPolyfill as unknown as typeof CloseEvent;
  }

  if (!globalAny.MessageEvent) {
    class MessageEventPolyfill<T = unknown> extends Event {
      readonly data: T;

      constructor(type: string, init?: { data?: T }) {
        super(type);
        this.data = init?.data as T;
      }
    }
    globalAny.MessageEvent = MessageEventPolyfill as unknown as typeof MessageEvent;
  }

  if (!globalAny.atob) {
    globalAny.atob = (value: string) => Buffer.from(value, "base64").toString("binary");
  }
  if (!globalAny.btoa) {
    globalAny.btoa = (value: string) => Buffer.from(value, "binary").toString("base64");
  }
}

async function loadKernelModule() {
  installRuntimePolyfills();
  const specifier = `../../../apps/taskpane/src/lib/runtime/inprocess-kernel.js?test=${Date.now()}-${Math.random()}`;
  return import(specifier);
}

async function setupOAuthConnector(runtime: { dispatchKernelRequest: (path: string, init?: RequestInit) => Promise<unknown> }, name: string) {
  return runtime.dispatchKernelRequest("/v1/connectors/setup/connect", {
    method: "POST",
    body: JSON.stringify({
      connectorId: "custom",
      name,
      enabled: true,
      favorite: false,
      scopeTarget: "global",
      setupKind: "remote_oauth",
      authMethod: "oauth",
      transport: "remote_http",
      credentialSource: "oauth",
      url: `https://${name.toLowerCase().replace(/[^a-z0-9]+/g, "-")}.example.test/mcp`,
    }),
  }) as Promise<{
    ok: true;
    status: {
      id: string;
      healthState: string;
      needsCredential: boolean;
      connected: boolean;
      credentialSource: string;
      lastError?: string;
    };
  }>;
}

async function startConnectorOAuth(runtime: { dispatchKernelRequest: (path: string, init?: RequestInit) => Promise<unknown> }, connectorId: string) {
  return runtime.dispatchKernelRequest("/v1/connectors/oauth/start", {
    method: "POST",
    body: JSON.stringify({ connectorId }),
  }) as Promise<{ ok: true; connectorId: string; state: string; expiresAt: string }>;
}

async function connectorStatus(runtime: { dispatchKernelRequest: (path: string, init?: RequestInit) => Promise<unknown> }, connectorId: string) {
  const statuses = await runtime.dispatchKernelRequest("/v1/connectors/status", {
    method: "POST",
    body: JSON.stringify({}),
  }) as {
    connectors: Array<{
      id: string;
      healthState: string;
      needsCredential: boolean;
      connected: boolean;
      credentialSource: string;
      lastError?: string;
    }>;
  };
  const status = statuses.connectors.find((entry) => entry.id === connectorId);
  assert.ok(status, `Expected connector ${connectorId} in status response.`);
  return status;
}

test("external gap closures keep refresh/skill/web helpers out of first-class Office tool inventory", () => {
  const inventory = new Set<string>(OFFICE_TOOL_NAMES);
  assert.equal(inventory.has("refresh_mcp_connectors"), false);
  assert.equal(inventory.has("read_skill"), false);
  assert.equal(inventory.has("web_search"), false);
  assert.equal(inventory.has("web_fetch"), false);
  assert.equal(TOOL_CATEGORY_MAP.refresh_mcp_connectors, undefined);
  assert.equal(TOOL_CATEGORY_MAP.read_skill, undefined);
  assert.equal(TOOL_CATEGORY_MAP.web_search, undefined);
  assert.equal(TOOL_CATEGORY_MAP.web_fetch, undefined);
});

test("connector refresh stays runtime/UI scoped through reverify route with saved-document gating", async () => {
  const runtime = await loadKernelModule();
  const setup = await runtime.dispatchKernelRequest("/v1/connectors/setup/connect", {
    method: "POST",
    body: JSON.stringify({
      connectorId: "custom",
      name: "External Context Connector",
      enabled: true,
      favorite: false,
      scopeTarget: "global",
      setupKind: "local_executable_or_docker",
      authMethod: "none",
      transport: "local_stdio",
      credentialSource: "none",
      command: "node",
      args: ["connector.js"],
    }),
  }) as {
    ok: boolean;
    status: { id: string; name: string };
  };

  assert.equal(setup.ok, true);
  assert.equal(typeof setup.status.id, "string");

  const reverify = await runtime.dispatchKernelRequest("/v1/connectors/reverify", {
    method: "POST",
    body: JSON.stringify({
      connectorId: setup.status.id,
      scopeContext: {
        host: "word",
        documentId: "doc-external-context",
        documentTitle: "External Context Test",
        documentSaved: false,
      },
    }),
  }) as {
    ok: boolean;
    status: { id: string };
  };

  assert.equal(typeof reverify.ok, "boolean");
  assert.equal(reverify.status.id, setup.status.id);

  await assert.rejects(
    () =>
      runtime.dispatchKernelRequest("/v1/connectors/scope", {
        method: "POST",
        body: JSON.stringify({
          connectorId: setup.status.id,
          enabled: true,
          scopeTarget: "document",
          scopeContext: {
            host: "word",
            documentId: "doc-external-context",
            documentTitle: "External Context Test",
            documentSaved: false,
          },
        }),
      }),
    /Document scope requires a saved document context\./,
  );
});

test("remote HTTP connectors are marked setup-only until browser execution exists", async () => {
  const runtime = await loadKernelModule();
  const setup = await runtime.dispatchKernelRequest("/v1/connectors/setup/connect", {
    method: "POST",
    body: JSON.stringify({
      connectorId: "custom",
      name: "Remote HTTP MCP",
      enabled: true,
      favorite: false,
      scopeTarget: "global",
      setupKind: "remote_oauth",
      authMethod: "none",
      transport: "remote_http",
      credentialSource: "none",
      url: "https://example.test/mcp",
    }),
  }) as {
    ok: boolean;
    status: {
      id: string;
      executionEnvironment?: string;
      executionAvailable?: boolean;
    };
  };

  assert.equal(setup.ok, true);
  assert.equal(setup.status.executionEnvironment, "browser");
  assert.equal(setup.status.executionAvailable, false);

  const statuses = await runtime.dispatchKernelRequest("/v1/connectors/status", {
    method: "POST",
    body: JSON.stringify({}),
  }) as {
    connectors: Array<{
      id: string;
      executionEnvironment?: string;
      executionAvailable?: boolean;
    }>;
  };

  const saved = statuses.connectors.find((connector) => connector.id === setup.status.id);
  assert.ok(saved, "saved remote connector should be returned by status route");
  assert.equal(saved.executionEnvironment, "browser");
  assert.equal(saved.executionAvailable, false);
});

test("connector OAuth callbacks cannot create connected state without a credential handoff", async () => {
  const runtime = await loadKernelModule();
  const setup = await setupOAuthConnector(runtime, "OAuth Contract");
  assert.equal(setup.status.healthState, "auth_required");
  assert.equal(setup.status.needsCredential, true);
  assert.equal(setup.status.connected, false);

  const missingToken = await startConnectorOAuth(runtime, setup.status.id);
  const missingTokenCallback = await runtime.dispatchKernelRequest("/v1/connectors/oauth/callback", {
    method: "POST",
    body: JSON.stringify({
      connectorId: setup.status.id,
      state: missingToken.state,
    }),
  }) as {
    ok: true;
    status: { healthState: string; needsCredential: boolean; connected: boolean; lastError?: string };
    diagnostics: Array<{ code: string; message: string }>;
  };
  assert.equal(missingTokenCallback.status.healthState, "auth_required");
  assert.equal(missingTokenCallback.status.needsCredential, true);
  assert.equal(missingTokenCallback.status.connected, false);
  assert.equal(missingTokenCallback.diagnostics[0]?.code, "oauth_not_completed");
  assert.match(missingTokenCallback.status.lastError ?? "", /verified credential handoff/i);

  const cancelled = await startConnectorOAuth(runtime, setup.status.id);
  const cancelledCallback = await runtime.dispatchKernelRequest("/v1/connectors/oauth/callback", {
    method: "POST",
    body: JSON.stringify({
      connectorId: setup.status.id,
      state: cancelled.state,
      error: "access_denied",
    }),
  }) as { status: { healthState: string; needsCredential: boolean; connected: boolean; lastError?: string } };
  assert.equal(cancelledCallback.status.healthState, "auth_required");
  assert.equal(cancelledCallback.status.needsCredential, true);
  assert.equal(cancelledCallback.status.connected, false);
  assert.match(cancelledCallback.status.lastError ?? "", /access_denied/);

  const mismatch = await startConnectorOAuth(runtime, setup.status.id);
  await assert.rejects(
    () =>
      runtime.dispatchKernelRequest("/v1/connectors/oauth/callback", {
        method: "POST",
        body: JSON.stringify({
          connectorId: setup.status.id,
          state: `${mismatch.state}-wrong`,
          credential: { accessToken: "wrong-state-token" },
        }),
      }),
    /OAuth callback state was not found/,
  );
  const afterMismatch = await connectorStatus(runtime, setup.status.id);
  assert.equal(afterMismatch.healthState, "auth_required");
  assert.equal(afterMismatch.needsCredential, true);
  assert.equal(afterMismatch.connected, false);

  const expired = await startConnectorOAuth(runtime, setup.status.id);
  const realDateNow = Date.now;
  Date.now = () => realDateNow() + 11 * 60 * 1000;
  try {
    await assert.rejects(
      () =>
        runtime.dispatchKernelRequest("/v1/connectors/oauth/callback", {
          method: "POST",
          body: JSON.stringify({
            connectorId: setup.status.id,
            state: expired.state,
            credential: { accessToken: "expired-token" },
          }),
        }),
      /OAuth callback state expired/,
    );
  } finally {
    Date.now = realDateNow;
  }
  const afterExpired = await connectorStatus(runtime, setup.status.id);
  assert.equal(afterExpired.healthState, "auth_required");
  assert.equal(afterExpired.needsCredential, true);
  assert.equal(afterExpired.connected, false);
  assert.match(afterExpired.lastError ?? "", /expired/i);

  const successful = await startConnectorOAuth(runtime, setup.status.id);
  const successCallback = await runtime.dispatchKernelRequest("/v1/connectors/oauth/callback", {
    method: "POST",
    body: JSON.stringify({
      connectorId: setup.status.id,
      state: successful.state,
      credential: {
        accessToken: "verified-access-token",
        tokenType: "Bearer",
        expiresInSeconds: 3600,
      },
    }),
  }) as {
    status: { healthState: string; needsCredential: boolean; connected: boolean; credentialSource: string };
    diagnostics: Array<{ code: string }>;
  };
  assert.equal(successCallback.status.healthState, "ready");
  assert.equal(successCallback.status.needsCredential, false);
  assert.equal(successCallback.status.connected, true);
  assert.equal(successCallback.status.credentialSource, "oauth");
  assert.equal(successCallback.diagnostics[0]?.code, "oauth_completed");

  const bundle = await runtime.dispatchKernelRequest("/v1/connectors/export") as {
    connectors: Array<{ connectorId: string; credentialSource: string; secret?: string }>;
    scopeOverrides: unknown[];
    favorites: string[];
    auditPreference: { enabled: boolean };
    version: number;
    exportedAt: string;
  };
  const exported = bundle.connectors.find((entry) => entry.credentialSource === "oauth");
  assert.ok(exported, "OAuth connector should be included in export bundle.");
  assert.equal(Object.prototype.hasOwnProperty.call(exported, "secret"), false);

  const importedRuntime = await loadKernelModule();
  const imported = await importedRuntime.dispatchKernelRequest("/v1/connectors/import/apply", {
    method: "POST",
    body: JSON.stringify({ bundle }),
  }) as { importedConnectorIds: string[] };
  assert.equal(imported.importedConnectorIds.length, 1);
  const importedStatus = await connectorStatus(importedRuntime, imported.importedConnectorIds[0]!);
  assert.equal(importedStatus.credentialSource, "oauth");
  assert.equal(importedStatus.healthState, "auth_required");
  assert.equal(importedStatus.needsCredential, true);
  assert.equal(importedStatus.connected, false);
  assert.match(importedStatus.lastError ?? "", /not included in connector imports/i);
});

test("skill closure path uses packaged skill injection plus explicit unsaved-document gating", () => {
  const skillPaths = getOfficeSkillPaths();
  assert.equal(skillPaths.some((entry) => entry.endsWith("office-host.SKILL.md")), true);
  assert.equal(skillPaths.some((entry) => entry.endsWith("workspace-handoff.SKILL.md")), true);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /When the document is unsaved, do not assume local file access is available/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /Read-only filesystem tools are only available when the optional local companion is connected/);
});

test("web-grounded context closure stays on hard-read-only research connectors", () => {
  const researchConnectors = listConnectorCatalog().filter((connector) => connector.category === "research");
  assert.ok(researchConnectors.length > 0);

  for (const connector of researchConnectors) {
    assert.equal(connector.readOnly, true);
    assert.equal(connector.readPolicy.mode, "hard-read-only");
    assert.equal(connector.readPolicy.allowPrompts, false);
    assert.equal(connector.readPolicy.blockPromptPatterns?.includes(".*"), true);
    assert.ok((connector.readPolicy.allowToolPatterns ?? []).length > 0);
  }
});

test("raw shell tools stay unavailable unless the companion sandbox reports available", async () => {
  const runtime = await loadKernelModule();
  const openResponse = await runtime.dispatchKernelRequest("/v1/sessions/open", {
    method: "POST",
    body: JSON.stringify({
      host: "word",
      documentId: "doc-no-raw-shell",
      saved: true,
      title: "No Raw Shell",
      documentPath: "C:/Users/manan/Documents/no-raw-shell.docx",
    }),
  }) as { sessionId: string };
  const socket = runtime.createLocalBridgeSocket(openResponse.sessionId);
  const session = (socket as unknown as { session: { agent: { state: { tools: Array<{ name: string }> } } } }).session;
  const toolNames = new Set(session.agent.state.tools.map((tool) => tool.name));

  assert.equal(toolNames.has("bash"), false);
  assert.equal(toolNames.has("edit"), false);
  assert.equal(toolNames.has("write"), false);
  assert.equal(toolNames.has("read"), false, "External file tools require a connected companion and saved-folder session binding.");
  socket.close();

  const companionServer = readFileSync(join(process.cwd(), "apps", "companion", "src", "server.ts"), "utf8");
  const shellSandbox = readFileSync(join(process.cwd(), "apps", "companion", "src", "shell-sandbox.ts"), "utf8");
  assert.match(companionServer, /\/v1\/sessions\/:sessionId\/shell\/execute/);
  assert.match(companionServer, /CompanionShellSandbox/);
  assert.match(companionServer, /\/v1\/sessions\/:sessionId\/files\/:toolName/);
  assert.match(shellSandbox, /createCompanionBashOperations/);
  assert.doesNotMatch(shellSandbox, /createLocalBashOperations/);
  assert.doesNotMatch(shellSandbox, /\bspawn\(/);
});
