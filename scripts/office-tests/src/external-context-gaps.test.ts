import assert from "node:assert/strict";
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

test("skill closure path uses packaged skill injection plus explicit unsaved-document gating", () => {
  const skillPaths = getOfficeSkillPaths();
  assert.equal(skillPaths.some((entry) => entry.endsWith("office-host.SKILL.md")), true);
  assert.equal(skillPaths.some((entry) => entry.endsWith("workspace-handoff.SKILL.md")), true);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /When the document is unsaved, do not assume local file access is available/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /Once the document is saved, workspace tools may become available/);
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
