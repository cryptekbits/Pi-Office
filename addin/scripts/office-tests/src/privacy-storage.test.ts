import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import test from "node:test";

import { CHAT_HISTORY_STORAGE_KEY, clearStoredChatHistory } from "../../../apps/taskpane/src/hooks/useChatHistory.js";

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

function installRuntimePolyfills(): Storage {
  const globalAny = globalThis as unknown as {
    localStorage?: Storage;
    window?: { location?: { origin?: string }; open?: (...args: unknown[]) => unknown; confirm?: (message?: string) => boolean };
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
  globalAny.window.confirm = globalAny.window.confirm ?? (() => true);

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

  return globalAny.localStorage;
}

async function loadKernelModule(seed?: (storage: Storage) => void) {
  const storage = installRuntimePolyfills();
  seed?.(storage);
  const specifier = `../../../apps/taskpane/src/lib/runtime/inprocess-kernel.js?test=${Date.now()}-${Math.random()}`;
  return { storage, runtime: await import(specifier) };
}

test("taskpane declares an explicit Office-compatible CSP policy", () => {
  const html = readFileSync(join(process.cwd(), "apps", "taskpane", "index.html"), "utf8");
  assert.match(html, /http-equiv="Content-Security-Policy"/);
  assert.match(html, /default-src 'self'/);
  assert.match(html, /https:\/\/appsforoffice\.microsoft\.com/);
  assert.match(html, /connect-src 'self' https:/);
  assert.match(html, /https:\/\/localhost:\*/);
  assert.match(html, /object-src 'none'/);
  assert.match(html, /form-action 'none'/);
});

test("provider auth clear-all removes encrypted credentials and the local crypto key", async () => {
  const { storage, runtime } = await loadKernelModule();

  await runtime.dispatchKernelRequest("/v1/auth/api-key", {
    method: "POST",
    body: JSON.stringify({ provider: "openai", apiKey: "test-provider-key" }),
  });
  const authEnvelope = storage.getItem("pi-office-auth");
  assert.ok(authEnvelope);
  assert.match(authEnvelope, /ciphertext/);
  assert.doesNotMatch(authEnvelope, /test-provider-key/);
  assert.ok(storage.getItem("pi-office-auth-key-v1"));

  await runtime.dispatchKernelRequest("/v1/auth", { method: "DELETE" });
  const authStatus = await runtime.dispatchKernelRequest("/v1/auth/status") as {
    storedProviders: string[];
    verifiedProviders: string[];
    providerStates: Array<{ provider: string; credentialStored: boolean; verifiedUsable: boolean }>;
  };

  assert.deepEqual(authStatus.storedProviders, []);
  assert.deepEqual(authStatus.verifiedProviders, []);
  assert.equal(authStatus.providerStates.find((entry) => entry.provider === "openai")?.credentialStored, false);
  assert.equal(storage.getItem("pi-office-auth"), null);
  assert.equal(storage.getItem("pi-office-auth-key-v1"), null);
});

test("connector clear-all removes config, scopes, OAuth state, logs, and the local crypto key", async () => {
  const { storage, runtime } = await loadKernelModule();

  const setup = await runtime.dispatchKernelRequest("/v1/connectors/setup/connect", {
    method: "POST",
    body: JSON.stringify({
      connectorId: "custom",
      name: "Privacy Storage Connector",
      enabled: true,
      favorite: true,
      scopeTarget: "global",
      setupKind: "remote_oauth",
      authMethod: "api_key",
      transport: "remote_http",
      credentialSource: "manual",
      url: "https://privacy-storage.example.test/mcp",
      secret: "connector-secret",
    }),
  }) as { ok: true; status: { id: string } };
  assert.equal(setup.ok, true);
  assert.ok(storage.getItem("pi-office-connectors"));
  assert.ok(storage.getItem("pi-office-connectors-key-v1"));

  await runtime.dispatchKernelRequest("/v1/connectors", { method: "DELETE" });
  const status = await runtime.dispatchKernelRequest("/v1/connectors/status", {
    method: "POST",
    body: JSON.stringify({}),
  }) as { connectors: unknown[] };
  const audit = await runtime.dispatchKernelRequest("/v1/connectors/audit") as {
    preference: { enabled: boolean };
  };

  assert.deepEqual(status.connectors, []);
  assert.equal(audit.preference.enabled, true);
  assert.equal(storage.getItem("pi-office-connectors"), null);
  assert.equal(storage.getItem("pi-office-connectors-key-v1"), null);
});

test("chat history storage exposes a clear helper for Privacy settings", () => {
  const storage = installRuntimePolyfills();
  storage.setItem(CHAT_HISTORY_STORAGE_KEY, JSON.stringify([{ chatId: "chat-1" }]));

  clearStoredChatHistory();

  assert.equal(storage.getItem(CHAT_HISTORY_STORAGE_KEY), null);
});
