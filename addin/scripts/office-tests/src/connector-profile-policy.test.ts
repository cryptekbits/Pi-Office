import assert from "node:assert/strict";
import test from "node:test";

import { listConnectorCatalog } from "../../../apps/taskpane/src/lib/runtime/connector-catalog.js";

class MemoryStorage {
  private readonly store = new Map<string, string>();

  clear(): void {
    this.store.clear();
  }

  getItem(key: string): string | null {
    return this.store.get(key) ?? null;
  }

  key(index: number): string | null {
    return Array.from(this.store.keys())[index] ?? null;
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
    atob?: (value: string) => string;
    btoa?: (value: string) => string;
  };
  globalAny.localStorage = new MemoryStorage() as unknown as Storage;
  globalAny.window = { location: { origin: "https://localhost:3443" }, open: () => null };
  globalAny.atob = (value: string) => Buffer.from(value, "base64").toString("binary");
  globalAny.btoa = (value: string) => Buffer.from(value, "binary").toString("base64");
}

async function loadBrowserConnectorRuntime() {
  installRuntimePolyfills();
  const module = await import(`../../../apps/taskpane/src/lib/runtime/browser-connectors.js?test=${Date.now()}-${Math.random()}`);
  return module.BrowserConnectorRuntime as typeof import("../../../apps/taskpane/src/lib/runtime/browser-connectors.js").BrowserConnectorRuntime;
}

test("all built-in connectors expose curated setup profiles", () => {
  const connectors = listConnectorCatalog();
  assert.ok(connectors.length >= 24);
  for (const connector of connectors) {
    assert.ok(connector.setupProfiles?.length, `${connector.id} should expose at least one setup profile`);
  }

  const parallel = connectors.find((connector) => connector.id === "parallel-web");
  assert.ok(parallel);
  assert.ok(parallel.setupProfiles?.some((profile) => profile.id === "parallel-search-anonymous" && profile.endpoint === "https://search.parallel.ai/mcp"));
  assert.ok(parallel.setupProfiles?.some((profile) => profile.id === "parallel-search-oauth-zdr" && profile.endpoint === "https://search.parallel.ai/mcp-oauth"));

  const obsidian = connectors.find((connector) => connector.id === "obsidian");
  assert.ok(obsidian);
  assert.equal(obsidian.setupProfiles?.every((profile) => profile.transport === "local_stdio"), true);

  const github = connectors.find((connector) => connector.id === "github");
  assert.equal(github?.setupProfiles?.some((profile) => profile.endpoint === "https://api.githubcopilot.com/mcp/readonly"), true);
});

test("browser connector runtime persists selected profile and tool policy overrides", async () => {
  const Runtime = await loadBrowserConnectorRuntime();
  const runtime = new Runtime();
  await runtime.ready;

  const response = await runtime.connectConnector({
    connectorId: "parallel-web",
    setupProfileId: "parallel-search-anonymous",
    name: "Parallel Web",
    enabled: true,
    authMethod: "none",
    transport: "remote_http",
    credentialSource: "none",
    url: "https://search.parallel.ai/mcp",
  });

  assert.equal(response.status.setupProfileId, "parallel-search-anonymous");
  const verified = await runtime.reverifyConnector(response.status.id);
  assert.ok(verified.status.capabilities?.toolInventory?.length);
  const tool = verified.status.capabilities.toolInventory[0]!;
  assert.equal(tool.enabled, true);

  const updated = await runtime.updateToolPolicy({
    connectorId: response.status.id,
    toolName: tool.name,
    enabled: false,
    warningAcknowledged: true,
    suppressWarning: true,
  });

  assert.equal(updated.status.suppressNonReadToolWarning, true);
  assert.equal(updated.status.capabilities?.allowedTools.includes(tool.name), false);
  assert.equal(updated.status.capabilities?.blockedTools.includes(tool.name), true);

  const exported = runtime.getExportBundle();
  const exportedConnector = exported.connectors.find((connector) => connector.connectorId === "parallel-web");
  assert.equal(exportedConnector?.setupProfileId, "parallel-search-anonymous");
  assert.equal(exportedConnector?.suppressNonReadToolWarning, true);
  assert.equal(exportedConnector?.toolPolicyOverrides?.[0]?.toolName, tool.name);
});
