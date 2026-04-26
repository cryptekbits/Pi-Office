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
    assert.ok(connector.docsUrl, `${connector.id} should expose connector docs provenance`);
    for (const profile of connector.setupProfiles ?? []) {
      assert.ok(profile.docsUrl, `${connector.id}/${profile.id} should expose docs provenance`);
      assert.ok(profile.endpointEvidenceUrl, `${connector.id}/${profile.id} should expose endpoint provenance`);
      assert.ok(profile.checkedAt, `${connector.id}/${profile.id} should expose checked date`);
      assert.ok(profile.availability, `${connector.id}/${profile.id} should expose availability`);
      assert.ok(profile.browserDirect, `${connector.id}/${profile.id} should expose browser-direct metadata`);
      assert.ok(Array.isArray(profile.riskNotes), `${connector.id}/${profile.id} should expose risk notes`);
      if (profile.authMethod !== "none") {
        assert.ok(profile.authEvidenceUrl, `${connector.id}/${profile.id} should expose auth provenance`);
      }
    }
  }

  const parallel = connectors.find((connector) => connector.id === "parallel-web");
  assert.ok(parallel);
  assert.ok(parallel.setupProfiles?.some((profile) => profile.id === "parallel-search-anonymous" && profile.endpoint === "https://search-mcp.parallel.ai/mcp"));
  assert.ok(parallel.setupProfiles?.some((profile) => profile.id === "parallel-search-oauth-zdr" && profile.endpoint === "https://search-mcp.parallel.ai/mcp"));

  const obsidian = connectors.find((connector) => connector.id === "obsidian");
  assert.ok(obsidian);
  assert.equal(obsidian.setupProfiles?.every((profile) => profile.transport === "local_stdio"), true);
  assert.equal(obsidian.vendor, "Community");
  assert.equal(obsidian.setupProfiles?.[0]?.officialness, "community");

  const github = connectors.find((connector) => connector.id === "github");
  assert.equal(github?.setupProfiles?.some((profile) => profile.endpoint === "https://api.githubcopilot.com/mcp/readonly"), true);

  const granola = connectors.find((connector) => connector.id === "granola");
  assert.ok(granola);
  assert.equal(granola.envHints.length, 0);
  assert.deepEqual(granola.setupProfiles?.map((profile) => profile.id), ["granola-hosted-oauth"]);
  assert.equal(granola.setupProfiles?.[0]?.authMethod, "oauth");
  assert.equal(granola.setupProfiles?.[0]?.requiresCompanion, false);
  assert.equal(granola.setupProfiles?.[0]?.browserDirect, "supported");

  const perplexity = connectors.find((connector) => connector.id === "perplexity");
  assert.ok(perplexity);
  assert.deepEqual(perplexity.setupProfiles?.map((profile) => profile.id), ["perplexity-local-stdio"]);
  assert.equal(perplexity.setupProfiles?.[0]?.requiresCompanion, true);

  const googleDrive = connectors.find((connector) => connector.id === "google-drive");
  assert.equal(googleDrive?.setupProfiles?.find((profile) => profile.id === "google-drive-official")?.setupDisabled, true);

  const postgres = connectors.find((connector) => connector.id === "postgresql-reader");
  assert.equal(postgres?.vendor, "Community");
  assert.equal(postgres?.setupProfiles?.[0]?.officialness, "provider_reference");
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
    url: "https://search-mcp.parallel.ai/mcp",
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

test("granola OAuth starts browser sign-in and browser-direct MCP verification enforces tool policy", async () => {
  const openedUrls: string[] = [];
  const globalAny = globalThis as unknown as {
    window?: { location?: { origin?: string }; open?: (...args: unknown[]) => unknown };
    fetch?: typeof fetch;
  };
  const Runtime = await loadBrowserConnectorRuntime();
  const runtime = new Runtime();
  await runtime.ready;
  globalAny.window = {
    location: { origin: "https://localhost:3443" },
    open: (url) => {
      openedUrls.push(String(url));
      return null;
    },
  };
  globalAny.fetch = async (input, init) => {
    const url = String(input);
    const method = init?.method ?? "GET";
    if (url === "https://mcp.granola.ai/.well-known/oauth-authorization-server") {
      return Response.json({
        authorization_endpoint: "https://mcp-auth.granola.ai/oauth2/authorize",
        token_endpoint: "https://mcp-auth.granola.ai/oauth2/token",
        registration_endpoint: "https://mcp-auth.granola.ai/oauth2/register",
      });
    }
    if (url === "https://mcp-auth.granola.ai/oauth2/register" && method === "POST") {
      return Response.json({ client_id: "pi-office-test-client" });
    }
    if (url === "https://mcp-auth.granola.ai/oauth2/token" && method === "POST") {
      return Response.json({
        access_token: "granola-access-token",
        refresh_token: "granola-refresh-token",
        token_type: "Bearer",
        expires_in: 3600,
      });
    }
    if (url === "https://mcp.granola.ai/mcp" && method === "POST") {
      const payload = JSON.parse(String(init?.body ?? "{}")) as { method?: string; id?: string };
      const headers = new Headers({ "content-type": "application/json", "mcp-session-id": "session-1" });
      if (payload.method === "initialize") {
        return Response.json({ jsonrpc: "2.0", id: payload.id, result: { protocolVersion: "2025-06-18", capabilities: {} } }, { headers });
      }
      if (payload.method === "notifications/initialized") {
        return new Response("", { status: 202, headers });
      }
      if (payload.method === "tools/list") {
        return Response.json({
          jsonrpc: "2.0",
          id: payload.id,
          result: {
            tools: [
              { name: "search_meetings", description: "Search meeting notes", annotations: { readOnlyHint: true, openWorldHint: true } },
              { name: "delete_meeting", description: "Delete meeting notes", annotations: { destructiveHint: true } },
            ],
          },
        }, { headers });
      }
      if (payload.method === "tools/call") {
        return Response.json({ jsonrpc: "2.0", id: payload.id, result: { content: [{ type: "text", text: "ok" }] } }, { headers });
      }
    }
    return new Response(`Unexpected ${method} ${url}`, { status: 500 });
  };

  const saved = await runtime.connectConnector({
    connectorId: "granola",
    setupProfileId: "granola-hosted-oauth",
    name: "Granola",
    enabled: true,
    authMethod: "oauth",
    transport: "remote_http",
    credentialSource: "oauth",
    url: "https://mcp.granola.ai/mcp",
  });
  assert.equal(saved.status.healthState, "auth_required");

  const started = await runtime.startOAuth(saved.status.id);
  assert.equal(started.callbackUrl, "https://localhost:3443/connector-oauth-callback");
  assert.equal(openedUrls.length, 0);
  assert.match(started.url ?? "", /^https:\/\/mcp-auth\.granola\.ai\/oauth2\/authorize/);
  assert.match(started.url ?? "", /code_challenge=/);

  const completed = await runtime.completeOAuth({ state: started.state, code: "auth-code" });
  assert.equal(completed.status.healthState, "ready");

  const verified = await runtime.reverifyConnector(saved.status.id);
  assert.equal(verified.ok, true);
  assert.deepEqual(verified.status.capabilities?.allowedTools, ["search_meetings"]);
  assert.deepEqual(verified.status.capabilities?.blockedTools, ["delete_meeting"]);

  await assert.rejects(
    () => runtime.executeBrowserMcpTool("delete_meeting", {}, { host: "word", documentSaved: false }),
    /not enabled|disabled/,
  );
  const executed = await runtime.executeBrowserMcpTool("search_meetings", { query: "recap" }, { host: "word", documentSaved: false });
  assert.ok(executed);
});
