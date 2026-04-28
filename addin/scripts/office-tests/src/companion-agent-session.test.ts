import assert from "node:assert/strict";
import { mkdtempSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import {
  DEFAULT_USER_PREFERENCES,
  TASKPANE_COMPANION_PROTOCOL,
  type CompanionSettingsSyncRequest,
} from "@pi-office/pi-office-pack/protocol";
import { CompanionAgentSessionManager } from "../../../../companion/src/agent-session.js";
import type { CompanionConfig } from "../../../../companion/src/config.js";
import { createCompanionSecretStore } from "../../../../companion/src/oauth-token-store.js";
import { CompanionProviderAuthStore } from "../../../../companion/src/provider-auth-store.js";

function companionConfig(dataDir: string): CompanionConfig {
  return {
    host: "127.0.0.1",
    port: 3444,
    endpoint: "https://127.0.0.1:3444",
    identity: "test-companion",
    repoRoot: process.cwd(),
    certDir: process.cwd(),
    dataDir,
    tls: {
      pfx: Buffer.alloc(0),
      passphrase: "",
    },
  };
}

function authStore(dataDir: string): CompanionProviderAuthStore {
  const config = companionConfig(dataDir);
  return new CompanionProviderAuthStore(
    createCompanionSecretStore(config, {
      legacyFileName: "provider-auth.json",
      encryptedFileName: "provider-auth.dpapi.json",
    }, {
      platform: "win32",
      protect: (value) => Buffer.from(value, "utf8").toString("base64"),
      unprotect: (value) => Buffer.from(value, "base64").toString("utf8"),
    }),
  );
}

function settings(mode: "basic" | "smart_auto" | "advanced"): CompanionSettingsSyncRequest {
  return {
    browserSessionId: "browser-session",
    preferences: {
      ...DEFAULT_USER_PREFERENCES,
      companionRuntimeMode: mode,
      defaultModelByProvider: { openai: "gpt-5.1" },
    },
    providerSelection: {
      enabledProviders: ["openai"],
      enabledModels: ["openai/gpt-5.1"],
      defaultModelByProvider: { openai: "gpt-5.1" },
    },
    connectorCount: 0,
    secretsIncluded: false,
  };
}

test("companion agent session preflight keeps Basic and Smart Auto on taskpane fallback", () => {
  const manager = new CompanionAgentSessionManager(authStore(mkdtempSync(join(tmpdir(), "pi-office-agent-"))));
  const response = manager.handlePrompt({
    id: "companion-session",
    browserSessionId: "browser-session",
    settings: settings("smart_auto"),
  }, {
    browserSessionId: "browser-session",
    text: "Draft a summary",
    secretsIncluded: false,
  });

  assert.equal(response.ok, false);
  assert.equal(response.status, "taskpane_fallback");
  assert.equal(response.taskpaneFallback, true);
  assert.equal(response.officeToolProxy, true);
  assert.equal(response.modelProvider, "openai");
  assert.equal(response.modelId, "gpt-5.1");
  assert.match(response.reason, /taskpane/i);
});

test("companion agent session requires explicit companion-held auth in Advanced mode", () => {
  const manager = new CompanionAgentSessionManager(authStore(mkdtempSync(join(tmpdir(), "pi-office-agent-"))));
  const response = manager.handlePrompt({
    id: "companion-session",
    browserSessionId: "browser-session",
    settings: settings("advanced"),
  }, {
    browserSessionId: "browser-session",
    text: "Draft a summary",
    secretsIncluded: false,
  });

  assert.equal(response.ok, false);
  assert.equal(response.status, "provider_auth_required");
  assert.equal(response.taskpaneFallback, true);
  assert.match(response.reason, /not silently migrated/i);
});

test("companion agent session remains fail-closed until the Pi agent runtime lands", () => {
  const store = authStore(mkdtempSync(join(tmpdir(), "pi-office-agent-")));
  store.setApiKey({ provider: "openai", apiKey: "test-provider-key", explicitUserAction: true });
  const manager = new CompanionAgentSessionManager(store);

  const response = manager.handlePrompt({
    id: "companion-session",
    browserSessionId: "browser-session",
    settings: settings("advanced"),
  }, {
    browserSessionId: "browser-session",
    text: "Draft a summary",
    secretsIncluded: false,
  });

  assert.equal(response.ok, false);
  assert.equal(response.status, "agent_runtime_unavailable");
  assert.equal(response.taskpaneFallback, true);
  assert.match(response.reason, /runtime is not enabled/i);
  assert.match(response.reason, /Office\.js execution remains taskpane-owned/i);
});

test("companion agent session rejects prompt and Office result payloads that claim secrets", () => {
  const manager = new CompanionAgentSessionManager(authStore(mkdtempSync(join(tmpdir(), "pi-office-agent-"))));
  const session = {
    id: "companion-session",
    browserSessionId: "browser-session",
    settings: settings("advanced"),
  };

  assert.throws(
    () => manager.handlePrompt(session, {
      browserSessionId: "browser-session",
      text: "Draft a summary",
      secretsIncluded: true as false,
    }),
    /must not include taskpane provider secrets/,
  );
  assert.throws(
    () => manager.handleOfficeToolResult(session, {
      browserSessionId: "browser-session",
      toolCallId: "tool-call-1",
      secretsIncluded: true as false,
    }),
    /must not include taskpane provider secrets/,
  );
});

test("companion Office tool result proxy rejects results without a pending companion request", () => {
  const manager = new CompanionAgentSessionManager(authStore(mkdtempSync(join(tmpdir(), "pi-office-agent-"))));
  const response = manager.handleOfficeToolResult({
    id: "companion-session",
    browserSessionId: "browser-session",
    settings: settings("advanced"),
  }, {
    browserSessionId: "browser-session",
    toolCallId: "tool-call-1",
    toolName: "verify_doc",
    result: { ok: true },
    secretsIncluded: false,
  });

  assert.equal(response.ok, false);
  assert.equal(response.status, "no_pending_office_tool_request");
  assert.equal(response.toolCallId, "tool-call-1");
  assert.match(response.reason, /Office\.js execution remains taskpane-owned/i);
});

test("companion agent protocol routes are declared and backed by the scaffold", () => {
  const protocol = TASKPANE_COMPANION_PROTOCOL;
  const chat = protocol.features.find((feature) => feature.id === "chat_streaming");
  const officeProxy = protocol.features.find((feature) => feature.id === "office_tool_execution");
  const serverSource = readFileSync(join(process.cwd(), "..", "companion", "src", "server.ts"), "utf8");
  const agentSource = readFileSync(join(process.cwd(), "..", "companion", "src", "agent-session.ts"), "utf8");

  assert.equal(chat?.state, "reserved");
  assert.equal(officeProxy?.state, "reserved");
  assert.match(chat?.routes[0]?.description ?? "", /preflight/i);
  assert.match(officeProxy?.routes[0]?.description ?? "", /pending companion Office request/i);
  assert.match(serverSource, /agentSessions\.handlePrompt/);
  assert.match(serverSource, /agentSessions\.handleOfficeToolResult/);
  assert.match(agentSource, /agent_runtime_unavailable/);
  assert.match(agentSource, /provider_auth_required/);
});
