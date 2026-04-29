import assert from "node:assert/strict";
import { mkdtempSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import {
  fauxAssistantMessage,
  registerFauxProvider,
  streamSimple,
} from "@mariozechner/pi-ai";
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

function settings(
  mode: "basic" | "smart_auto" | "advanced",
  provider = "openai",
  modelId = "gpt-5.1",
): CompanionSettingsSyncRequest {
  return {
    browserSessionId: "browser-session",
    preferences: {
      ...DEFAULT_USER_PREFERENCES,
      companionRuntimeMode: mode,
      defaultModelByProvider: { [provider]: modelId },
    },
    providerSelection: {
      enabledProviders: [provider],
      enabledModels: [`${provider}/${modelId}`],
      defaultModelByProvider: { [provider]: modelId },
    },
    connectorCount: 0,
    secretsIncluded: false,
  };
}

test("companion agent session preflight keeps Basic and Smart Auto on taskpane fallback", async () => {
  const manager = new CompanionAgentSessionManager(authStore(mkdtempSync(join(tmpdir(), "pi-office-agent-"))));
  const response = await manager.handlePrompt({
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

test("companion agent session requires explicit companion-held auth in Advanced mode", async () => {
  const manager = new CompanionAgentSessionManager(authStore(mkdtempSync(join(tmpdir(), "pi-office-agent-"))));
  const response = await manager.handlePrompt({
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

test("companion agent session runs a companion-owned Pi agent in Advanced mode with companion-held auth", async () => {
  const provider = "faux-companion-agent";
  const modelId = "companion-faux";
  const faux = registerFauxProvider({
    provider,
    models: [{ id: modelId, name: "Companion Faux" }],
    tokensPerSecond: 0,
  });
  const store = authStore(mkdtempSync(join(tmpdir(), "pi-office-agent-")));
  store.setApiKey({ provider, apiKey: "test-provider-key", explicitUserAction: true });
  faux.setResponses([
    (context) => {
      assert.match(context.systemPrompt ?? "", /Advanced companion agent/);
      assert.match(context.systemPrompt ?? "", /Office\.js document reads and writes remain taskpane-owned/);
      assert.equal(context.tools?.length ?? 0, 0);
      assert.match(String(context.messages[0]?.content ?? ""), /Draft a summary/);
      assert.match(String(context.messages[0]?.content ?? ""), /Quarterly Plan/);
      return fauxAssistantMessage("Companion summary ready.");
    },
  ]);

  try {
    const manager = new CompanionAgentSessionManager(store, {
      resolveModel: (candidateProvider, candidateModelId) => {
        if (candidateProvider !== provider) return undefined;
        return faux.getModel(candidateModelId ?? modelId) as never;
      },
      streamFn: streamSimple as never,
    });

    const response = await manager.handlePrompt({
      id: "companion-session",
      browserSessionId: "browser-session",
      host: "word",
      documentId: "doc-1",
      title: "Quarterly Plan",
      saved: true,
      workspaceDir: "C:\\Docs",
      settings: settings("advanced", provider, modelId),
    }, {
      browserSessionId: "browser-session",
      text: "Draft a summary",
      secretsIncluded: false,
    });

    assert.equal(response.ok, true, response.reason);
    assert.equal(response.status, "accepted");
    assert.equal(response.taskpaneFallback, false);
    assert.equal(response.officeToolProxy, true);
    assert.equal(response.completed, true);
    assert.equal(response.assistantText, "Companion summary ready.");
    assert.equal(response.toolCallCount, 0);
    assert.equal(store.status([provider]).verifiedProviders.includes(provider), true);
    assert.equal(faux.state.callCount, 1);
  } finally {
    faux.unregister();
  }
});

test("companion agent session demotes provider auth after companion runtime auth failures", async () => {
  const provider = "faux-companion-auth-failure";
  const modelId = "companion-faux";
  const faux = registerFauxProvider({
    provider,
    models: [{ id: modelId, name: "Companion Faux" }],
    tokensPerSecond: 0,
  });
  const store = authStore(mkdtempSync(join(tmpdir(), "pi-office-agent-")));
  store.setApiKey({ provider, apiKey: "test-provider-key", explicitUserAction: true });
  faux.setResponses([
    fauxAssistantMessage([], {
      stopReason: "error",
      errorMessage: "401 invalid api key",
    }),
  ]);

  try {
    const manager = new CompanionAgentSessionManager(store, {
      resolveModel: (candidateProvider, candidateModelId) => {
        if (candidateProvider !== provider) return undefined;
        return faux.getModel(candidateModelId ?? modelId) as never;
      },
      streamFn: streamSimple as never,
    });

    const response = await manager.handlePrompt({
      id: "companion-session",
      browserSessionId: "browser-session",
      settings: settings("advanced", provider, modelId),
    }, {
      browserSessionId: "browser-session",
      text: "Draft a summary",
      secretsIncluded: false,
    });

    assert.equal(response.ok, false);
    assert.equal(response.status, "agent_runtime_unavailable");
    assert.match(response.reason, /401 invalid api key/);
    assert.equal(store.status([provider]).verificationFailedProviders.includes(provider), true);
  } finally {
    faux.unregister();
  }
});

test("companion agent session rejects prompt and Office result payloads that claim secrets", async () => {
  const manager = new CompanionAgentSessionManager(authStore(mkdtempSync(join(tmpdir(), "pi-office-agent-"))));
  const session = {
    id: "companion-session",
    browserSessionId: "browser-session",
    settings: settings("advanced"),
  };

  await assert.rejects(
    manager.handlePrompt(session, {
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

test("companion agent protocol routes are declared and backed by the runtime", () => {
  const protocol = TASKPANE_COMPANION_PROTOCOL;
  const chat = protocol.features.find((feature) => feature.id === "chat_streaming");
  const officeProxy = protocol.features.find((feature) => feature.id === "office_tool_execution");
  const serverSource = readFileSync(join(process.cwd(), "..", "companion", "src", "server.ts"), "utf8");
  const agentSource = readFileSync(join(process.cwd(), "..", "companion", "src", "agent-session.ts"), "utf8");

  assert.equal(chat?.state, "available");
  assert.equal(officeProxy?.state, "reserved");
  assert.match(chat?.routes[0]?.description ?? "", /companion-owned Pi agent/i);
  assert.match(officeProxy?.routes[0]?.description ?? "", /pending companion Office request/i);
  assert.match(serverSource, /await this\.agentSessions\.handlePrompt/);
  assert.match(serverSource, /agentSessions\.handleOfficeToolResult/);
  assert.match(agentSource, /new Agent/);
  assert.match(agentSource, /toolExecution: "sequential"/);
  assert.match(agentSource, /provider_auth_required/);
});
