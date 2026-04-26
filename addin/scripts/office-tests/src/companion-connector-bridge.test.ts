import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";

import type { CompanionConnectorDefinition } from "@pi-office/pi-office-pack/protocol";
import {
  CompanionConnectorBridge,
  buildRemoteHttpRequestHeaders,
  classifyConnectorToolForPolicy,
  resolveLocalStdioProcessEnvironment,
} from "../../../../companion/src/connector-bridge.js";
import { CompanionOAuthBroker } from "../../../../companion/src/oauth-broker.js";
import type { CompanionConfig } from "../../../../companion/src/config.js";

const READ_POLICY: CompanionConnectorDefinition["readPolicy"] = {
  mode: "hard-read-only",
  allowResources: true,
  allowPrompts: true,
  allowToolPatterns: [".*"],
  blockToolPatterns: [],
  allowPromptPatterns: [".*"],
  blockPromptPatterns: [],
};

function connector(
  overrides: Partial<CompanionConnectorDefinition> = {},
): CompanionConnectorDefinition {
  return {
    id: "test-local-connector",
    connectorId: "custom",
    name: "Test Local Connector",
    source: "custom",
    category: "knowledge",
    maturity: "custom_mcp_only",
    setupKind: "local_executable_or_docker",
    authMethod: "api_key",
    transport: "local_stdio",
    credentialSource: "manual",
    command: "node",
    args: ["server.js"],
    readOnly: true,
    readPolicy: READ_POLICY,
    ...overrides,
  };
}

function companionConfig(dataDir: string): CompanionConfig {
  return {
    host: "localhost",
    port: 3444,
    endpoint: "https://localhost:3444",
    identity: "test-companion",
    repoRoot: dataDir,
    certDir: dataDir,
    dataDir,
    tls: {
      pfx: Buffer.from(""),
      passphrase: "",
    },
  };
}

test("local stdio env merges safe inherited variables, connector env, and manual credential target", () => {
  const resolved = resolveLocalStdioProcessEnvironment(
    connector({
      secret: "manual-secret",
      secretEnvKey: "PI_OFFICE_TEST_TOKEN",
      env: {
        CONNECTOR_MODE: "readonly",
      },
    }),
    {
      PATH: "C:\\Windows\\System32",
      USERPROFILE: "C:\\Users\\tester",
      OPENAI_API_KEY: "must-not-leak",
    },
  );

  assert.equal(resolved.credentialInjected, true);
  assert.equal(resolved.env.PATH, "C:\\Windows\\System32");
  assert.equal(resolved.env.USERPROFILE, "C:\\Users\\tester");
  assert.equal(resolved.env.CONNECTOR_MODE, "readonly");
  assert.equal(resolved.env.PI_OFFICE_TEST_TOKEN, "manual-secret");
  assert.equal(resolved.env.OPENAI_API_KEY, undefined);
  assert.equal(resolved.missingCredentialReason, undefined);
});

test("detected env credentials are injected without inheriting unrelated secrets", () => {
  const resolved = resolveLocalStdioProcessEnvironment(
    connector({
      credentialSource: "detected_env",
      useDetectedEnvKey: "GITHUB_TOKEN",
      secret: "ignored-manual-secret",
    }),
    {
      PATH: "/usr/bin",
      GITHUB_TOKEN: "detected-secret",
      AWS_SECRET_ACCESS_KEY: "must-not-leak",
    },
  );

  assert.equal(resolved.credential.source, "detected_env");
  assert.equal(resolved.credentialInjected, true);
  assert.equal(resolved.env.PATH, "/usr/bin");
  assert.equal(resolved.env.GITHUB_TOKEN, "detected-secret");
  assert.equal(resolved.env.AWS_SECRET_ACCESS_KEY, undefined);
  assert.equal(resolved.env.GITHUB_TOKEN === resolved.credential.value, true);
});

test("stdio env passthrough copies named host variables into the process environment", () => {
  const resolved = resolveLocalStdioProcessEnvironment(
    connector({
      stdioEnvPassthrough: ["CUSTOM_PASSTHROUGH", "EMPTY_MISSING"],
      env: {
        CONNECTOR_MODE: "readonly",
      },
    }),
    {
      PATH: "/usr/bin",
      CUSTOM_PASSTHROUGH: "passed-through",
    },
  );

  assert.equal(resolved.env.CONNECTOR_MODE, "readonly");
  assert.equal(resolved.env.CUSTOM_PASSTHROUGH, "passed-through");
  assert.equal(resolved.env.EMPTY_MISSING, undefined);
});

test("remote HTTP request headers merge env-sourced and static entries; static wins on duplicate names", () => {
  const headers = buildRemoteHttpRequestHeaders(
    {
      remoteHttpHeadersFromEnv: [
        { name: "X-From-Env", envVarName: "MCP_HDR_A" },
        { name: "X-Overlap", envVarName: "MCP_HDR_B" },
      ],
      remoteHttpHeaders: [
        { name: "X-Static", value: "s" },
        { name: "X-Overlap", value: "static-wins" },
      ],
    },
    {
      MCP_HDR_A: "env-a",
      MCP_HDR_B: "env-b",
    },
  );

  assert.equal(headers["X-From-Env"], "env-a");
  assert.equal(headers["X-Static"], "s");
  assert.equal(headers["X-Overlap"], "static-wins");
});

test("env-var credentials stay auth-required when the named env value is missing", () => {
  const resolved = resolveLocalStdioProcessEnvironment(
    connector({
      credentialSource: "env",
      secretEnvKey: "MISSING_CONNECTOR_TOKEN",
    }),
    {
      PATH: "/usr/bin",
    },
  );

  assert.equal(resolved.credential.source, "env");
  assert.equal(resolved.credential.envKey, "MISSING_CONNECTOR_TOKEN");
  assert.equal(resolved.credential.value, undefined);
  assert.equal(resolved.credentialInjected, false);
  assert.equal(resolved.missingCredentialReason, "missing_value");
});

test("manual local stdio credentials need an env target and never appear in status diagnostics", async () => {
  const bridge = new CompanionConnectorBridge();
  const result = await bridge.probeConnector(
    connector({
      secret: "manual-secret-without-target",
      secretEnvKey: undefined,
      useDetectedEnvKey: undefined,
    }),
  );

  assert.equal(result.ok, false);
  assert.equal(result.status.healthState, "auth_required");
  assert.equal(result.status.needsCredential, true);
  assert.equal(result.diagnostics.some((item) => item.code === "credential_env_key_required"), true);
  assert.doesNotMatch(JSON.stringify(result), /manual-secret-without-target/);
});

test("connector status and diagnostics expose env key names but not detected secret values", async () => {
  const key = "PI_OFFICE_DETECTED_TEST_TOKEN";
  const previous = process.env[key];
  process.env[key] = "detected-secret-value";
  try {
    const bridge = new CompanionConnectorBridge();
    const result = await bridge.probeConnector(
      connector({
        command: undefined,
        credentialSource: "detected_env",
        useDetectedEnvKey: key,
      }),
    );

    assert.equal(result.status.detectedEnvKey, key);
    assert.equal(result.status.needsCredential, false);
    assert.equal(result.status.healthState, "unverified");
    assert.doesNotMatch(JSON.stringify(result), /detected-secret-value/);
  } finally {
    if (previous === undefined) {
      delete process.env[key];
    } else {
      process.env[key] = previous;
    }
  }
});

test("connector tool classification uses annotations before conservative name rules", () => {
  assert.deepEqual(
    classifyConnectorToolForPolicy(connector(), {
      name: "anything_generate_invoice",
      annotations: { readOnlyHint: true },
    }),
    {
      classification: "read",
      defaultEnabled: true,
      reason: "MCP annotations mark this tool as read-only.",
    },
  );

  const destructive = classifyConnectorToolForPolicy(connector(), {
    name: "list_customers",
    annotations: { destructiveHint: true },
  });
  assert.equal(destructive.classification, "destructive");
  assert.equal(destructive.defaultEnabled, false);

  const unknown = classifyConnectorToolForPolicy(connector({
    readPolicy: {
      ...READ_POLICY,
      allowToolPatterns: ["^get_"],
      blockToolPatterns: ["^delete_"],
    },
  }), {
    name: "launch_remote_task",
  });
  assert.equal(unknown.classification, "unknown");
  assert.equal(unknown.defaultEnabled, false);
});

test("companion OAuth broker performs DCR, callback exchange, and token persistence", async () => {
  const dataDir = mkdtempSync(join(tmpdir(), "pi-office-oauth-"));
  const originalFetch = globalThis.fetch;
  try {
    globalThis.fetch = (async (input, init) => {
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
        const body = JSON.parse(String(init?.body ?? "{}")) as { redirect_uris?: string[] };
        assert.deepEqual(body.redirect_uris, ["https://localhost:3444/v1/connectors/oauth/callback"]);
        return Response.json({ client_id: "pi-office-test-client" }, { status: 201 });
      }
      if (url === "https://mcp-auth.granola.ai/oauth2/token" && method === "POST") {
        assert.match(String(init?.body ?? ""), /code_verifier=/);
        return Response.json({
          access_token: "granola-access-token",
          refresh_token: "granola-refresh-token",
          token_type: "Bearer",
          expires_in: 3600,
        });
      }
      return new Response(`Unexpected ${method} ${url}`, { status: 500 });
    }) as typeof fetch;

    const broker = new CompanionOAuthBroker(companionConfig(dataDir));
    const definition = connector({
      id: "granola",
      connectorId: "granola",
      name: "Granola",
      source: "library",
      category: "knowledge",
      maturity: "beta",
      setupKind: "remote_oauth",
      authMethod: "oauth",
      transport: "remote_http",
      credentialSource: "oauth",
      url: "https://mcp.granola.ai/mcp",
      command: undefined,
      args: undefined,
      setupProfileId: "granola-hosted-oauth",
      oauth: {
        broker: "companion",
        launchMode: "system_browser",
        metadataUrl: "https://mcp.granola.ai/.well-known/oauth-authorization-server",
        redirectPath: "/v1/connectors/oauth/callback",
        clientName: "Pi-Office",
      },
    });

    const started = await broker.start(definition);
    assert.equal(started.callbackUrl, "https://localhost:3444/v1/connectors/oauth/callback");
    assert.equal(started.openMode, "system_browser");
    assert.match(started.url ?? "", /^https:\/\/mcp-auth\.granola\.ai\/oauth2\/authorize/);
    assert.match(started.url ?? "", /code_challenge=/);

    const callback = await broker.completeCallback({ state: started.state, code: "auth-code" });
    assert.equal(callback.statusCode, 200);
    assert.equal(await broker.getAccessToken(definition), "granola-access-token");

    const restoredBroker = new CompanionOAuthBroker(companionConfig(dataDir));
    assert.equal(await restoredBroker.getAccessToken(definition), "granola-access-token");
  } finally {
    globalThis.fetch = originalFetch;
    rmSync(dataDir, { recursive: true, force: true });
  }
});

test("companion bridge uses brokered OAuth tokens instead of reporting missing credentials", async () => {
  const bridge = new CompanionConnectorBridge({
    getAccessToken: async () => "broker-token",
  });
  const result = await bridge.probeConnector(connector({
    id: "remote-oauth",
    connectorId: "granola",
    name: "Remote OAuth",
    source: "library",
    category: "knowledge",
    maturity: "beta",
    setupKind: "remote_oauth",
    authMethod: "oauth",
    transport: "remote_http",
    credentialSource: "oauth",
    url: "http://127.0.0.1:1/mcp",
    command: undefined,
    args: undefined,
    oauth: {
      broker: "companion",
      launchMode: "system_browser",
    },
  }));

  assert.notEqual(result.status.healthState, "auth_required");
  assert.equal(result.status.needsCredential, false);
  assert.equal(result.diagnostics.some((item) => item.code === "credential_required"), false);
});
