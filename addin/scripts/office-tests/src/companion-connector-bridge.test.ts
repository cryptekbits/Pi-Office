import assert from "node:assert/strict";
import test from "node:test";

import type { CompanionConnectorDefinition } from "@pi-office/pi-office-pack/protocol";
import {
  CompanionConnectorBridge,
  buildRemoteHttpRequestHeaders,
  classifyConnectorToolForPolicy,
  resolveLocalStdioProcessEnvironment,
} from "../../../../companion/src/connector-bridge.js";

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
