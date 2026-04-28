import assert from "node:assert/strict";
import { mkdtempSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import {
  TASKPANE_COMPANION_PROTOCOL,
  type CompanionProtocolFeature,
} from "@pi-office/pi-office-pack/protocol";
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

function authStore(dataDir: string, platform: NodeJS.Platform): CompanionProviderAuthStore {
  const config = companionConfig(dataDir);
  return new CompanionProviderAuthStore(
    createCompanionSecretStore(config, {
      legacyFileName: "provider-auth.json",
      encryptedFileName: "provider-auth.dpapi.json",
    }, {
      platform,
      protect: (value) => Buffer.from(value, "utf8").toString("base64"),
      unprotect: (value) => Buffer.from(value, "base64").toString("utf8"),
    }),
  );
}

function protocolFeature(id: string): CompanionProtocolFeature {
  const feature = TASKPANE_COMPANION_PROTOCOL.features.find((entry) => entry.id === id);
  assert.ok(feature, `Expected protocol feature ${id}.`);
  return feature;
}

test("companion provider auth stores API keys only behind secure storage", () => {
  const dataDir = mkdtempSync(join(tmpdir(), "pi-office-provider-auth-"));
  const store = authStore(dataDir, "win32");

  const saved = store.setApiKey({
    provider: "openai",
    apiKey: "test-provider-key",
    explicitUserAction: true,
  });
  const envelope = readFileSync(join(dataDir, "provider-auth.dpapi.json"), "utf8");

  assert.equal(saved.secureStorage, true);
  assert.equal(saved.storageKind, "windows-dpapi");
  assert.equal(saved.storedProviders.includes("openai"), true);
  assert.equal(saved.unverifiedProviders.includes("openai"), true);
  assert.doesNotMatch(envelope, /test-provider-key/);
  assert.equal(store.getApiKey("openai"), "test-provider-key");

  const capability = store.getCapability();
  assert.equal(capability.state, "available");
  assert.equal(capability.available, true);
  assert.equal(capability.secureStorage, true);
  assert.equal(capability.configuredProviderCount, 1);
  assert.equal(capability.explicitMigrationRequired, true);
});

test("companion provider auth refuses writes without explicit user action", () => {
  const dataDir = mkdtempSync(join(tmpdir(), "pi-office-provider-auth-"));
  const store = authStore(dataDir, "win32");

  assert.throws(
    () => store.setApiKey({
      provider: "openai",
      apiKey: "test-provider-key",
      explicitUserAction: false as true,
    }),
    /explicitUserAction=true/,
  );
  assert.deepEqual(store.status().storedProviders, []);
});

test("companion provider auth fails closed when secure storage is unavailable", () => {
  const dataDir = mkdtempSync(join(tmpdir(), "pi-office-provider-auth-"));
  const store = authStore(dataDir, "linux");

  assert.throws(
    () => store.setApiKey({
      provider: "openai",
      apiKey: "test-provider-key",
      explicitUserAction: true,
    }),
    /Secure companion provider auth storage is unavailable/,
  );

  const capability = store.getCapability();
  assert.equal(capability.state, "unavailable");
  assert.equal(capability.available, false);
  assert.equal(capability.secureStorage, false);
  assert.equal(capability.storageKind, "local-json-unsupported");
  assert.deepEqual(store.status().storedProviders, []);
});

test("companion provider auth clear removes one provider or all providers", () => {
  const dataDir = mkdtempSync(join(tmpdir(), "pi-office-provider-auth-"));
  const store = authStore(dataDir, "win32");

  store.setApiKey({ provider: "openai", apiKey: "openai-key", explicitUserAction: true });
  store.setApiKey({ provider: "anthropic", apiKey: "anthropic-key", explicitUserAction: true });

  const afterOne = store.clear({ provider: "openai" });
  assert.equal(afterOne.storedProviders.includes("openai"), false);
  assert.equal(afterOne.storedProviders.includes("anthropic"), true);
  assert.equal(store.getApiKey("anthropic"), "anthropic-key");

  const afterAll = store.clear({});
  assert.deepEqual(afterAll.storedProviders, []);
});

test("companion provider auth protocol routes are declared and implemented", () => {
  const feature = protocolFeature("auth_migration");
  const serverSource = readFileSync(join(process.cwd(), "..", "companion", "src", "server.ts"), "utf8");

  assert.equal(feature.state, "available");
  assert.match(feature.summary, /explicit user action/i);
  assert.equal(feature.routes.some((route) => route.path === "/v1/provider-auth/status"), true);
  assert.equal(feature.routes.some((route) => route.path === "/v1/provider-auth/api-key"), true);
  assert.equal(feature.routes.some((route) => route.path === "/v1/provider-auth"), true);

  for (const route of feature.routes) {
    assert.match(serverSource, new RegExp(route.path.replaceAll("/", "\\/")));
  }
});

test("taskpane companion provider auth setup is visible and explicit", () => {
  const clientSource = readFileSync(join(process.cwd(), "apps", "taskpane", "src", "lib", "runtime", "companion-client.ts"), "utf8");
  const kernelSource = readFileSync(join(process.cwd(), "apps", "taskpane", "src", "lib", "runtime", "inprocess-kernel.ts"), "utf8");
  const appSource = readFileSync(join(process.cwd(), "apps", "taskpane", "src", "app", "App.tsx"), "utf8");
  const settingsSource = readFileSync(join(process.cwd(), "apps", "taskpane", "src", "app", "components", "SettingsPage.tsx"), "utf8");
  const syncStart = kernelSource.indexOf("private buildCompanionSettingsSyncRequest");
  const syncBody = kernelSource.slice(syncStart, syncStart + 900);

  assert.match(clientSource, /\/v1\/provider-auth\/status/);
  assert.match(clientSource, /\/v1\/provider-auth\/api-key/);
  assert.match(clientSource, /\/v1\/provider-auth/);
  assert.match(kernelSource, /\/v1\/companion\/provider-auth\/copy-api-key/);
  assert.match(kernelSource, /explicitUserAction=true is required before copying/);
  assert.match(kernelSource, /this\.modelRegistry\.getApiKey\(provider\)/);
  assert.match(
    kernelSource,
    /this\.companionClient\.setProviderApiKey\(\{[\s\S]*provider,[\s\S]*apiKey,[\s\S]*explicitUserAction: true/,
  );
  assert.match(settingsSource, /onCopyProviderAuthToCompanion/);
  assert.match(settingsSource, /Copy to companion/);
  assert.match(settingsSource, /window\.confirm/);
  assert.match(appSource, /Cleared all stored provider credentials from this taskpane and any reachable companion/);
  assert.match(syncBody, /secretsIncluded: false/);
  assert.doesNotMatch(syncBody, /apiKey|oauth|manualCredential/i);
});
