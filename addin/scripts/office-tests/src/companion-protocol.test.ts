import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import test from "node:test";
import {
  COMPANION_PROTOCOL_FEATURE_IDS,
  COMPANION_PROTOCOL_VERSION,
  TASKPANE_COMPANION_PROTOCOL,
  type CompanionProtocolFeature,
  type CompanionProtocolFeatureId,
} from "@pi-office/pi-office-pack/protocol";

function protocolFeature(id: CompanionProtocolFeatureId): CompanionProtocolFeature {
  const feature = TASKPANE_COMPANION_PROTOCOL.features.find((entry) => entry.id === id);
  assert.ok(feature, `Expected companion protocol feature ${id}.`);
  return feature;
}

test("advanced companion protocol covers the required ownership seams", () => {
  assert.equal(TASKPANE_COMPANION_PROTOCOL.version, COMPANION_PROTOCOL_VERSION);
  assert.equal(TASKPANE_COMPANION_PROTOCOL.advancedMode.officeJsExecutor, "taskpane");
  assert.equal(TASKPANE_COMPANION_PROTOCOL.advancedMode.secretMigration, "explicit_user_action");
  assert.equal(TASKPANE_COMPANION_PROTOCOL.advancedMode.fallback, "taskpane_basic_mode");
  assert.ok(TASKPANE_COMPANION_PROTOCOL.advancedMode.taskpaneRole.some((role) => /Office\.js executor/i.test(role)));
  assert.ok(TASKPANE_COMPANION_PROTOCOL.advancedMode.companionRole.some((role) => /model inference/i.test(role)));

  const featureIds = TASKPANE_COMPANION_PROTOCOL.features.map((feature) => feature.id).sort();
  assert.deepEqual(featureIds, [...COMPANION_PROTOCOL_FEATURE_IDS].sort());

  assert.equal(protocolFeature("capability_discovery").state, "available");
  assert.equal(protocolFeature("tool_requests").state, "available");
  assert.equal(protocolFeature("chat_streaming").state, "available");
  assert.equal(protocolFeature("office_tool_execution").owner, "taskpane");
  assert.equal(protocolFeature("settings_sync").state, "available");
  assert.equal(protocolFeature("auth_migration").state, "available");
  assert.match(protocolFeature("auth_migration").summary, /explicit user action/i);
});

test("companion protocol routes are declared and implemented", () => {
  const serverSource = readFileSync(join(process.cwd(), "..", "companion", "src", "server.ts"), "utf8");

  for (const featureId of ["settings_sync", "chat_streaming", "office_tool_execution", "auth_migration"] as const) {
    const feature = protocolFeature(featureId);
    for (const route of feature.routes) {
      assert.match(serverSource, new RegExp(route.path.replaceAll("/", "\\/")));
    }
  }

  assert.equal(protocolFeature("settings_sync").routes.some((route) => route.path === "/v1/sessions/:sessionId/settings/sync"), true);
  assert.equal(protocolFeature("chat_streaming").state, "available");
  assert.equal(protocolFeature("office_tool_execution").state, "reserved");
});

test("available companion tool routes stay listed in the shared protocol descriptor", () => {
  const toolRoutes = new Set(protocolFeature("tool_requests").routes.map((route) => route.path));

  assert.equal(toolRoutes.has("/v1/sessions/:sessionId/files/:toolName"), true);
  assert.equal(toolRoutes.has("/v1/sessions/:sessionId/mcp/execute"), true);
  assert.equal(toolRoutes.has("/v1/sessions/:sessionId/mcp/search"), true);
  assert.equal(toolRoutes.has("/v1/sessions/:sessionId/shell/execute"), true);
  assert.equal(toolRoutes.has("/v1/sessions/:sessionId/native-capture/viewport"), true);
});
