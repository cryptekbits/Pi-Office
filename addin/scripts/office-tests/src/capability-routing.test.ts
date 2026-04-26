import assert from "node:assert/strict";
import test from "node:test";
import {
  getAvailableToolNames,
  getResolvedCapability,
  resolvePiOfficeCapabilities,
} from "@pi-office/pi-office-pack/capabilities";
import type { CompanionState } from "@pi-office/pi-office-pack/protocol";

function disconnectedCompanion(): CompanionState {
  return {
    status: "unavailable",
    connectorToolNames: [],
    capabilities: {
      fileRead: false,
      localMcp: false,
      nativeCapture: {
        state: "unavailable",
        available: false,
        hosts: [],
        trueViewportScreenshot: false,
        includeWindowFrame: false,
      },
      agent: {
        state: "unavailable",
        available: false,
        officeToolProxy: true,
        providerAuth: false,
        smartAuto: true,
      },
      providerAuth: {
        state: "unavailable",
        available: false,
        explicitMigrationRequired: true,
      },
      mcp: {
        state: "unavailable",
        available: false,
        readOnly: true,
        toolCount: 0,
      },
    },
  };
}

function connectedCompanion(): CompanionState {
  return {
    status: "connected",
    endpoint: "https://localhost:3444",
    identity: "test-companion",
    sessionId: "companion-session",
    connectorToolNames: ["parallel_web.search"],
    capabilities: {
      fileRead: true,
      localMcp: true,
      endpoint: "https://localhost:3444",
      nativeCapture: {
        state: "available",
        available: true,
        hosts: ["word", "excel"],
        trueViewportScreenshot: true,
        includeWindowFrame: true,
      },
      agent: {
        state: "available",
        available: true,
        officeToolProxy: true,
        providerAuth: true,
        smartAuto: true,
      },
      providerAuth: {
        state: "available",
        available: true,
        explicitMigrationRequired: true,
        supportedAuthMethods: ["oauth", "api_key"],
      },
      mcp: {
        state: "available",
        available: true,
        readOnly: true,
        toolCount: 1,
      },
      memory: {
        state: "available",
        available: true,
      },
    },
  };
}

test("capability routing keeps taskpane-only mode usable and hides companion-only tools", () => {
  const capabilities = resolvePiOfficeCapabilities({
    host: "word",
    documentSaved: true,
    companion: disconnectedCompanion(),
  });
  const tools = getAvailableToolNames(capabilities);

  assert.equal(getResolvedCapability(capabilities, "office_context").activeRuntime, "addin");
  assert.equal(getResolvedCapability(capabilities, "inference").activeRuntime, "addin");
  assert.equal(getResolvedCapability(capabilities, "native_viewport_capture").available, false);
  assert.equal(getResolvedCapability(capabilities, "local_files").available, false);
  assert.equal(tools.has("office_capture_snapshot"), true);
  assert.equal(tools.has("office_capture_viewport"), false);
  assert.equal(tools.has("mcp"), false);
});

test("capability routing prefers companion for eligible capabilities when advertised", () => {
  const capabilities = resolvePiOfficeCapabilities({
    host: "word",
    documentSaved: true,
    companion: connectedCompanion(),
  });
  const tools = getAvailableToolNames(capabilities);

  assert.equal(getResolvedCapability(capabilities, "inference").preferredRuntime, "companion");
  assert.equal(getResolvedCapability(capabilities, "inference").fallbackRuntime, "addin");
  assert.equal(getResolvedCapability(capabilities, "native_viewport_capture").activeRuntime, "companion");
  assert.equal(getResolvedCapability(capabilities, "local_files").activeRuntime, "companion");
  assert.equal(tools.has("office_capture_viewport"), true);
  assert.equal(tools.has("mcp"), true);
  assert.equal(tools.has("read"), true);
});

test("capability routing keeps Office tools taskpane-owned even with companion connected", () => {
  const capabilities = resolvePiOfficeCapabilities({
    host: "powerpoint",
    documentSaved: true,
    companion: connectedCompanion(),
  });

  assert.equal(getResolvedCapability(capabilities, "office_write").activeRuntime, "addin");
  assert.equal(getResolvedCapability(capabilities, "office_review").activeRuntime, "addin");
  assert.equal(getResolvedCapability(capabilities, "powerpoint_native_visual").activeRuntime, "addin");
  assert.equal(getAvailableToolNames(capabilities).has("verify_slide_visual"), true);
});

test("companion local file routing requires a saved document", () => {
  const capabilities = resolvePiOfficeCapabilities({
    host: "word",
    documentSaved: false,
    companion: connectedCompanion(),
  });

  assert.equal(getResolvedCapability(capabilities, "local_files").available, false);
  assert.match(getResolvedCapability(capabilities, "local_files").reason ?? "", /Save the document/i);
  assert.equal(getResolvedCapability(capabilities, "mcp_connectors").available, true);
});
