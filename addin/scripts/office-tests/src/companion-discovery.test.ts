import assert from "node:assert/strict";
import test from "node:test";
import {
  buildCompanionDiscoveryCandidates,
  COMPANION_DISCOVERY_TIMEOUT_MS,
  describeCompanionDiscoveryError,
  isCompanionHealthResponse,
} from "../../../apps/taskpane/src/lib/runtime/companion-client.js";
import { isAllowedCompanionOrigin } from "../../../../companion/src/http.js";
import { captureNativeViewport, createNativeCaptureCapability } from "../../../../companion/src/native-capture.js";
import { createCompanionRuntimeDiagnostics } from "../../../../companion/src/runtime-diagnostics.js";
import { TASKPANE_COMPANION_PROTOCOL } from "@pi-office/pi-office-pack/protocol";

test("companion CORS allows loopback taskpane origins and rejects non-loopback origins", () => {
  assert.equal(isAllowedCompanionOrigin("https://localhost:3443"), true);
  assert.equal(isAllowedCompanionOrigin("https://127.0.0.1:3443"), true);
  assert.equal(isAllowedCompanionOrigin("http://localhost:3443"), true);
  assert.equal(isAllowedCompanionOrigin("https://example.com"), false);
  assert.equal(isAllowedCompanionOrigin("not a url"), false);
  assert.equal(isAllowedCompanionOrigin(undefined), false);
});

test("companion runtime diagnostics report machine checks instead of browser-only missing state", () => {
  const diagnostics = createCompanionRuntimeDiagnostics((command) => {
    if (command === "node") return { ok: true, detail: "node: v99.0.0" };
    if (command === "npx") return { ok: true, detail: "npx: 99.0.0" };
    return { ok: false, detail: `${command} missing` };
  }, () => new Date("2026-04-26T00:00:00.000Z"));

  assert.equal(diagnostics.generatedAt, "2026-04-26T00:00:00.000Z");
  assert.equal(diagnostics.diagnostics[0]?.code, "companion_runtime");
  assert.equal(diagnostics.runtimes.find((entry) => entry.key === "node")?.ok, true);
  assert.equal(diagnostics.runtimes.find((entry) => entry.key === "npx")?.ok, true);
  assert.equal(diagnostics.runtimes.find((entry) => entry.key === "npm")?.ok, false);
  assert.match(diagnostics.runtimes.find((entry) => entry.key === "node")?.detail ?? "", /v99\.0\.0/);
});

test("companion native capture reports Windows-only support honestly", () => {
  const linuxCapability = createNativeCaptureCapability("linux");
  assert.equal(linuxCapability.state, "unavailable");
  assert.equal(linuxCapability.available, false);
  assert.equal(linuxCapability.trueViewportScreenshot, false);

  const windowsCapability = createNativeCaptureCapability("win32");
  assert.equal(windowsCapability.state, "available");
  assert.equal(windowsCapability.available, true);
  assert.deepEqual(windowsCapability.hosts, ["word", "excel"]);
  assert.equal(windowsCapability.trueViewportScreenshot, true);
});

test("companion native capture fails closed on unsupported platforms", () => {
  const response = captureNativeViewport(
    { host: "word", title: "Unsupported capture test" },
    { host: "word", includeWindowFrame: true },
    "linux",
  );

  assert.equal(response.ok, false);
  assert.match(response.error ?? "", /Windows only|unavailable/i);
});

test("companion discovery tries saved endpoints before localhost and 127 fallback", () => {
  assert.deepEqual(
    buildCompanionDiscoveryCandidates(" localhost:3444/v1/health ", "https://localhost:3444"),
    ["https://localhost:3444", "https://127.0.0.1:3444"],
  );

  assert.deepEqual(
    buildCompanionDiscoveryCandidates("http://127.0.0.1:4500", undefined),
    ["http://127.0.0.1:4500", "https://localhost:3444", "https://127.0.0.1:3444"],
  );
});

test("companion discovery normalizes abort errors into actionable copy", () => {
  const abort = Object.assign(new Error("signal is aborted without reason"), { name: "AbortError" });
  const message = describeCompanionDiscoveryError(abort, "https://localhost:3444", COMPANION_DISCOVERY_TIMEOUT_MS);

  assert.match(message, /Timed out contacting the Pi-Office companion/i);
  assert.match(message, /npm run dev:companion/);
  assert.match(message, /3444/);
  assert.doesNotMatch(message, /signal is aborted without reason/i);
});

test("companion discovery validates the health response shape before connecting", () => {
  assert.equal(isCompanionHealthResponse({
    ok: true,
    endpoint: "https://localhost:3444",
    identity: "pi-office-companion@localhost:3444",
    capabilities: { fileRead: true, localMcp: true, protocol: TASKPANE_COMPANION_PROTOCOL },
  }), true);

  assert.equal(isCompanionHealthResponse({
    ok: true,
    endpoint: "https://localhost:3443",
    identity: "vite-dev-server",
  }), false);
});
