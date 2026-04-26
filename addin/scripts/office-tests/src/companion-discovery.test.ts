import assert from "node:assert/strict";
import test from "node:test";
import { isAllowedCompanionOrigin } from "../../../../companion/src/http.js";
import { createCompanionRuntimeDiagnostics } from "../../../../companion/src/runtime-diagnostics.js";

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
