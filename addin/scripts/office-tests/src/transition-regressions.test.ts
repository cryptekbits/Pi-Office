import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import test from "node:test";

function source(path: string): string {
  return readFileSync(join(process.cwd(), path), "utf8");
}

test("transition regression coverage keeps P0 remediation work under automated guards", () => {
  const privacyStorageTests = source("scripts/office-tests/src/privacy-storage.test.ts");
  const defaultsAutonomyTests = source("scripts/office-tests/src/defaults-autonomy.test.ts");
  const protocolParityTests = source("scripts/office-tests/src/protocol-parity.test.ts");
  const officeRefreshTests = source("scripts/office-tests/src/office-refresh-policy.test.ts");
  const externalContextTests = source("scripts/office-tests/src/external-context-gaps.test.ts");

  assert.match(privacyStorageTests, /doesNotMatch\(authEnvelope, \/test-provider-key\//);
  assert.match(defaultsAutonomyTests, /only lists workspace tools when the session exposes them/);
  assert.match(protocolParityTests, /survive session reopen/);
  assert.match(protocolParityTests, /disconnect cancels pending interactive requests/);
  assert.match(protocolParityTests, /Smart Auto hides viewport screenshots until companion native capture is available/);
  assert.match(officeRefreshTests, /latest active attempt/i);
  assert.match(externalContextTests, /raw shell tools stay unavailable/i);
});

test("dev and sideload preflight scripts encode current runtime assumptions", () => {
  const devPreflight = source("scripts/preflight-dev-runtime.mjs");
  const sideloadPreflight = source("scripts/preflight-sideload.mjs");

  assert.match(devPreflight, /const requiredPort = 3443/);
  assert.match(devPreflight, /localhost\.pfx/);
  assert.match(devPreflight, /passphrase\.txt/);
  assert.match(devPreflight, /thumbprint\.txt/);
  assert.match(devPreflight, /Cert:\\\\CurrentUser\\\\Root/);
  assert.match(devPreflight, /ensurePortAvailable\(requiredPort\)/);

  assert.match(sideloadPreflight, /https:\/\/localhost:3443\//);
  assert.match(sideloadPreflight, /shortcuts\.json/);
  assert.match(sideloadPreflight, /brand\/icon-80\.png\?v=20260421/);
  assert.match(sideloadPreflight, /optional companion on https:\/\/localhost:3444 is not required for add-in load/);
});
