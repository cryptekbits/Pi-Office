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

test("release runtime docs match current scripts and manifest assumptions", () => {
  const releaseDoc = source("../docs/release-runtime.md");
  const readme = source("../README.md");
  const rootPackage = JSON.parse(source("../package.json")) as { scripts: Record<string, string> };
  const addinPackage = JSON.parse(source("package.json")) as { scripts: Record<string, string> };
  const companionPackage = JSON.parse(source("../companion/package.json")) as { scripts: Record<string, string> };
  const manifests = ["word.xml", "excel.xml", "powerpoint.xml"].map((name) => source(`manifests/${name}`));

  assert.match(readme, /docs\/release-runtime\.md/);
  assert.equal(rootPackage.scripts.dev, "npm --prefix addin run dev");
  assert.equal(rootPackage.scripts["dev:companion"], "npm --prefix companion run dev");
  assert.equal(rootPackage.scripts["preflight:sideload"], "npm --prefix addin run preflight:sideload");
  assert.equal(addinPackage.scripts.dev, "npm run prepare:certs && npm run preflight:dev && npm-run-all --parallel dev:pack dev:taskpane");
  assert.equal(companionPackage.scripts.dev, "tsx watch src/index.ts");

  for (const command of [
    "npm run dev",
    "npm run dev:companion",
    "npm run sideload:word",
    "npm run sideload:excel",
    "npm run sideload:powerpoint",
    "npm run validate:manifests",
  ]) {
    assert.match(releaseDoc, new RegExp(command.replaceAll(" ", "\\s+")));
  }

  assert.match(releaseDoc, /https:\/\/localhost:3443/);
  assert.match(releaseDoc, /https:\/\/localhost:3444/);
  assert.match(releaseDoc, /source package, not a published npm package, zip, binary, service installer, or auto-updater yet/);
  assert.match(releaseDoc, /does not produce a signed Office Store package/);
  assert.match(releaseDoc, /1\.0\.0\.2/);
  assert.match(releaseDoc, /20260421/);
  assert.match(releaseDoc, /shortcuts\.json/);

  for (const manifest of manifests) {
    assert.match(manifest, /<Version>1\.0\.0\.2<\/Version>/);
    assert.match(manifest, /https:\/\/localhost:3443\//);
    assert.match(manifest, /brand\/icon-32\.png\?v=20260421/);
    assert.match(manifest, /ExtendedOverrides Url="https:\/\/localhost:3443\/shortcuts\.json"/);
  }
});
