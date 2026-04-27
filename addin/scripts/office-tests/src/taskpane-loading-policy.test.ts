import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import test from "node:test";

function readProjectFile(relativePath: string): string {
  return readFileSync(join(process.cwd(), relativePath), "utf8");
}

test("taskpane startup keeps connector runtime and settings UI behind lazy load boundaries", () => {
  const kernelSource = readProjectFile("apps/taskpane/src/lib/runtime/inprocess-kernel.ts");
  const appSource = readProjectFile("apps/taskpane/src/app/App.tsx");
  const iconsSource = readProjectFile("apps/taskpane/src/lib/icons.tsx");
  const taskpanePackage = JSON.parse(readProjectFile("apps/taskpane/package.json")) as {
    dependencies?: Record<string, string>;
  };

  assert.match(
    kernelSource,
    /import type \{ BrowserConnectorRuntime \} from "\.\/browser-connectors";/,
    "The browser connector runtime should be type-only in the startup kernel.",
  );
  assert.doesNotMatch(
    kernelSource,
    /import \{ BrowserConnectorRuntime \} from "\.\/browser-connectors";/,
    "The browser connector runtime must not be statically imported by the startup kernel.",
  );
  assert.match(
    kernelSource,
    /import\("\.\/browser-connectors"\)/,
    "Standalone connector support should load the browser connector runtime dynamically.",
  );

  assert.match(
    appSource,
    /const SettingsPage = lazy\(\(\) =>\s*import\("\.\/components\/SettingsPage"\)/s,
    "Settings and Integrations UI should stay out of the first chat render.",
  );
  assert.doesNotMatch(
    appSource,
    /import \{ SettingsPage \} from "\.\/components\/SettingsPage";/,
    "SettingsPage must not be eagerly imported by App.",
  );

  assert.doesNotMatch(iconsSource, /from "simple-icons"/, "Connector icons should not import the full simple-icons package.");
  assert.ok(!taskpanePackage.dependencies?.["simple-icons"], "simple-icons should not remain a taskpane runtime dependency.");
});
