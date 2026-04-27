import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import test from "node:test";

test("model sync suppresses stale-session Unknown session errors in the chat transcript", () => {
  const appSource = readFileSync(join(process.cwd(), "apps", "taskpane", "src", "app", "App.tsx"), "utf8");

  assert.match(appSource, /const targetSessionId = sessionId;/);
  assert.match(appSource, /sessionIdRef\.current === targetSessionId/);
  assert.match(
    appSource,
    /sessionIdRef\.current === targetSessionId\)[\s\S]*pushErrorMessage\(`Model change failed:/,
  );
  assert.doesNotMatch(appSource, /if \(active\) pushErrorMessage\(`Model change failed:/);
});
