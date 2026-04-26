import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import test from "node:test";

test("tool permission timeouts fail closed instead of allowing execution", () => {
  const runtimeSource = readFileSync(
    join(process.cwd(), "apps", "taskpane", "src", "lib", "runtime", "inprocess-kernel.ts"),
    "utf8",
  );

  assert.match(runtimeSource, /tool_permission_expired/);
  assert.match(runtimeSource, /allowed:\s*false/);
  assert.doesNotMatch(runtimeSource, /allowed:\s*true,\s*scope:\s*"once"\s*\}\);/);
});
