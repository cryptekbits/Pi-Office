import { readdirSync, rmSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

const rootDir = join(fileURLToPath(new URL(".", import.meta.url)), "..");
const outDir = join(rootDir, ".codex-office-tests-dist");

function run(command, args) {
  const result = spawnSync(command, args, {
    cwd: rootDir,
    stdio: "inherit",
    shell: process.platform === "win32",
  });
  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
}

rmSync(outDir, { recursive: true, force: true });

try {
  run("npm", ["run", "build", "--workspace", "@pi-office/pi-office-pack"]);
  run("npx", ["tsc", "-p", "scripts/office-tests/tsconfig.json", "--outDir", outDir]);
  const testDir = join(outDir, "scripts/office-tests/src");
  const testFiles = readdirSync(testDir)
    .filter((entry) => entry.endsWith(".test.js"))
    .map((entry) => join(testDir, entry));
  run("node", ["--test", ...testFiles]);
} finally {
  rmSync(outDir, { recursive: true, force: true });
}
