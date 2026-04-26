import { spawnSync } from "node:child_process";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = dirname(fileURLToPath(new URL("../package.json", import.meta.url)));
const manifestDir = join(repoRoot, "manifests");
const files = ["word.xml", "excel.xml", "powerpoint.xml"].map((name) => join(manifestDir, name));

for (const file of files) {
  console.log(`Validating ${file}`);
  const result =
    process.platform === "win32"
      ? spawnSync("cmd.exe", ["/d", "/s", "/c", "npx", "office-addin-manifest", "validate", file], {
          cwd: repoRoot,
          stdio: "inherit",
        })
      : spawnSync("npx", ["office-addin-manifest", "validate", file], {
          cwd: repoRoot,
          stdio: "inherit",
        });

  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
}

console.log("All manifests validated.");
