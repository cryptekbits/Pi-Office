import { spawnSync } from "node:child_process";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = dirname(fileURLToPath(new URL("../package.json", import.meta.url)));
const manifestDir = join(repoRoot, "manifests");
const files = ["word.xml", "excel.xml", "powerpoint.xml"].map((name) => join(manifestDir, name));
const manifestCliPackage = "office-addin-manifest@2.1.3";

for (const file of files) {
  console.log(`Validating ${file}`);
  const npxArgs = ["--yes", manifestCliPackage, "validate", file];
  const result = process.platform === "win32"
    ? spawnSync("cmd.exe", ["/d", "/s", "/c", "npx", ...npxArgs], {
        cwd: repoRoot,
        stdio: "inherit",
      })
    : spawnSync("npx", npxArgs, {
        cwd: repoRoot,
        stdio: "inherit",
      });

  if (result.error) {
    console.error(result.error.message);
    process.exit(1);
  }

  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
}

console.log("All manifests validated.");
