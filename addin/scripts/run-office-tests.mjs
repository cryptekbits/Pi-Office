import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync, symlinkSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

const rootDir = join(fileURLToPath(new URL(".", import.meta.url)), "..");
const repoRoot = join(rootDir, "..");
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

function listJsFiles(directory) {
  const entries = readdirSync(directory);
  const files = [];
  for (const entry of entries) {
    const fullPath = join(directory, entry);
    const stats = statSync(fullPath);
    if (stats.isDirectory()) {
      files.push(...listJsFiles(fullPath));
      continue;
    }
    if (entry.endsWith(".js")) {
      files.push(fullPath);
    }
  }
  return files;
}

function normalizeRelativeImports(filePath) {
  const source = readFileSync(filePath, "utf8");
  const rewriteSpecifier = (specifier) => {
    if (!specifier.startsWith(".") || /\/$/.test(specifier)) return specifier;
    if (/\.[a-z0-9]+$/i.test(specifier)) return specifier;
    return `${specifier}.js`;
  };

  const rewritten = source
    .replace(/from\s+(['"])(\.{1,2}\/[^'"]+)\1/g, (match, quote, specifier) => {
      const next = rewriteSpecifier(specifier);
      return next === specifier ? match : `from ${quote}${next}${quote}`;
    })
    .replace(/import\(\s*(['"])(\.{1,2}\/[^'"]+)\1\s*\)/g, (match, quote, specifier) => {
      const next = rewriteSpecifier(specifier);
      return next === specifier ? match : `import(${quote}${next}${quote})`;
    });

  if (rewritten !== source) {
    writeFileSync(filePath, rewritten, "utf8");
  }
}

function rewriteCompiledRelativeImports(directory) {
  for (const filePath of listJsFiles(directory)) {
    normalizeRelativeImports(filePath);
  }
}

function linkCompanionDependencies() {
  const companionNodeModules = join(repoRoot, "companion", "node_modules");
  if (!existsSync(companionNodeModules)) return;

  const compiledCompanionDir = join(outDir, "companion");
  const compiledCompanionNodeModules = join(compiledCompanionDir, "node_modules");
  mkdirSync(compiledCompanionDir, { recursive: true });
  if (!existsSync(compiledCompanionNodeModules)) {
    symlinkSync(companionNodeModules, compiledCompanionNodeModules, process.platform === "win32" ? "junction" : "dir");
  }
}

rmSync(outDir, { recursive: true, force: true });

try {
  run("npm", ["run", "build", "--workspace", "@pi-office/pi-office-pack"]);
  run("npx", ["tsc", "-p", "scripts/office-tests/tsconfig.json", "--outDir", outDir]);
  rewriteCompiledRelativeImports(outDir);
  linkCompanionDependencies();
  const testDir = join(outDir, "addin/scripts/office-tests/src");
  const testFiles = readdirSync(testDir)
    .filter((entry) => entry.endsWith(".test.js"))
    .map((entry) => join(testDir, entry));
  run("node", ["--test", ...testFiles]);
} finally {
  rmSync(outDir, { recursive: true, force: true });
}
