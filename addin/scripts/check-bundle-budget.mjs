import { readdirSync, statSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const rootDir = join(fileURLToPath(new URL(".", import.meta.url)), "..");
const assetsDir = join(rootDir, "apps", "taskpane", "dist", "assets");

const MAX_MAIN_JS_BYTES = 3_000_000;
const MAX_MAIN_CSS_BYTES = 130_000;
const MAX_SINGLE_CHUNK_BYTES = 900_000;

function readAssetSize(fileName) {
  const fullPath = join(assetsDir, fileName);
  return statSync(fullPath).size;
}

function formatSize(bytes) {
  return `${(bytes / 1024).toFixed(1)} KiB`;
}

function fail(message) {
  console.error(`[bundle-budget] ${message}`);
  process.exit(1);
}

const files = readdirSync(assetsDir);
const mainJs = files.find((file) => /^index-.*\.js$/.test(file));
const mainCss = files.find((file) => /^index-.*\.css$/.test(file));
if (!mainJs || !mainCss) {
  fail("Could not find taskpane index asset files. Run build before checking budgets.");
}

const mainJsBytes = readAssetSize(mainJs);
const mainCssBytes = readAssetSize(mainCss);

if (mainJsBytes > MAX_MAIN_JS_BYTES) {
  fail(`Main JS bundle exceeds budget: ${formatSize(mainJsBytes)} > ${formatSize(MAX_MAIN_JS_BYTES)}.`);
}
if (mainCssBytes > MAX_MAIN_CSS_BYTES) {
  fail(`Main CSS bundle exceeds budget: ${formatSize(mainCssBytes)} > ${formatSize(MAX_MAIN_CSS_BYTES)}.`);
}

const chunkViolations = files
  .filter((file) => file.endsWith(".js"))
  .filter((file) => file !== mainJs)
  .map((file) => ({ file, size: readAssetSize(file) }))
  .filter(({ size }) => size > MAX_SINGLE_CHUNK_BYTES)
  .sort((left, right) => right.size - left.size);

if (chunkViolations.length) {
  const details = chunkViolations
    .map(({ file, size }) => `${file} (${formatSize(size)})`)
    .join(", ");
  fail(`One or more JS chunks exceed budget ${formatSize(MAX_SINGLE_CHUNK_BYTES)}: ${details}`);
}

console.log(
  `[bundle-budget] OK main.js=${formatSize(mainJsBytes)} main.css=${formatSize(mainCssBytes)} maxChunk<=${formatSize(MAX_SINGLE_CHUNK_BYTES)}`,
);
