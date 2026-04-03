import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Resvg } from "@resvg/resvg-js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const brandDir = path.resolve(__dirname, "..", "apps", "taskpane", "public", "brand");
const sourcePath = path.join(brandDir, "icon.svg");
const sizes = [16, 32, 80];

const svg = await readFile(sourcePath, "utf8");

for (const size of sizes) {
  const outputPath = path.join(brandDir, `icon-${size}.png`);
  const png = new Resvg(svg, {
    fitTo: { mode: "width", value: size },
  })
    .render()
    .asPng();

  await writeFile(outputPath, png);
  console.log(`Wrote ${path.relative(process.cwd(), outputPath)}`);
}
