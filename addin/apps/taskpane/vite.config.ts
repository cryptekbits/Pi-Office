import { createRequire } from "node:module";
import { readFileSync } from "node:fs";
import { fileURLToPath, pathToFileURL, URL } from "node:url";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

const require = createRequire(import.meta.url);
const vscodeJsonRpcCommonDir = new URL("./lib/common/", pathToFileURL(require.resolve("vscode-jsonrpc/package.json")));
const certDir = new URL("../../certs/", import.meta.url);
const certPfxPath = fileURLToPath(new URL("localhost.pfx", certDir));
const certPassphrasePath = fileURLToPath(new URL("passphrase.txt", certDir));
const chunkSizeWarningLimitKb = 1600;

function loadDevServerConfig() {
  return {
    host: "localhost",
    port: 3443,
    strictPort: true,
    https: {
      pfx: readFileSync(certPfxPath),
      passphrase: readFileSync(certPassphrasePath, "utf8").trim(),
    },
  };
}

export default defineConfig(({ command }) => ({
  plugins: [react()],
  ...(command === "serve" ? { server: loadDevServerConfig() } : {}),
  resolve: {
    alias: {
      "vscode-jsonrpc/lib/common/cancellation.js": fileURLToPath(new URL("cancellation.js", vscodeJsonRpcCommonDir)),
      "vscode-jsonrpc/lib/common/events.js": fileURLToPath(new URL("events.js", vscodeJsonRpcCommonDir)),
    },
  },
  build: {
    outDir: "dist",
    sourcemap: true,
    chunkSizeWarningLimit: chunkSizeWarningLimitKb,
  },
}));
