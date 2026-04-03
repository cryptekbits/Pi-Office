import { fileURLToPath, URL } from "node:url";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

const vscodeJsonRpcCommonDir = new URL(
  "../../node_modules/vscode-languageserver-protocol/node_modules/vscode-jsonrpc/lib/common/",
  import.meta.url,
);

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      "vscode-jsonrpc/lib/common/cancellation.js": fileURLToPath(new URL("cancellation.js", vscodeJsonRpcCommonDir)),
      "vscode-jsonrpc/lib/common/events.js": fileURLToPath(new URL("events.js", vscodeJsonRpcCommonDir)),
    },
  },
  build: {
    outDir: "dist",
    sourcemap: true,
  },
});
