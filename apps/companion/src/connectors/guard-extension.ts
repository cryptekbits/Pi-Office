import type { ExtensionAPI } from "@mariozechner/pi-coding-agent";
import type { ConnectorGuardSnapshot } from "./runtime.js";

function findServerByTool(toolName: string, snapshot: ConnectorGuardSnapshot) {
  return snapshot.servers.find(
    (server) => server.allowedTools.includes(toolName) || server.blockedTools.includes(toolName),
  );
}

export function createConnectorGuardExtension(
  getSnapshot: () => ConnectorGuardSnapshot,
  onToolUse?: (toolName: string, blocked: boolean) => void,
) {
  return (pi: ExtensionAPI) => {
    pi.on("before_agent_start", async (event) => {
      const snapshot = getSnapshot();
      if (!snapshot.servers.length) {
        return undefined;
      }

      const summary = snapshot.servers
        .map((server) => {
          const promptSummary = server.allowedPrompts.length ? `, ${server.allowedPrompts.length} prompts` : "";
          return `- ${server.connectorName}: ${server.allowedTools.length} read-safe tools${promptSummary}`;
        })
        .join("\n");

      return {
        systemPrompt: `${event.systemPrompt}

Only use MCP connectors in read-only mode.
Search before calling a connector tool when you are unsure which one applies.
Never attempt write-like connector actions such as create, update, delete, send, reply, comment, merge, refund, or approve.

Verified connector access:
${summary}`,
      };
    });

    pi.on("tool_call", async (event) => {
      const snapshot = getSnapshot();
      if (!snapshot.servers.length) {
        return undefined;
      }

      if (event.toolName === "mcp") {
        const input = event.input as Record<string, unknown>;
        const action = typeof input.action === "string" ? input.action : undefined;
        if (action === "ui-messages") {
          return { block: true, reason: "Connector UIs are disabled in this add-in." };
        }

        const requestedTool = typeof input.tool === "string" ? input.tool : undefined;
        if (!requestedTool) {
          return undefined;
        }

        const explicitServer = typeof input.server === "string" ? input.server : undefined;
        const policy = explicitServer
          ? snapshot.servers.find((server) => server.serverName === explicitServer)
          : findServerByTool(requestedTool, snapshot);

        if (!policy) {
          onToolUse?.(requestedTool, true);
          return { block: true, reason: "This connector tool has not been verified for read-only use." };
        }

        if (!policy.allowedTools.includes(requestedTool)) {
          onToolUse?.(requestedTool, true);
          return { block: true, reason: `${policy.connectorName} is connected in read-only mode.` };
        }

        onToolUse?.(requestedTool, false);
        return undefined;
      }

      const server = findServerByTool(event.toolName, snapshot);
      if (server && !server.allowedTools.includes(event.toolName)) {
        onToolUse?.(event.toolName, true);
        return { block: true, reason: `${server.connectorName} is connected in read-only mode.` };
      }

      if (server) {
        onToolUse?.(event.toolName, false);
      }

      return undefined;
    });
  };
}
