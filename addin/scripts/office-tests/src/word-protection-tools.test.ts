import assert from "node:assert/strict";
import test from "node:test";

import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import { WORD_OFFICE_TOOL_DEFINITIONS } from "../../../apps/taskpane/src/lib/office/tools/word.js";
import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
import { OFFICE_TOOL_NAMES, TOOL_CATEGORY_MAP, type OfficeToolRequest } from "../../../packages/pi-office-pack/src/protocol.js";

test("Word protection protocol exposes collaboration diagnostics", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("word_collab_guard"));
  assert.equal(TOOL_CATEGORY_MAP.word_collab_guard, "read");
  const definition = WORD_OFFICE_TOOL_DEFINITIONS.find((tool) => tool.name === "word_collab_guard");
  assert.ok(definition);
  assert.match(definition.description, /protection/i);
  assert.ok(definition.discovery?.capabilityIds?.includes("word.protection"));
});

test("Word protection bridge dispatches non-mutating protection action", async () => {
  const calls: Array<{ host: string; action: Record<string, unknown> }> = [];
  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async () => ({ ok: true }),
    applyHostAction: async (host, action) => {
      calls.push({ host, action });
      return { ok: true, action: action.type };
    },
    navigateOfficeAnchor: async () => ({ ok: true }),
    readDocumentSection: async () => ({ ok: true }),
    executeOfficeJs: async () => ({ ok: true }),
    proposeEdits: async () => ({ ok: true }),
  });

  const result = await executeOfficeTool({
    requestId: "protect",
    toolName: "word_collab_guard" as OfficeToolRequest["toolName"],
    host: "word",
    params: { operation: "diagnostics" },
  } as OfficeToolRequest);

  assert.equal(result.success, true);
  assert.equal(calls[0]?.action.type, "collaborationGuard");
});

test("Word protection guidance warns before risky writes", () => {
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bword_collab_guard\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /protected or conflicted ranges/i);
});
