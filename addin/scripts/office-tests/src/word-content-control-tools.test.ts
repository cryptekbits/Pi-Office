import assert from "node:assert/strict";
import test from "node:test";

import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import { WORD_OFFICE_TOOL_DEFINITIONS } from "../../../apps/taskpane/src/lib/office/tools/word.js";
import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
import { OFFICE_TOOL_NAMES, TOOL_CATEGORY_MAP, type OfficeToolRequest } from "../../../packages/pi-office-pack/src/protocol.js";

test("Word content-control protocol and registry expose template filling", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("word_content_control"));
  assert.equal(TOOL_CATEGORY_MAP.word_content_control, "write-doc");
  const definition = WORD_OFFICE_TOOL_DEFINITIONS.find((tool) => tool.name === "word_content_control");
  assert.ok(definition);
  assert.match(definition.description, /content controls/i);
  assert.ok(definition.discovery?.capabilityIds?.includes("word.contentControls"));
});

test("Word content-control bridge dispatches structured template operations", async () => {
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
    requestId: "template-fill",
    toolName: "word_content_control" as OfficeToolRequest["toolName"],
    host: "word",
    params: {
      operation: "fillText",
      target: { kind: "contentControl", id: "contentControl:42" },
      text: "Filled answer",
      title: "Client Name",
      cannotEdit: true,
    },
  } as OfficeToolRequest);

  assert.equal(result.success, true);
  assert.equal(calls[0]?.action.type, "contentControlEdit");
  assert.equal((calls[0]?.action.options as Record<string, unknown>).operation, "fillText");
});

test("Word content-control guidance prefers explicit template mapping", () => {
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bword_content_control\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /title, or tag/i);
});
