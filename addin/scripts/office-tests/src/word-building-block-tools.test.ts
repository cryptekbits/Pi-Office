import assert from "node:assert/strict";
import test from "node:test";

import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import { WORD_OFFICE_TOOL_DEFINITIONS } from "../../../apps/taskpane/src/lib/office/tools/word.js";
import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
import { OFFICE_TOOL_NAMES, TOOL_CATEGORY_MAP, type OfficeToolRequest } from "../../../packages/pi-office-pack/src/protocol.js";

test("Word building block protocol exposes provenance-gated reusable content tooling", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("word_building_block"));
  assert.equal(TOOL_CATEGORY_MAP.word_building_block, "write-doc");
  const definition = WORD_OFFICE_TOOL_DEFINITIONS.find((tool) => tool.name === "word_building_block");
  assert.ok(definition);
  assert.match(definition.description, /building block/i);
  assert.ok(definition.discovery?.capabilityIds?.includes("word.buildingBlocks"));
});

test("Word building block bridge dispatches structured insert action", async () => {
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
    requestId: "building-block",
    toolName: "word_building_block" as OfficeToolRequest["toolName"],
    host: "word",
    params: {
      operation: "insert",
      name: "Approved Clause",
      richText: true,
      provenanceApproved: true,
    },
  } as OfficeToolRequest);

  assert.equal(result.success, true);
  assert.equal(calls[0]?.action.type, "buildingBlock");
  assert.equal((calls[0]?.action.options as Record<string, unknown>).provenanceApproved, true);
});

test("Word building block guidance prevents fabricated reusable content", () => {
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bword_building_block\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /Do not invent approved reusable content/i);
});
