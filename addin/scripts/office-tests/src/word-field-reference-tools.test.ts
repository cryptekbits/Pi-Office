import assert from "node:assert/strict";
import test from "node:test";

import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import { WORD_OFFICE_TOOL_DEFINITIONS } from "../../../apps/taskpane/src/lib/office/tools/word.js";
import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
import { OFFICE_TOOL_NAMES, TOOL_CATEGORY_MAP, type OfficeToolRequest } from "../../../packages/pi-office-pack/src/protocol.js";

test("Word field protocol and registry expose structural reference tooling", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("word_field_reference"));
  assert.equal(TOOL_CATEGORY_MAP.word_field_reference, "write-doc");
  const definition = WORD_OFFICE_TOOL_DEFINITIONS.find((tool) => tool.name === "word_field_reference");
  assert.ok(definition);
  assert.match(definition.description, /field/i);
  assert.ok(definition.discovery?.capabilityIds?.includes("word.fields"));
});

test("Word field bridge dispatches structured field actions", async () => {
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
    requestId: "field-update",
    toolName: "word_field_reference" as OfficeToolRequest["toolName"],
    host: "word",
    params: { operation: "updateField", fieldId: "field:1" },
  } as OfficeToolRequest);

  assert.equal(result.success, true);
  assert.equal(calls[0]?.host, "word");
  assert.equal(calls[0]?.action.type, "fieldAction");
  assert.equal((calls[0]?.action.options as Record<string, unknown>).operation, "update");
  assert.deepEqual(calls[0]?.action.target, { kind: "field", id: "field:1" });
});

test("Word field guidance separates fields and generated references from text rewrites", () => {
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bword_field_reference\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /inventory\/update\/lock\/select\/insert fields/i);
});
