import assert from "node:assert/strict";
import test from "node:test";

import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import { WORD_OFFICE_TOOL_DEFINITIONS } from "../../../apps/taskpane/src/lib/office/tools/word.js";
import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
import { OFFICE_TOOL_NAMES, TOOL_CATEGORY_MAP, type OfficeToolRequest } from "../../../packages/pi-office-pack/src/protocol.js";

test("Word annotation protocol exposes native critique fallback tooling", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("word_annotation_review"));
  assert.equal(TOOL_CATEGORY_MAP.word_annotation_review, "write-doc");
  const definition = WORD_OFFICE_TOOL_DEFINITIONS.find((tool) => tool.name === "word_annotation_review");
  assert.ok(definition);
  assert.match(definition.description, /critique/i);
  assert.ok(definition.discovery?.capabilityIds?.includes("word.annotations"));
});

test("Word annotation bridge dispatches structured annotation action", async () => {
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
    requestId: "annotation",
    toolName: "word_annotation_review" as OfficeToolRequest["toolName"],
    host: "word",
    params: { operation: "insert", critiques: [{ text: "Tighten wording", start: 0, length: 5 }] },
  } as OfficeToolRequest);

  assert.equal(result.success, true);
  assert.equal(calls[0]?.action.type, "critiqueAnnotation");
  assert.equal((calls[0]?.action.options as { operation?: string }).operation, "propose");
});

test("Word annotation guidance preserves sidepane fallback", () => {
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bword_annotation_review\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /office_propose_edits/i);
});
