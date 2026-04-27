import assert from "node:assert/strict";
import test from "node:test";

import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import { WORD_OFFICE_TOOL_DEFINITIONS } from "../../../apps/taskpane/src/lib/office/tools/word.js";
import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
import { OFFICE_TOOL_NAMES, TOOL_CATEGORY_MAP, type OfficeToolRequest } from "../../../packages/pi-office-pack/src/protocol.js";

test("Word redline protocol exposes compare/review exchange tooling", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("word_redline_review"));
  assert.equal(TOOL_CATEGORY_MAP.word_redline_review, "write-doc");
  const definition = WORD_OFFICE_TOOL_DEFINITIONS.find((tool) => tool.name === "word_redline_review");
  assert.ok(definition);
  assert.match(definition.description, /compare/i);
  assert.ok(definition.discovery?.capabilityIds?.includes("word.redline"));
});

test("Word redline bridge dispatches structured review exchange actions", async () => {
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
    requestId: "redline",
    toolName: "word_redline_review" as OfficeToolRequest["toolName"],
    host: "word",
    params: { operation: "compare", baselinePath: "C:/Docs/baseline.docx", compareTarget: "Current" },
  } as OfficeToolRequest);

  assert.equal(result.success, true);
  assert.equal(calls[0]?.action.type, "reviewExchange");
  assert.equal((calls[0]?.action.options as Record<string, unknown>).operation, "compare");
});

test("Word redline guidance preserves review boundaries", () => {
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bword_redline_review\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /compare\/redline\/review exchange/i);
});
