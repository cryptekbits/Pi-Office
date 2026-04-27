import assert from "node:assert/strict";
import test from "node:test";

import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import { WORD_OFFICE_TOOL_DEFINITIONS } from "../../../apps/taskpane/src/lib/office/tools/word.js";
import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
import { OFFICE_TOOL_NAMES, TOOL_CATEGORY_MAP, type OfficeToolRequest } from "../../../packages/pi-office-pack/src/protocol.js";

test("Word table protocol and registry expose structured table operations", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("word_table"));
  assert.equal(TOOL_CATEGORY_MAP.word_table, "write-doc");
  const definition = WORD_OFFICE_TOOL_DEFINITIONS.find((tool) => tool.name === "word_table");
  assert.ok(definition);
  assert.match(definition.description, /cell/i);
  assert.ok(definition.discovery?.capabilityIds?.includes("word.tables"));
});

test("Word table bridge dispatches structured table actions and enforces Word-only scope", async () => {
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
    requestId: "word-table",
    toolName: "word_table" as OfficeToolRequest["toolName"],
    host: "word",
    params: { operation: "setCellText", tableIndex: 1, rowIndex: 2, columnIndex: 1, text: "Updated" },
  } as OfficeToolRequest);

  assert.equal(result.success, true);
  assert.equal(calls[0]?.host, "word");
  assert.equal(calls[0]?.action.type, "tableEdit");
  assert.equal((calls[0]?.action.options as Record<string, unknown>).operation, "setCellText");

  const unsupported = await executeOfficeTool({
    requestId: "word-table-excel",
    toolName: "word_table" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {},
  } as OfficeToolRequest);
  assert.equal(unsupported.success, false);
  assert.match(unsupported.error ?? "", /only available for Word/i);
});

test("Word table guidance prefers structured table tools over OOXML rewrites", () => {
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bword_table\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /cell text, row\/column changes/i);
});
