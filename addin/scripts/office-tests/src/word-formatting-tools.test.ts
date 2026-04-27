import assert from "node:assert/strict";
import test from "node:test";

import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import { WORD_OFFICE_TOOL_DEFINITIONS } from "../../../apps/taskpane/src/lib/office/tools/word.js";
import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
import { OFFICE_TOOL_NAMES, TOOL_CATEGORY_MAP, type OfficeToolRequest } from "../../../packages/pi-office-pack/src/protocol.js";

test("Word formatting protocol and registry expose structured formatting", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("word_format_text"));
  assert.equal(TOOL_CATEGORY_MAP.word_format_text, "write-doc");

  const definition = WORD_OFFICE_TOOL_DEFINITIONS.find((tool) => tool.name === "word_format_text");
  assert.ok(definition);
  assert.equal(definition.category, "write-doc");
  assert.match(definition.description, /style/i);
  assert.ok(definition.discovery?.capabilityIds?.includes("word.formatting"));
  assert.ok(definition.discovery?.keywords?.includes("paragraph"));
});

test("Word format bridge dispatches structured format action and enforces Word-only scope", async () => {
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
    requestId: "word-format",
    toolName: "word_format_text" as OfficeToolRequest["toolName"],
    host: "word",
    params: {
      target: { kind: "paragraph", paragraphId: "p-1" },
      style: "Heading 2",
      alignment: "centered",
      font: { bold: true, color: "#123456" },
    },
  } as OfficeToolRequest);

  assert.equal(result.success, true);
  assert.equal(calls[0]?.host, "word");
  assert.equal(calls[0]?.action.type, "applyTextFormat");
  const options = calls[0]?.action.options as Record<string, unknown>;
  assert.equal(options.bold, true);
  assert.equal(options.color, "#123456");

  const unsupported = await executeOfficeTool({
    requestId: "word-format-excel",
    toolName: "word_format_text" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {},
  } as OfficeToolRequest);
  assert.equal(unsupported.success, false);
  assert.match(unsupported.error ?? "", /only available for Word/i);
});

test("Word formatting guidance prefers structured tool over raw Office.js", () => {
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bword_format_text\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /style, font, highlight, alignment, spacing, indentation/i);
});
