import assert from "node:assert/strict";
import test from "node:test";

import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import { WORD_OFFICE_TOOL_DEFINITIONS } from "../../../apps/taskpane/src/lib/office/tools/word.js";
import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
import { OFFICE_TOOL_NAMES, TOOL_CATEGORY_MAP, type OfficeToolRequest } from "../../../packages/pi-office-pack/src/protocol.js";

test("Word list protocol and registry expose structured list formatting", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("word_list_format"));
  assert.equal(TOOL_CATEGORY_MAP.word_list_format, "write-doc");
  const definition = WORD_OFFICE_TOOL_DEFINITIONS.find((tool) => tool.name === "word_list_format");
  assert.ok(definition);
  assert.match(definition.description, /bullet/i);
  assert.ok(definition.discovery?.capabilityIds?.includes("word.list"));
});

test("Word list bridge dispatches structured list action and enforces Word-only scope", async () => {
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
    requestId: "word-list",
    toolName: "word_list_format" as OfficeToolRequest["toolName"],
    host: "word",
    params: {
      target: { kind: "paragraph", paragraphId: "p-1" },
      listKind: "bullet",
      level: 2,
      confirmBroadChange: true,
    },
  } as OfficeToolRequest);

  assert.equal(result.success, true);
  assert.equal(calls[0]?.host, "word");
  assert.equal(calls[0]?.action.type, "formatList");
  assert.equal((calls[0]?.action.options as Record<string, unknown>).listType, "bullet");

  const unsupported = await executeOfficeTool({
    requestId: "word-list-excel",
    toolName: "word_list_format" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {},
  } as OfficeToolRequest);
  assert.equal(unsupported.success, false);
  assert.match(unsupported.error ?? "", /only available for Word/i);
});

test("Word list guidance prefers structured tool before text rewrites", () => {
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bword_list_format\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /bullets, numbering, list levels/i);
});
