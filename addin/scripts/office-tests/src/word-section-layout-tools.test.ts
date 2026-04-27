import assert from "node:assert/strict";
import test from "node:test";

import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import {
  countWordOoxmlBreaks,
  normalizeWordBreakType,
  wordOoxmlHasAdjacentPageBreak,
} from "../../../apps/taskpane/src/lib/office/word-actions.js";
import { WORD_OFFICE_TOOL_DEFINITIONS } from "../../../apps/taskpane/src/lib/office/tools/word.js";
import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
import { OFFICE_TOOL_NAMES, TOOL_CATEGORY_MAP, type OfficeToolRequest } from "../../../packages/pi-office-pack/src/protocol.js";

test("Word section layout protocol and registry expose headers, footers, page setup, and breaks", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("word_section_layout"));
  assert.equal(TOOL_CATEGORY_MAP.word_section_layout, "write-doc");
  const definition = WORD_OFFICE_TOOL_DEFINITIONS.find((tool) => tool.name === "word_section_layout");
  assert.ok(definition);
  assert.match(definition.description, /header\/footer/i);
  assert.ok(definition.discovery?.capabilityIds?.includes("word.sections"));
});

test("Word section layout bridge dispatches structured actions and enforces Word-only scope", async () => {
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
    requestId: "word-section",
    toolName: "word_section_layout" as OfficeToolRequest["toolName"],
    host: "word",
    params: { operation: "setHeader", sectionIndex: 1, headerFooterType: "primary", text: "Confidential" },
  } as OfficeToolRequest);

  assert.equal(result.success, true);
  assert.equal(calls[0]?.host, "word");
  assert.equal(calls[0]?.action.type, "sectionLayout");
  assert.equal((calls[0]?.action.options as Record<string, unknown>).operation, "setHeader");

  const breakResult = await executeOfficeTool({
    requestId: "word-page-break",
    toolName: "word_section_layout" as OfficeToolRequest["toolName"],
    host: "word",
    params: {
      operation: "insertBreak",
      breakType: "page",
      placement: "before",
      target: { kind: "heading", text: "7. Conclusion" },
    },
  } as OfficeToolRequest);

  assert.equal(breakResult.success, true);
  assert.equal(calls[1]?.action.type, "sectionLayout");
  assert.equal((calls[1]?.action.target as { kind?: string }).kind, "heading");
  assert.equal((calls[1]?.action.options as Record<string, unknown>).operation, "insertBreak");
  assert.equal((calls[1]?.action.options as Record<string, unknown>).breakType, "page");
  assert.equal((calls[1]?.action.options as Record<string, unknown>).allowDuplicatePageBreak, undefined);
  assert.equal(calls[1]?.action.placement, "before");

  const unsupported = await executeOfficeTool({
    requestId: "word-section-excel",
    toolName: "word_section_layout" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {},
  } as OfficeToolRequest);
  assert.equal(unsupported.success, false);
  assert.match(unsupported.error ?? "", /only available for Word/i);
});

test("Word page-break helpers normalize Office.js casing and count persisted OOXML evidence", () => {
  assert.equal(normalizeWordBreakType("page"), "Page");
  assert.equal(normalizeWordBreakType("Section_Next"), "SectionNext");
  assert.equal(normalizeWordBreakType("section continuous"), "SectionContinuous");
  assert.throws(() => normalizeWordBreakType("chapter"), /Unsupported Word breakType/);

  const ooxml = '<w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:r><w:br/></w:r></w:p><w:br w:type="page"></w:br>';
  assert.equal(countWordOoxmlBreaks(ooxml, "Page"), 2);
  assert.equal(countWordOoxmlBreaks(ooxml, "SectionNext"), 0);
});

test("Word page-break helper detects an existing adjacent break before a heading", () => {
  const ooxml = [
    '<w:p><w:r><w:t>1. Executive Summary</w:t></w:r></w:p>',
    '<w:p><w:r><w:br w:type="page"/></w:r></w:p>',
    '<w:p><w:r><w:t>7. Conclusion</w:t></w:r></w:p>',
  ].join("");

  assert.equal(wordOoxmlHasAdjacentPageBreak(ooxml, "7. Conclusion", "before"), true);
  assert.equal(wordOoxmlHasAdjacentPageBreak(ooxml, "1. Executive Summary", "before"), false);
});

test("Word section layout guidance separates headers and footers from document body edits", () => {
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bword_section_layout\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /headers\/footers/i);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /Do not claim a page break landed/i);
});
