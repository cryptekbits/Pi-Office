import assert from "node:assert/strict";
import test from "node:test";

import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import { WORD_OFFICE_TOOL_DEFINITIONS } from "../../../apps/taskpane/src/lib/office/tools/word.js";
import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
import {
  OFFICE_ANCHOR_KINDS,
  OFFICE_TOOL_NAMES,
  TOOL_CATEGORY_MAP,
  type OfficeToolRequest,
} from "../../../packages/pi-office-pack/src/protocol.js";

test("Word search protocol and registry expose first-class search anchors", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("word_search"));
  assert.equal(TOOL_CATEGORY_MAP.word_search, "read");
  assert.ok(OFFICE_ANCHOR_KINDS.includes("searchResult"));

  const definition = WORD_OFFICE_TOOL_DEFINITIONS.find((tool) => tool.name === "word_search");
  assert.ok(definition);
  assert.equal(definition.category, "read");
  assert.match(definition.description, /paragraphs, headings, comments/i);
  assert.ok(definition.discovery?.capabilityIds?.includes("word.search"));
  assert.ok(definition.discovery?.keywords?.includes("anchor"));
});

test("Word search bridge dispatches to dependency and enforces Word-only scope", async () => {
  const calls: Array<Record<string, unknown>> = [];
  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async () => ({ ok: true }),
    applyHostAction: async () => ({ ok: true }),
    navigateOfficeAnchor: async () => ({ ok: true }),
    readDocumentSection: async () => ({ ok: true }),
    executeOfficeJs: async () => ({ ok: true }),
    proposeEdits: async () => ({ ok: true }),
    searchWordDocument: async (params) => {
      calls.push(params);
      return {
        ok: true,
        query: params.query,
        results: [
          {
            anchor: { kind: "searchResult", id: "word-search:paragraph:1", paragraphId: "p1", text: "Alpha" },
            objectType: "paragraph",
          },
        ],
      };
    },
  });

  const result = await executeOfficeTool({
    requestId: "word-search",
    toolName: "word_search" as OfficeToolRequest["toolName"],
    host: "word",
    params: { query: "Alpha", objectTypes: ["paragraph"], maxResults: 5 },
  } as OfficeToolRequest);

  assert.equal(result.success, true);
  assert.equal(calls[0]?.query, "Alpha");
  assert.equal((result.content as { results: Array<{ anchor: { kind: string } }> }).results[0]?.anchor.kind, "searchResult");

  const unsupported = await executeOfficeTool({
    requestId: "word-search-excel",
    toolName: "word_search" as OfficeToolRequest["toolName"],
    host: "excel",
    params: { query: "Alpha" },
  } as OfficeToolRequest);
  assert.equal(unsupported.success, false);
  assert.match(String(unsupported.error), /only available for Word/);
});

test("Word search guidance tells model to search before deterministic edits", () => {
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /word_search/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /paragraphId/i);
});
