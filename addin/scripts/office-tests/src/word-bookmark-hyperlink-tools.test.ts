import assert from "node:assert/strict";
import test from "node:test";

import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import { WORD_OFFICE_TOOL_DEFINITIONS } from "../../../apps/taskpane/src/lib/office/tools/word.js";
import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
import { OFFICE_ANCHOR_KINDS, OFFICE_TOOL_NAMES, TOOL_CATEGORY_MAP, type OfficeToolRequest } from "../../../packages/pi-office-pack/src/protocol.js";

test("Word bookmark and hyperlink protocol exposes first-class navigation anchors", () => {
  assert.ok(OFFICE_ANCHOR_KINDS.includes("bookmark"));
  assert.ok(OFFICE_ANCHOR_KINDS.includes("hyperlink"));
  assert.ok(OFFICE_TOOL_NAMES.includes("word_reference_inventory"));
  assert.ok(OFFICE_TOOL_NAMES.includes("word_hyperlink"));
  assert.equal(TOOL_CATEGORY_MAP.word_reference_inventory, "read");
  assert.equal(TOOL_CATEGORY_MAP.word_hyperlink, "write-doc");

  assert.ok(WORD_OFFICE_TOOL_DEFINITIONS.find((tool) => tool.name === "word_reference_inventory")?.discovery?.capabilityIds?.includes("word.bookmarks"));
  assert.ok(WORD_OFFICE_TOOL_DEFINITIONS.find((tool) => tool.name === "word_hyperlink")?.discovery?.capabilityIds?.includes("word.hyperlinks"));
});

test("Word bookmark/hyperlink bridge dispatches structured actions", async () => {
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

  const bookmark = await executeOfficeTool({
    requestId: "bookmark",
    toolName: "word_reference_inventory" as OfficeToolRequest["toolName"],
    host: "word",
    params: { operation: "add", name: "PiOffice_Target_1", target: { kind: "selection" } },
  } as OfficeToolRequest);
  const hyperlink = await executeOfficeTool({
    requestId: "hyperlink",
    toolName: "word_hyperlink" as OfficeToolRequest["toolName"],
    host: "word",
    params: { operation: "add", address: "https://example.com", textToDisplay: "Example" },
  } as OfficeToolRequest);

  assert.equal(bookmark.success, true);
  assert.equal(hyperlink.success, true);
  assert.equal(calls[0]?.action.type, "bookmark");
  assert.equal(calls[1]?.action.type, "hyperlink");
  assert.equal((calls[1]?.action.options as Record<string, unknown>).address, "https://example.com");
});

test("Word bookmark/hyperlink guidance prefers structured anchors", () => {
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /word_reference_inventory/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /word_hyperlink/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /durable long-document anchors/i);
});
