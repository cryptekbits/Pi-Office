import assert from "node:assert/strict";
import test from "node:test";

import { McpResultStore } from "../../../apps/taskpane/src/lib/runtime/mcp-result-store.js";
import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
import { OFFICE_TOOL_NAMES, TOOL_CATEGORY_MAP } from "../../../packages/pi-office-pack/src/protocol.js";

test("protocol exposes MCP result handle tools with safe categories", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("mcp_result_get"));
  assert.ok(OFFICE_TOOL_NAMES.includes("mcp_result_summarize"));
  assert.ok(OFFICE_TOOL_NAMES.includes("mcp_result_clear"));
  assert.equal(TOOL_CATEGORY_MAP.mcp_result_get, "read");
  assert.equal(TOOL_CATEGORY_MAP.mcp_result_summarize, "read");
  assert.equal(TOOL_CATEGORY_MAP.mcp_result_clear, "connector");
});

test("MCP result store pages, summarizes, and clears large connector payloads", () => {
  const store = new McpResultStore();
  const value = { rows: Array.from({ length: 400 }, (_, index) => ({ index, text: `row ${index} alpha beta gamma` })) };
  const stored = store.store({
    source: "browser",
    connectorId: "parallel",
    connectorName: "Parallel Search",
    toolName: "parallel_search",
    value,
    pageSizeBytes: 1200,
  });

  assert.ok(stored.handle.pageCount > 1);
  assert.equal(stored.handle.connectorId, "parallel");
  assert.match(stored.handle.handleId, /^mcp-result-browser-/);
  assert.ok("resultHandle" in (stored.content as Record<string, unknown>));

  const page = store.getPage({ handleId: stored.handle.handleId, page: 1 });
  assert.equal(page.ok, true);
  assert.equal(page.page, 1);
  assert.ok(page.content.length > 0);

  const summary = store.summarize({ handleId: stored.handle.handleId, query: "row 399", maxChars: 500 });
  assert.equal(summary.ok, true);
  assert.match(summary.summary, /row 399/);

  const clearOne = store.clear({ handleId: stored.handle.handleId });
  assert.equal(clearOne.cleared, 1);
  assert.throws(() => store.getPage({ handleId: stored.handle.handleId }), /Unknown or expired/);
});

test("MCP result store distinguishes companion handles for taskpane routing", () => {
  const store = new McpResultStore();
  const stored = store.store({
    source: "companion",
    connectorId: "local-docs",
    connectorName: "Local Docs",
    toolName: "local_docs_search",
    value: "x".repeat(3_000),
    pageSizeBytes: 1024,
  });

  assert.match(stored.handle.handleId, /^mcp-result-companion-/);
  assert.equal(stored.handle.source, "companion");
  assert.ok("resultHandle" in (stored.content as Record<string, unknown>));
});

test("MCP result guidance tells the model to use handles instead of flooding chat", () => {
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /mcp_result_get/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /mcp_result_summarize/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /Large MCP responses/i);
});
