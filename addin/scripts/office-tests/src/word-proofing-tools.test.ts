import assert from "node:assert/strict";
import test from "node:test";

import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import { WORD_OFFICE_TOOL_DEFINITIONS } from "../../../apps/taskpane/src/lib/office/tools/word.js";
import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
import { OFFICE_TOOL_NAMES, TOOL_CATEGORY_MAP, type OfficeToolRequest } from "../../../packages/pi-office-pack/src/protocol.js";

test("Word proofing protocol exposes non-mutating statistics tooling", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("word_proofing_stats"));
  assert.equal(TOOL_CATEGORY_MAP.word_proofing_stats, "read");
  const definition = WORD_OFFICE_TOOL_DEFINITIONS.find((tool) => tool.name === "word_proofing_stats");
  assert.ok(definition);
  assert.match(definition.description, /readability/i);
  assert.ok(definition.discovery?.capabilityIds?.includes("word.proofing"));
});

test("Word proofing bridge dispatches read-only proofing action", async () => {
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
    requestId: "proofing",
    toolName: "word_proofing_stats" as OfficeToolRequest["toolName"],
    host: "word",
    params: { scope: "document", includeReadability: true },
  } as OfficeToolRequest);

  assert.equal(result.success, true);
  assert.equal(calls[0]?.action.type, "proofingStats");
  assert.equal((calls[0]?.action.options as Record<string, unknown>).includeReadability, true);
});

test("Word proofing guidance distinguishes native metrics from model judgment", () => {
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bword_proofing_stats\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /native metrics/i);
});
