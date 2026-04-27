import assert from "node:assert/strict";
import test from "node:test";

import { executeOfficeBatchPlan } from "../../../apps/taskpane/src/lib/runtime/batch-executor.js";
import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import { OFFICE_TOOL_NAMES, TOOL_CATEGORY_MAP, type OfficeToolRequest } from "../../../packages/pi-office-pack/src/protocol.js";

test("protocol exposes typed batch tools with safe categories", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("office_batch_execute"));
  assert.ok(OFFICE_TOOL_NAMES.includes("mcp_batch_execute"));
  assert.equal(TOOL_CATEGORY_MAP.office_batch_execute, "write-doc");
  assert.equal(TOOL_CATEGORY_MAP.mcp_batch_execute, "connector");
});

test("office batch execution runs allowlisted steps and rejects raw Office.js", async () => {
  const calls: Array<{ toolName: string; params: Record<string, unknown> }> = [];
  const result = await executeOfficeBatchPlan({
    host: "word",
    request: {
      steps: [
        { id: "inspect", type: "tool", toolName: "office_get_context", arguments: { scope: "document" } },
        { id: "verify", type: "tool", toolName: "verify_doc", arguments: { scope: "document" } },
      ],
    },
    invokeOfficeTool: async (toolName: OfficeToolRequest["toolName"], params: Record<string, unknown>) => {
      calls.push({ toolName, params });
      return { ok: true, toolName, params };
    },
  });

  assert.equal(result.ok, true);
  assert.match(result.summary, /2\/2/);
  assert.deepEqual(calls.map((entry) => entry.toolName), ["office_get_context", "verify_doc"]);

  const rejected = await executeOfficeBatchPlan({
    host: "word",
    request: {
      steps: [
        { id: "raw", type: "tool", toolName: "office_execute_js", arguments: { code: "Word.run(async()=>{})" } },
      ],
    },
    invokeOfficeTool: async () => {
      throw new Error("raw code should not execute");
    },
  });
  assert.equal(rejected.ok, false);
  assert.equal(rejected.steps[0]?.status, "blocked");
  assert.match(rejected.steps[0]?.error ?? "", /not allowed inside structured batches/i);
});

test("office batch execution relies on outer write-doc permission instead of model-supplied approval flags", async () => {
  const calls: string[] = [];
  const result = await executeOfficeBatchPlan({
    host: "word",
    request: {
      steps: [
        { id: "write", type: "tool", toolName: "office_apply_edit", arguments: { action: { type: "insertText", content: "x" } } },
      ],
    },
    invokeOfficeTool: async (toolName) => {
      calls.push(toolName);
      return { ok: true };
    },
  });

  assert.equal(result.ok, true);
  assert.deepEqual(calls, ["office_apply_edit"]);
});

test("taskpane bridge keeps office_batch_execute runtime-owned", async () => {
  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async () => ({ ok: true }),
    applyHostAction: async () => ({ ok: true }),
    navigateOfficeAnchor: async () => ({ ok: true }),
    readDocumentSection: async () => ({ ok: true }),
    executeOfficeJs: async () => ({ ok: true }),
    proposeEdits: async () => ({ ok: true }),
  });

  const result = await executeOfficeTool({
    requestId: "batch-bridge",
    toolName: "office_batch_execute" as OfficeToolRequest["toolName"],
    host: "word",
    params: {
      steps: [
        { id: "context", type: "tool", toolName: "office_get_context", arguments: { scope: "document" } },
      ],
    },
  } as OfficeToolRequest);

  assert.equal(result.success, false);
  assert.match(result.error ?? "", /taskpane runtime/i);
});
