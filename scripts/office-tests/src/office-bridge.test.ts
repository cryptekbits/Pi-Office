import assert from "node:assert/strict";
import test from "node:test";

import { createOfficeToolExecutor, summarizeOfficeToolError, toAnchor, toHostAction } from "../../../apps/taskpane/src/lib/office-bridge.js";
import type { OfficeToolRequest } from "../../../packages/pi-office-pack/src/protocol.js";

test("toAnchor resolves Word, Excel, and PowerPoint anchor kinds from structured params", () => {
  assert.deepEqual(toAnchor({ commentId: "comment-7" }), {
    kind: "comment",
    commentId: "comment-7",
    text: undefined,
    label: undefined,
    id: undefined,
    sheetName: undefined,
    address: undefined,
    paragraphId: undefined,
    revisionId: undefined,
    slideId: undefined,
    slideIndex: undefined,
    shapeId: undefined,
    tableName: undefined,
    chartName: undefined,
    pivotTableName: undefined,
    namedItemName: undefined,
  });

  assert.equal(toAnchor({ sheetName: "Budget", address: "A1:C3" }).kind, "range");
  assert.equal(toAnchor({ sheetName: "Budget", address: "B2" }).kind, "cell");
  assert.equal(toAnchor({ layoutId: "layout-2", layoutName: "Two Content" }).kind, "layout");
  assert.equal(toAnchor({ id: "notes:3", slideIndex: 3 }).kind, "notesRegion");
  assert.equal(toAnchor({ id: "contentControl:42" }).kind, "contentControl");
});

test("toHostAction creates Excel matrix actions from legacy params", () => {
  const action = toHostAction({
    mode: "setRangeValues",
    format: "matrix",
    content: "[[1,2],[3,4]]",
    sheetName: "Budget",
    address: "A1:B2",
  });

  assert.equal(action.type, "setRangeValues");
  assert.equal(action.target?.kind, "range");
  assert.equal(action.target?.sheetName, "Budget");
  assert.equal(action.target?.address, "A1:B2");
  assert.deepEqual(action.values, [
    [1, 2],
    [3, 4],
  ]);
});

test("toHostAction preserves structured PowerPoint actions and normalizes the target anchor", () => {
  const action = toHostAction({
    action: {
      type: "addSlideChart",
      target: { slideId: "slide-1" },
      chartType: "column",
      options: { confirmDestructive: true },
    },
  });

  assert.equal(action.type, "addSlideChart");
  assert.equal(action.target?.kind, "slide");
  assert.equal(action.target?.slideId, "slide-1");
  assert.equal(action.chartType, "column");
  assert.deepEqual(action.options, { confirmDestructive: true });
});

test("createOfficeToolExecutor dispatches apply-edit and capture requests through the adapter layer", async () => {
  const calls: Array<{ name: string; payload: unknown }> = [];
  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async (host, options) => {
      calls.push({ name: "collectOfficeContext", payload: { host, options } });
      return { host, options };
    },
    applyHostAction: async (host, action) => {
      calls.push({ name: "applyHostAction", payload: { host, action } });
      return { host, action };
    },
    navigateOfficeAnchor: async (host, anchor) => {
      calls.push({ name: "navigateOfficeAnchor", payload: { host, anchor } });
      return { host, anchor };
    },
  });

  const applyRequest: OfficeToolRequest = {
    requestId: "apply-1",
    toolName: "office_apply_edit",
    host: "excel",
    params: {
      mode: "setRangeValues",
      format: "matrix",
      content: "[[11]]",
      sheetName: "Budget",
      address: "B2",
    },
  };
  const captureRequest: OfficeToolRequest = {
    requestId: "capture-1",
    toolName: "office_capture_snapshot",
    host: "powerpoint",
    params: {
      scope: "slide",
      includeFormatting: false,
      maxImages: 99,
    },
  };

  const applyResult = await executeOfficeTool(applyRequest);
  const captureResult = await executeOfficeTool(captureRequest);

  assert.equal(applyResult.success, true);
  assert.equal(captureResult.success, true);
  assert.equal(calls[0]?.name, "applyHostAction");
  assert.equal((calls[0]?.payload as { action: { type: string } }).action.type, "setRangeValues");
  assert.equal(calls[1]?.name, "collectOfficeContext");
  assert.deepEqual((calls[1]?.payload as { options: { includeFormatting: boolean; maxImages: number; scope?: string } }).options, {
    includeFormatting: false,
    maxImages: 4,
    scope: "slide",
  });
});

test("createOfficeToolExecutor dispatches Word navigation and formats Office runtime errors", async () => {
  const logEntries: unknown[] = [];
  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async () => ({ ok: true }),
    applyHostAction: async () => {
      throw { message: "No selection", code: "ItemNotFound", debugInfo: { errorLocation: "Range.getText" } };
    },
    navigateOfficeAnchor: async (host, anchor) => ({ host, anchor }),
    logger: {
      error: (...args: unknown[]) => {
        logEntries.push(args);
      },
    },
  });

  const navigateResult = await executeOfficeTool({
    requestId: "nav-1",
    toolName: "office_navigate",
    host: "word",
    params: {
      paragraphId: "paragraph-9",
    },
  });
  const failingResult = await executeOfficeTool({
    requestId: "apply-2",
    toolName: "office_apply_edit",
    host: "word",
    params: {
      mode: "replaceSelection",
      content: "Updated text",
    },
  });

  assert.equal(navigateResult.success, true);
  assert.equal((navigateResult.content as { anchor: { kind: string } }).anchor.kind, "paragraph");
  assert.equal(failingResult.success, false);
  assert.match(String(failingResult.error), /office_apply_edit failed: No selection/);
  assert.match(String(failingResult.error), /code=ItemNotFound/);
  assert.equal(logEntries.length, 1);
});

test("createOfficeToolExecutor forwards get-context scope to the host adapter", async () => {
  const calls: Array<{ name: string; payload: unknown }> = [];
  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async (host, options) => {
      calls.push({ name: "collectOfficeContext", payload: { host, options } });
      return { host, options };
    },
    applyHostAction: async () => ({ ok: true }),
    navigateOfficeAnchor: async () => ({ ok: true }),
  });

  const result = await executeOfficeTool({
    requestId: "ctx-1",
    toolName: "office_get_context",
    host: "excel",
    params: {
      scope: "workbook",
    },
  });

  assert.equal(result.success, true);
  assert.equal(calls[0]?.name, "collectOfficeContext");
  assert.deepEqual((calls[0]?.payload as { options: { includeFormatting: boolean; maxImages: number; scope?: string } }).options, {
    includeFormatting: true,
    maxImages: 0,
    scope: "workbook",
  });
});

test("summarizeOfficeToolError includes code, location, statement, and traces when present", () => {
  const summary = summarizeOfficeToolError(
    {
      message: "Host rejected the request",
      code: "GeneralException",
      debugInfo: {
        errorLocation: "Worksheet.getRange",
        statement: "worksheet.getRange(\"A1\")",
      },
      traceMessages: ["sync", "load"],
    },
    "office_apply_edit",
  );

  assert.match(summary, /Host rejected the request/);
  assert.match(summary, /code=GeneralException/);
  assert.match(summary, /location=Worksheet.getRange/);
  assert.match(summary, /statement=worksheet.getRange/);
  assert.match(summary, /trace=sync \| load/);
});
