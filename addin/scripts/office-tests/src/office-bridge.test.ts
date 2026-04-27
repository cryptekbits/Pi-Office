import assert from "node:assert/strict";
import test from "node:test";

import { createOfficeToolExecutor, summarizeOfficeToolError, toAnchor, toHostAction } from "../../../apps/taskpane/src/lib/office-bridge.js";
import type { OfficeToolRequest } from "../../../packages/pi-office-pack/src/protocol.js";

test("toAnchor resolves Word, Excel, and PowerPoint anchor kinds from structured params", () => {
  const commentAnchor = toAnchor({ commentId: "comment-7" });
  assert.equal(commentAnchor.kind, "comment");
  assert.equal(commentAnchor.commentId, "comment-7");

  assert.equal(toAnchor({ sheetName: "Budget", address: "A1:C3" }).kind, "range");
  assert.equal(toAnchor({ sheetName: "Budget", address: "B2" }).kind, "cell");
  assert.equal(toAnchor({ layoutId: "layout-2", layoutName: "Two Content" }).kind, "layout");
  assert.equal(toAnchor({ id: "notes:3", slideIndex: 3 }).kind, "notesRegion");
  assert.equal(toAnchor({ id: "contentControl:42" }).kind, "contentControl");
});

test("toAnchor preserves Word durable targeting metadata for deferred tools", () => {
  const searchAnchor = toAnchor({
    searchResultId: "search:7",
    searchResultIndex: 6,
    searchQuery: "indemnity",
    objectType: "field",
    occurrenceIndex: 2,
  });
  assert.equal(searchAnchor.kind, "searchResult");
  assert.equal(searchAnchor.searchResultId, "search:7");
  assert.equal(searchAnchor.searchResultIndex, 6);
  assert.equal(searchAnchor.searchQuery, "indemnity");
  assert.equal(searchAnchor.objectType, "field");
  assert.equal(searchAnchor.occurrenceIndex, 2);

  const noteAnchor = toAnchor({ kind: "footnote", id: "footnote:2", noteTarget: "body" });
  assert.equal(noteAnchor.noteTarget, "body");

  const bookmarkAnchor = toAnchor({ bookmarkName: "PiOffice_Target_1" });
  assert.equal(bookmarkAnchor.kind, "bookmark");
  assert.equal(bookmarkAnchor.bookmarkName, "PiOffice_Target_1");

  const hyperlinkAnchor = toAnchor({ hyperlinkId: "hyperlink:3", hyperlinkAddress: "https://example.com" });
  assert.equal(hyperlinkAnchor.kind, "hyperlink");
  assert.equal(hyperlinkAnchor.hyperlinkId, "hyperlink:3");
  assert.equal(hyperlinkAnchor.hyperlinkAddress, "https://example.com");
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

test("toHostAction supports operation/text/html aliases from model-generated params", () => {
  const textAction = toHostAction({
    operation: "insertText",
    text: "Hello from alias",
  });
  assert.equal(textAction.type, "insertText");
  assert.equal(textAction.content, "Hello from alias");
  assert.equal(textAction.placement, "replace");

  const htmlAction = toHostAction({
    operation: "insertHtml",
    html: "<b>Hello</b>",
    mode: "insert_after_selection",
  });
  assert.equal(htmlAction.type, "insertHtml");
  assert.equal(htmlAction.content, "<b>Hello</b>");
  assert.equal(htmlAction.placement, "after");

  const bareHtmlAction = toHostAction({
    html: "<h1>Executive Brief</h1>",
  });
  assert.equal(bareHtmlAction.type, "insertHtml");
  assert.equal(bareHtmlAction.content, "<h1>Executive Brief</h1>");
});

test("toHostAction rejects model-generated HTML shapes that would become literal Word text", () => {
  assert.throws(
    () => toHostAction({ content: "<h1>Executive Brief</h1>" }),
    /HTML-looking content/i,
  );

  const literalText = toHostAction({
    operation: "insertText",
    content: "<h1>Show this literal tag</h1>",
  });
  assert.equal(literalText.type, "insertText");
  assert.equal(literalText.content, "<h1>Show this literal tag</h1>");

  assert.throws(
    () => toHostAction({ operation: "insertHtml", values: [{ html: "<p>Wrong shape</p>" }] }),
    /values\[\]\.html/i,
  );

  assert.throws(
    () => toHostAction({ operation: "insertHtml", html: "" }),
    /requires non-empty content/i,
  );

  assert.throws(
    () => toHostAction({ operation: "insertHtml", html: "<p>$$E=mc^2$$</p>" }),
    /LaTeX\/Markdown math.*word_equation/i,
  );
});

test("toHostAction rejects Gemini Pro JSON-string action without falling back to empty insertText", () => {
  assert.throws(
    () => toHostAction({
      action: "{\"type\": \"insertHtml\", \"content\": \"<h1>Executive Briefing: The Fast Fourier Transform (FFT)</h1>\"}",
    }),
    /action must be an object, not a JSON string/i,
  );
});

test("toHostAction supports explicit document-scope HTML replacement and append", () => {
  const replace = toHostAction({
    operation: "replaceDocumentHtml",
    html: "<h1>Executive Brief</h1><p>One write.</p>",
  });
  assert.equal(replace.type, "insertHtml");
  assert.equal(replace.target?.kind, "document");
  assert.equal(replace.placement, "replace");

  const append = toHostAction({
    action: {
      type: "appendDocumentHtml",
      content: "<h2>7. Conclusion</h2><p>Closing section.</p>",
    },
  });
  assert.equal(append.type, "insertHtml");
  assert.equal(append.target?.kind, "document");
  assert.equal(append.placement, "end");
});

test("toHostAction blocks ambiguous follow-up numbered Word sections at stale selection", () => {
  assert.throws(
    () => toHostAction({
      operation: "insertHtml",
      html: "<h2>7. Conclusion</h2><p>Closing section inserted after the full document draft.</p>",
    }),
    /top-level section.*explicit target|stale active Word selection/i,
  );
});

test("toHostAction supports direct matrix values payload", () => {
  const action = toHostAction({
    operation: "setRangeValues",
    values: [["A", "B"]],
    sheetName: "Sheet1",
    address: "A1:B1",
  });

  assert.equal(action.type, "setRangeValues");
  assert.deepEqual(action.values, [["A", "B"]]);
  assert.equal(action.target?.sheetName, "Sheet1");
  assert.equal(action.target?.address, "A1:B1");
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
    readDocumentSection: async () => ({ ok: true }),
    executeOfficeJs: async () => ({ ok: true }),
    proposeEdits: async () => ({ ok: true }),
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

test("createOfficeToolExecutor accepts plural office_tool_get params from the public schema", async () => {
  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async () => ({ ok: true }),
    applyHostAction: async () => ({ ok: true }),
    navigateOfficeAnchor: async () => ({ ok: true }),
    readDocumentSection: async () => ({ ok: true }),
    executeOfficeJs: async () => ({ ok: true }),
    proposeEdits: async () => ({ ok: true }),
  });

  const byIds = await executeOfficeTool({
    requestId: "tool-get-ids",
    toolName: "office_tool_get",
    host: "word",
    params: {
      ids: ["word_section_layout"],
    },
  });
  const byToolNames = await executeOfficeTool({
    requestId: "tool-get-toolnames",
    toolName: "office_tool_get",
    host: "word",
    params: {
      toolNames: ["word_section_layout"],
    },
  });
  const multiple = await executeOfficeTool({
    requestId: "tool-get-multiple",
    toolName: "office_tool_get",
    host: "word",
    params: {
      toolNames: ["word_section_layout", "word_table"],
    },
  });

  assert.equal(byIds.success, true);
  assert.equal((byIds.content as { detail: { toolName: string } }).detail.toolName, "word_section_layout");
  assert.equal(byToolNames.success, true);
  assert.equal((byToolNames.content as { detail: { toolName: string } }).detail.toolName, "word_section_layout");
  assert.equal(multiple.success, true);
  assert.deepEqual(
    (multiple.content as { details: Array<{ toolName: string }> }).details.map((detail) => detail.toolName),
    ["word_section_layout", "word_table"],
  );
});

test("createOfficeToolExecutor honors office_tool_search maxResults schema", async () => {
  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async () => ({ ok: true }),
    applyHostAction: async () => ({ ok: true }),
    navigateOfficeAnchor: async () => ({ ok: true }),
    readDocumentSection: async () => ({ ok: true }),
    executeOfficeJs: async () => ({ ok: true }),
    proposeEdits: async () => ({ ok: true }),
  });

  const result = await executeOfficeTool({
    requestId: "tool-search-max",
    toolName: "office_tool_search",
    host: "word",
    params: {
      query: "word",
      maxResults: 2,
    },
  });

  assert.equal(result.success, true);
  assert.equal((result.content as { results: unknown[] }).results.length, 2);
});

test("createOfficeToolExecutor dispatches Word navigation and formats Office runtime errors", async () => {
  const logEntries: unknown[] = [];
  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async () => ({ ok: true }),
    applyHostAction: async () => {
      throw { message: "No selection", code: "ItemNotFound", debugInfo: { errorLocation: "Range.getText" } };
    },
    navigateOfficeAnchor: async (host, anchor) => ({ host, anchor }),
    readDocumentSection: async () => ({ ok: true }),
    executeOfficeJs: async () => ({ ok: true }),
    proposeEdits: async () => ({ ok: true }),
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

test("createOfficeToolExecutor accepts legacy aliases for read-section and execute-js tools", async () => {
  const calls: Array<{ name: string; payload: unknown }> = [];
  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async () => ({ ok: true }),
    applyHostAction: async () => ({ ok: true }),
    navigateOfficeAnchor: async () => ({ ok: true }),
    readDocumentSection: async (_host, startIndex, endIndex, includeStyles) => {
      calls.push({ name: "readDocumentSection", payload: { startIndex, endIndex, includeStyles } });
      return { ok: true };
    },
    executeOfficeJs: async (_host, code) => {
      calls.push({ name: "executeOfficeJs", payload: { code } });
      return { ok: true };
    },
    proposeEdits: async () => ({ ok: true }),
  });

  const readResult = await executeOfficeTool({
    requestId: "read-1",
    toolName: "office_read_section",
    host: "word",
    params: {
      start: 4,
      end: 9,
      includeStyles: false,
    },
  });
  const executeResult = await executeOfficeTool({
    requestId: "js-1",
    toolName: "office_execute_js",
    host: "excel",
    params: {
      script: "return 1 + 1;",
    },
  });

  assert.equal(readResult.success, true);
  assert.equal(executeResult.success, true);
  assert.deepEqual(calls[0], {
    name: "readDocumentSection",
    payload: { startIndex: 4, endIndex: 9, includeStyles: false },
  });
  assert.deepEqual(calls[1], {
    name: "executeOfficeJs",
    payload: { code: "return 1 + 1;" },
  });
});

test("createOfficeToolExecutor normalizes payload-aware failures for read-section, execute-js, and propose-edits", async () => {
  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async () => ({ ok: true }),
    applyHostAction: async () => ({ ok: true }),
    navigateOfficeAnchor: async () => ({ ok: true }),
    readDocumentSection: async () => ({ error: "office_read_section is only supported for Word documents." }),
    executeOfficeJs: async () => ({ ok: false, error: "Code blocked: contains disallowed pattern." }),
    proposeEdits: async () => ({ error: "office_propose_edits is only supported for Word documents." }),
  });

  const readResult = await executeOfficeTool({
    requestId: "read-unsupported",
    toolName: "office_read_section",
    host: "excel",
    params: {},
  });
  const executeResult = await executeOfficeTool({
    requestId: "execute-blocked",
    toolName: "office_execute_js",
    host: "word",
    params: {
      code: "fetch('https://example.com')",
    },
  });
  const proposeResult = await executeOfficeTool({
    requestId: "propose-unsupported",
    toolName: "office_propose_edits",
    host: "powerpoint",
    params: {
      edits: [{ searchText: "A", newText: "B" }],
    },
  });

  assert.deepEqual(readResult, {
    requestId: "read-unsupported",
    success: false,
    error: "office_read_section is only supported for Word documents.",
  });
  assert.deepEqual(executeResult, {
    requestId: "execute-blocked",
    success: false,
    error: "Code blocked: contains disallowed pattern.",
  });
  assert.deepEqual(proposeResult, {
    requestId: "propose-unsupported",
    success: false,
    error: "office_propose_edits is only supported for Word documents.",
  });
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
    readDocumentSection: async () => ({ ok: true }),
    executeOfficeJs: async () => ({ ok: true }),
    proposeEdits: async () => ({ ok: true }),
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

test("createOfficeToolExecutor keeps true viewport screenshots companion-only", async () => {
  const calls: Array<{ host: string; options: unknown }> = [];
  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async (host, options) => {
      calls.push({ host, options });
      return { summary: "Base summary", formatting: { existing: true }, state: { host } };
    },
    applyHostAction: async () => ({ ok: true }),
    navigateOfficeAnchor: async () => ({ ok: true }),
    readDocumentSection: async () => ({ ok: true }),
    executeOfficeJs: async () => ({ ok: true }),
    proposeEdits: async () => ({ ok: true }),
  });

  const result = await executeOfficeTool({
    requestId: "vp-1",
    toolName: "office_capture_viewport",
    host: "word",
    params: {
      includeFormatting: true,
      includeWindowFrame: true,
    },
  });

  assert.equal(result.success, false);
  assert.equal(calls.length, 0);
  assert.match(String(result.error), /requires companion native capture/i);
  assert.match(String(result.error), /office_capture_snapshot/i);
});

test("createOfficeToolExecutor returns structured payloads for first-class Word verification tools", async () => {
  const calls: Array<{ host: string; options: unknown }> = [];
  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async (host, options) => {
      calls.push({ host, options });
      return {
        summary: "Word context captured",
        state: {
          host,
          selection: { label: "Paragraph 1" },
        },
        snippets: {
          documentStructure: {
            paragraphs: 3,
          },
        },
        formatting: {
          viewport: {
            pagesEnclosingViewport: [{ index: 1 }],
          },
          viewportCapture: {
            mode: "officejs-context",
          },
        },
        visuals: [
          {
            kind: "viewport",
            data: "ZmFrZQ==",
            mimeType: "image/png",
          },
        ],
      };
    },
    applyHostAction: async () => ({ ok: true }),
    navigateOfficeAnchor: async () => ({ ok: true }),
    readDocumentSection: async () => ({ ok: true }),
    executeOfficeJs: async () => ({ ok: true }),
    proposeEdits: async () => ({ ok: true }),
  });

  const verifyResult = await executeOfficeTool({
    requestId: "verify-1",
    toolName: "verify_doc" as OfficeToolRequest["toolName"],
    host: "word",
    params: {
      scope: "document",
    },
  } as OfficeToolRequest);
  const verifyVisualResult = await executeOfficeTool({
    requestId: "verify-visual-1",
    toolName: "verify_doc_visual" as OfficeToolRequest["toolName"],
    host: "word",
    params: {
      includeFormatting: true,
      includeWindowFrame: true,
    },
  } as OfficeToolRequest);
  const verifyVisualUnsupported = await executeOfficeTool({
    requestId: "verify-visual-unsupported",
    toolName: "verify_doc_visual" as OfficeToolRequest["toolName"],
    host: "powerpoint",
    params: {},
  } as OfficeToolRequest);

  assert.equal(verifyResult.success, true);
  assert.deepEqual(calls[0], {
    host: "word",
    options: {
      includeFormatting: true,
      maxImages: 0,
      scope: "document",
    },
  });
  const verifyPayload = verifyResult.content as {
    summary: string;
    details: { kind: string; mutating: boolean; context: { snippets?: { documentStructure?: { paragraphs?: number } } } };
  };
  assert.match(verifyPayload.summary, /Word context captured/);
  assert.equal(verifyPayload.details.kind, "word-document-verification");
  assert.equal(verifyPayload.details.mutating, false);
  assert.equal(verifyPayload.details.context.snippets?.documentStructure?.paragraphs, 3);

  assert.equal(verifyVisualResult.success, true);
  assert.deepEqual(calls[1], {
    host: "word",
    options: {
      includeFormatting: true,
      maxImages: 1,
      scope: "viewport",
    },
  });
  const verifyVisualPayload = verifyVisualResult.content as {
    visual: {
      kind: string;
      captureMode: string;
      includeWindowFrameRequested: boolean;
      includeWindowFrameCaptured: boolean;
    };
    details: { kind: string; mutating: boolean };
    visuals: unknown[];
  };
  assert.equal(verifyVisualPayload.visual.kind, "word-viewport");
  assert.equal(verifyVisualPayload.visual.captureMode, "officejs-context");
  assert.equal(verifyVisualPayload.visual.includeWindowFrameRequested, true);
  assert.equal(verifyVisualPayload.visual.includeWindowFrameCaptured, false);
  assert.equal(verifyVisualPayload.details.kind, "word-visual-verification");
  assert.equal(verifyVisualPayload.details.mutating, false);
  assert.equal(Array.isArray(verifyVisualPayload.visuals), true);

  assert.equal(verifyVisualUnsupported.success, false);
  assert.match(String(verifyVisualUnsupported.error), /only available for Word/);
});

test("verify_doc flags corrupted numbered heading order from saved document context", async () => {
  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async () => ({
      summary: "Word context captured",
      state: { host: "word" },
      snippets: {
        headings: [
          { text: "7. Conclusion", styleBuiltIn: "Heading2" },
          { text: "1. Executive Summary", styleBuiltIn: "Heading2" },
        ],
        paragraphs: [
          {
            text: "7. Conclusion This closing section accidentally landed at the top of the document before the executive summary and now reads like body prose rather than a real heading.",
            styleBuiltIn: "Heading1",
          },
        ],
      },
    }),
    applyHostAction: async () => ({ ok: true }),
    navigateOfficeAnchor: async () => ({ ok: true }),
    readDocumentSection: async () => ({ ok: true }),
    executeOfficeJs: async () => ({ ok: true }),
    proposeEdits: async () => ({ ok: true }),
  });

  const result = await executeOfficeTool({
    requestId: "verify-heading-order",
    toolName: "verify_doc" as OfficeToolRequest["toolName"],
    host: "word",
    params: { scope: "document" },
  } as OfficeToolRequest);

  assert.equal(result.success, true);
  const payload = result.content as { summary: string; warnings?: string[]; details: { warnings?: string[] } };
  assert.match(payload.summary, /7\. Conclusion/);
  assert.match(payload.summary, /before a section 1 heading|appears after/i);
  assert.ok(payload.warnings?.some((warning) => /Malformed heading order/.test(warning)));
  assert.ok(payload.details.warnings?.some((warning) => /Heading1 but looks like body text/.test(warning)));
});

test("verify_doc flags literal LaTeX and empty heading paragraphs in Word context", async () => {
  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async () => ({
      summary: "Word context captured",
      state: { host: "word" },
      snippets: {
        headings: [
          { text: "", styleBuiltIn: "Heading2" },
          { text: "1. Method", styleBuiltIn: "Heading2" },
        ],
        paragraphs: [
          { text: "The result is $$E=mc^2$$.", styleBuiltIn: "Normal" },
        ],
      },
    }),
    applyHostAction: async () => ({ ok: true }),
    navigateOfficeAnchor: async () => ({ ok: true }),
    readDocumentSection: async () => ({ ok: true }),
    executeOfficeJs: async () => ({ ok: true }),
    proposeEdits: async () => ({ ok: true }),
  });

  const result = await executeOfficeTool({
    requestId: "verify-literal-latex",
    toolName: "verify_doc" as OfficeToolRequest["toolName"],
    host: "word",
    params: { scope: "document" },
  } as OfficeToolRequest);

  assert.equal(result.success, true);
  const payload = result.content as { warnings?: string[]; details: { warnings?: string[] } };
  assert.ok(payload.warnings?.some((warning) => /literal LaTeX/i.test(warning)));
  assert.ok(payload.details.warnings?.some((warning) => /heading but has no visible text/i.test(warning)));
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
