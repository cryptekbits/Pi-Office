#!/usr/bin/env node

const hostOrder = ["word", "excel", "powerpoint"];

const smokeScenarios = {
  word: {
    prerequisites: [
      "Open a document with headings, tracked changes, comments, footnotes, endnotes, fields, and content controls.",
      "Open the Pi-Office taskpane and connect it to the active document.",
    ],
    scenarios: [
      {
        id: "word-review-anchors",
        title: "Review-anchor navigation",
        covers: [
          "heading anchors",
          "paragraph anchors",
          "comment anchors",
          "revision anchors",
          "footnote/endnote anchors",
          "field/content-control anchors",
        ],
        setup: "Seed at least one instance of each review artifact in the document.",
        steps: [
          "Run `office_get_context` with `includeFormatting=true` and confirm the returned anchors include the review artifacts above.",
          "Run `office_navigate` against one anchor of each kind and confirm Word moves selection to the expected location.",
        ],
        expected: [
          "Each anchor resolves without a backend fallback.",
          "Word selection lands on the matching artifact or its reference location.",
        ],
      },
      {
        id: "word-structured-edits",
        title: "Structured Word edits",
        covers: [
          "native text/html insertion",
          "comment add/reply/resolve/delete",
          "field insertion",
          "content-control insertion",
          "revision accept/reject",
        ],
        setup: "Place the caret in a writable paragraph and leave one unresolved comment plus one pending revision available.",
        steps: [
          "Run `office_apply_edit` with `insertText`, `insertHtml`, `insertComment`, `replyToComment`, and `resolveComment` actions.",
          "Run `office_apply_edit` with `insertField`, `insertContentControl`, `acceptRevision`, and `rejectRevision` on concrete anchors.",
        ],
        expected: [
          "Word creates or updates native document objects instead of flattening edits to plain text.",
          "Action results return touched objects matching the edited comment, revision, field, or content control.",
        ],
      },
      {
        id: "word-visual-verification",
        title: "Word document + visual verification",
        covers: [
          "structured Word verification context",
          "viewport/visual verification payload",
          "Word-only visual verification constraints",
        ],
        setup: "Open a document where the current visible viewport includes at least one heading, one list block, and one paragraph with tracked changes.",
        steps: [
          "Run `verify_doc` with `scope=document` and confirm it returns summary text plus a structured `details` payload containing host/context snippets.",
          "Run `verify_doc_visual` with `includeFormatting=true` and confirm it returns `visual`, `details`, and `visuals` payloads sourced from the viewport path.",
          "From a non-Word host session, run `verify_doc_visual` and confirm it reports the Word-only availability constraint.",
        ],
        expected: [
          "`verify_doc` remains non-mutating and returns structured verification context for review workflows.",
          "`verify_doc_visual` returns viewport-oriented visual metadata with explicit window-frame limitations.",
          "Word-only constraints are explicit and do not silently fall through on non-Word hosts.",
        ],
      },
      {
        id: "word-taskpane-stability",
        title: "Taskpane focus, scroll, prompt, and connector stability",
        covers: [
          "taskpane keyboard focus",
          "chat vertical scroll behavior",
          "prompt + streaming stability",
          "connector action execution",
          "session reconnect behavior",
        ],
        setup: "Keep the taskpane open with an active model and at least one connected connector.",
        steps: [
          "Type a long prompt in the composer and confirm keyboard input remains in the taskpane (not Word document body).",
          "Send the prompt, stream a long response, manually scroll up mid-stream, then verify auto-scroll does not fight manual scroll.",
          "Trigger one connector-backed request and verify success/error feedback appears inline in chat/settings.",
          "Close and reopen Word (or force taskpane reconnect), then verify session reconnects and a new prompt can complete.",
        ],
        expected: [
          "Composer focus remains stable while typing and after sending prompts.",
          "Chat scroll stays user-controlled and resumes normally at the bottom.",
          "Connector calls surface clear success/failure outcomes without silent drops.",
          "Reconnected session remains usable for subsequent prompts and tool calls.",
        ],
      },
    ],
  },
  excel: {
    prerequisites: [
      "Open a workbook with multiple worksheets, one formatted table, one chart, and one PivotTable.",
      "Open the Pi-Office taskpane and connect it to the active workbook.",
    ],
    scenarios: [
      {
        id: "excel-context-citations",
        title: "Workbook context and citations",
        covers: [
          "workbook anchors",
          "sheet snapshots",
          "cell/range citations",
          "named-item navigation",
        ],
        setup: "Select a populated range containing formulas and visible values.",
        steps: [
          "Run `office_get_context` with `scope=workbook` and `includeFormatting=true`.",
          "Inspect `selectionRangeCitation`, `selectionCellCitations`, `activeWorksheetSnapshot`, and `workbookSheetSnapshots`.",
          "Run `office_navigate` to a cited `cell`, `range`, `sheet`, and `namedItem` anchor.",
        ],
        expected: [
          "Context includes bounded workbook summaries plus cell-level citations for the current selection.",
          "Navigation activates the requested worksheet object and preserves the referenced address.",
        ],
      },
      {
        id: "excel-formatting-and-tables",
        title: "Range formatting, borders, and table controls",
        covers: [
          "font/fill/alignment formatting",
          "range borders",
          "table style updates",
          "table filters",
          "gridlines/headings/print area",
          "data validation",
        ],
        setup: "Select a normal range and identify one existing Excel table.",
        steps: [
          "Run `office_apply_edit` with `formatRange` using font, fill, alignment, and structured border payloads.",
          "Run `office_apply_edit` with `formatTable` or `updateTableStyle` to toggle headers, totals, banding, highlight columns, and style.",
          "Run `applyTableFilter`, `clearTableFilters`, `setWorksheetGridlines`, `setWorksheetHeadings`, `setPrintArea`, and `setDataValidation`.",
        ],
        expected: [
          "Range borders and table style flags update natively in Excel.",
          "Filter and worksheet view controls operate without leaving the current workbook model.",
        ],
      },
      {
        id: "excel-charts-and-pivots",
        title: "Existing chart and PivotTable editing",
        covers: [
          "chart create/update",
          "axis and label edits",
          "PivotTable create/update",
          "pivot filters",
          "pivot sorting",
        ],
        setup: "Prepare one existing chart and one existing PivotTable with editable source data.",
        steps: [
          "Run `createChart` or `updateChart` with title, legend, axis, gridline, display-unit, and data-label options.",
          "Run `createPivotTable` or `updatePivotTable` with hierarchy, layout, filter, refresh, and sort instructions.",
          "Run `sortPivotField`, `sortPivotByLabels`, and `sortPivotByValues` on the same PivotTable.",
        ],
        expected: [
          "Excel updates the existing chart instead of replacing it with an image or screenshot.",
          "PivotTable schema, filters, refresh, and sorts all execute through native workbook objects.",
        ],
      },
    ],
  },
  powerpoint: {
    prerequisites: [
      "Open a deck with multiple slides, at least one custom layout or master, at least one editable table, and at least one chart.",
      "Open the Pi-Office taskpane and connect it to the active presentation.",
    ],
    scenarios: [
      {
        id: "powerpoint-anchor-navigation",
        title: "Slide, layout, shape, and notes anchors",
        covers: [
          "slide anchors",
          "shape anchors",
          "layout anchors",
          "slide master anchors",
          "notes-region anchors",
        ],
        setup: "Select a slide that has notes and at least one named shape.",
        steps: [
          "Run `office_get_context` with `includeFormatting=true` and confirm anchors include slides, selected shapes, layouts, masters, and `notesRegion`.",
          "Run `office_navigate` for one anchor of each kind.",
        ],
        expected: [
          "Slide and shape navigation selects the exact target.",
          "Layout, slide master, and notes-region navigation return partial completion with explicit fallback reporting to slide selection.",
        ],
      },
      {
        id: "powerpoint-slide-and-shape-authoring",
        title: "Native slide and shape operations",
        covers: [
          "slide create/reorder/duplicate/delete/combine",
          "shape text updates",
          "shape image replacement",
          "table edits",
          "flow/diagram generation",
        ],
        setup: "Use a disposable presentation copy so slide lifecycle actions can be exercised safely.",
        steps: [
          "Run `addAgendaSlide`, `addTransitionSlide`, `duplicateSlide`, `reorderSlides`, `combineSlides`, and `deleteSlide` with `confirmDestructive=true` where required.",
          "Run `setShapeText`, `updateShapeProperties`, `replaceShapeImage`, and table update actions such as `setTableValues` or `updateTableCell`.",
          "Run the process-flow authoring action that builds editable shapes and connectors.",
        ],
        expected: [
          "Slides are inserted, moved, duplicated, or deleted as real deck objects.",
          "Shapes, images, and tables remain editable inside PowerPoint after the action completes.",
        ],
      },
      {
        id: "powerpoint-notes-and-charts",
        title: "Serialized notes and chart workflows",
        covers: [
          "presentation-package inspection",
          "slide notes read/write",
          "chart inspect/create/update",
          "embedded workbook synchronization",
        ],
        setup: "Identify one slide with notes and one slide that can accept a chart insertion.",
        steps: [
          "Run `getSlideNotes` and `setSlideNotes` on the notes slide, then reopen the notes anchor in context.",
          "Run `inspectPresentationPackage` or `getPresentationTheme` to confirm package metadata is available.",
          "Run `getSlideCharts`, `addSlideChart`, and `updateSlideChart` on chart-bearing slides.",
        ],
        expected: [
          "Notes serialization returns the updated note text and a replacement slide selection.",
          "Chart inspection and updates return chart metadata, preserve editability, and keep embedded workbook data synchronized.",
        ],
      },
    ],
  },
};

function parseArgs(argv) {
  let host;
  let json = false;
  let list = false;

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === "--json") {
      json = true;
      continue;
    }
    if (arg === "--list") {
      list = true;
      continue;
    }
    if (arg === "--host") {
      host = argv[index + 1];
      index += 1;
      continue;
    }
    if (arg?.startsWith("--host=")) {
      host = arg.slice("--host=".length);
      continue;
    }
    throw new Error(`Unknown argument: ${arg}`);
  }

  if (host && !hostOrder.includes(host)) {
    throw new Error(`Unsupported host "${host}". Expected one of: ${hostOrder.join(", ")}.`);
  }

  return { host, json, list };
}

function validateSmokeScenarios(matrix) {
  for (const host of hostOrder) {
    const hostPlan = matrix[host];
    if (!hostPlan) {
      throw new Error(`Missing smoke scenarios for host: ${host}`);
    }
    if (!Array.isArray(hostPlan.prerequisites) || hostPlan.prerequisites.length === 0) {
      throw new Error(`Host ${host} must define prerequisites.`);
    }
    if (!Array.isArray(hostPlan.scenarios) || hostPlan.scenarios.length === 0) {
      throw new Error(`Host ${host} must define at least one scenario.`);
    }
    for (const scenario of hostPlan.scenarios) {
      const requiredKeys = ["id", "title", "covers", "setup", "steps", "expected"];
      for (const key of requiredKeys) {
        const value = scenario[key];
        if ((Array.isArray(value) && value.length === 0) || (!Array.isArray(value) && typeof value !== "string")) {
          throw new Error(`Scenario ${host}:${scenario.id} is missing ${key}.`);
        }
      }
    }
  }
}

function renderHost(host, plan, options) {
  if (options.list) {
    return [`${host}:`, ...plan.scenarios.map((scenario) => `- ${scenario.id}: ${scenario.title}`)].join("\n");
  }

  const lines = [`# ${host.toUpperCase()} Smoke Plan`, "", "Prerequisites:"];
  for (const entry of plan.prerequisites) {
    lines.push(`- ${entry}`);
  }

  for (const scenario of plan.scenarios) {
    lines.push("");
    lines.push(`## ${scenario.id} - ${scenario.title}`);
    lines.push(`Covers: ${scenario.covers.join("; ")}`);
    lines.push(`Setup: ${scenario.setup}`);
    lines.push("Steps:");
    for (const step of scenario.steps) {
      lines.push(`- ${step}`);
    }
    lines.push("Expected:");
    for (const expectation of scenario.expected) {
      lines.push(`- ${expectation}`);
    }
  }

  return lines.join("\n");
}

function main() {
  const options = parseArgs(process.argv.slice(2));
  validateSmokeScenarios(smokeScenarios);

  const hosts = options.host ? [options.host] : hostOrder;
  const selected = Object.fromEntries(hosts.map((host) => [host, smokeScenarios[host]]));

  if (options.json) {
    process.stdout.write(`${JSON.stringify(selected, null, 2)}\n`);
    return;
  }

  process.stdout.write(`${hosts.map((host) => renderHost(host, smokeScenarios[host], options)).join("\n\n")}\n`);
}

main();
