import type { OfficeAnchor, OfficeHost, OfficeHostAction, OfficeToolRequest, OfficeToolResult } from "@pi-office/pi-office-pack/protocol";

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === "object" && !Array.isArray(value));
}

function getErrorRecord(error: unknown): Record<string, unknown> | undefined {
  return error && typeof error === "object" ? (error as Record<string, unknown>) : undefined;
}

function trimString(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function isSingleCellAddress(value: string | undefined): boolean {
  const normalized = trimString(value)?.replace(/\$/g, "");
  if (!normalized) {
    return false;
  }

  const localAddress = normalized.includes("!") ? normalized.split("!").pop() : normalized;
  if (!localAddress) {
    return false;
  }

  const [start, end] = localAddress.split(":");
  return !end || start === end;
}

function normalizeExcelSheetName(value: unknown): string | undefined {
  const raw = trimString(value);
  if (!raw) return undefined;
  return raw.replace(/^'+|'+$/g, "").toLowerCase();
}

function normalizeExcelAddress(value: unknown): string | undefined {
  const raw = trimString(value);
  if (!raw) return undefined;
  const address = raw.includes("!") ? raw.split("!").pop() : raw;
  return address?.replace(/\$/g, "").replace(/\s+/g, "").toUpperCase();
}

function parseExcelSelectionLabel(value: unknown): { sheetName?: string | undefined; address?: string | undefined } {
  const label = trimString(value);
  if (!label) return {};
  const bangIndex = label.lastIndexOf("!");
  if (bangIndex === -1) {
    return { address: normalizeExcelAddress(label) };
  }
  return {
    sheetName: normalizeExcelSheetName(label.slice(0, bangIndex)),
    address: normalizeExcelAddress(label.slice(bangIndex + 1)),
  };
}

function excelSelectionMatchesRequest(
  selectionLabel: unknown,
  requestedSheetName: string | undefined,
  requestedAddress: string | undefined,
): boolean | undefined {
  const requestedSheet = normalizeExcelSheetName(requestedSheetName);
  const requestedRange = normalizeExcelAddress(requestedAddress);
  if (!requestedSheet && !requestedRange) return true;

  const selection = parseExcelSelectionLabel(selectionLabel);
  if (!selection.sheetName && !selection.address) return undefined;

  if (requestedSheet && selection.sheetName && requestedSheet !== selection.sheetName) return false;
  if (requestedRange && selection.address && requestedRange !== selection.address) return false;
  if (requestedSheet && !selection.sheetName) return undefined;
  if (requestedRange && !selection.address) return undefined;
  return true;
}

function formatExcelRangeHint(sheetName: string | undefined, address: string | undefined): string {
  if (sheetName && address) return `${sheetName}!${address}`;
  return sheetName ?? address ?? "the requested range";
}

function stringifyDetail(value: unknown): string | undefined {
  if (value == null) return undefined;
  if (typeof value === "string" && value.trim()) return value.trim();
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  if (Array.isArray(value)) {
    const parts = value.map((entry) => stringifyDetail(entry)).filter((entry): entry is string => Boolean(entry));
    return parts.length ? parts.join(" | ") : undefined;
  }

  try {
    return JSON.stringify(value);
  } catch {
    return undefined;
  }
}

function extractPayloadFailure(payload: unknown): string | undefined {
  if (!isRecord(payload)) {
    return undefined;
  }

  const hasErrorField = Object.prototype.hasOwnProperty.call(payload, "error");
  const payloadError = stringifyDetail(payload.error);

  if (payload.ok === false) {
    return payloadError ?? "Office tool returned ok=false.";
  }

  if (hasErrorField && payloadError) {
    return payloadError;
  }

  return undefined;
}

function toPayloadAwareResult(requestId: string, payload: unknown): OfficeToolResult {
  const payloadFailure = extractPayloadFailure(payload);
  if (payloadFailure) {
    return {
      requestId,
      success: false,
      error: payloadFailure,
    };
  }

  return {
    requestId,
    success: true,
    content: payload,
  };
}

function appendViewportCaptureSummary(payload: unknown, includeWindowFrameRequested: boolean): unknown {
  if (!isRecord(payload)) {
    return payload;
  }

  const formatting = isRecord(payload.formatting) ? payload.formatting : {};
  const viewportCapture = {
    mode: "officejs-context",
    includeWindowFrameRequested,
    includeWindowFrameCaptured: false,
    note: includeWindowFrameRequested
      ? "Full window-frame capture is unavailable in browser-only runtime."
      : "Viewport metadata captured from Office.js context.",
  };

  const summaryParts = [
    trimString(payload.summary),
    "Visible Word viewport metadata captured.",
    includeWindowFrameRequested ? "Full window-frame capture is unavailable in browser-only runtime." : undefined,
  ].filter((entry): entry is string => Boolean(entry));

  return {
    ...payload,
    summary: summaryParts.join("\n"),
    formatting: {
      ...formatting,
      viewportCapture,
    },
  };
}

function toWordDocumentVerificationPayload(payload: unknown, scope: string | undefined): unknown {
  if (!isRecord(payload)) {
    return payload;
  }

  return {
    summary: trimString(payload.summary) ?? "Word verification context captured.",
    details: {
      kind: "word-document-verification",
      mutating: false,
      host: "word",
      scope: scope ?? "document",
      context: {
        state: payload.state,
        anchors: payload.anchors,
        snippets: payload.snippets,
        formatting: payload.formatting,
      },
    },
  };
}

function toWordVisualVerificationPayload(payload: unknown, includeWindowFrameRequested: boolean): unknown {
  if (!isRecord(payload)) {
    return payload;
  }

  const visuals = Array.isArray(payload.visuals) ? payload.visuals : [];
  const formatting = isRecord(payload.formatting) ? payload.formatting : {};
  const viewportCapture = isRecord(formatting.viewportCapture) ? formatting.viewportCapture : {};
  const state = isRecord(payload.state) ? payload.state : {};
  const captureMode = trimString(viewportCapture.mode) ?? "officejs-context";
  const includeWindowFrameCaptured = viewportCapture.includeWindowFrameCaptured === true;
  const note =
    trimString(viewportCapture.note) ??
    (includeWindowFrameRequested
      ? "Full window-frame capture is unavailable in browser-only runtime."
      : "Viewport metadata captured from Office.js context.");

  return {
    summary: trimString(payload.summary) ?? "Visible Word viewport metadata captured.",
    visual: {
      kind: "word-viewport",
      captureMode,
      imageCount: visuals.length,
      includeWindowFrameRequested,
      includeWindowFrameCaptured,
      note,
    },
    details: {
      kind: "word-visual-verification",
      mutating: false,
      host: "word",
      viewport: formatting.viewport,
      viewportCapture: {
        mode: captureMode,
        includeWindowFrameRequested,
        includeWindowFrameCaptured,
        note,
      },
      selection: state.selection,
    },
    visuals,
  };
}

function toPowerPointSlidesVerificationPayload(payload: unknown, scope: string | undefined): unknown {
  if (!isRecord(payload)) {
    return payload;
  }

  return {
    summary: trimString(payload.summary) ?? "PowerPoint slide verification context captured.",
    details: {
      kind: "powerpoint-slide-verification",
      mutating: false,
      host: "powerpoint",
      scope: scope ?? "presentation",
      structure: {
        slideCount: payload.slideCount,
        selectedSlideCount: payload.selectedSlideCount,
        selectedSlideIds: payload.selectedSlideIds,
        slides: payload.slides,
        slidePreviews: payload.slidePreviews,
        slideMasters: payload.slideMasters,
      },
    },
  };
}

function toPowerPointVisualVerificationPayload(payload: unknown, maxImages: number): unknown {
  if (!isRecord(payload)) {
    return payload;
  }

  const visuals = Array.isArray(payload.visuals) ? payload.visuals : [];
  const snippets = isRecord(payload.snippets) ? payload.snippets : {};
  const formatting = isRecord(payload.formatting) ? payload.formatting : {};

  return {
    summary: trimString(payload.summary) ?? "PowerPoint visual verification context captured.",
    visual: {
      kind: "powerpoint-slide-snapshot",
      captureMode: "officejs-slide-snapshot",
      imageCount: visuals.length,
      maxImagesRequested: maxImages,
      note: "Visual verification uses supported Office.js slide/shape snapshot paths.",
    },
    details: {
      kind: "powerpoint-slide-visual-verification",
      mutating: false,
      host: "powerpoint",
      selectedSlides: snippets.selectedSlides,
      selectedShapeDescriptors: snippets.selectedShapeDescriptors,
      formatting: {
        selectedSlides: formatting.selectedSlides,
        selectedShapes: formatting.selectedShapes,
      },
    },
    visuals,
  };
}

type ExcelObjectKind = "table" | "chart" | "pivotTable" | "namedItem" | "worksheet" | "cell";

interface ExcelInventoryEntry {
  kind: ExcelObjectKind;
  label: string;
  name?: string | undefined;
  sheetName?: string | undefined;
  address?: string | undefined;
  id?: string | undefined;
  type?: string | undefined;
  text?: string | undefined;
  formula?: string | undefined;
  anchor: OfficeAnchor;
}

function normalizeExcelObjectKind(value: unknown): ExcelObjectKind | undefined {
  const normalized = trimString(value)?.replace(/[\s_-]/g, "").toLowerCase();
  switch (normalized) {
    case "table":
    case "tables":
      return "table";
    case "chart":
    case "charts":
      return "chart";
    case "pivottable":
    case "pivot":
    case "pivots":
    case "pivottables":
      return "pivotTable";
    case "nameditem":
    case "nameditems":
    case "name":
    case "names":
      return "namedItem";
    case "worksheet":
    case "worksheets":
    case "sheet":
    case "sheets":
      return "worksheet";
    case "cell":
    case "cells":
    case "range":
    case "ranges":
      return "cell";
    default:
      return undefined;
  }
}

function toExcelObjectKindSet(value: unknown): Set<ExcelObjectKind> | undefined {
  const entries = Array.isArray(value) ? value : value == null ? [] : [value];
  const kinds = new Set<ExcelObjectKind>();
  for (const entry of entries) {
    const kind = normalizeExcelObjectKind(entry);
    if (kind) {
      kinds.add(kind);
    }
  }
  return kinds.size ? kinds : undefined;
}

function toExcelObjectActionType(value: string | undefined): string | undefined {
  if (!value) {
    return undefined;
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  switch (compact) {
    case "formatrange":
      return "formatRange";
    case "createtable":
      return "createTable";
    case "formattable":
    case "updatetablestyle":
    case "configuretable":
    case "settablestyle":
      return "formatTable";
    case "applytablefilter":
    case "filtertable":
      return "applyTableFilter";
    case "cleartablefilter":
      return "clearTableFilter";
    case "cleartablefilters":
      return "clearTableFilters";
    case "reapplytablefilters":
      return "reapplyTableFilters";
    case "createchart":
      return "createChart";
    case "updatechart":
    case "formatchart":
    case "setchartaxes":
    case "setchartdatalabels":
      return "updateChart";
    case "createpivottable":
      return "createPivotTable";
    case "updatepivottable":
    case "configurepivottable":
    case "applypivotfilter":
      return "updatePivotTable";
    case "sortpivotfield":
      return "sortPivotField";
    case "sortpivotbylabels":
      return "sortPivotByLabels";
    case "sortpivotbyvalues":
      return "sortPivotByValues";
    case "refreshpivottable":
      return "refreshPivotTable";
    case "setworksheetgridlines":
      return "setWorksheetGridlines";
    case "setworksheetheadings":
      return "setWorksheetHeadings";
    case "setprintarea":
      return "setPrintArea";
    case "setdatavalidation":
      return "setDataValidation";
    case "cleardatavalidation":
      return "clearDataValidation";
    case "addconditionalformat":
      return "addConditionalFormat";
    case "clearconditionalformats":
      return "clearConditionalFormats";
    case "insertinlinepicture":
      return "insertInlinePicture";
    case "removeduplicates":
      return "removeDuplicates";
    default:
      return undefined;
  }
}

function toExcelObjectTarget(params: Record<string, unknown>): OfficeAnchor | undefined {
  const sheetName = trimString(params.sheetName);
  const address = trimString(params.address);
  const tableName = trimString(params.tableName);
  const chartName = trimString(params.chartName);
  const pivotTableName = trimString(params.pivotTableName);
  const namedItemName = trimString(params.namedItemName);
  const anchor = isRecord(params.anchor) && Object.keys(params.anchor).length > 0 ? params.anchor : undefined;

  if (!sheetName && !address && !tableName && !chartName && !pivotTableName && !namedItemName && !anchor) {
    return undefined;
  }

  return toAnchor({
    ...(anchor ? { anchor } : {}),
    ...(sheetName ? { sheetName } : {}),
    ...(address ? { address } : {}),
    ...(tableName ? { tableName } : {}),
    ...(chartName ? { chartName } : {}),
    ...(pivotTableName ? { pivotTableName } : {}),
    ...(namedItemName ? { namedItemName } : {}),
    ...(trimString(params.id) ? { id: trimString(params.id) } : {}),
  });
}

function pushExcelInventoryEntry(entries: ExcelInventoryEntry[], seen: Set<string>, entry: ExcelInventoryEntry | undefined): void {
  if (!entry) {
    return;
  }
  const key = `${entry.kind}:${entry.sheetName ?? ""}:${entry.name ?? entry.label}:${entry.address ?? ""}:${entry.id ?? ""}`;
  if (seen.has(key)) {
    return;
  }
  seen.add(key);
  entries.push(entry);
}

function toExcelInventoryEntries(payload: unknown): ExcelInventoryEntry[] {
  if (!isRecord(payload)) {
    return [];
  }

  const snippets = isRecord(payload.snippets) ? payload.snippets : {};
  const workbookObjects = isRecord(snippets.workbookObjects) ? snippets.workbookObjects : {};
  const entries: ExcelInventoryEntry[] = [];
  const seen = new Set<string>();

  for (const table of Array.isArray(workbookObjects.tables) ? workbookObjects.tables : []) {
    if (!isRecord(table)) continue;
    const name = trimString(table.name) ?? trimString(table.tableName);
    if (!name) continue;
    const sheetName = trimString(table.sheetName);
    pushExcelInventoryEntry(entries, seen, {
      kind: "table",
      label: sheetName ? `${sheetName}!${name}` : name,
      name,
      sheetName,
      id: trimString(table.id) ?? trimString(table.tableId),
      anchor: toAnchor({
        kind: "table",
        tableName: name,
        ...(sheetName ? { sheetName } : {}),
      }),
    });
  }

  for (const chart of Array.isArray(workbookObjects.charts) ? workbookObjects.charts : []) {
    if (!isRecord(chart)) continue;
    const name = trimString(chart.name) ?? trimString(chart.chartName);
    if (!name) continue;
    const sheetName = trimString(chart.sheetName);
    pushExcelInventoryEntry(entries, seen, {
      kind: "chart",
      label: sheetName ? `${sheetName}!${name}` : name,
      name,
      sheetName,
      id: trimString(chart.id) ?? trimString(chart.chartId),
      anchor: toAnchor({
        kind: "chart",
        chartName: name,
        ...(sheetName ? { sheetName } : {}),
      }),
    });
  }

  for (const pivotTable of Array.isArray(workbookObjects.pivotTables) ? workbookObjects.pivotTables : []) {
    if (!isRecord(pivotTable)) continue;
    const name = trimString(pivotTable.name) ?? trimString(pivotTable.pivotTableName);
    if (!name) continue;
    const sheetName = trimString(pivotTable.sheetName);
    pushExcelInventoryEntry(entries, seen, {
      kind: "pivotTable",
      label: sheetName ? `${sheetName}!${name}` : name,
      name,
      sheetName,
      id: trimString(pivotTable.id) ?? trimString(pivotTable.pivotTableId),
      anchor: toAnchor({
        kind: "pivotTable",
        pivotTableName: name,
        ...(sheetName ? { sheetName } : {}),
      }),
    });
  }

  for (const namedItem of Array.isArray(snippets.namedItems) ? snippets.namedItems : []) {
    if (!isRecord(namedItem)) continue;
    const name = trimString(namedItem.name) ?? trimString(namedItem.namedItemName);
    if (!name) continue;
    pushExcelInventoryEntry(entries, seen, {
      kind: "namedItem",
      label: name,
      name,
      type: trimString(namedItem.type),
      anchor: toAnchor({
        kind: "namedItem",
        namedItemName: name,
        label: name,
      }),
    });
  }

  for (const worksheet of Array.isArray(snippets.workbookSheets) ? snippets.workbookSheets : []) {
    if (!isRecord(worksheet)) continue;
    const name = trimString(worksheet.name);
    if (!name) continue;
    pushExcelInventoryEntry(entries, seen, {
      kind: "worksheet",
      label: name,
      name,
      sheetName: name,
      id: trimString(worksheet.id),
      anchor: toAnchor({
        kind: "sheet",
        sheetName: name,
        label: name,
      }),
    });
  }

  for (const citation of Array.isArray(snippets.selectionCellCitations) ? snippets.selectionCellCitations : []) {
    if (!isRecord(citation)) continue;
    const citationAnchor = isRecord(citation.anchor) ? citation.anchor : undefined;
    const address = trimString(citationAnchor?.address) ?? trimString(citation.address);
    const sheetName = trimString(citationAnchor?.sheetName) ?? trimString(citation.sheetName);
    if (!address && !trimString(citation.label)) {
      continue;
    }
    pushExcelInventoryEntry(entries, seen, {
      kind: "cell",
      label: trimString(citation.label) ?? (sheetName && address ? `${sheetName}!${address}` : address ?? "Cell"),
      sheetName,
      address,
      text: trimString(citation.text),
      formula: trimString(citation.formula),
      anchor: toAnchor({
        ...(citationAnchor ? { anchor: citationAnchor } : {}),
        ...(sheetName ? { sheetName } : {}),
        ...(address ? { address } : {}),
      }),
    });
  }

  return entries;
}

function serializeExcelInventoryEntry(entry: ExcelInventoryEntry): Record<string, unknown> {
  return {
    kind: entry.kind,
    label: entry.label,
    name: entry.name,
    sheetName: entry.sheetName,
    address: entry.address,
    id: entry.id,
    type: entry.type,
    text: entry.text,
    formula: entry.formula,
    anchor: entry.anchor,
  };
}

function toCsvCell(value: unknown, delimiter: string, quoteValues: boolean): string {
  const text = value == null ? "" : String(value);
  const escaped = text.replace(/"/g, "\"\"");
  const shouldQuote =
    quoteValues || escaped.includes("\"") || escaped.includes("\n") || escaped.includes("\r") || escaped.includes(delimiter);
  return shouldQuote ? `"${escaped}"` : escaped;
}

function toCsv(rows: unknown[][], delimiter: string, quoteValues: boolean): string {
  return rows.map((row) => row.map((cell) => toCsvCell(cell, delimiter, quoteValues)).join(delimiter)).join("\n");
}

function toExcelMatrixFromPayload(payload: unknown, includeFormulas: boolean): unknown[][] {
  if (!isRecord(payload)) {
    return [];
  }

  const data = isRecord(payload.data) ? payload.data : payload;
  const values = Array.isArray(data.values) ? data.values : undefined;
  const text = Array.isArray(data.text) ? data.text : undefined;
  const formulas = Array.isArray(data.formulas) ? data.formulas : undefined;
  const selected = includeFormulas ? formulas ?? values ?? text : values ?? text ?? formulas;
  if (!Array.isArray(selected)) {
    return [];
  }
  return selected.map((row) => (Array.isArray(row) ? row : [row]));
}

export function summarizeOfficeToolError(error: unknown, toolName: OfficeToolRequest["toolName"]): string {
  const record = getErrorRecord(error);
  const message =
    trimString(record?.message) ||
    (error instanceof Error ? trimString(error.message) : undefined) ||
    trimString(error) ||
    "Unknown Office host error.";
  const code = trimString(record?.code);
  const debugInfo = getErrorRecord(record?.debugInfo);
  const errorLocation = trimString(debugInfo?.errorLocation);
  const statement = trimString(debugInfo?.statement);
  const traceMessages = stringifyDetail(record?.traceMessages);

  const details = [
    code ? `code=${code}` : undefined,
    errorLocation ? `location=${errorLocation}` : undefined,
    statement ? `statement=${statement}` : undefined,
    traceMessages ? `trace=${traceMessages}` : undefined,
  ].filter((entry): entry is string => Boolean(entry));

  return details.length ? `${toolName} failed: ${message} (${details.join(", ")})` : `${toolName} failed: ${message}`;
}

export function serializeOfficeToolError(error: unknown): Record<string, unknown> {
  const record = getErrorRecord(error);
  const debugInfo = getErrorRecord(record?.debugInfo);

  return {
    name: trimString(record?.name) || (error instanceof Error ? error.name : undefined),
    message:
      trimString(record?.message) ||
      (error instanceof Error ? trimString(error.message) : undefined) ||
      trimString(error) ||
      "Unknown Office host error.",
    code: trimString(record?.code),
    stack: trimString(record?.stack) || (error instanceof Error ? trimString(error.stack) : undefined),
    traceMessages: record?.traceMessages,
    debugInfo,
    details: getErrorRecord(record?.details) ?? getErrorRecord((error as { details?: unknown } | undefined)?.details),
    error,
  };
}

export function toAnchor(params: Record<string, unknown>): OfficeAnchor {
  const anchorRecord = isRecord(params.anchor) ? params.anchor : params;
  const explicitKind = trimString(anchorRecord.kind);
  const layoutId = trimString(anchorRecord.layoutId);
  const slideMasterId = trimString(anchorRecord.slideMasterId);
  const anchorId = trimString(anchorRecord.id);
  const isNotesAnchorId =
    typeof anchorId === "string" &&
    (anchorId.startsWith("notes:") || anchorId.startsWith("notesRegion:") || anchorId.startsWith("slideNotes:"));
  const inferredKind =
    layoutId
      ? "layout"
      : slideMasterId
        ? "slideMaster"
        : isNotesAnchorId
          ? "notesRegion"
        : typeof anchorRecord.slideIndex === "number"
          ? "slide"
            : trimString(anchorRecord.slideId)
              ? "slide"
            : trimString(anchorRecord.shapeId)
              ? "shape"
              : trimString(anchorRecord.address)
                ? isSingleCellAddress(trimString(anchorRecord.address)) ? "cell" : "range"
                : trimString(anchorRecord.chartName)
                  ? "chart"
                  : trimString(anchorRecord.pivotTableName)
                    ? "pivotTable"
                    : trimString(anchorRecord.tableName)
                      ? "table"
                      : trimString(anchorRecord.sheetName)
                        ? "sheet"
                        : trimString(anchorRecord.paragraphId)
                          ? "paragraph"
                          : trimString(anchorRecord.commentId)
                            ? "comment"
                            : trimString(anchorRecord.revisionId)
                              ? "revision"
                              : anchorId?.startsWith("contentControl:")
                                ? "contentControl"
                                : anchorId?.startsWith("field:")
                                  ? "field"
                                  : "selection";

  return {
    kind: (explicitKind ?? inferredKind) as OfficeAnchor["kind"],
    label: trimString(anchorRecord.label) ?? trimString(anchorRecord.layoutName) ?? trimString(anchorRecord.slideMasterName),
    id: anchorId ?? layoutId ?? slideMasterId,
    text: trimString(anchorRecord.text) ?? trimString(params.target),
    sheetName: trimString(anchorRecord.sheetName),
    address: trimString(anchorRecord.address),
    paragraphId: trimString(anchorRecord.paragraphId),
    commentId: trimString(anchorRecord.commentId),
    revisionId: trimString(anchorRecord.revisionId),
    slideId: trimString(anchorRecord.slideId),
    slideIndex: typeof anchorRecord.slideIndex === "number" ? anchorRecord.slideIndex : undefined,
    shapeId: trimString(anchorRecord.shapeId),
    tableName: trimString(anchorRecord.tableName),
    chartName: trimString(anchorRecord.chartName),
    pivotTableName: trimString(anchorRecord.pivotTableName),
    namedItemName: trimString(anchorRecord.namedItemName),
  };
}

function parseMatrixContent(content: string): unknown {
  try {
    return JSON.parse(content);
  } catch {
    return undefined;
  }
}

function firstString(...values: unknown[]): string | undefined {
  for (const value of values) {
    if (typeof value === "string") {
      return value;
    }
  }
  return undefined;
}

function normalizeActionType(value: string | undefined): string | undefined {
  if (!value) {
    return undefined;
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  if (compact === "replaceselection" || compact === "insertafterselection" || compact === "insertbeforeselection") {
    return undefined;
  }
  if (compact === "inserttext") {
    return "insertText";
  }
  if (compact === "inserthtml") {
    return "insertHtml";
  }
  if (compact === "setrangevalues") {
    return "setRangeValues";
  }

  return value;
}

function normalizeMode(value: string | undefined): string {
  if (!value) {
    return "replaceSelection";
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  if (compact === "insertafterselection" || compact === "insertafter") {
    return "insertAfterSelection";
  }
  if (compact === "insertbeforeselection" || compact === "insertbefore") {
    return "insertBeforeSelection";
  }
  if (compact === "setrangevalues") {
    return "setRangeValues";
  }
  if (compact === "replaceselection" || compact === "replace") {
    return "replaceSelection";
  }

  return value;
}

function normalizeFormat(value: string | undefined): string | undefined {
  if (!value) {
    return undefined;
  }

  const normalized = value.trim().toLowerCase();
  if (!normalized) {
    return undefined;
  }
  if (normalized === "htm") {
    return "html";
  }
  if (normalized === "values" || normalized === "table") {
    return "matrix";
  }

  return normalized;
}

function toStringArray(value: unknown): string[] {
  if (Array.isArray(value)) {
    return value
      .map((entry) => trimString(entry))
      .filter((entry): entry is string => Boolean(entry));
  }

  const single = trimString(value);
  return single ? [single] : [];
}

function toExcelRangeTarget(
  params: Record<string, unknown>,
  keys: { sheetNameKey?: string; addressKey?: string; anchorKey?: string } = {},
): OfficeAnchor | undefined {
  const sheetName = trimString(params[keys.sheetNameKey ?? "sheetName"]);
  const address = trimString(params[keys.addressKey ?? "address"]);
  const anchorCandidate = params[keys.anchorKey ?? "anchor"];
  const explicitAnchor = isRecord(anchorCandidate) && Object.keys(anchorCandidate).length > 0
    ? (anchorCandidate as Record<string, unknown>)
    : undefined;

  if (!sheetName && !address && !explicitAnchor) {
    return undefined;
  }

  const next: Record<string, unknown> = {};
  if (explicitAnchor) {
    next.anchor = explicitAnchor;
  }
  if (sheetName) {
    next.sheetName = sheetName;
  }
  if (address) {
    next.address = address;
  }

  return toAnchor(next);
}

function toExcelToolOptions(params: Record<string, unknown>, excludedKeys: string[]): Record<string, unknown> {
  const next: Record<string, unknown> = {
    ...(isRecord(params.options) ? params.options : {}),
  };
  const excluded = new Set<string>(["options", ...excludedKeys]);

  for (const [key, value] of Object.entries(params)) {
    if (excluded.has(key)) {
      continue;
    }
    next[key] = value;
  }

  return next;
}

function toExcelSheetStructureActionType(value: string | undefined): string | undefined {
  if (!value) {
    return undefined;
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  switch (compact) {
    case "createworksheet":
    case "addworksheet":
    case "createsheet":
    case "addsheet":
      return "createWorksheet";
    case "renameworksheet":
    case "renamesheet":
      return "renameWorksheet";
    case "duplicateworksheet":
    case "duplicatesheet":
    case "copyworksheet":
    case "copysheet":
      return "duplicateWorksheet";
    case "deleteworksheet":
    case "deletesheet":
    case "removeworksheet":
    case "removesheet":
      return "deleteWorksheet";
    default:
      return undefined;
  }
}

function toPowerPointStructureActionType(value: string | undefined): string | undefined {
  if (!value) {
    return undefined;
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  switch (compact) {
    case "addslide":
      return "addSlide";
    case "moveslide":
      return "moveSlide";
    case "reorderslides":
    case "reorderstoryline":
      return "reorderSlides";
    case "deleteslide":
      return "deleteSlide";
    case "deleteslides":
      return "deleteSlides";
    case "applylayout":
      return "applyLayout";
    case "selectslides":
      return "selectSlides";
    case "addagendaslide":
      return "addAgendaSlide";
    case "addtransitionslide":
      return "addTransitionSlide";
    case "combineslides":
      return "combineSlides";
    case "importslidesfrombase64":
    case "mergepresentationfrombase64":
      return "importSlidesFromBase64";
    default:
      return undefined;
  }
}

function toPowerPointElementInsertActionType(value: string | undefined): string | undefined {
  if (!value) {
    return undefined;
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  switch (compact) {
    case "addtextbox":
    case "inserttextbox":
      return "addTextBox";
    case "addgeometricshape":
    case "insertgeometricshape":
    case "addshape":
      return "addGeometricShape";
    case "addtable":
    case "inserttable":
      return "addTable";
    case "addline":
    case "insertline":
      return "addLine";
    case "addprocessflow":
      return "addProcessFlow";
    case "addsimplediagram":
      return "addSimpleDiagram";
    case "insertinlinepicture":
    case "insertimage":
    case "addimage":
      return "insertInlinePicture";
    default:
      return undefined;
  }
}

function toPowerPointElementRemoveActionType(value: string | undefined): string | undefined {
  if (!value) {
    return "deleteShape";
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  switch (compact) {
    case "removeshape":
    case "deleteshape":
      return "deleteShape";
    case "removeshapes":
    case "deleteshapes":
      return "deleteShapes";
    case "clearshapetext":
      return "clearShapeText";
    default:
      return undefined;
  }
}

function toPowerPointTextAction(
  value: string | undefined,
  params: Record<string, unknown>,
): { type: string; placement?: string | undefined } | undefined {
  if (!value) {
    return trimString(params.shapeId) ? { type: "setShapeText" } : { type: "insertText" };
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  switch (compact) {
    case "setshapetext":
    case "editshapetext":
    case "replaceshapetext":
      return { type: "setShapeText" };
    case "appendshapetext":
    case "appendtext":
      return { type: "setShapeText", placement: "after" };
    case "clearshapetext":
      return { type: "clearShapeText" };
    case "inserttext":
    case "editselectedtext":
    case "setselectedtext":
      return { type: "insertText" };
    default:
      return undefined;
  }
}

function toPowerPointXmlActionType(value: string | undefined): string | undefined {
  if (!value) {
    return undefined;
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  switch (compact) {
    case "inspectpresentationpackage":
    case "readpresentationpackage":
      return "inspectPresentationPackage";
    case "getpresentationtheme":
    case "inspecttheme":
      return "getPresentationTheme";
    case "getslidenotes":
    case "readslidenotes":
    case "inspectslidenotes":
      return "getSlideNotes";
    case "setslidenotes":
    case "editslidenotes":
      return "setSlideNotes";
    case "replaceslidenotes":
      return "replaceSlideNotes";
    case "importslidesfrombase64":
      return "importSlidesFromBase64";
    case "mergepresentationfrombase64":
      return "mergePresentationFromBase64";
    case "exportslidesasbase64":
      return "exportSlidesAsBase64";
    default:
      return undefined;
  }
}

function toPowerPointMasterActionType(value: string | undefined): string | undefined {
  if (!value) {
    return "applyLayout";
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  switch (compact) {
    case "applylayout":
    case "setlayout":
    case "editslidelayout":
      return "applyLayout";
    default:
      return undefined;
  }
}

function toPowerPointChartActionType(value: string | undefined): string | undefined {
  if (!value) {
    return "updateSlideChart";
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  switch (compact) {
    case "getslidecharts":
    case "inspectslidecharts":
    case "readslidecharts":
    case "inspectcharts":
    case "readcharts":
    case "listcharts":
      return "getSlideCharts";
    case "addslidechart":
    case "createslidechart":
    case "addcharttoslide":
    case "insertslidechart":
    case "createchart":
    case "addchart":
      return "addSlideChart";
    case "updateslidechart":
    case "setchartdata":
    case "updatechartdata":
    case "replacechartdata":
    case "editslidechart":
    case "editchart":
      return "updateSlideChart";
    default:
      return undefined;
  }
}

function toPowerPointStructureOptions(params: Record<string, unknown>): Record<string, unknown> {
  const next: Record<string, unknown> = {
    ...(isRecord(params.options) ? params.options : {}),
  };

  for (const [key, value] of Object.entries(params)) {
    if (
      key === "options" ||
      key === "operation" ||
      key === "mode" ||
      key === "type" ||
      key === "anchor" ||
      key === "target" ||
      key === "content"
    ) {
      continue;
    }
    next[key] = value;
  }

  return next;
}

function toPowerPointStructureTarget(params: Record<string, unknown>): OfficeAnchor | undefined {
  const hasSlideTarget = Boolean(trimString(params.slideId)) || typeof params.slideIndex === "number";
  const hasShapeTarget = Boolean(trimString(params.shapeId));
  const hasLayoutTarget = Boolean(trimString(params.layoutId) || trimString(params.layoutName));
  const hasMasterTarget = Boolean(trimString(params.slideMasterId) || trimString(params.slideMasterName));
  const hasExplicitAnchor = isRecord(params.anchor) && Object.keys(params.anchor).length > 0;

  if (!hasSlideTarget && !hasShapeTarget && !hasLayoutTarget && !hasMasterTarget && !hasExplicitAnchor) {
    return undefined;
  }

  return toAnchor(params);
}

export function toHostAction(params: Record<string, unknown>): OfficeHostAction {
  if (isRecord(params.action) && typeof params.action.type === "string") {
    const action = { ...(params.action as Record<string, unknown>) } as OfficeHostAction;
    if (isRecord(action.target)) {
      action.target = toAnchor({ anchor: action.target });
    }
    return action;
  }

  const requestedType = normalizeActionType(trimString(params.operation) ?? trimString(params.type));
  const mode = normalizeMode(trimString(params.mode));
  const inferredFormatFromType =
    requestedType === "insertHtml" ? "html" : requestedType === "setRangeValues" ? "matrix" : undefined;
  const format = normalizeFormat(trimString(params.format)) ?? inferredFormatFromType;
  const content = firstString(
    params.content,
    format === "html" ? params.html : undefined,
    params.text,
    params.html,
  ) ?? "";
  const target = toAnchor(params);
  const directValues = Array.isArray(params.values) ? (params.values as OfficeHostAction["values"]) : undefined;

  if (requestedType === "setRangeValues" || mode === "setRangeValues" || format === "matrix") {
    return {
      type: "setRangeValues",
      target,
      values: directValues ?? (parseMatrixContent(content) as OfficeHostAction["values"]),
      content,
      format,
    };
  }

  const placement = trimString(params.placement) ?? (mode === "insertAfterSelection" ? "after" : mode === "insertBeforeSelection" ? "before" : "replace");

  return {
    type: requestedType ?? (format === "html" ? "insertHtml" : "insertText"),
    target,
    content,
    format,
    placement,
  };
}

export interface OfficeToolExecutorDependencies {
  collectOfficeContext: (
    host: OfficeHost,
    options: { includeFormatting?: boolean; maxImages?: number; scope?: string },
  ) => Promise<unknown>;
  applyHostAction: (host: OfficeHost, action: OfficeHostAction) => Promise<unknown>;
  navigateOfficeAnchor: (host: OfficeHost, anchor: OfficeAnchor) => Promise<unknown>;
  readDocumentSection: (host: OfficeHost, startIndex: number, endIndex: number, includeStyles: boolean) => Promise<unknown>;
  executeOfficeJs: (host: OfficeHost, code: string) => Promise<unknown>;
  proposeEdits: (host: OfficeHost, params: Record<string, unknown>) => Promise<unknown>;
  logger?: Pick<Console, "error"> | undefined;
}

export function createOfficeToolExecutor(dependencies: OfficeToolExecutorDependencies) {
  return async function executeOfficeTool(request: OfficeToolRequest): Promise<OfficeToolResult> {
    try {
      if (request.toolName === "office_get_context") {
        const scope = trimString(request.params.scope);
        const state = await dependencies.collectOfficeContext(request.host, {
          includeFormatting: true,
          maxImages: 0,
          ...(scope ? { scope } : {}),
        });
        return {
          requestId: request.requestId,
          success: true,
          content: state,
        };
      }

      if (request.toolName === "office_apply_edit") {
        const result = await dependencies.applyHostAction(request.host, toHostAction(request.params));
        return {
          requestId: request.requestId,
          success: true,
          content: result,
        };
      }

      if (request.toolName === "edit_doc_text") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "edit_doc_text is only available for Word.",
          };
        }

        const result = await dependencies.applyHostAction(request.host, toHostAction(request.params));
        return {
          requestId: request.requestId,
          success: true,
          content: result,
        };
      }

      if (request.toolName === "edit_doc_list") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "edit_doc_list is only available for Word.",
          };
        }

        const result = await dependencies.proposeEdits(request.host, request.params);
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "office_navigate") {
        const result = await dependencies.navigateOfficeAnchor(request.host, toAnchor(request.params));
        return {
          requestId: request.requestId,
          success: true,
          content: result,
        };
      }

      if (request.toolName === "office_capture_snapshot") {
        const includeFormatting = request.params.includeFormatting !== false;
        const scope = trimString(request.params.scope);
        const maxImages =
          typeof request.params.maxImages === "number" && Number.isFinite(request.params.maxImages)
            ? Math.max(0, Math.min(4, Math.trunc(request.params.maxImages)))
            : 2;
        const result = await dependencies.collectOfficeContext(request.host, {
          includeFormatting,
          maxImages,
          ...(scope ? { scope } : {}),
        });
        return {
          requestId: request.requestId,
          success: true,
          content: result,
        };
      }

      if (request.toolName === "office_capture_viewport") {
        return {
          requestId: request.requestId,
          success: false,
          error:
            "office_capture_viewport is a compatibility tool for true viewport/window screenshots and requires companion native capture. " +
            "Use office_capture_snapshot for Office.js context snapshots, verify_doc_visual for Word metadata/selection visuals, read_range_image for Excel active-selection snapshots, or verify_slide_visual for PowerPoint native slide/shape snapshots.",
        };
      }

      if (request.toolName === "office_read_section") {
        const startIndexValue = request.params.startIndex ?? request.params.start;
        const endIndexValue = request.params.endIndex ?? request.params.end;
        const startIndex = typeof startIndexValue === "number" ? Math.max(0, Math.trunc(startIndexValue)) : 0;
        const endIndex = typeof endIndexValue === "number" ? Math.trunc(endIndexValue) : startIndex + 20;
        const includeStyles = request.params.includeStyles !== false;
        const result = await dependencies.readDocumentSection(request.host, startIndex, endIndex, includeStyles);
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "verify_doc") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "verify_doc is only available for Word.",
          };
        }

        const scope = trimString(request.params.scope);
        const includeFormatting = request.params.includeFormatting !== false;
        const rawResult = await dependencies.collectOfficeContext(request.host, {
          includeFormatting,
          maxImages: 0,
          ...(scope ? { scope } : {}),
        });
        const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
        if (!payloadAware.success) {
          return payloadAware;
        }

        return {
          requestId: request.requestId,
          success: true,
          content: toWordDocumentVerificationPayload(payloadAware.content, scope),
        };
      }

      if (request.toolName === "verify_doc_visual") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "verify_doc_visual is only available for Word.",
          };
        }

        const includeFormatting = request.params.includeFormatting !== false;
        const includeWindowFrameRequested = request.params.includeWindowFrame === true;
        const rawResult = await dependencies.collectOfficeContext(request.host, {
          includeFormatting,
          maxImages: 1,
          scope: "viewport",
        });
        const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
        if (!payloadAware.success) {
          return payloadAware;
        }

        const viewportPayload = appendViewportCaptureSummary(payloadAware.content, includeWindowFrameRequested);
        return {
          requestId: request.requestId,
          success: true,
          content: toWordVisualVerificationPayload(viewportPayload, includeWindowFrameRequested),
        };
      }

      if (request.toolName === "get_cell_ranges") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "get_cell_ranges is only available for Excel.",
          };
        }

        const target = toExcelRangeTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: "getRangeValues",
          ...(target ? { target } : {}),
          options: toExcelToolOptions(request.params, ["anchor", "sheetName", "address"]),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "set_cell_range") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "set_cell_range is only available for Excel.",
          };
        }

        const action = toHostAction({
          ...request.params,
          operation: "setRangeValues",
          mode: "setRangeValues",
          format: "matrix",
        });
        const result = await dependencies.applyHostAction(request.host, {
          ...action,
          type: "setRangeValues",
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "clear_cell_range") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "clear_cell_range is only available for Excel.",
          };
        }

        const target = toExcelRangeTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: "clearRange",
          ...(target ? { target } : {}),
          options: toExcelToolOptions(request.params, ["anchor", "sheetName", "address"]),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "resize_range") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "resize_range is only available for Excel.",
          };
        }

        const target = toExcelRangeTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: "resizeRange",
          ...(target ? { target } : {}),
          options: toExcelToolOptions(request.params, ["anchor", "sheetName", "address"]),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "copy_to") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "copy_to is only available for Excel.",
          };
        }

        const sourceTarget =
          toExcelRangeTarget(request.params, {
            sheetNameKey: "sourceSheetName",
            addressKey: "sourceAddress",
            anchorKey: "sourceAnchor",
          }) ?? toExcelRangeTarget(request.params);
        const destinationTarget = toExcelRangeTarget(request.params, {
          sheetNameKey: "destinationSheetName",
          addressKey: "destinationAddress",
          anchorKey: "destinationAnchor",
        });
        if (!destinationTarget?.address) {
          return {
            requestId: request.requestId,
            success: false,
            error: "copy_to requires destinationAddress (and optional destinationSheetName).",
          };
        }

        const options = toExcelToolOptions(request.params, [
          "anchor",
          "sheetName",
          "address",
          "sourceAnchor",
          "sourceSheetName",
          "sourceAddress",
          "destinationAnchor",
          "destinationSheetName",
          "destinationAddress",
        ]);
        options.destinationAddress = destinationTarget.address;
        if (destinationTarget.sheetName) {
          options.destinationSheetName = destinationTarget.sheetName;
        }

        const result = await dependencies.applyHostAction(request.host, {
          type: "copyRange",
          ...(sourceTarget ? { target: sourceTarget } : {}),
          options,
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "modify_sheet_structure") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "modify_sheet_structure is only available for Excel.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toExcelSheetStructureActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "modify_sheet_structure requires a supported operation (create_worksheet, rename_worksheet, duplicate_worksheet, delete_worksheet).",
          };
        }

        const sheetName = trimString(request.params.sheetName) ?? trimString(request.params.sourceSheetName);
        const anchor = isRecord(request.params.anchor) && Object.keys(request.params.anchor).length > 0
          ? request.params.anchor
          : undefined;
        const target = sheetName || anchor ? toAnchor({
          ...(anchor ? { anchor } : {}),
          ...(sheetName ? { sheetName } : {}),
        }) : undefined;

        if (actionType !== "createWorksheet" && !target) {
          return {
            requestId: request.requestId,
            success: false,
            error: "modify_sheet_structure requires sheetName or anchor for rename, duplicate, and delete operations.",
          };
        }

        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          options: toExcelToolOptions(request.params, ["operation", "mode", "type", "sheetName", "sourceSheetName", "anchor"]),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "modify_object") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "modify_object is only available for Excel.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toExcelObjectActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "modify_object requires a supported operation (format_range, create_table, format_table, apply_table_filter, clear_table_filter, clear_table_filters, reapply_table_filters, create_chart, update_chart, create_pivot_table, update_pivot_table, sort_pivot_field, sort_pivot_by_labels, sort_pivot_by_values, refresh_pivot_table, set_worksheet_gridlines, set_worksheet_headings, set_print_area, set_data_validation, clear_data_validation, add_conditional_format, clear_conditional_formats, insert_inline_picture).",
          };
        }

        const target = toExcelObjectTarget(request.params);
        const content = firstString(
          request.params.content,
          request.params.text,
          request.params.base64,
          request.params.imageBase64,
        );
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          ...(typeof content === "string" ? { content } : {}),
          options: toExcelToolOptions(request.params, [
            "operation",
            "mode",
            "type",
            "anchor",
            "sheetName",
            "address",
            "tableName",
            "chartName",
            "pivotTableName",
            "namedItemName",
            "content",
            "text",
            "base64",
            "imageBase64",
          ]),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "get_all_objects") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "get_all_objects is only available for Excel.",
          };
        }

        const scope = trimString(request.params.scope) ?? "workbook";
        const includeFormatting = request.params.includeFormatting !== false;
        const rawResult = await dependencies.collectOfficeContext(request.host, {
          includeFormatting,
          maxImages: 0,
          scope,
        });
        const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
        if (!payloadAware.success) {
          return payloadAware;
        }

        const requestedKinds = toExcelObjectKindSet(request.params.objectTypes);
        const entries = toExcelInventoryEntries(payloadAware.content).filter((entry) =>
          requestedKinds ? requestedKinds.has(entry.kind) : true,
        );
        const limit =
          typeof request.params.limit === "number" && Number.isFinite(request.params.limit)
            ? Math.max(1, Math.min(200, Math.trunc(request.params.limit)))
            : entries.length || 200;
        const limited = entries.slice(0, limit);

        return {
          requestId: request.requestId,
          success: true,
          content: {
            summary: `Excel object inventory captured (${limited.length} item${limited.length === 1 ? "" : "s"}).`,
            details: {
              kind: "excel-object-inventory",
              mutating: false,
              scope,
              requestedObjectTypes: requestedKinds ? Array.from(requestedKinds) : undefined,
              matchCount: limited.length,
              totalCount: entries.length,
              objects: {
                tables: limited.filter((entry) => entry.kind === "table").map((entry) => serializeExcelInventoryEntry(entry)),
                charts: limited.filter((entry) => entry.kind === "chart").map((entry) => serializeExcelInventoryEntry(entry)),
                pivotTables: limited.filter((entry) => entry.kind === "pivotTable").map((entry) => serializeExcelInventoryEntry(entry)),
                namedItems: limited.filter((entry) => entry.kind === "namedItem").map((entry) => serializeExcelInventoryEntry(entry)),
                worksheets: limited.filter((entry) => entry.kind === "worksheet").map((entry) => serializeExcelInventoryEntry(entry)),
                cells: limited.filter((entry) => entry.kind === "cell").map((entry) => serializeExcelInventoryEntry(entry)),
              },
            },
          },
        };
      }

      if (request.toolName === "search_data") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "search_data is only available for Excel.",
          };
        }

        const query = trimString(request.params.query) ?? trimString(request.params.search) ?? trimString(request.params.text);
        if (!query) {
          return {
            requestId: request.requestId,
            success: false,
            error: "search_data requires a non-empty query.",
          };
        }

        const scope = trimString(request.params.scope) ?? "workbook";
        const rawResult = await dependencies.collectOfficeContext(request.host, {
          includeFormatting: false,
          maxImages: 0,
          scope,
        });
        const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
        if (!payloadAware.success) {
          return payloadAware;
        }

        const requestedKinds = toExcelObjectKindSet(request.params.objectTypes);
        const queryLower = query.toLowerCase();
        const matches = toExcelInventoryEntries(payloadAware.content)
          .filter((entry) => (requestedKinds ? requestedKinds.has(entry.kind) : true))
          .filter((entry) => {
            const fields = [entry.label, entry.name, entry.sheetName, entry.address, entry.type, entry.text, entry.formula]
              .filter((value): value is string => typeof value === "string");
            return fields.some((value) => value.toLowerCase().includes(queryLower));
          });

        const limit =
          typeof request.params.limit === "number" && Number.isFinite(request.params.limit)
            ? Math.max(1, Math.min(200, Math.trunc(request.params.limit)))
            : 40;
        const limited = matches.slice(0, limit);

        return {
          requestId: request.requestId,
          success: true,
          content: {
            summary: `Excel search found ${limited.length} match${limited.length === 1 ? "" : "es"} for "${query}".`,
            details: {
              kind: "excel-data-search",
              mutating: false,
              query,
              scope,
              requestedObjectTypes: requestedKinds ? Array.from(requestedKinds) : undefined,
              matchCount: limited.length,
              totalMatches: matches.length,
              matches: limited.map((entry) => serializeExcelInventoryEntry(entry)),
            },
          },
        };
      }

      if (request.toolName === "get_range_as_csv") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "get_range_as_csv is only available for Excel.",
          };
        }

        const target = toExcelRangeTarget(request.params);
        const includeFormulas = request.params.includeFormulas === true;
        const delimiter = trimString(request.params.delimiter) ?? ",";
        const quoteValues = request.params.quoteValues === true;
        const includeHeaders = request.params.includeHeaders !== false;
        const rawResult = await dependencies.applyHostAction(request.host, {
          type: "getRangeValues",
          ...(target ? { target } : {}),
          options: {
            ...toExcelToolOptions(request.params, [
              "anchor",
              "sheetName",
              "address",
              "delimiter",
              "quoteValues",
              "includeHeaders",
              "includeFormulas",
            ]),
            includeValues: true,
            includeText: true,
            includeFormulas: true,
          },
        });
        const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
        if (!payloadAware.success) {
          return payloadAware;
        }

        const matrix = toExcelMatrixFromPayload(payloadAware.content, includeFormulas);
        const rows = includeHeaders ? matrix : matrix.slice(1);
        const csv = toCsv(rows, delimiter, quoteValues);
        const payloadRecord = isRecord(payloadAware.content) ? payloadAware.content : {};
        const payloadData = isRecord(payloadRecord.data) ? payloadRecord.data : payloadRecord;
        const sheetName = trimString(payloadData.sheetName) ?? target?.sheetName;
        const address = trimString(payloadData.address) ?? target?.address;
        const columnCount = rows.length && Array.isArray(rows[0]) ? rows[0].length : 0;

        return {
          requestId: request.requestId,
          success: true,
          content: {
            summary: `Exported ${rows.length} row${rows.length === 1 ? "" : "s"} from Excel range as CSV.`,
            csv,
            details: {
              kind: "excel-range-csv-export",
              mutating: false,
              sheetName,
              address,
              rowCount: rows.length,
              columnCount,
              delimiter,
              includeFormulas,
            },
          },
        };
      }

      if (request.toolName === "read_range_image") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "read_range_image is only available for Excel.",
          };
        }

        const includeFormatting = request.params.includeFormatting !== false;
        const maxImages =
          typeof request.params.maxImages === "number" && Number.isFinite(request.params.maxImages)
            ? Math.max(1, Math.min(4, Math.trunc(request.params.maxImages)))
            : 1;
        const scope = trimString(request.params.scope) ?? "selection";
        const rawResult = await dependencies.collectOfficeContext(request.host, {
          includeFormatting,
          maxImages,
          scope,
        });
        const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
        if (!payloadAware.success) {
          return payloadAware;
        }

        const payloadRecord = isRecord(payloadAware.content) ? payloadAware.content : {};
        const visuals = Array.isArray(payloadRecord.visuals) ? payloadRecord.visuals : [];
        const requestedSheetName = trimString(request.params.sheetName);
        const requestedAddress = trimString(request.params.address);
        const state = isRecord(payloadRecord.state) ? payloadRecord.state : {};
        const selection = isRecord(state.selection) ? state.selection : {};
        const activeSelectionLabel = trimString(selection.label);
        const requestedRangeHonored = excelSelectionMatchesRequest(
          activeSelectionLabel,
          requestedSheetName,
          requestedAddress,
        );
        if (requestedRangeHonored === false) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              `read_range_image can only capture the active Excel selection as an Office.js image snapshot. ` +
              `The active selection is ${activeSelectionLabel ?? "unknown"}, not ${formatExcelRangeHint(requestedSheetName, requestedAddress)}. ` +
              "Navigate to or select the target range first, then retry.",
          };
        }
        if (!visuals.length) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "read_range_image could not capture an Office.js image snapshot of the active Excel selection. " +
              "Browser-only runtime does not render arbitrary ranges by address; select the target range first and retry.",
          };
        }

        const captureNote =
          requestedRangeHonored === true
            ? "Captured the active Excel selection through Office.js image coercion."
            : "Captured the active Excel selection through Office.js image coercion; requested sheet/address values are advisory unless they match the active selection.";

        return {
          requestId: request.requestId,
          success: true,
          content: {
            summary:
              trimString(payloadRecord.summary) ??
              `Excel active-selection visual snapshot captured (${visuals.length} image${visuals.length === 1 ? "" : "s"}). ${captureNote}`,
            visual: {
              kind: "excel-selection-snapshot",
              captureMode: "officejs-selection-snapshot",
              imageCount: visuals.length,
              scopeRequested: scope,
              requestedRangeHonored,
              note: captureNote,
            },
            details: {
              kind: "excel-selection-snapshot-read",
              mutating: false,
              host: "excel",
              scope,
              requestedRange: {
                sheetName: requestedSheetName,
                address: requestedAddress,
                honored: requestedRangeHonored,
              },
              selection,
              formatting: payloadRecord.formatting,
            },
            visuals,
          },
        };
      }

      if (request.toolName === "extract_chart_xml") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "extract_chart_xml is only available for Excel.",
          };
        }

        const chartName = trimString(request.params.chartName) ?? trimString(request.params.name);
        const chartId = trimString(request.params.chartId) ?? trimString(request.params.id);
        const chartIndex =
          typeof request.params.chartIndex === "number" && Number.isFinite(request.params.chartIndex)
            ? Math.max(1, Math.trunc(request.params.chartIndex))
            : undefined;
        const hasAnchor = isRecord(request.params.anchor) && Object.keys(request.params.anchor).length > 0;
        if (!chartName && !chartId && typeof chartIndex !== "number" && !hasAnchor) {
          return {
            requestId: request.requestId,
            success: false,
            error: "extract_chart_xml requires chartName, chartId, chartIndex, or a chart anchor.",
          };
        }

        const target = toExcelObjectTarget({
          ...request.params,
          ...(chartName ? { chartName } : {}),
          ...(chartId ? { id: chartId } : {}),
        });
        const rawResult = await dependencies.applyHostAction(request.host, {
          type: "extractChartXml",
          ...(target ? { target } : {}),
          options: toExcelToolOptions(request.params, ["anchor", "sheetName", "chartName", "chartId", "chartIndex"]),
        });
        const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
        if (!payloadAware.success) {
          return payloadAware;
        }

        const payloadRecord = isRecord(payloadAware.content) ? payloadAware.content : {};
        const payloadData = isRecord(payloadRecord.data) ? payloadRecord.data : payloadRecord;
        const xml = trimString(payloadData.chartXml) ?? trimString(payloadData.xml) ?? trimString(payloadRecord.chartXml);
        if (!xml) {
          return {
            requestId: request.requestId,
            success: false,
            error: "extract_chart_xml did not return chart XML content.",
          };
        }

        const resolvedChartName = trimString(payloadData.chartName) ?? chartName;
        const resolvedSheetName = trimString(payloadData.sheetName) ?? trimString(request.params.sheetName) ?? target?.sheetName;

        return {
          requestId: request.requestId,
          success: true,
          content: {
            summary: trimString(payloadRecord.summary) ?? `Extracted chart XML for ${resolvedChartName ?? "the target chart"}.`,
            xml,
            details: {
              kind: "excel-chart-xml",
              mutating: false,
              chartName: resolvedChartName,
              sheetName: resolvedSheetName,
              extraction: "runtime-generated-chart-metadata-xml",
            },
          },
        };
      }

      if (request.toolName === "get_presentation_structure") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "get_presentation_structure is only available for PowerPoint.",
          };
        }

        const result = await dependencies.applyHostAction(request.host, {
          type: "getPresentationStructure",
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "get_slide") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "get_slide is only available for PowerPoint.",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: "getSlide",
          ...(target ? { target } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "list_slide_shapes") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "list_slide_shapes is only available for PowerPoint.",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: "listSlideShapes",
          ...(target ? { target } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "modify_presentation_structure") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "modify_presentation_structure is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toPowerPointStructureActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "modify_presentation_structure requires a supported operation (add_slide, move_slide, reorder_slides, delete_slide, apply_layout, select_slides, add_agenda_slide, add_transition_slide, combine_slides, import_slides_from_base64).",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          ...(typeof request.params.content === "string" ? { content: request.params.content } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "duplicate_slide") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "duplicate_slide is only available for PowerPoint.",
          };
        }

        const requestedSlideIds = toStringArray(request.params.slideIds);
        const fallbackSlideId = trimString(request.params.slideId);
        const slideIds = requestedSlideIds.length ? requestedSlideIds : fallbackSlideId ? [fallbackSlideId] : [];
        const targetParams =
          slideIds.length && !fallbackSlideId && typeof request.params.slideIndex !== "number"
            ? { ...request.params, slideId: slideIds[0] }
            : request.params;
        const target = toPowerPointStructureTarget(targetParams);
        const result = await dependencies.applyHostAction(request.host, {
          type: "duplicateSlide",
          ...(target ? { target } : {}),
          options: {
            ...toPowerPointStructureOptions(request.params),
            ...(slideIds.length ? { slideIds } : {}),
          },
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "insert_slide_element") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "insert_slide_element is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toPowerPointElementInsertActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "insert_slide_element requires a supported operation (add_text_box, add_geometric_shape, add_table, add_line, add_process_flow, add_simple_diagram, insert_inline_picture).",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const content = firstString(request.params.content, request.params.text, request.params.base64);
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          ...(typeof content === "string" ? { content } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "remove_slide_element") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "remove_slide_element is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toPowerPointElementRemoveActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "remove_slide_element requires a supported operation (remove_shape, remove_shapes, clear_shape_text).",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "edit_slide_text") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "edit_slide_text is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const textAction = toPowerPointTextAction(operation, request.params);
        if (!textAction) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "edit_slide_text requires a supported operation (set_shape_text, append_shape_text, clear_shape_text, insert_text).",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const content = firstString(request.params.content, request.params.text, request.params.newText) ?? "";
        const actionOptions = toPowerPointStructureOptions(request.params);
        const placement = trimString(request.params.placement) ?? textAction.placement;
        if (placement && !("placement" in actionOptions)) {
          actionOptions.placement = placement;
        }
        const result = await dependencies.applyHostAction(request.host, {
          type: textAction.type,
          ...(target ? { target } : {}),
          ...(textAction.type !== "clearShapeText" ? { content } : {}),
          ...(placement ? { placement } : {}),
          options: actionOptions,
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "edit_slide_xml") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "edit_slide_xml is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toPowerPointXmlActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "edit_slide_xml requires a supported operation (inspect_presentation_package, get_presentation_theme, get_slide_notes, set_slide_notes, replace_slide_notes, import_slides_from_base64, merge_presentation_from_base64, export_slides_as_base64).",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const content = firstString(request.params.content, request.params.text, request.params.base64);
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          ...(typeof content === "string" ? { content } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "edit_slide_master") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "edit_slide_master is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toPowerPointMasterActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "edit_slide_master is a legacy-named layout application tool. It currently supports only apply_layout/set_layout; it does not edit slide masters.",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "edit_slide_chart") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "edit_slide_chart is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toPowerPointChartActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "edit_slide_chart requires a supported operation (get_slide_charts, add_slide_chart, update_slide_chart).",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const content = firstString(request.params.content, request.params.title);
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          ...(typeof content === "string" ? { content } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "copy_image_between_slides") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "copy_image_between_slides is only available for PowerPoint.",
          };
        }

        const targetParams: Record<string, unknown> = {
          ...request.params,
          slideId: trimString(request.params.targetSlideId) ?? trimString(request.params.slideId),
          slideIndex:
            typeof request.params.targetSlideIndex === "number"
              ? request.params.targetSlideIndex
              : request.params.slideIndex,
          shapeId: trimString(request.params.targetShapeId) ?? trimString(request.params.shapeId),
        };
        const target = toPowerPointStructureTarget(targetParams);
        const content = firstString(request.params.sourceImageBase64, request.params.base64, request.params.content);
        const result = await dependencies.applyHostAction(request.host, {
          type: "copyImageBetweenSlides",
          ...(target ? { target } : {}),
          ...(typeof content === "string" ? { content } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "search_icons") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "search_icons is only available for PowerPoint.",
          };
        }

        const query = firstString(request.params.query, request.params.search, request.params.content);
        if (!query?.trim()) {
          return {
            requestId: request.requestId,
            success: false,
            error: "search_icons requires a non-empty query.",
          };
        }

        const result = await dependencies.applyHostAction(request.host, {
          type: "searchIcons",
          content: query,
          options: {
            ...toPowerPointStructureOptions(request.params),
            query,
          },
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "insert_icon") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "insert_icon is only available for PowerPoint.",
          };
        }

        const iconId = firstString(request.params.iconId, request.params.iconName, request.params.query, request.params.content);
        if (!iconId?.trim()) {
          return {
            requestId: request.requestId,
            success: false,
            error: "insert_icon requires iconId, iconName, query, or content.",
          };
        }

        const targetParams: Record<string, unknown> = {
          ...request.params,
          slideId: trimString(request.params.targetSlideId) ?? trimString(request.params.slideId),
          slideIndex:
            typeof request.params.targetSlideIndex === "number"
              ? request.params.targetSlideIndex
              : request.params.slideIndex,
          shapeId: trimString(request.params.targetShapeId) ?? trimString(request.params.shapeId),
        };
        const target = toPowerPointStructureTarget(targetParams);
        const result = await dependencies.applyHostAction(request.host, {
          type: "insertIcon",
          ...(target ? { target } : {}),
          content: iconId,
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "verify_slides") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "verify_slides is only available for PowerPoint.",
          };
        }

        const scope = trimString(request.params.scope);
        const result = await dependencies.applyHostAction(request.host, {
          type: "getPresentationStructure",
          options: toPowerPointStructureOptions(request.params),
        });
        const payloadAware = toPayloadAwareResult(request.requestId, result);
        if (!payloadAware.success) {
          return payloadAware;
        }

        return {
          requestId: request.requestId,
          success: true,
          content: toPowerPointSlidesVerificationPayload(payloadAware.content, scope),
        };
      }

      if (request.toolName === "verify_slide_visual") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "verify_slide_visual is only available for PowerPoint.",
          };
        }

        const includeFormatting = request.params.includeFormatting !== false;
        const maxImages =
          typeof request.params.maxImages === "number" && Number.isFinite(request.params.maxImages)
            ? Math.max(1, Math.min(4, Math.trunc(request.params.maxImages)))
            : 1;
        const scope = trimString(request.params.scope) ?? "slide";
        const rawResult = await dependencies.collectOfficeContext(request.host, {
          includeFormatting,
          maxImages,
          scope,
        });
        const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
        if (!payloadAware.success) {
          return payloadAware;
        }

        return {
          requestId: request.requestId,
          success: true,
          content: toPowerPointVisualVerificationPayload(payloadAware.content, maxImages),
        };
      }

      if (request.toolName === "office_execute_js") {
        const code = firstString(request.params.code, request.params.script) ?? "";
        if (!code.trim()) {
          return { requestId: request.requestId, success: false, error: "No code provided." };
        }
        const result = await dependencies.executeOfficeJs(request.host, code);
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "office_propose_edits") {
        const result = await dependencies.proposeEdits(request.host, request.params);
        return toPayloadAwareResult(request.requestId, result);
      }

      return {
        requestId: request.requestId,
        success: false,
        error: `${request.toolName} is not supported by the taskpane bridge.`,
      };
    } catch (error) {
      dependencies.logger?.error?.(`[office-tool] ${request.toolName} failed`, {
        requestId: request.requestId,
        host: request.host,
        params: request.params,
        error: serializeOfficeToolError(error),
      });

      return {
        requestId: request.requestId,
        success: false,
        error: summarizeOfficeToolError(error, request.toolName),
      };
    }
  };
}
