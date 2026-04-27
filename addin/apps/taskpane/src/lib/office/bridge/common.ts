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
  searchWordDocument?: (params: Record<string, unknown>) => Promise<unknown>;
  applyHostAction: (host: OfficeHost, action: OfficeHostAction) => Promise<unknown>;
  navigateOfficeAnchor: (host: OfficeHost, anchor: OfficeAnchor) => Promise<unknown>;
  readDocumentSection: (host: OfficeHost, startIndex: number, endIndex: number, includeStyles: boolean) => Promise<unknown>;
  executeOfficeJs: (host: OfficeHost, code: string) => Promise<unknown>;
  proposeEdits: (host: OfficeHost, params: Record<string, unknown>) => Promise<unknown>;
  logger?: Pick<Console, "error"> | undefined;
}


export {
  appendViewportCaptureSummary,
  excelSelectionMatchesRequest,
  firstString,
  formatExcelRangeHint,
  isRecord,
  isSingleCellAddress,
  serializeExcelInventoryEntry,
  toCsv,
  toExcelInventoryEntries,
  toExcelMatrixFromPayload,
  toExcelObjectActionType,
  toExcelObjectKindSet,
  toExcelObjectTarget,
  toExcelRangeTarget,
  toExcelSheetStructureActionType,
  toExcelToolOptions,
  toPayloadAwareResult,
  toPowerPointChartActionType,
  toPowerPointElementInsertActionType,
  toPowerPointElementRemoveActionType,
  toPowerPointMasterActionType,
  toPowerPointSlidesVerificationPayload,
  toPowerPointStructureActionType,
  toPowerPointStructureOptions,
  toPowerPointStructureTarget,
  toPowerPointTextAction,
  toPowerPointVisualVerificationPayload,
  toPowerPointXmlActionType,
  toStringArray,
  toWordDocumentVerificationPayload,
  toWordVisualVerificationPayload,
  trimString,
};
