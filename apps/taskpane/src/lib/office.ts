import type {
  ConnectorScopeContext,
  OfficeActionCompletion,
  OfficeActionResult,
  OfficeAnchor,
  OfficeCellValue,
  OfficeContextPayload,
  OfficeHost,
  OfficeHostAction,
  OfficeObjectReference,
  OfficeSelectionMeta,
  OfficeSelectionSummary,
  OfficeSessionOpenRequest,
  OfficeStateUpdate,
  OfficeVisualSnapshot,
} from "@pi-office/pi-office-pack/protocol";
import { HOST_LABELS } from "@pi-office/pi-office-pack/defaults";
import { assertDestructiveActionAllowed } from "./office-action-policy";
import type { OfficeCaptureOptions, OfficeHostAdapter } from "./office-host-adapter-types";
import { createExcelOfficeHostAdapter } from "./office-excel-adapter";
import {
  createPowerPointChartInPresentationBase64,
  inspectPowerPointPresentationBase64,
  replaceSlideNotesInPowerPointPresentationBase64,
  updatePowerPointChartInPresentationBase64,
} from "./powerpoint-transform";
import { createPowerPointOfficeHostAdapter } from "./office-powerpoint-adapter";
import { createWordOfficeHostAdapter } from "./office-word-adapter";

export interface OfficeThemeSnapshot {
  bodyBackgroundColor: string;
  bodyForegroundColor: string;
  controlBackgroundColor: string;
  controlForegroundColor: string;
  isDarkTheme: boolean;
}

interface ExcelCitationRecord {
  anchor: OfficeAnchor;
  label: string;
  sheetName: string;
  address: string;
  textPreview?: string | undefined;
  formula?: string | undefined;
  numberFormat?: string | undefined;
}

interface ExcelWorksheetSnapshot {
  worksheetId: string;
  worksheetName: string;
  usedRange?: { address: string; rowCount: number; columnCount: number } | undefined;
  previewAddress: string;
  previewRows: string[];
  citedCells: ExcelCitationRecord[];
}

interface ExcelBorderUpdateOptions {
  color?: string | undefined;
  style?: Excel.BorderLineStyle | undefined;
  weight?: Excel.BorderWeight | undefined;
  tintAndShade?: number | undefined;
}

function mapHost(host: Office.HostType | string): OfficeHost {
  switch (host) {
    case Office.HostType.Word:
      return "word";
    case Office.HostType.Excel:
      return "excel";
    case Office.HostType.PowerPoint:
      return "powerpoint";
    default:
      throw new Error(`Unsupported Office host: ${String(host)}`);
  }
}

function basename(input: string): string {
  const parts = input.split(/[\\/]/g).filter(Boolean);
  return parts[parts.length - 1] ?? input;
}

function toDocumentPath(documentUrl: string | undefined): string | undefined {
  if (!documentUrl) {
    return undefined;
  }

  try {
    const parsed = new URL(documentUrl);
    if (parsed.protocol !== "file:") {
      return undefined;
    }

    return decodeURIComponent(parsed.pathname.replace(/^\//, "")).replace(/\//g, "\\");
  } catch {
    return undefined;
  }
}

function getUnsavedDocumentId(host: OfficeHost): string {
  const key = `pi-office-unsaved:${host}`;
  const existing = sessionStorage.getItem(key);
  if (existing) {
    return existing;
  }

  const created = `${host}-${crypto.randomUUID()}`;
  sessionStorage.setItem(key, created);
  return created;
}

function supportsRequirementSet(name: string, version: string): boolean {
  try {
    return Office.context.requirements.isSetSupported(name, version);
  } catch {
    return false;
  }
}

function normalizeTextPreview(value: string | undefined): string | undefined {
  const normalized = value?.replace(/\s+/g, " ").trim();
  return normalized ? normalized.slice(0, 420) : undefined;
}

function truncateText(value: string | undefined, limit = 2400): string | undefined {
  const normalized = value?.trim();
  if (!normalized) {
    return undefined;
  }

  return normalized.length > limit ? `${normalized.slice(0, limit)}\n…` : normalized;
}

function pluralize(count: number, singular: string, plural = `${singular}s`): string {
  return `${count} ${count === 1 ? singular : plural}`;
}

function uniqueDetails(details: Array<string | undefined>): string[] | undefined {
  const next = details.filter((detail): detail is string => Boolean(detail?.trim())).map((detail) => detail.trim());
  return next.length ? Array.from(new Set(next)).slice(0, 6) : undefined;
}

function containsImageMarkup(markup: string | undefined): boolean {
  return /<(?:img|w:drawing|pic:pic|a:blip|v:imagedata|v:shape)\b/i.test(markup ?? "");
}

function stripDataUrlPrefix(input: string): { data: string; mimeType: string } {
  const match = input.match(/^data:([^;]+);base64,(.+)$/);
  if (!match) {
    return { data: input, mimeType: "image/png" };
  }

  return {
    mimeType: match[1] ?? "image/png",
    data: match[2] ?? "",
  };
}

function formatPoints(value: number | undefined): string | undefined {
  if (typeof value !== "number" || Number.isNaN(value)) {
    return undefined;
  }

  return `${Math.round(value * 10) / 10} pt`;
}

function formatPercent(value: number | undefined): number | undefined {
  if (typeof value !== "number" || Number.isNaN(value)) {
    return undefined;
  }

  return Math.round(value * 10) / 10;
}

function pointsToPixels(value: number | undefined): number | undefined {
  if (typeof value !== "number" || Number.isNaN(value)) {
    return undefined;
  }

  return Math.round((value * 96) / 72);
}

function serializeWordPages(pages: Array<{ index: number; width: number; height: number }>): Array<Record<string, unknown>> {
  return pages.map((page) => ({
    index: page.index,
    width: formatPoints(page.width),
    height: formatPoints(page.height),
  }));
}

function buildSelectionSummary(input: {
  textPreview?: string | undefined;
  imageCount?: number | undefined;
  objectCount?: number | undefined;
  altPreview?: string | undefined;
  details?: string[] | undefined;
  emptyLabel: string;
}): OfficeSelectionSummary {
  const textPreview = normalizeTextPreview(input.textPreview);
  const imageCount = input.imageCount ?? 0;
  const objectCount = input.objectCount ?? 0;
  const hasText = Boolean(textPreview);
  const hasImages = imageCount > 0;
  const hasObjects = objectCount > 0;

  let label = input.emptyLabel;
  let kind: OfficeSelectionSummary["kind"] = "empty";
  if (hasText && hasImages) {
    label = `Mixed selection · ${textPreview!.length} chars + ${pluralize(imageCount, "image")}`;
    kind = "mixed";
  } else if (hasText && hasObjects) {
    label = `Mixed selection · ${textPreview!.length} chars + ${pluralize(objectCount, "object")}`;
    kind = "mixed";
  } else if (hasImages) {
    label = `${pluralize(imageCount, "image")} selected`;
    kind = "images";
  } else if (hasObjects) {
    label = `${pluralize(objectCount, "object")} selected`;
    kind = "objects";
  } else if (hasText) {
    label = `Selection · ${textPreview!.length} chars`;
    kind = "text";
  }

  return {
    label,
    kind,
    imageCount,
    objectCount,
    textPreview: textPreview ?? input.altPreview ?? (hasImages ? `${pluralize(imageCount, "image")} selected.` : undefined),
    details: input.details?.length ? input.details : undefined,
  };
}

function wordStyleToMarkdownPrefix(style: string | undefined): string {
  const normalized = (style ?? "").toLowerCase().replace(/\s+/g, "");
  if (/^heading1$/i.test(normalized)) return "# ";
  if (/^heading2$/i.test(normalized)) return "## ";
  if (/^heading3$/i.test(normalized)) return "### ";
  if (/^heading[4-9]$/i.test(normalized)) return "#### ";
  if (/^title$/i.test(normalized)) return "# ";
  if (/^subtitle$/i.test(normalized)) return "## ";
  if (/listbullet|listparagraph/i.test(normalized)) return "- ";
  if (/listnumber/i.test(normalized)) return "1. ";
  if (/quote|blockquote|intensequote/i.test(normalized)) return "> ";
  return "";
}

function friendlyStyleName(style: string | undefined): string {
  const raw = (style ?? "Normal").replace(/([a-z])([A-Z])/g, "$1 $2");
  return raw.charAt(0).toUpperCase() + raw.slice(1);
}

function buildStructuredSelectionPreview(
  paragraphs: Array<{ text?: string | undefined; style?: string | undefined; styleBuiltIn?: string | undefined }>,
  limit = 1500,
): string | undefined {
  if (!paragraphs.length) return undefined;
  const lines: string[] = [];
  let length = 0;
  for (const p of paragraphs) {
    if (length >= limit) break;
    const text = p.text?.trim() ?? "";
    const style = String(p.styleBuiltIn || p.style || "");
    if (!text) {
      lines.push("");
      continue;
    }
    const prefix = wordStyleToMarkdownPrefix(style);
    const line = `${prefix}${text}`;
    lines.push(line);
    length += line.length + 1;
  }
  const result = lines.join("\n").trim();
  return result || undefined;
}

function buildSelectionMeta(
  paragraphs: Array<{ text?: string | undefined; style?: string | undefined; styleBuiltIn?: string | undefined }>,
  font?: { name?: string; size?: number; color?: string; bold?: boolean; italic?: boolean } | undefined,
): OfficeSelectionMeta {
  const histogram: Record<string, number> = {};
  let isListItem = false;
  for (const p of paragraphs) {
    const style = friendlyStyleName(p.styleBuiltIn || p.style);
    histogram[style] = (histogram[style] ?? 0) + 1;
    if (/list/i.test(p.styleBuiltIn || p.style || "")) isListItem = true;
  }
  const firstStyle = paragraphs[0] ? friendlyStyleName(paragraphs[0].styleBuiltIn || paragraphs[0].style) : undefined;
  const enclosing = paragraphs[0]?.text?.trim();

  return {
    paragraphCount: paragraphs.length,
    firstParagraphStyle: firstStyle,
    styleHistogram: Object.keys(histogram).length > 0 ? histogram : undefined,
    isListItem: isListItem || undefined,
    enclosingParagraphText: enclosing ? (enclosing.length > 300 ? `${enclosing.slice(0, 300)}…` : enclosing) : undefined,
    fontSummary: font?.name || font?.size
      ? {
          name: font.name || undefined,
          size: font.size || undefined,
          bold: font.bold || undefined,
          italic: font.italic || undefined,
          color: font.color || undefined,
        }
      : undefined,
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === "object" && !Array.isArray(value));
}

function trimString(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function truncateLabel(value: string | undefined, limit = 96): string | undefined {
  const normalized = normalizeTextPreview(value);
  if (!normalized) {
    return undefined;
  }

  return normalized.length > limit ? `${normalized.slice(0, limit)}…` : normalized;
}

function matchesTextQuery(value: string | undefined, query: string | undefined): boolean {
  const needle = normalizeTextPreview(query)?.toLowerCase();
  if (!needle) {
    return false;
  }

  return (normalizeTextPreview(value)?.toLowerCase() ?? "").includes(needle);
}

function parseIndexedAnchorNumber(value: string | undefined, prefix: string): number | undefined {
  const normalized = trimString(value);
  if (!normalized) {
    return undefined;
  }

  const matcher = new RegExp(`^${prefix}[\\s:]+(\\d+)$`, "i");
  const match = normalized.match(matcher);
  if (!match) {
    return undefined;
  }

  const index = Number(match[1]);
  return Number.isInteger(index) && index > 0 ? index : undefined;
}

function parsePositiveInteger(value: unknown): number | undefined {
  const numeric = toNumber(value);
  return typeof numeric === "number" && Number.isInteger(numeric) && numeric > 0 ? numeric : undefined;
}

function parseAnchorOrdinal(value: string | undefined, prefix: string): number | undefined {
  return parsePositiveInteger(value) ?? parseIndexedAnchorNumber(value, prefix);
}

function clampPercentage(value: unknown): number | undefined {
  const numeric = toNumber(value);
  if (typeof numeric !== "number") {
    return undefined;
  }

  return Math.min(1, Math.max(0, numeric));
}

function anchorKey(anchor: OfficeAnchor): string {
  return [
    anchor.kind,
    anchor.id,
    anchor.paragraphId,
    anchor.commentId,
    anchor.revisionId,
    anchor.slideId,
    anchor.slideIndex,
    anchor.shapeId,
    anchor.sheetName,
    anchor.address,
    anchor.tableName,
    anchor.chartName,
    anchor.pivotTableName,
    anchor.namedItemName,
    anchor.text,
    anchor.label,
  ]
    .filter((value) => value != null && value !== "")
    .join("|");
}

function uniqueAnchors(anchors: OfficeAnchor[]): OfficeAnchor[] {
  const seen = new Set<string>();
  const result: OfficeAnchor[] = [];
  for (const anchor of anchors) {
    const key = anchorKey(anchor);
    if (!key || seen.has(key)) {
      continue;
    }
    seen.add(key);
    result.push(anchor);
  }

  return result;
}

function getActionOptions(action: OfficeHostAction): Record<string, unknown> {
  return isRecord(action.options) ? action.options : {};
}

function toNumber(value: unknown): number | undefined {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  if (typeof value === "string" && value.trim()) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : undefined;
  }

  return undefined;
}

function toBoolean(value: unknown): boolean | undefined {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "string") {
    if (value === "true") return true;
    if (value === "false") return false;
  }

  return undefined;
}

function toStringMatrix(value: unknown): string[][] | undefined {
  if (!Array.isArray(value)) {
    return undefined;
  }

  return value.map((row) =>
    Array.isArray(row)
      ? row.map((cell) => (cell == null ? "" : typeof cell === "string" ? cell : String(cell)))
      : [row == null ? "" : typeof row === "string" ? row : String(row)],
  );
}

function toCellValue(value: unknown): OfficeCellValue {
  if (value == null || typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return value as OfficeCellValue;
  }

  return JSON.stringify(value);
}

function toValueMatrix(value: unknown, fallbackContent?: string): OfficeCellValue[][] {
  if (Array.isArray(value)) {
    return value.map((row) => (Array.isArray(row) ? row.map((cell) => toCellValue(cell)) : [toCellValue(row)]));
  }

  if (typeof value === "string") {
    try {
      const parsed = JSON.parse(value);
      if (Array.isArray(parsed)) {
        return toValueMatrix(parsed);
      }
    } catch {
      // Fall through to scalar handling.
    }
  }

  return [[toCellValue(fallbackContent ?? value ?? "")]];
}

function splitSheetAddress(
  address: string | undefined,
  fallbackSheetName?: string | undefined,
): { sheetName?: string | undefined; address?: string | undefined } {
  const normalizedAddress = trimString(address);
  if (!normalizedAddress) {
    return { sheetName: fallbackSheetName };
  }

  const bangIndex = normalizedAddress.indexOf("!");
  if (bangIndex < 0) {
    return {
      sheetName: fallbackSheetName,
      address: normalizedAddress,
    };
  }

  const rawSheetName = normalizedAddress.slice(0, bangIndex).replace(/^'/, "").replace(/'$/, "");
  return {
    sheetName: rawSheetName || fallbackSheetName,
    address: normalizedAddress.slice(bangIndex + 1),
  };
}

function withSheetName(sheetName: string | undefined, address: string | undefined): string | undefined {
  const normalizedAddress = trimString(address);
  if (!normalizedAddress) {
    return undefined;
  }
  if (!sheetName || normalizedAddress.includes("!")) {
    return normalizedAddress;
  }

  return `${sheetName}!${normalizedAddress}`;
}

function isSingleCellAddress(address: string | undefined): boolean {
  const normalized = trimString(address)?.replace(/\$/g, "");
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

function excelColumnName(index: number): string {
  let next = index + 1;
  let result = "";
  while (next > 0) {
    const remainder = (next - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    next = Math.floor((next - 1) / 26);
  }
  return result;
}

function excelCellAddress(rowIndex: number, columnIndex: number): string {
  return `${excelColumnName(columnIndex)}${rowIndex + 1}`;
}

function firstMatrixString(value: unknown): string | undefined {
  if (Array.isArray(value)) {
    const first = value[0];
    if (Array.isArray(first)) {
      return trimString(first[0]);
    }
    return trimString(first);
  }

  return trimString(value);
}

function buildExcelCitationRecord(input: {
  kind: "cell" | "range";
  sheetName: string;
  address: string;
  text?: string | undefined;
  formula?: string | undefined;
  numberFormat?: string | undefined;
}): ExcelCitationRecord {
  const label = `${input.sheetName}!${input.address}`;
  return {
    anchor: {
      kind: input.kind,
      label,
      sheetName: input.sheetName,
      address: input.address,
      text: truncateLabel(input.text, 160),
    } satisfies OfficeAnchor,
    label,
    sheetName: input.sheetName,
    address: input.address,
    textPreview: truncateLabel(input.text, 160),
    formula: trimString(input.formula),
    numberFormat: trimString(input.numberFormat),
  };
}

function wordInsertLocationFromPlacement(placement: string | undefined): Word.InsertLocation {
  switch (placement) {
    case "after":
      return Word.InsertLocation.after;
    case "before":
      return Word.InsertLocation.before;
    case "start":
      return Word.InsertLocation.start;
    case "end":
      return Word.InsertLocation.end;
    default:
      return Word.InsertLocation.replace;
  }
}

function compactStrings(values: Array<string | undefined>): string[] | undefined {
  const filtered = values.map((value) => trimString(value)).filter((value): value is string => Boolean(value));
  return filtered.length ? filtered : undefined;
}

function toActionCompletion(value: unknown): OfficeActionCompletion | undefined {
  return value === "native" || value === "fallback" || value === "partial" ? value : undefined;
}

function toObjectReference(host: OfficeHost, anchor: OfficeAnchor | undefined): OfficeObjectReference | undefined {
  if (!anchor) {
    return undefined;
  }

  return {
    host,
    kind: anchor.kind,
    label: anchor.label,
    id: anchor.id,
    text: anchor.text,
    sheetName: anchor.sheetName,
    address: anchor.address,
    paragraphId: anchor.paragraphId,
    commentId: anchor.commentId,
    revisionId: anchor.revisionId,
    slideId: anchor.slideId,
    slideIndex: anchor.slideIndex,
    shapeId: anchor.shapeId,
    tableName: anchor.tableName,
    chartName: anchor.chartName,
    pivotTableName: anchor.pivotTableName,
    namedItemName: anchor.namedItemName,
  };
}

function mergeObjectReferences(references: Array<OfficeObjectReference | undefined>): OfficeObjectReference[] | undefined {
  const seen = new Set<string>();
  const merged: OfficeObjectReference[] = [];
  for (const reference of references) {
    if (!reference) {
      continue;
    }
    const key = `${reference.host}|${anchorKey(reference)}`;
    if (!key || seen.has(key)) {
      continue;
    }
    seen.add(key);
    merged.push(reference);
  }

  return merged.length ? merged : undefined;
}

function extractObjectReferencesFromRecord(host: OfficeHost, record: Record<string, unknown>): OfficeObjectReference[] {
  const refs: OfficeObjectReference[] = [];

  const push = (anchor: OfficeAnchor | undefined) => {
    const reference = toObjectReference(host, anchor);
    if (reference) {
      refs.push(reference);
    }
  };
  const pushSlides = (value: unknown) => {
    for (const slide of getRecordArray(value)) {
      push({
        kind: "slide",
        label: trimString(slide.label) ?? (typeof slide.slideIndex === "number" ? `Slide ${slide.slideIndex}` : undefined),
        slideId: trimString(slide.slideId),
        slideIndex: typeof slide.slideIndex === "number" ? slide.slideIndex : undefined,
      });
    }
  };
  const pushShapes = (value: unknown) => {
    for (const shape of getRecordArray(value)) {
      push({
        kind: "shape",
        label: trimString(shape.shapeName) ?? trimString(shape.label),
        shapeId: trimString(shape.shapeId),
        slideId: trimString(shape.slideId),
        slideIndex: typeof shape.slideIndex === "number" ? shape.slideIndex : undefined,
      });
    }
  };

  push(
    trimString(record.namedItemName)
      ? {
          kind: "namedItem",
          label: trimString(record.namedItemName),
          namedItemName: trimString(record.namedItemName),
          sheetName: trimString(record.sheetName),
        }
      : undefined,
  );
  push(
    trimString(record.tableName)
      ? {
          kind: "table",
          label: trimString(record.tableName),
          tableName: trimString(record.tableName),
          id: trimString(record.tableId) ?? trimString(record.id),
          sheetName: trimString(record.sheetName),
          address: trimString(record.address),
        }
      : undefined,
  );
  push(
    trimString(record.chartName)
      ? {
          kind: "chart",
          label: trimString(record.chartName),
          chartName: trimString(record.chartName),
          id: trimString(record.chartId) ?? trimString(record.id),
          sheetName: trimString(record.sheetName),
        }
      : undefined,
  );
  push(
    trimString(record.pivotTableName)
      ? {
          kind: "pivotTable",
          label: trimString(record.pivotTableName),
          pivotTableName: trimString(record.pivotTableName),
          id: trimString(record.pivotTableId) ?? trimString(record.id),
          sheetName: trimString(record.sheetName),
          address: trimString(record.address),
        }
      : undefined,
  );
  push(
    trimString(record.sheetName) && !trimString(record.address) && !trimString(record.tableName) && !trimString(record.chartName) && !trimString(record.pivotTableName)
      ? {
          kind: "sheet",
          label: trimString(record.sheetName),
          sheetName: trimString(record.sheetName),
        }
      : undefined,
  );
  push(
    trimString(record.address)
      ? {
          kind: "range",
          label: trimString(record.sheetName) ? `${trimString(record.sheetName)}!${trimString(record.address)}` : trimString(record.address),
          sheetName: trimString(record.sheetName),
          address: trimString(record.address),
        }
      : undefined,
  );
  push(
    trimString(record.commentId)
      ? {
          kind: "comment",
          label: trimString(record.content) ?? trimString(record.label),
          commentId: trimString(record.commentId),
        }
      : undefined,
  );
  push(
    trimString(record.paragraphId)
      ? {
          kind: "paragraph",
          label: trimString(record.text) ?? trimString(record.label),
          paragraphId: trimString(record.paragraphId),
        }
      : undefined,
  );
  push(
    trimString(record.revisionId)
      ? {
          kind: "revision",
          label: trimString(record.text) ?? trimString(record.label),
          revisionId: trimString(record.revisionId),
        }
      : undefined,
  );
  push(
    parsePositiveInteger(record.contentControlId)
      ? {
          kind: "contentControl",
          id: `contentControl:${parsePositiveInteger(record.contentControlId)}`,
          label: trimString(record.title) ?? trimString(record.tag) ?? trimString(record.text) ?? trimString(record.label),
          text: trimString(record.text),
        }
      : undefined,
  );
  push(
    parsePositiveInteger(record.fieldIndex) || trimString(record.fieldId)
      ? {
          kind: "field",
          id: trimString(record.fieldId) ?? `field:${parsePositiveInteger(record.fieldIndex)}`,
          label: trimString(record.fieldType) ?? trimString(record.fieldCode) ?? trimString(record.text) ?? trimString(record.label),
          text: trimString(record.fieldCode) ?? trimString(record.text),
        }
      : undefined,
  );
  pushSlides(record.slides);
  pushSlides(record.createdSlides);
  pushSlides(record.deletedSlides);
  pushShapes(record.shapes);
  pushShapes(record.createdShapes);
  pushShapes(record.deletedShapes);
  push(
    trimString(record.slideId) || typeof record.slideIndex === "number"
      ? {
          kind: "slide",
          label: trimString(record.label) ?? (typeof record.slideIndex === "number" ? `Slide ${record.slideIndex}` : undefined),
          slideId: trimString(record.slideId),
          slideIndex: typeof record.slideIndex === "number" ? record.slideIndex : undefined,
        }
      : undefined,
  );
  push(
    trimString(record.shapeId)
      ? {
          kind: "shape",
          label: trimString(record.shapeName) ?? trimString(record.label),
          shapeId: trimString(record.shapeId),
          slideId: trimString(record.slideId),
          slideIndex: typeof record.slideIndex === "number" ? record.slideIndex : undefined,
        }
      : undefined,
  );

  return refs;
}

function isCreationOperation(operation: string): boolean {
  return /^(add|create|insert(?!Text|Html)|duplicate|merge)/i.test(operation);
}

function buildOfficeActionResult(
  host: OfficeHost,
  operation: string,
  raw: unknown,
  options: { target?: OfficeAnchor | undefined; navigation?: OfficeAnchor | undefined } = {},
): OfficeActionResult {
  const record = isRecord(raw) ? raw : { value: raw };
  const summary =
    trimString(record.summary) ??
    `${HOST_LABELS[host]} ${operation.replace(/([A-Z])/g, " $1").trim().toLowerCase()} completed.`;
  const targetReference = toObjectReference(host, options.target);
  const recordReferences = extractObjectReferencesFromRecord(host, record);
  const createdObjects = isCreationOperation(operation) ? mergeObjectReferences(recordReferences) : undefined;
  const touchedObjects = mergeObjectReferences([
    targetReference,
    ...(isCreationOperation(operation) ? [] : recordReferences),
  ]);
  const fallbackStrategy = trimString(record.fallbackStrategy) ?? trimString(record.fallback);
  const nativeFailure =
    trimString(record.nativeFailure) ??
    (isRecord(record.nativeError) ? trimString(record.nativeError.message) ?? trimString(record.nativeError.code) : undefined);
  const completion =
    toActionCompletion(record.completion) ??
    (fallbackStrategy ? "fallback" : trimString(record.partialFailure) ? "partial" : "native");
  const nativeAttempted =
    typeof record.nativeAttempted === "boolean" ? record.nativeAttempted : Boolean(fallbackStrategy || nativeFailure || completion !== "native");
  const warnings = compactStrings(
    [
      ...(Array.isArray(record.warnings) ? record.warnings.map((entry) => (typeof entry === "string" ? entry : undefined)) : []),
      trimString(record.warning),
      completion === "fallback" && fallbackStrategy ? `Completed via fallback path: ${fallbackStrategy}.` : undefined,
      completion === "partial" ? trimString(record.partialFailure) ?? "The host completed the request partially." : undefined,
      nativeFailure && fallbackStrategy ? `Native path detail: ${nativeFailure}` : undefined,
    ],
  );

  return {
    ok: true,
    host,
    operation,
    summary,
    touchedObjects,
    createdObjects,
    navigation: options.navigation,
    completion,
    fallbackStrategy,
    nativeAttempted,
    nativeFailure,
    warnings,
    data: record,
  };
}

function isPowerPointImageShape(type: unknown): boolean {
  return String(type ?? "").toLowerCase() === "image";
}

function isPowerPointTableShape(type: unknown): boolean {
  return String(type ?? "").toLowerCase() === "table";
}

function isPowerPointGroupShape(type: unknown): boolean {
  return String(type ?? "").toLowerCase() === "group";
}

function getPowerPointShapeContentKind(type: unknown): string {
  const normalized = String(type ?? "").toLowerCase();
  if (!normalized) {
    return "shape";
  }

  if (normalized === "table") return "table";
  if (normalized === "chart") return "chart";
  if (normalized === "group") return "group";
  if (normalized === "textbox" || normalized === "placeholder") return "text";
  if (normalized === "line" || normalized.includes("connector")) return "line";
  if (isPowerPointImageShape(normalized)) return "image";
  if (normalized.includes("geometric") || normalized.includes("arrow") || normalized.includes("callout")) return "geometric";
  return normalized;
}

function truncateStringMatrix(value: string[][] | undefined, rowLimit = 4, columnLimit = 6, cellLimit = 80): string[][] | undefined {
  if (!value?.length) {
    return undefined;
  }

  return value.slice(0, rowLimit).map((row) => row.slice(0, columnLimit).map((cell) => truncateText(cell, cellLimit) ?? ""));
}

function resolveZeroBasedIndex(indexValue: unknown, ordinalValue: unknown): number | undefined {
  const direct = toNumber(indexValue);
  if (typeof direct === "number" && Number.isInteger(direct) && direct >= 0) {
    return direct;
  }

  const ordinal = parsePositiveInteger(ordinalValue);
  return typeof ordinal === "number" ? ordinal - 1 : undefined;
}

function resolvePositiveCount(value: unknown, fallback = 1): number {
  return parsePositiveInteger(value) ?? fallback;
}

function isWordPictureShape(type: unknown): boolean {
  return String(type ?? "").toLowerCase() === "picture";
}

function countSelectedWordImages(inlinePictureCount: number, floatingImageShapeCount: number): number {
  // Word can surface a selected inline picture through both inlinePictures and shapes.
  return inlinePictureCount > 0 ? inlinePictureCount : floatingImageShapeCount;
}

async function getSelectedTextAsync(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error.message));
        return;
      }

      resolve(String(result.value ?? ""));
    });
  });
}

function getSelectedMarkupAsync(coercionType: Office.CoercionType.Html | Office.CoercionType.Ooxml): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(coercionType, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error.message));
        return;
      }

      resolve(String(result.value ?? ""));
    });
  });
}

function setSelectedTextAsync(value: string): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      value,
      { coercionType: Office.CoercionType.Text },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(result.error.message));
          return;
        }

        resolve();
      },
    );
  });
}

function setSelectedImageAsync(value: string, options: Partial<Office.SetSelectedDataOptions> = {}): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      value,
      {
        ...options,
        coercionType: Office.CoercionType.Image as never,
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(result.error.message));
          return;
        }

        resolve();
      },
    );
  });
}

async function buildBaseState(host: OfficeHost): Promise<OfficeStateUpdate> {
  const documentUrl = Office.context.document.url || undefined;
  const documentPath = toDocumentPath(documentUrl);
  const title = documentPath ? basename(documentPath) : `Unsaved ${HOST_LABELS[host]}`;
  const id = documentUrl || getUnsavedDocumentId(host);

  return {
    host,
    document: {
      id,
      title,
      saved: Boolean(documentPath),
      documentUrl,
      documentPath,
      workspaceDir: documentPath ? documentPath.replace(/[\\/][^\\/]+$/, "") : undefined,
    },
    selection: {
      label: "Selection unavailable",
      kind: "empty",
      imageCount: 0,
      objectCount: 0,
    },
    capabilities: [],
    timestamp: new Date().toISOString(),
  };
}

async function collectWordState(base: OfficeStateUpdate): Promise<OfficeStateUpdate> {
  try {
    return await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const paragraphs = selection.paragraphs;
      const inlinePictures = selection.inlinePictures;
      const supportsShapes = supportsRequirementSet("WordApiDesktop", "1.2");
      const supportsViewportPages = supportsRequirementSet("WordApiDesktop", "1.2");
      const supportsWindowMetadata = supportsRequirementSet("WordApiDesktop", "1.4");
      const shapes = supportsShapes ? selection.shapes : undefined;

      selection.load("text");
      paragraphs.load("items/text,items/style,items/styleBuiltIn");
      inlinePictures.load("items/altTextTitle,items/altTextDescription");
      shapes?.load("items/type,items/name,items/altTextDescription");
      await context.sync();

      const preview = normalizeTextPreview(selection.text);
      const details = paragraphs.items.map((paragraph) => paragraph.text.trim()).filter(Boolean).slice(0, 2);
      const imageShapes = shapes?.items.filter((shape) => isWordPictureShape(shape.type)) ?? [];
      const imageCount = countSelectedWordImages(inlinePictures.items.length, imageShapes.length);
      const objectCount = shapes?.items.length ?? 0;
      const altPreview =
        inlinePictures.items
          .map((picture) => picture.altTextTitle || picture.altTextDescription)
          .find((value) => Boolean(value?.trim())) ??
        imageShapes
          .map((shape) => shape.altTextDescription || shape.name)
          .find((value) => Boolean(value?.trim()));
      const structuredPreview = buildStructuredSelectionPreview(paragraphs.items);
      const selectionMeta = paragraphs.items.length > 0 ? buildSelectionMeta(paragraphs.items) : undefined;
      const selectionSummary: OfficeSelectionSummary = {
        ...buildSelectionSummary({
          textPreview: preview,
          imageCount,
          objectCount: imageCount > 0 ? Math.max(objectCount - imageShapes.length, 0) : objectCount,
          altPreview: normalizeTextPreview(altPreview),
          details: uniqueDetails([
            ...details,
            imageCount > 0 ? pluralize(imageCount, "image") : undefined,
            objectCount > imageShapes.length ? pluralize(objectCount - imageShapes.length, "shape") : undefined,
          ]),
          emptyLabel: "Insertion point",
        }),
        structuredPreview,
        selectionMeta,
      };

      return {
        ...base,
        selection: selectionSummary,
        capabilities: [
          "word.selection",
          "word.inlinePictures",
          ...(supportsShapes ? ["word.shapes"] : []),
          ...(supportsViewportPages ? ["word.viewportPages", "word.selectionPages"] : []),
          ...(supportsWindowMetadata ? ["word.activeWindow", "word.view"] : []),
          "word.insertText",
          "word.insertHtml",
          "office.getSelectedData",
        ],
      };
    });
  } catch {
    const [preview, html, ooxml] = await Promise.all([
      getSelectedTextAsync().catch(() => ""),
      getSelectedMarkupAsync(Office.CoercionType.Html).catch(() => ""),
      getSelectedMarkupAsync(Office.CoercionType.Ooxml).catch(() => ""),
    ]);
    const imageCount = containsImageMarkup(html) || containsImageMarkup(ooxml) ? 1 : 0;
    return {
      ...base,
      selection: buildSelectionSummary({
        textPreview: preview,
        imageCount,
        details: uniqueDetails([imageCount > 0 ? "Image content detected" : undefined]),
        emptyLabel: "Insertion point",
      }),
      capabilities: ["office.getSelectedData", "office.setSelectedData"],
    };
  }
}

async function collectExcelState(base: OfficeStateUpdate): Promise<OfficeStateUpdate> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["address", "rowCount", "columnCount", "text"]);
    range.worksheet.load("name");
    await context.sync();

    const preview = range.text.flat().join(" | ").trim();
    const previewRows = range.text.slice(0, 8).map((row) => row.join("\t"));
    const excelStructuredPreview = previewRows.length > 0 ? previewRows.join("\n") : undefined;
    const excelMeta: OfficeSelectionMeta = {
      paragraphCount: range.rowCount,
      firstParagraphStyle: range.rowCount === 1 && range.columnCount === 1 ? "cell" : "range",
      styleHistogram: { [`${range.rowCount}x${range.columnCount}`]: 1 },
    };
    return {
      ...base,
      selection: {
        label: `${range.worksheet.name}!${range.address}`,
        kind: preview ? "text" : "empty",
        imageCount: 0,
        objectCount: 0,
        textPreview: preview || undefined,
        structuredPreview: excelStructuredPreview,
        selectionMeta: excelMeta,
        details: [`${range.rowCount} rows`, `${range.columnCount} columns`],
      },
      capabilities: ["excel.range", "excel.values", "excel.formulas", "excel.format"],
    };
  });
}

async function collectPowerPointState(base: OfficeStateUpdate): Promise<OfficeStateUpdate> {
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.getSelectedSlides();
      const supportsRichSelection = supportsRequirementSet("PowerPointApi", "1.5");
      const shapes = supportsRichSelection ? context.presentation.getSelectedShapes() : undefined;
      const textRange = supportsRichSelection ? context.presentation.getSelectedTextRangeOrNullObject() : undefined;

      slides.load("items/id,items/index");
      shapes?.load("items/type,items/name,items/id");
      textRange?.load("text,isNullObject");
      await context.sync();

      const preview =
        textRange && !textRange.isNullObject
          ? normalizeTextPreview(textRange.text)
          : normalizeTextPreview(await getSelectedTextAsync().catch(() => ""));
      const imageShapeCount = shapes?.items.filter((shape) => isPowerPointImageShape(shape.type)).length ?? 0;
      const objectCount = shapes?.items.length ?? 0;
      const details = uniqueDetails([
        ...slides.items.map((slide) => `Slide ${slide.index + 1}`).slice(0, 2),
        objectCount > 0 ? pluralize(objectCount, "shape") : undefined,
        imageShapeCount > 0 ? pluralize(imageShapeCount, "image") : undefined,
      ]);

      const pptShapeTypes = shapes?.items.map((s) => String(s.type ?? "shape")) ?? [];
      const pptStyleHistogram: Record<string, number> = {};
      for (const t of pptShapeTypes) {
        const kind = getPowerPointShapeContentKind(t);
        pptStyleHistogram[kind] = (pptStyleHistogram[kind] ?? 0) + 1;
      }
      const pptMeta: OfficeSelectionMeta | undefined = objectCount > 0 || preview
        ? {
            paragraphCount: preview ? 1 : undefined,
            firstParagraphStyle: pptShapeTypes[0] ? getPowerPointShapeContentKind(pptShapeTypes[0]) : undefined,
            styleHistogram: Object.keys(pptStyleHistogram).length > 0 ? pptStyleHistogram : undefined,
          }
        : undefined;

      return {
        ...base,
        selection: {
          ...buildSelectionSummary({
            textPreview: preview,
            imageCount: imageShapeCount,
            objectCount: imageShapeCount > 0 ? Math.max(objectCount - imageShapeCount, 0) : objectCount,
            altPreview: shapes?.items[0]?.name,
            details,
            emptyLabel: slides.items.length > 0 ? `${pluralize(slides.items.length, "slide")} selected` : "Slide selection",
          }),
          selectionMeta: pptMeta,
        },
        capabilities: [
          "powerpoint.selection",
          ...(supportsRichSelection ? ["powerpoint.shapes", "powerpoint.textRange"] : []),
          "powerpoint.slideSnapshot",
          "office.setSelectedData",
        ],
      };
    });
  } catch {
    const preview = await getSelectedTextAsync().catch(() => "");
    return {
      ...base,
      selection: buildSelectionSummary({
        textPreview: preview,
        emptyLabel: "Slide selection",
      }),
      capabilities: ["office.getSelectedData", "office.setSelectedData"],
    };
  }
}

function parseSelectedImageValue(value: unknown): { data: string; mimeType: string } | undefined {
  if (typeof value === "string") {
    return stripDataUrlPrefix(value);
  }

  if (!value || typeof value !== "object") {
    return undefined;
  }

  const record = value as Record<string, unknown>;
  const candidate = [record.data, record.base64, record.imageData, record.value].find(
    (entry): entry is string => typeof entry === "string" && entry.length > 0,
  );
  if (!candidate) {
    return undefined;
  }

  const parsed = stripDataUrlPrefix(candidate);
  return {
    data: parsed.data,
    mimeType: typeof record.mimeType === "string" && record.mimeType.trim() ? record.mimeType : parsed.mimeType,
  };
}

function getActionImagePayload(
  action: OfficeHostAction,
  options: Record<string, unknown>,
): { data: string; mimeType: string } | undefined {
  const candidates = [
    action.content,
    action.image,
    action.imageData,
    action.base64,
    action.data,
    options.image,
    options.imageData,
    options.base64,
    options.data,
    options.fillImageBase64,
  ];

  for (const candidate of candidates) {
    const parsed = parseSelectedImageValue(candidate);
    if (parsed?.data) {
      return parsed;
    }
  }

  return undefined;
}

function getSelectedImageAsync(): Promise<OfficeVisualSnapshot | undefined> {
  return new Promise((resolve, reject) => {
    try {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Image as never, (result: Office.AsyncResult<unknown>) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(result.error.message));
          return;
        }

        const parsed = parseSelectedImageValue(result.value);
        if (!parsed?.data) {
          resolve(undefined);
          return;
        }

        resolve({
          kind: "selection",
          label: "Selection snapshot",
          data: parsed.data,
          mimeType: parsed.mimeType,
        });
      });
    } catch (error) {
      reject(error instanceof Error ? error : new Error(String(error)));
    }
  });
}

function loadImageElement(sourceUrl: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("Failed to decode Office image payload."));
    image.src = sourceUrl;
  });
}

async function optimizeVisual(
  visual: OfficeVisualSnapshot,
  options?: { maxDimension?: number; preferredMimeType?: "image/png" | "image/jpeg" },
): Promise<OfficeVisualSnapshot> {
  if (typeof document === "undefined" || visual.mimeType === "image/svg+xml") {
    return visual;
  }

  try {
    const image = await loadImageElement(`data:${visual.mimeType};base64,${visual.data}`);
    const maxDimension = options?.maxDimension ?? 1400;
    const scale = Math.min(1, maxDimension / Math.max(image.naturalWidth, image.naturalHeight));
    const width = Math.max(1, Math.round(image.naturalWidth * scale));
    const height = Math.max(1, Math.round(image.naturalHeight * scale));
    if (scale === 1) {
      return { ...visual, width: visual.width ?? image.naturalWidth, height: visual.height ?? image.naturalHeight };
    }

    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d");
    if (!ctx) {
      return visual;
    }

    ctx.drawImage(image, 0, 0, width, height);
    const preferredMimeType =
      options?.preferredMimeType ?? (visual.mimeType === "image/jpeg" || visual.mimeType === "image/jpg" ? "image/jpeg" : "image/png");
    const dataUrl = canvas.toDataURL(preferredMimeType, preferredMimeType === "image/jpeg" ? 0.84 : undefined);
    const parsed = stripDataUrlPrefix(dataUrl);
    return {
      ...visual,
      data: parsed.data,
      mimeType: parsed.mimeType,
      width,
      height,
    };
  } catch {
    return visual;
  }
}

async function waitForSnapshotLayout(): Promise<void> {
  if (typeof document === "undefined") {
    return;
  }

  try {
    await document.fonts?.ready;
  } catch {
    // Font readiness is best-effort.
  }

  await new Promise<void>((resolve) => {
    requestAnimationFrame(() => {
      requestAnimationFrame(() => resolve());
    });
  });
}

function getWordSnapshotWidthPx(pageWidth: number | undefined, leftMargin: number | undefined, rightMargin: number | undefined): number {
  const contentWidth = typeof pageWidth === "number" ? pageWidth - (leftMargin ?? 0) - (rightMargin ?? 0) : undefined;
  return Math.max(520, Math.min(960, pointsToPixels(contentWidth) ?? 760));
}

async function renderHtmlSelectionSnapshot(params: {
  html: string | undefined;
  label: string;
  widthPx?: number;
}): Promise<OfficeVisualSnapshot | undefined> {
  if (
    typeof document === "undefined" ||
    typeof DOMParser === "undefined" ||
    typeof XMLSerializer === "undefined" ||
    typeof Blob === "undefined" ||
    !params.html?.trim()
  ) {
    return undefined;
  }

  const parsed = new DOMParser().parseFromString(params.html, "text/html");
  const embeddedStyles = Array.from(parsed.querySelectorAll("style"))
    .map((node) => node.textContent?.trim())
    .filter((value): value is string => Boolean(value));
  const bodyMarkup = parsed.body?.innerHTML?.trim() || parsed.documentElement?.innerHTML?.trim() || params.html.trim();
  if (!bodyMarkup) {
    return undefined;
  }

  const measurementHost = document.createElement("div");
  measurementHost.style.cssText = [
    "position:fixed",
    "left:-100000px",
    "top:0",
    "visibility:hidden",
    "pointer-events:none",
    "z-index:-1",
  ].join(";");

  const snapshotRoot = document.createElement("div");
  snapshotRoot.setAttribute("xmlns", "http://www.w3.org/1999/xhtml");
  snapshotRoot.style.width = `${Math.max(480, Math.round(params.widthPx ?? 760))}px`;
  snapshotRoot.style.padding = "30px 34px";
  snapshotRoot.style.boxSizing = "border-box";
  snapshotRoot.style.background = "#FFFFFF";
  snapshotRoot.style.color = "#201F1E";
  snapshotRoot.style.fontFamily = '"Aptos", "Segoe UI", sans-serif';
  snapshotRoot.style.lineHeight = "1.45";

  const baseStyles = document.createElement("style");
  baseStyles.textContent = `
    * { box-sizing: border-box; }
    .pi-office-html-snapshot {
      width: 100%;
      background: #FFFFFF;
      color: #201F1E;
      font-family: "Aptos", "Segoe UI", sans-serif;
      line-height: 1.45;
    }
    .pi-office-html-snapshot table {
      max-width: 100%;
      border-collapse: collapse;
    }
    .pi-office-html-snapshot img {
      max-width: 100%;
      height: auto;
    }
  `;
  snapshotRoot.append(baseStyles);

  if (embeddedStyles.length) {
    const officeStyles = document.createElement("style");
    officeStyles.textContent = embeddedStyles.join("\n\n");
    snapshotRoot.append(officeStyles);
  }

  const contentRoot = document.createElement("div");
  contentRoot.className = "pi-office-html-snapshot";
  contentRoot.innerHTML = bodyMarkup;
  contentRoot.querySelectorAll("script, iframe, object").forEach((node) => node.remove());
  contentRoot.querySelectorAll("img").forEach((image) => {
    const src = image.getAttribute("src")?.trim() ?? "";
    if (src && !src.startsWith("data:")) {
      image.removeAttribute("src");
    }
  });
  snapshotRoot.append(contentRoot);
  measurementHost.append(snapshotRoot);
  document.body.append(measurementHost);

  try {
    await waitForSnapshotLayout();
    const width = Math.max(1, Math.ceil(snapshotRoot.scrollWidth));
    const height = Math.max(1, Math.ceil(snapshotRoot.scrollHeight));
    if (height <= 8) {
      return undefined;
    }

    const serialized = new XMLSerializer().serializeToString(snapshotRoot);
    const svg = `
      <svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
        <foreignObject width="100%" height="100%">${serialized}</foreignObject>
      </svg>
    `;
    const svgUrl = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml;charset=utf-8" }));

    try {
      const image = await loadImageElement(svgUrl);
      const scale = Math.min(window.devicePixelRatio || 1, 2);
      const canvas = document.createElement("canvas");
      canvas.width = Math.max(1, Math.round(width * scale));
      canvas.height = Math.max(1, Math.round(height * scale));
      const ctx = canvas.getContext("2d");
      if (!ctx) {
        return undefined;
      }

      ctx.scale(scale, scale);
      ctx.fillStyle = "#FFFFFF";
      ctx.fillRect(0, 0, width, height);
      ctx.drawImage(image, 0, 0, width, height);

      const dataUrl = canvas.toDataURL("image/png");
      const parsedImage = stripDataUrlPrefix(dataUrl);
      return {
        kind: "selection",
        label: params.label,
        data: parsedImage.data,
        mimeType: parsedImage.mimeType,
        width,
        height,
      };
    } finally {
      URL.revokeObjectURL(svgUrl);
    }
  } catch {
    return undefined;
  } finally {
    measurementHost.remove();
  }
}

function createSummary(payload: OfficeContextPayload): string {
  const lines = [
    `${HOST_LABELS[payload.state.host]} context for "${payload.state.document.title}"`,
    `Selection: ${payload.state.selection.label}`,
    payload.state.selection.textPreview ? `Preview: ${payload.state.selection.textPreview}` : undefined,
    payload.visuals?.length ? `${payload.visuals.length} visual snapshot(s) attached.` : undefined,
  ];
  return lines.filter(Boolean).join("\n");
}

function serializeOfficeRuntimeError(error: unknown): Record<string, unknown> {
  if (!error || typeof error !== "object") {
    return {
      message: error instanceof Error ? error.message : String(error),
    };
  }

  const record = error as Record<string, unknown>;
  return {
    name: typeof record.name === "string" ? record.name : undefined,
    message:
      typeof record.message === "string"
        ? record.message
        : error instanceof Error
          ? error.message
          : String(error),
    code: typeof record.code === "string" ? record.code : undefined,
    traceMessages: record.traceMessages,
    debugInfo: record.debugInfo,
    stack: typeof record.stack === "string" ? record.stack : error instanceof Error ? error.stack : undefined,
  };
}

async function collectWordContext(base: OfficeStateUpdate, options: OfficeCaptureOptions): Promise<OfficeContextPayload> {
  const maxImages = Math.max(0, Math.min(options.maxImages ?? 0, 4));
  const payload = await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const body = context.document.body;
    const font = selection.font;
    const paragraphs = selection.paragraphs;
    const bodyParagraphs = body.paragraphs;
    const inlinePictures = selection.inlinePictures;
    const supportsShapes = supportsRequirementSet("WordApiDesktop", "1.2");
    const supportsViewportPages = supportsRequirementSet("WordApiDesktop", "1.2");
    const supportsPageSetup = supportsRequirementSet("WordApiDesktop", "1.3");
    const supportsWindowMetadata = supportsRequirementSet("WordApiDesktop", "1.4");
    const supportsComments = supportsRequirementSet("WordApi", "1.4");
    const supportsNotes = supportsRequirementSet("WordApi", "1.5");
    const supportsTrackedChanges = supportsRequirementSet("WordApi", "1.6");
    const supportsParagraphIds = supportsRequirementSet("WordApi", "1.6");
    const supportsReviewedText = supportsRequirementSet("WordApi", "1.4");
    const supportsFields = supportsRequirementSet("WordApi", "1.4");
    const supportsFieldMetadata = supportsRequirementSet("WordApi", "1.5");
    const supportsContentControls = supportsRequirementSet("WordApi", "1.1");
    const supportsContentControlSubtypes = supportsRequirementSet("WordApi", "1.3");
    const shapes = supportsShapes ? selection.shapes : undefined;
    const comments = supportsComments ? selection.getComments() : undefined;
    const footnotes = supportsNotes ? body.footnotes : undefined;
    const endnotes = supportsNotes ? body.endnotes : undefined;
    const trackedChanges = supportsTrackedChanges ? selection.getTrackedChanges() : undefined;
    const documentComments = supportsComments ? body.getComments() : undefined;
    const documentTrackedChanges = supportsTrackedChanges ? body.getTrackedChanges() : undefined;
    const selectionFields = supportsFields ? selection.fields : undefined;
    const documentFields = supportsFields ? body.fields : undefined;
    const selectionContentControls = supportsContentControls ? selection.contentControls : undefined;
    const documentContentControls = supportsContentControls ? body.contentControls : undefined;
    const pageSetup = supportsPageSetup ? context.document.pageSetup : undefined;
    const activeWindow = supportsViewportPages ? context.document.activeWindow : undefined;
    const activePane = activeWindow?.activePane;
    const viewportPages = activePane?.pagesEnclosingViewport;
    const selectionPages = supportsViewportPages ? selection.pages : undefined;
    const view = supportsWindowMetadata ? activeWindow?.view : undefined;
    const reviewedSelectionCurrent = supportsReviewedText ? selection.getReviewedText("Current") : undefined;
    const reviewedSelectionOriginal = supportsReviewedText ? selection.getReviewedText("Original") : undefined;
    const reviewedDocumentCurrent = supportsReviewedText ? body.getReviewedText("Current") : undefined;
    const reviewedDocumentOriginal = supportsReviewedText ? body.getReviewedText("Original") : undefined;

    selection.load("text");
    font.load(["name", "size", "color", "bold", "italic", "underline", "underlineColor", "highlightColor"]);
    paragraphs.load(
      supportsParagraphIds
        ? "items/text,items/uniqueLocalId,items/style,items/styleBuiltIn,items/alignment,items/leftIndent,items/rightIndent,items/firstLineIndent,items/lineSpacing,items/spaceBefore,items/spaceAfter"
        : "items/text,items/style,items/styleBuiltIn,items/alignment,items/leftIndent,items/rightIndent,items/firstLineIndent,items/lineSpacing,items/spaceBefore,items/spaceAfter",
    );
    bodyParagraphs.load(
      supportsParagraphIds
        ? "items/text,items/uniqueLocalId,items/style,items/styleBuiltIn"
        : "items/text,items/style,items/styleBuiltIn",
    );
    inlinePictures.load("items/altTextTitle,items/altTextDescription,items/width,items/height,items/imageFormat");
    shapes?.load("items/type,items/name,items/left,items/top,items/width,items/height,items/rotation,items/altTextDescription");
    comments?.load("items/id,items/authorName,items/content,items/resolved,items/creationDate");
    footnotes?.load("items/type,items/body/text,items/reference/text");
    endnotes?.load("items/type,items/body/text,items/reference/text");
    trackedChanges?.load("items/author,items/date,items/text,items/type");
    documentComments?.load("items/id,items/authorName,items/content,items/resolved,items/creationDate");
    documentTrackedChanges?.load("items/author,items/date,items/text,items/type");
    selectionFields?.load(
      supportsFieldMetadata ? "items/code,items/type,items/locked,items/result/text" : "items/code,items/result/text",
    );
    documentFields?.load(
      supportsFieldMetadata ? "items/code,items/type,items/locked,items/result/text" : "items/code,items/result/text",
    );
    selectionContentControls?.load(
      supportsContentControlSubtypes
        ? "items/id,items/title,items/tag,items/type,items/subtype,items/appearance,items/cannotDelete,items/cannotEdit,items/removeWhenEdited,items/placeholderText,items/text"
        : "items/id,items/title,items/tag,items/type,items/appearance,items/cannotDelete,items/cannotEdit,items/removeWhenEdited,items/placeholderText,items/text",
    );
    documentContentControls?.load(
      supportsContentControlSubtypes
        ? "items/id,items/title,items/tag,items/type,items/subtype,items/appearance,items/cannotDelete,items/cannotEdit,items/removeWhenEdited,items/placeholderText,items/text"
        : "items/id,items/title,items/tag,items/type,items/appearance,items/cannotDelete,items/cannotEdit,items/removeWhenEdited,items/placeholderText,items/text",
    );
    pageSetup?.load("topMargin,bottomMargin,leftMargin,rightMargin,pageWidth,pageHeight");
    viewportPages?.load("items/index,items/width,items/height");
    selectionPages?.load("items/index,items/width,items/height");
    if (supportsWindowMetadata) {
      activeWindow?.load([
        "caption",
        "height",
        "width",
        "usableHeight",
        "usableWidth",
        "horizontalPercentScrolled",
        "verticalPercentScrolled",
        "areRulersDisplayed",
        "areThumbnailsDisplayed",
        "isSplit",
        "isHorizontalScrollBarDisplayed",
        "isVerticalScrollBarDisplayed",
        "isVisible",
        "left",
        "top",
        "windowNumber",
      ]);
    }
    view?.load("type,seekView");
    await context.sync();

    const inlineImages = inlinePictures.items.slice(0, maxImages);
    const base64Results = inlineImages.map((picture) => picture.getBase64ImageSrc());
    if (base64Results.length) {
      await context.sync();
    }

    const readSelectionSnippet = async (
      kind: "html" | "ooxml",
      getter: () => OfficeExtension.ClientResult<string>,
      maxLength: number,
    ): Promise<string | undefined> => {
      try {
        const result = getter();
        await context.sync();
        return truncateText(result.value, maxLength);
      } catch (error) {
        console.warn(`[office-word] Failed to read selection ${kind}.`, serializeOfficeRuntimeError(error));
        return undefined;
      }
    };

    const hasSelectedContent = Boolean(selection.text?.trim()) || inlinePictures.items.length > 0;
    const html = hasSelectedContent ? await readSelectionSnippet("html", () => selection.getHtml(), 2000) : undefined;
    const ooxml = hasSelectedContent ? await readSelectionSnippet("ooxml", () => selection.getOoxml(), 3000) : undefined;

    const preview = normalizeTextPreview(selection.text);
    const details = paragraphs.items.map((paragraph) => paragraph.text.trim()).filter(Boolean).slice(0, 2);
    const imageShapes = shapes?.items.filter((shape) => isWordPictureShape(shape.type)) ?? [];
    const imageCount = countSelectedWordImages(inlinePictures.items.length, imageShapes.length);
    const objectCount = shapes?.items.length ?? 0;
    const altPreview =
      inlinePictures.items
        .map((picture) => picture.altTextTitle || picture.altTextDescription)
        .find((value) => Boolean(value?.trim())) ??
      imageShapes
        .map((shape) => shape.altTextDescription || shape.name)
        .find((value) => Boolean(value?.trim()));
    const contextStructuredPreview = buildStructuredSelectionPreview(paragraphs.items);
    const contextSelectionMeta = paragraphs.items.length > 0
      ? buildSelectionMeta(paragraphs.items, { name: font.name, size: font.size, color: font.color, bold: font.bold, italic: font.italic })
      : undefined;
    const state: OfficeStateUpdate = {
      ...base,
      selection: {
        ...buildSelectionSummary({
          textPreview: preview,
          imageCount,
          objectCount: imageCount > 0 ? Math.max(objectCount - imageShapes.length, 0) : objectCount,
          altPreview: normalizeTextPreview(altPreview),
          details: uniqueDetails([
            ...details,
            imageCount > 0 ? pluralize(imageCount, "image") : undefined,
            objectCount > imageShapes.length ? pluralize(objectCount - imageShapes.length, "shape") : undefined,
          ]),
          emptyLabel: "Insertion point",
        }),
        structuredPreview: contextStructuredPreview,
        selectionMeta: contextSelectionMeta,
      },
      capabilities: [
        "word.selection",
        "word.document",
        "word.inlinePictures",
        ...(supportsShapes ? ["word.shapes"] : []),
        ...(supportsPageSetup ? ["word.pageSetup"] : []),
        ...(supportsViewportPages ? ["word.viewportPages", "word.selectionPages"] : []),
        ...(supportsWindowMetadata ? ["word.activeWindow", "word.view"] : []),
        ...(supportsComments ? ["word.comments"] : []),
        ...(supportsNotes ? ["word.footnotes", "word.endnotes"] : []),
        ...(supportsTrackedChanges ? ["word.trackedChanges"] : []),
        ...(supportsReviewedText ? ["word.reviewedText"] : []),
        ...(supportsFields ? ["word.fields"] : []),
        ...(supportsContentControls ? ["word.contentControls"] : []),
        "word.insertText",
        "word.insertHtml",
        "word.insertOoxml",
        "word.insertFileFromBase64",
        ...(supportsContentControls ? ["word.insertContentControl"] : []),
        ...(supportsFields && supportsFieldMetadata ? ["word.insertField"] : []),
        ...(supportsComments ? ["word.commentThreads"] : []),
        ...(supportsTrackedChanges ? ["word.revisionActions"] : []),
        "office.getSelectedData",
      ],
    };

    const headingParagraphs = bodyParagraphs.items.filter((paragraph) => /heading/i.test(String(paragraph.styleBuiltIn || paragraph.style || "")));
    const reviewComments = documentComments?.items ?? comments?.items ?? [];
    const reviewChanges = documentTrackedChanges?.items ?? trackedChanges?.items ?? [];
    const documentFootnotes = footnotes?.items ?? [];
    const documentEndnotes = endnotes?.items ?? [];
    const documentFieldsList = documentFields?.items ?? selectionFields?.items ?? [];
    const selectedFields = selectionFields?.items ?? [];
    const documentContentControlList = documentContentControls?.items ?? selectionContentControls?.items ?? [];
    const selectedContentControls = selectionContentControls?.items ?? [];
    const paragraphMap = bodyParagraphs.items.slice(0, 40).map((paragraph) => ({
      text: truncateLabel(paragraph.text, 180),
      paragraphId: supportsParagraphIds ? paragraph.uniqueLocalId : undefined,
      style: paragraph.style || paragraph.styleBuiltIn,
      isHeading: /heading/i.test(String(paragraph.styleBuiltIn || paragraph.style || "")),
    }));
    const reviewedText = supportsReviewedText
      ? {
          selectionCurrentPreview: truncateText(reviewedSelectionCurrent?.value, 320),
          selectionOriginalPreview: truncateText(reviewedSelectionOriginal?.value, 320),
          documentCurrentPreview: truncateText(reviewedDocumentCurrent?.value, 480),
          documentOriginalPreview: truncateText(reviewedDocumentOriginal?.value, 480),
          selectionHasTrackedDifference:
            normalizeTextPreview(reviewedSelectionCurrent?.value) !== normalizeTextPreview(reviewedSelectionOriginal?.value),
          documentHasTrackedDifference:
            normalizeTextPreview(reviewedDocumentCurrent?.value) !== normalizeTextPreview(reviewedDocumentOriginal?.value),
        }
      : undefined;

    let visuals = await Promise.all(
      inlineImages.map(async (picture, index) =>
        optimizeVisual(
          {
            kind: "inline-picture",
            label: picture.altTextTitle || picture.altTextDescription || `Selected inline picture ${index + 1}`,
            data: base64Results[index]?.value ?? "",
            mimeType: String(picture.imageFormat ?? "png").toLowerCase() === "jpeg" ? "image/jpeg" : "image/png",
            width: picture.width,
            height: picture.height,
          },
          { maxDimension: 1400 },
        ),
      ),
    );

    if (!visuals.length && maxImages > 0) {
      const renderedSelection = await renderHtmlSelectionSnapshot({
        html: html ?? "",
        label: "Rendered selection snapshot",
        widthPx: getWordSnapshotWidthPx(pageSetup?.pageWidth, pageSetup?.leftMargin, pageSetup?.rightMargin),
      });
      if (renderedSelection) {
        visuals = [await optimizeVisual(renderedSelection, { maxDimension: 1600 })];
      }
    }

    const formatting = options.includeFormatting
      ? {
          selectionFont: {
            name: font.name,
            size: font.size,
            color: font.color,
            bold: font.bold,
            italic: font.italic,
            underline: font.underline,
            underlineColor: font.underlineColor,
            highlightColor: font.highlightColor,
          },
          pageSetup: pageSetup
            ? {
                pageWidth: formatPoints(pageSetup.pageWidth),
                pageHeight: formatPoints(pageSetup.pageHeight),
                topMargin: formatPoints(pageSetup.topMargin),
                bottomMargin: formatPoints(pageSetup.bottomMargin),
                leftMargin: formatPoints(pageSetup.leftMargin),
                rightMargin: formatPoints(pageSetup.rightMargin),
              }
            : undefined,
          viewport:
            supportsViewportPages && activeWindow
              ? {
                  window: supportsWindowMetadata
                    ? {
                        caption: activeWindow.caption,
                        width: formatPoints(activeWindow.width),
                        height: formatPoints(activeWindow.height),
                        usableWidth: formatPoints(activeWindow.usableWidth),
                        usableHeight: formatPoints(activeWindow.usableHeight),
                        left: formatPoints(activeWindow.left),
                        top: formatPoints(activeWindow.top),
                        horizontalPercentScrolled: formatPercent(activeWindow.horizontalPercentScrolled),
                        verticalPercentScrolled: formatPercent(activeWindow.verticalPercentScrolled),
                        areRulersDisplayed: activeWindow.areRulersDisplayed,
                        areThumbnailsDisplayed: activeWindow.areThumbnailsDisplayed,
                        isSplit: activeWindow.isSplit,
                        isHorizontalScrollBarDisplayed: activeWindow.isHorizontalScrollBarDisplayed,
                        isVerticalScrollBarDisplayed: activeWindow.isVerticalScrollBarDisplayed,
                        isVisible: activeWindow.isVisible,
                        windowNumber: activeWindow.windowNumber,
                      }
                    : undefined,
                  view: supportsWindowMetadata && view ? { type: view.type, seekView: view.seekView } : undefined,
                  pagesEnclosingViewport: viewportPages ? serializeWordPages(viewportPages.items) : [],
                  selectionPages: selectionPages ? serializeWordPages(selectionPages.items) : [],
                }
              : undefined,
          paragraphs: paragraphs.items.slice(0, 3).map((paragraph) => ({
            text: truncateText(paragraph.text, 280),
            style: paragraph.style || paragraph.styleBuiltIn,
            alignment: paragraph.alignment,
            firstLineIndent: formatPoints(paragraph.firstLineIndent),
            leftIndent: formatPoints(paragraph.leftIndent),
            rightIndent: formatPoints(paragraph.rightIndent),
            lineSpacing: formatPoints(paragraph.lineSpacing),
            spaceBefore: formatPoints(paragraph.spaceBefore),
            spaceAfter: formatPoints(paragraph.spaceAfter),
          })),
          selectedShapes: shapes?.items.slice(0, 4).map((shape) => ({
            name: shape.name,
            type: shape.type,
            left: formatPoints(shape.left),
            top: formatPoints(shape.top),
            width: formatPoints(shape.width),
            height: formatPoints(shape.height),
            rotation: shape.rotation,
            altTextDescription: shape.altTextDescription,
          })),
        }
      : undefined;

    const result: OfficeContextPayload = {
      summary: "",
      state,
      anchors: uniqueAnchors([
        ...headingParagraphs.slice(0, 20).map((paragraph) => ({
          kind: "heading",
          label: truncateLabel(paragraph.text),
          text: truncateLabel(paragraph.text, 240),
          paragraphId: supportsParagraphIds ? paragraph.uniqueLocalId : undefined,
        }) as OfficeAnchor),
        ...bodyParagraphs.items.slice(0, 24).map((paragraph) => {
          const style = String(paragraph.styleBuiltIn || paragraph.style || "");
          const kind = /heading/i.test(style) ? "heading" : "paragraph";
          return {
            kind: kind as OfficeAnchor["kind"],
            label: truncateLabel(paragraph.text),
            text: truncateLabel(paragraph.text, 240),
            paragraphId: supportsParagraphIds ? paragraph.uniqueLocalId : undefined,
          } as OfficeAnchor;
        }),
        ...(reviewComments.slice(0, 12).map((comment) => ({
          kind: "comment",
          label: truncateLabel(comment.content) || `Comment by ${comment.authorName || "unknown author"}`,
          text: truncateLabel(comment.content, 240),
          commentId: comment.id,
        }) as OfficeAnchor) ?? []),
        ...(reviewChanges.slice(0, 12).map((change, index) => ({
          kind: "revision",
          id: `revision:${index + 1}`,
          label: truncateLabel(change.text) || `${change.type} revision`,
          text: truncateLabel(change.text, 240),
          revisionId: `revision:${index + 1}`,
        }) as OfficeAnchor) ?? []),
        ...documentFootnotes.slice(0, 12).map((note, index) => ({
          kind: "footnote",
          id: `footnote:${index + 1}`,
          label: `Footnote ${index + 1}`,
          text: truncateLabel(note.body.text || note.reference.text, 240),
        }) as OfficeAnchor),
        ...documentEndnotes.slice(0, 12).map((note, index) => ({
          kind: "endnote",
          id: `endnote:${index + 1}`,
          label: `Endnote ${index + 1}`,
          text: truncateLabel(note.body.text || note.reference.text, 240),
        }) as OfficeAnchor),
        ...documentContentControlList.slice(0, 12).map((control) => ({
          kind: "contentControl",
          id: `contentControl:${control.id}`,
          label: truncateLabel(control.title || control.tag || control.text || `Content control ${control.id}`),
          text: truncateLabel(control.text, 240),
        }) as OfficeAnchor),
        ...documentFieldsList.slice(0, 12).map((field, index) => ({
          kind: "field",
          id: `field:${index + 1}`,
          label: truncateLabel(
            `${supportsFieldMetadata ? `${field.type || "Field"} · ` : ""}${field.code || field.result.text || `Field ${index + 1}`}`,
            120,
          ),
          text: truncateLabel(field.code || field.result.text, 240),
        }) as OfficeAnchor),
      ]),
      formatting,
      snippets: {
        html,
        ooxml,
        documentStructure: {
          paragraphs: bodyParagraphs.items.length,
          headings: headingParagraphs.length,
          comments: reviewComments.length,
          footnotes: documentFootnotes.length,
          endnotes: documentEndnotes.length,
          revisions: reviewChanges.length,
          fields: documentFieldsList.length,
          contentControls: documentContentControlList.length,
        },
        comments: reviewComments.slice(0, 8).map((comment) => ({
          id: comment.id,
          author: comment.authorName,
          content: truncateLabel(comment.content, 160),
          resolved: comment.resolved,
          creationDate: comment.creationDate?.toISOString?.(),
        })),
        trackedChanges: reviewChanges.slice(0, 8).map((change, index) => ({
          id: `revision:${index + 1}`,
          author: change.author,
          date: change.date?.toISOString?.(),
          type: change.type,
          text: truncateLabel(change.text, 160),
        })),
        headings: headingParagraphs.slice(0, 20).map((paragraph) => ({
          text: truncateLabel(paragraph.text, 160),
          paragraphId: supportsParagraphIds ? paragraph.uniqueLocalId : undefined,
        })),
        paragraphs: paragraphMap,
        footnotes: documentFootnotes.slice(0, 8).map((note, index) => ({
          id: `footnote:${index + 1}`,
          referenceText: truncateLabel(note.reference.text, 80),
          text: truncateLabel(note.body.text, 160),
        })),
        endnotes: documentEndnotes.slice(0, 8).map((note, index) => ({
          id: `endnote:${index + 1}`,
          referenceText: truncateLabel(note.reference.text, 80),
          text: truncateLabel(note.body.text, 160),
        })),
        fields: documentFieldsList.slice(0, 12).map((field, index) => ({
          id: `field:${index + 1}`,
          type: supportsFieldMetadata ? field.type : undefined,
          code: truncateLabel(field.code, 180),
          resultText: truncateLabel(field.result.text, 160),
          locked: supportsFieldMetadata ? field.locked : undefined,
        })),
        selectedFields: selectedFields.slice(0, 8).map((field, index) => ({
          id: `selectionField:${index + 1}`,
          type: supportsFieldMetadata ? field.type : undefined,
          code: truncateLabel(field.code, 160),
          resultText: truncateLabel(field.result.text, 120),
        })),
        contentControls: documentContentControlList.slice(0, 12).map((control) => ({
          id: `contentControl:${control.id}`,
          title: trimString(control.title),
          tag: trimString(control.tag),
          type: control.type,
          subtype: supportsContentControlSubtypes ? control.subtype : undefined,
          text: truncateLabel(control.text, 160),
          placeholderText: trimString(control.placeholderText),
          appearance: control.appearance,
          cannotDelete: control.cannotDelete,
          cannotEdit: control.cannotEdit,
          removeWhenEdited: control.removeWhenEdited,
        })),
        selectedContentControls: selectedContentControls.slice(0, 8).map((control) => ({
          id: `contentControl:${control.id}`,
          title: trimString(control.title),
          tag: trimString(control.tag),
          type: control.type,
          text: truncateLabel(control.text, 120),
        })),
        reviewedText,
      },
      visuals,
    };
    result.summary = createSummary(result);
    return result;
  });

  if (!payload.visuals?.length && maxImages > 0 && (payload.state.selection.imageCount ?? 0) > 0) {
    const fallback = await getSelectedImageAsync().catch(() => undefined);
    if (fallback) {
      payload.visuals = [await optimizeVisual(fallback, { maxDimension: 1400 })];
      payload.summary = createSummary(payload);
    }
  }

  return payload;
}

async function collectExcelContext(base: OfficeStateUpdate, options: OfficeCaptureOptions): Promise<OfficeContextPayload> {
  const payload = await Excel.run(async (context) => {
    const normalizedScope = trimString(options.scope)?.toLowerCase();
    const supportsWorksheetView = supportsRequirementSet("ExcelApi", "1.8");
    const supportsPageLayout = supportsRequirementSet("ExcelApi", "1.9");
    const range = context.workbook.getSelectedRange();
    const workbook = context.workbook;
    const worksheets = workbook.worksheets;
    const namedItems = workbook.names;
    const activeWorksheet = range.worksheet;
    const format = range.format;
    const font = format.font;
    const fill = format.fill;
    const tables = activeWorksheet.tables;
    const charts = activeWorksheet.charts;
    const pivotTables = activeWorksheet.pivotTables;
    const activePrintArea = supportsPageLayout ? activeWorksheet.pageLayout.getPrintAreaOrNullObject() : undefined;

    range.load(["address", "rowCount", "columnCount", "text", "formulas", "numberFormat"]);
    activeWorksheet.load("name,id");
    if (supportsWorksheetView) {
      activeWorksheet.load("showGridlines,showHeadings");
    }
    worksheets.load("items/name,items/id,items/position,items/visibility");
    namedItems.load("items/name,items/type");
    tables.load("items/name,items/id");
    charts.load("items/name,items/id");
    pivotTables.load("items/name,items/id");
    format.load(["horizontalAlignment", "verticalAlignment", "wrapText", "rowHeight", "columnWidth"]);
    font.load(["name", "size", "color", "bold", "italic", "underline"]);
    fill.load("color");
    activePrintArea?.load("isNullObject,address");
    await context.sync();

    const usedRanges = worksheets.items.map((worksheet) => worksheet.getUsedRangeOrNullObject(true));
    const worksheetTables = worksheets.items.map((worksheet) => worksheet.tables);
    const worksheetCharts = worksheets.items.map((worksheet) => worksheet.charts);
    const worksheetPivotTables = worksheets.items.map((worksheet) => worksheet.pivotTables);
    const worksheetPrintAreas = supportsPageLayout ? worksheets.items.map((worksheet) => worksheet.pageLayout.getPrintAreaOrNullObject()) : [];
    for (const usedRange of usedRanges) {
      usedRange.load("isNullObject,address,rowCount,columnCount,rowIndex,columnIndex");
    }
    for (const worksheet of worksheets.items) {
      if (supportsWorksheetView) {
        worksheet.load("showGridlines,showHeadings");
      }
    }
    for (const tableCollection of worksheetTables) {
      tableCollection.load("items/name,items/id");
    }
    for (const chartCollection of worksheetCharts) {
      chartCollection.load("items/name,items/id");
    }
    for (const pivotCollection of worksheetPivotTables) {
      pivotCollection.load("items/name,items/id");
    }
    for (const printArea of worksheetPrintAreas) {
      printArea.load("isNullObject,address");
    }
    await context.sync();

    const selectionCellRanges: Excel.Range[] = [];
    const maxSelectionCells = normalizedScope === "selection" ? 16 : 12;
    let selectionCellBudget = maxSelectionCells;
    for (let rowIndex = 0; rowIndex < range.rowCount && selectionCellBudget > 0; rowIndex += 1) {
      for (let columnIndex = 0; columnIndex < range.columnCount && selectionCellBudget > 0; columnIndex += 1) {
        const cell = range.getCell(rowIndex, columnIndex);
        cell.load(["address", "text", "formulas", "numberFormat"]);
        selectionCellRanges.push(cell);
        selectionCellBudget -= 1;
      }
    }

    const activeWorksheetIndex = worksheets.items.findIndex((worksheet) => worksheet.id === activeWorksheet.id || worksheet.name === activeWorksheet.name);
    const snapshotIndexes = new Set<number>();
    if (activeWorksheetIndex >= 0) {
      snapshotIndexes.add(activeWorksheetIndex);
    }
    if (normalizedScope !== "selection" && normalizedScope !== "worksheet") {
      for (let index = 0; index < worksheets.items.length && snapshotIndexes.size < 6; index += 1) {
        const usedRange = usedRanges[index];
        if (index !== activeWorksheetIndex && usedRange && !usedRange.isNullObject && usedRange.address) {
          snapshotIndexes.add(index);
        }
      }
    }

    const worksheetSnapshotRanges: Array<{
      worksheetIndex: number;
      previewRange: Excel.Range;
      startRowIndex: number;
      startColumnIndex: number;
      rowCount: number;
      columnCount: number;
    }> = [];
    for (const worksheetIndex of snapshotIndexes) {
      const usedRange = usedRanges[worksheetIndex];
      const worksheet = worksheets.items[worksheetIndex];
      if (!worksheet || !usedRange || usedRange.isNullObject) {
        continue;
      }

      const previewRowCount = Math.min(usedRange.rowCount, 3);
      const previewColumnCount = Math.min(usedRange.columnCount, 4);
      if (previewRowCount < 1 || previewColumnCount < 1) {
        continue;
      }

      const previewRange = worksheet.getRangeByIndexes(
        usedRange.rowIndex,
        usedRange.columnIndex,
        previewRowCount,
        previewColumnCount,
      );
      previewRange.load(["address", "text", "formulas", "numberFormat"]);
      worksheetSnapshotRanges.push({
        worksheetIndex,
        previewRange,
        startRowIndex: usedRange.rowIndex,
        startColumnIndex: usedRange.columnIndex,
        rowCount: previewRowCount,
        columnCount: previewColumnCount,
      });
    }
    await context.sync();

    const worksheetSummaries = worksheets.items.slice(0, 24).map((worksheet, index) => ({
      id: worksheet.id,
      name: worksheet.name,
      position: worksheet.position + 1,
      visibility: worksheet.visibility,
      usedRange:
        usedRanges[index] && !usedRanges[index].isNullObject
          ? {
              address: usedRanges[index].address,
              rowCount: usedRanges[index].rowCount,
              columnCount: usedRanges[index].columnCount,
            }
          : undefined,
      showGridlines: supportsWorksheetView ? worksheet.showGridlines : undefined,
      showHeadings: supportsWorksheetView ? worksheet.showHeadings : undefined,
      printArea:
        supportsPageLayout && worksheetPrintAreas[index] && !worksheetPrintAreas[index].isNullObject
          ? worksheetPrintAreas[index].address
          : undefined,
      tableCount: worksheetTables[index]?.items.length ?? 0,
      chartCount: worksheetCharts[index]?.items.length ?? 0,
      pivotTableCount: worksheetPivotTables[index]?.items.length ?? 0,
      tables: worksheetTables[index]?.items.slice(0, 6).map((table) => table.name) ?? [],
      charts: worksheetCharts[index]?.items.slice(0, 6).map((chart) => chart.name) ?? [],
      pivotTables: worksheetPivotTables[index]?.items.slice(0, 6).map((pivotTable) => pivotTable.name) ?? [],
    }));
    const workbookTables = worksheets.items.flatMap((worksheet, index) =>
      (worksheetTables[index]?.items ?? []).map((table) => ({
        sheetName: worksheet.name,
        name: table.name,
        id: table.id,
      })),
    );
    const workbookCharts = worksheets.items.flatMap((worksheet, index) =>
      (worksheetCharts[index]?.items ?? []).map((chart) => ({
        sheetName: worksheet.name,
        name: chart.name,
        id: chart.id,
      })),
    );
    const workbookPivotTables = worksheets.items.flatMap((worksheet, index) =>
      (worksheetPivotTables[index]?.items ?? []).map((pivotTable) => ({
        sheetName: worksheet.name,
        name: pivotTable.name,
        id: pivotTable.id,
      })),
    );

    const preview = range.text.flat().join(" | ").trim();
    const selectionRangeCitation = buildExcelCitationRecord({
      kind: isSingleCellAddress(range.address) ? "cell" : "range",
      sheetName: activeWorksheet.name,
      address: range.address,
      text: preview || undefined,
      formula: range.rowCount === 1 && range.columnCount === 1 ? trimString(range.formulas[0]?.[0]) : undefined,
      numberFormat: range.rowCount === 1 && range.columnCount === 1 ? trimString(range.numberFormat[0]?.[0]) : undefined,
    });
    const selectionCellCitations = selectionCellRanges.map((cell) =>
      buildExcelCitationRecord({
        kind: "cell",
        sheetName: activeWorksheet.name,
        address: cell.address,
        text: firstMatrixString(cell.text),
        formula: firstMatrixString(cell.formulas),
        numberFormat: firstMatrixString(cell.numberFormat),
      }),
    );
    const workbookSheetSnapshots: ExcelWorksheetSnapshot[] = worksheetSnapshotRanges.flatMap((snapshot) => {
      const worksheetSummary = worksheetSummaries[snapshot.worksheetIndex];
      if (!worksheetSummary) {
        return [];
      }
      const previewRows = snapshot.previewRange.text
        .slice(0, 3)
        .map((row) => row.map((value) => trimString(value) ?? "").join(" | "))
        .filter(Boolean);
      const citedCells: ExcelCitationRecord[] = [];
      for (let rowOffset = 0; rowOffset < snapshot.rowCount && citedCells.length < 6; rowOffset += 1) {
        for (let columnOffset = 0; columnOffset < snapshot.columnCount && citedCells.length < 6; columnOffset += 1) {
          citedCells.push(
            buildExcelCitationRecord({
              kind: "cell",
              sheetName: worksheetSummary.name,
              address: excelCellAddress(snapshot.startRowIndex + rowOffset, snapshot.startColumnIndex + columnOffset),
              text: snapshot.previewRange.text[rowOffset]?.[columnOffset],
              formula: trimString(snapshot.previewRange.formulas[rowOffset]?.[columnOffset]),
              numberFormat: trimString(snapshot.previewRange.numberFormat[rowOffset]?.[columnOffset]),
            }),
          );
        }
      }
      return [{
        worksheetId: worksheetSummary.id,
        worksheetName: worksheetSummary.name,
        usedRange: worksheetSummary.usedRange,
        previewAddress: snapshot.previewRange.address,
        previewRows,
        citedCells,
      }];
    });
    const activeWorksheetSnapshot =
      workbookSheetSnapshots.find((snapshot) => snapshot.worksheetId === activeWorksheet.id) ?? workbookSheetSnapshots[0];

    const state: OfficeStateUpdate = {
      ...base,
      selection: {
        label: `${activeWorksheet.name}!${range.address}`,
        kind: preview ? "text" : "empty",
        imageCount: 0,
        objectCount: 0,
        textPreview: preview || undefined,
        details: [`${range.rowCount} rows`, `${range.columnCount} columns`],
      },
      capabilities: [
        "excel.workbook",
        "excel.range",
        "excel.values",
        "excel.formulas",
        "excel.format",
        "excel.citations",
        "excel.sheetSnapshots",
        ...(supportsWorksheetView ? ["excel.worksheetView"] : []),
        ...(supportsPageLayout ? ["excel.pageLayout"] : []),
      ],
    };

    const result: OfficeContextPayload = {
      summary: "",
      state,
      anchors: uniqueAnchors([
        {
          kind: "workbook",
          label: base.document.title || "Workbook",
        } as OfficeAnchor,
        {
          kind: selectionRangeCitation.anchor.kind,
          label: selectionRangeCitation.label,
          sheetName: activeWorksheet.name,
          address: range.address,
        } as OfficeAnchor,
        ...selectionCellCitations.map((citation) => citation.anchor as OfficeAnchor),
        ...worksheets.items.slice(0, 24).map((worksheet) => ({
          kind: "sheet",
          label: worksheet.name,
          sheetName: worksheet.name,
        }) as OfficeAnchor),
        ...worksheetSummaries
          .filter((worksheet) => Boolean(worksheet.usedRange?.address))
          .map((worksheet) => ({
            kind: "range",
            label: `${worksheet.name}!${worksheet.usedRange!.address}`,
            sheetName: worksheet.name,
            address: worksheet.usedRange!.address,
          }) as OfficeAnchor),
        ...namedItems.items.slice(0, 24).map((namedItem) => ({
          kind: "namedItem",
          label: namedItem.name,
          namedItemName: namedItem.name,
        }) as OfficeAnchor),
        ...workbookTables.slice(0, 24).map((table) => ({
          kind: "table",
          label: table.name,
          sheetName: table.sheetName,
          tableName: table.name,
          id: table.id,
        }) as OfficeAnchor),
        ...workbookCharts.slice(0, 24).map((chart) => ({
          kind: "chart",
          label: chart.name,
          sheetName: chart.sheetName,
          chartName: chart.name,
          id: chart.id,
        }) as OfficeAnchor),
        ...workbookPivotTables.slice(0, 24).map((pivotTable) => ({
          kind: "pivotTable",
          label: pivotTable.name,
          sheetName: pivotTable.sheetName,
          pivotTableName: pivotTable.name,
          id: pivotTable.id,
        }) as OfficeAnchor),
        ...workbookSheetSnapshots.flatMap((snapshot) => snapshot.citedCells.slice(0, 4).map((citation) => citation.anchor as OfficeAnchor)).slice(0, 24),
      ]),
      formatting: options.includeFormatting
        ? {
            selectionFormat: {
              fontName: font.name,
              fontSize: font.size,
              fontColor: font.color,
              bold: font.bold,
              italic: font.italic,
              underline: font.underline,
              fillColor: fill.color,
              horizontalAlignment: format.horizontalAlignment,
              verticalAlignment: format.verticalAlignment,
              wrapText: format.wrapText,
              rowHeight: format.rowHeight,
              columnWidth: format.columnWidth,
            },
            worksheetView: {
              worksheetId: activeWorksheet.id,
              showGridlines: supportsWorksheetView ? activeWorksheet.showGridlines : undefined,
              showHeadings: supportsWorksheetView ? activeWorksheet.showHeadings : undefined,
              printArea: supportsPageLayout && activePrintArea && !activePrintArea.isNullObject ? activePrintArea.address : undefined,
            },
          }
        : undefined,
      snippets: {
        documentStructure: {
          worksheets: worksheets.items.length,
          namedItems: namedItems.items.length,
          workbookTables: workbookTables.length,
          workbookCharts: workbookCharts.length,
          workbookPivotTables: workbookPivotTables.length,
          activeSheetTables: tables.items.length,
          activeSheetCharts: charts.items.length,
          activeSheetPivotTables: pivotTables.items.length,
        },
        workbookSheets: worksheetSummaries,
        workbookSheetSnapshots,
        activeWorksheetSnapshot,
        namedItems: namedItems.items.slice(0, 24).map((namedItem) => ({
          name: namedItem.name,
          type: namedItem.type,
        })),
        selectionRangeCitation,
        selectionCellCitations,
        textPreviewRows: range.text.slice(0, 4).map((row) => row.join(" | ")),
        formulas: range.formulas.slice(0, 4),
        numberFormats: range.numberFormat.slice(0, 4),
        workbookObjects: {
          tables: workbookTables.slice(0, 24),
          charts: workbookCharts.slice(0, 24),
          pivotTables: workbookPivotTables.slice(0, 24),
        },
        activeWorksheetObjects: {
          worksheetId: activeWorksheet.id,
          worksheetName: activeWorksheet.name,
          tables: tables.items.map((table) => table.name).slice(0, 12),
          charts: charts.items.map((chart) => chart.name).slice(0, 12),
          pivotTables: pivotTables.items.map((pivotTable) => pivotTable.name).slice(0, 12),
        },
        activeWorksheetView: {
          showGridlines: supportsWorksheetView ? activeWorksheet.showGridlines : undefined,
          showHeadings: supportsWorksheetView ? activeWorksheet.showHeadings : undefined,
          printArea: supportsPageLayout && activePrintArea && !activePrintArea.isNullObject ? activePrintArea.address : undefined,
        },
      },
    };
    result.summary = createSummary(result);
    return result;
  });

  if ((options.maxImages ?? 0) > 0) {
    const fallback = await getSelectedImageAsync().catch(() => undefined);
    if (fallback) {
      payload.visuals = [await optimizeVisual({ ...fallback, kind: "worksheet", label: "Worksheet selection snapshot" })];
      payload.summary = createSummary(payload);
    }
  }

  return payload;
}

async function collectPowerPointContext(base: OfficeStateUpdate, options: OfficeCaptureOptions): Promise<OfficeContextPayload> {
  const maxImages = Math.max(0, Math.min(options.maxImages ?? 0, 4));
  const payload = await PowerPoint.run(async (context) => {
    const presentation = context.presentation;
    const slides = presentation.getSelectedSlides();
    const presentationSlides = presentation.slides;
    const supportsRichSelection = supportsRequirementSet("PowerPointApi", "1.5");
    const supportsLayoutMetadata = supportsRequirementSet("PowerPointApi", "1.3");
    const supportsSlideSnapshots = supportsRequirementSet("PowerPointApi", "1.8");
    const supportsShapeMetadata = supportsRequirementSet("PowerPointApi", "1.8");
    const supportsTableShapes = supportsRequirementSet("PowerPointApi", "1.8");
    const supportsTableEditing = supportsRequirementSet("PowerPointApi", "1.9");
    const supportsPageSetup = supportsRequirementSet("PowerPointApi", "1.10");
    const supportsShapeSnapshots = supportsRequirementSet("PowerPointApi", "1.10");
    const slideMasters = supportsLayoutMetadata ? presentation.slideMasters : undefined;
    const shapes = supportsRichSelection ? presentation.getSelectedShapes() : undefined;
    const textRange = supportsRichSelection ? presentation.getSelectedTextRangeOrNullObject() : undefined;
    const pageSetup = supportsPageSetup ? presentation.pageSetup : undefined;

    slides.load("items/id,items/index");
    presentationSlides.load("items/id,items/index");
    slideMasters?.load("items/id,items/name");
    const shapeLoadProperties = [
      "items/id",
      "items/name",
      "items/type",
      "items/left",
      "items/top",
      "items/width",
      "items/height",
      "items/rotation",
      "items/altTextTitle",
      "items/altTextDescription",
      ...(supportsShapeMetadata ? ["items/zOrderPosition"] : []),
      ...(supportsShapeSnapshots ? ["items/visible", "items/creationId"] : []),
    ];
    shapes?.load(shapeLoadProperties.join(","));
    textRange?.load(
      "isNullObject,text,start,length,font/name,font/size,font/color,font/bold,font/italic,font/underline,paragraphFormat/horizontalAlignment,paragraphFormat/indentLevel",
    );
    pageSetup?.load("slideHeight,slideWidth");
    await context.sync();

    if (supportsLayoutMetadata) {
      for (const slide of presentationSlides.items) {
        slide.layout.load("id,name,type");
        slide.slideMaster.load("id,name");
      }
      for (const slideMaster of slideMasters?.items ?? []) {
        slideMaster.layouts.load("items/id,items/name,items/type");
      }
    }

    const selectedShapes = shapes?.items.slice(0, Math.max(maxImages, 4)) ?? [];
    const textFrames =
      supportsShapeSnapshots && selectedShapes.length
        ? selectedShapes.map((shape) => shape.getTextFrameOrNullObject())
        : [];
    const tables =
      supportsTableShapes && selectedShapes.length
        ? selectedShapes.map((shape) => (isPowerPointTableShape(shape.type) ? shape.getTable() : undefined))
        : [];
    for (const textFrame of textFrames) {
      textFrame.load(
        "isNullObject,hasText,autoSizeSetting,topMargin,leftMargin,rightMargin,bottomMargin,verticalAlignment,wordWrap,textRange/text,textRange/font/name,textRange/font/size,textRange/font/color,textRange/font/bold,textRange/font/italic,textRange/font/underline,textRange/paragraphFormat/horizontalAlignment,textRange/paragraphFormat/indentLevel",
      );
    }
    for (const table of tables) {
      table?.load("rowCount,columnCount,values");
      if (supportsTableEditing) {
        table?.styleSettings.load(
          "style,areColumnsBanded,areRowsBanded,isFirstColumnHighlighted,isFirstRowHighlighted,isLastColumnHighlighted,isLastRowHighlighted",
        );
      }
    }

    const slideImageResults =
      maxImages > 0 && supportsSlideSnapshots && slides.items.length > 0 && selectedShapes.length === 0
        ? slides.items.slice(0, maxImages).map((slide) => slide.getImageAsBase64({ width: 1400 }))
        : [];
    const shapeImageResults =
      maxImages > 0 && supportsShapeSnapshots && selectedShapes.length > 0
        ? selectedShapes.slice(0, maxImages).map((shape) => shape.getImageAsBase64({ format: "Png", width: 1400 }))
        : [];

    if (supportsLayoutMetadata || textFrames.length || tables.some(Boolean) || slideImageResults.length || shapeImageResults.length) {
      await context.sync();
    }

    const slideMetadataById = new Map(
      presentationSlides.items.map((slide) => [
        slide.id,
        {
          id: slide.id,
          index: slide.index + 1,
          layoutId: supportsLayoutMetadata ? slide.layout.id : undefined,
          layoutName: supportsLayoutMetadata ? slide.layout.name : undefined,
          layoutType: supportsLayoutMetadata ? slide.layout.type : undefined,
          slideMasterId: supportsLayoutMetadata ? slide.slideMaster.id : undefined,
          slideMasterName: supportsLayoutMetadata ? slide.slideMaster.name : undefined,
        },
      ]),
    );
    const slideMetadataByIndex = new Map(Array.from(slideMetadataById.values()).map((slide) => [slide.index, slide]));
    const selectedSlideSummaries = slides.items.map(
      (slide) =>
        slideMetadataById.get(slide.id) ?? {
          id: slide.id,
          index: slide.index + 1,
          layoutId: undefined,
          layoutName: undefined,
          layoutType: undefined,
          slideMasterId: undefined,
          slideMasterName: undefined,
        },
    );
    const selectedSlideNotesAnchors = selectedSlideSummaries.slice(0, 12).map(
      (slide) =>
        ({
          kind: "notesRegion",
          id: `notes:${slide.index}`,
          label: `Notes for Slide ${slide.index}`,
          slideId: slide.id,
          slideIndex: slide.index,
          text: `Speaker notes for slide ${slide.index}`,
        }) as OfficeAnchor,
    );
    const masterSummaries = (slideMasters?.items ?? []).map((slideMaster) => ({
      id: slideMaster.id,
      name: slideMaster.name,
      layoutCount: slideMaster.layouts.items.length,
      layouts: slideMaster.layouts.items.map((layout) => ({ id: layout.id, name: layout.name, type: layout.type })),
    }));
    const layoutUsageMap = new Map<
      string,
      {
        id: string | undefined;
        name: string | undefined;
        type: string | undefined;
        slideMasterId: string | undefined;
        slideMasterName: string | undefined;
        count: number;
      }
    >();
    for (const slide of slideMetadataById.values()) {
      const key = slide.layoutId ?? slide.layoutName ?? `${slide.index}`;
      const existing = layoutUsageMap.get(key);
      if (existing) {
        existing.count += 1;
        continue;
      }

      layoutUsageMap.set(key, {
        id: slide.layoutId,
        name: slide.layoutName,
        type: slide.layoutType,
        slideMasterId: slide.slideMasterId,
        slideMasterName: slide.slideMasterName,
        count: 1,
      });
    }
    const layoutUsage = Array.from(layoutUsageMap.values()).sort((left, right) => right.count - left.count);
    const selectedShapeDescriptors = selectedShapes.slice(0, 8).map((shape, index) => {
      const textFrame = textFrames[index];
      const table = tables[index];
      const slideSummary = slides.items[0] ? slideMetadataById.get(slides.items[0].id) : undefined;
      const textPreview =
        textFrame && !textFrame.isNullObject && textFrame.hasText ? truncateText(textFrame.textRange.text, 400) : undefined;

      return {
        id: shape.id,
        name: shape.name,
        label: shape.name || truncateLabel(shape.altTextTitle || shape.altTextDescription) || `Shape ${shape.id}`,
        type: shape.type,
        contentKind: getPowerPointShapeContentKind(shape.type),
        slideId: slideSummary?.id,
        slideIndex: slideSummary?.index,
        layoutId: slideSummary?.layoutId,
        layoutName: slideSummary?.layoutName,
        slideMasterId: slideSummary?.slideMasterId,
        slideMasterName: slideSummary?.slideMasterName,
        left: formatPoints(shape.left),
        top: formatPoints(shape.top),
        width: formatPoints(shape.width),
        height: formatPoints(shape.height),
        rotation: shape.rotation,
        visible: supportsShapeSnapshots ? shape.visible : undefined,
        zOrderPosition: supportsShapeMetadata ? shape.zOrderPosition : undefined,
        creationId: supportsShapeSnapshots ? shape.creationId : undefined,
        altTextTitle: shape.altTextTitle,
        altTextDescription: shape.altTextDescription,
        textPreview,
        textFrame:
          textFrame && !textFrame.isNullObject
            ? {
                hasText: textFrame.hasText,
                autoSizeSetting: textFrame.autoSizeSetting,
                topMargin: formatPoints(textFrame.topMargin),
                rightMargin: formatPoints(textFrame.rightMargin),
                bottomMargin: formatPoints(textFrame.bottomMargin),
                leftMargin: formatPoints(textFrame.leftMargin),
                verticalAlignment: textFrame.verticalAlignment,
                wordWrap: textFrame.wordWrap,
                fontName: textFrame.hasText ? textFrame.textRange.font.name : undefined,
                fontSize: textFrame.hasText ? textFrame.textRange.font.size : undefined,
                fontColor: textFrame.hasText ? textFrame.textRange.font.color : undefined,
                bold: textFrame.hasText ? textFrame.textRange.font.bold : undefined,
                italic: textFrame.hasText ? textFrame.textRange.font.italic : undefined,
                underline: textFrame.hasText ? textFrame.textRange.font.underline : undefined,
                horizontalAlignment: textFrame.hasText ? textFrame.textRange.paragraphFormat.horizontalAlignment : undefined,
                indentLevel: textFrame.hasText ? textFrame.textRange.paragraphFormat.indentLevel : undefined,
              }
            : undefined,
        table:
          table && isPowerPointTableShape(shape.type)
            ? {
                rowCount: table.rowCount,
                columnCount: table.columnCount,
                previewValues: truncateStringMatrix(table.values),
                styleSettings: supportsTableEditing
                  ? {
                      style: table.styleSettings.style,
                      areRowsBanded: table.styleSettings.areRowsBanded,
                      areColumnsBanded: table.styleSettings.areColumnsBanded,
                      isFirstRowHighlighted: table.styleSettings.isFirstRowHighlighted,
                      isFirstColumnHighlighted: table.styleSettings.isFirstColumnHighlighted,
                      isLastRowHighlighted: table.styleSettings.isLastRowHighlighted,
                      isLastColumnHighlighted: table.styleSettings.isLastColumnHighlighted,
                    }
                  : undefined,
              }
            : undefined,
      };
    });

    const preview =
      textRange && !textRange.isNullObject
        ? normalizeTextPreview(textRange.text)
        : normalizeTextPreview(await getSelectedTextAsync().catch(() => ""));
    const imageShapeCount = selectedShapes.filter((shape) => isPowerPointImageShape(shape.type)).length;
    const objectCount = selectedShapes.length;
    const details = uniqueDetails([
      ...selectedSlideSummaries
        .slice(0, 2)
        .map((slide) => (slide.layoutName ? `Slide ${slide.index} · ${slide.layoutName}` : `Slide ${slide.index}`)),
      objectCount > 0 ? pluralize(objectCount, "shape") : undefined,
      imageShapeCount > 0 ? pluralize(imageShapeCount, "image") : undefined,
    ]);
    const ctxPptShapeTypes = selectedShapes.map((s) => String(s.type ?? "shape"));
    const ctxPptStyleHistogram: Record<string, number> = {};
    for (const t of ctxPptShapeTypes) {
      const kind = getPowerPointShapeContentKind(t);
      ctxPptStyleHistogram[kind] = (ctxPptStyleHistogram[kind] ?? 0) + 1;
    }
    const ctxPptMeta: OfficeSelectionMeta | undefined = objectCount > 0 || preview
      ? {
          paragraphCount: preview ? 1 : undefined,
          firstParagraphStyle: ctxPptShapeTypes[0] ? getPowerPointShapeContentKind(ctxPptShapeTypes[0]) : undefined,
          styleHistogram: Object.keys(ctxPptStyleHistogram).length > 0 ? ctxPptStyleHistogram : undefined,
        }
      : undefined;
    const state: OfficeStateUpdate = {
      ...base,
      selection: {
        ...buildSelectionSummary({
          textPreview: preview,
          imageCount: imageShapeCount,
          objectCount: imageShapeCount > 0 ? Math.max(objectCount - imageShapeCount, 0) : objectCount,
          altPreview: selectedShapeDescriptors[0]?.textPreview ?? selectedShapes[0]?.name,
          details,
          emptyLabel: slides.items.length > 0 ? `${pluralize(slides.items.length, "slide")} selected` : "Slide selection",
        }),
        selectionMeta: ctxPptMeta,
      },
      capabilities: [
        "powerpoint.selection",
        ...(supportsRichSelection ? ["powerpoint.shapes", "powerpoint.textRange"] : []),
        ...(supportsLayoutMetadata ? ["powerpoint.slideMasters", "powerpoint.layouts"] : []),
        ...(supportsTableShapes ? ["powerpoint.tables"] : []),
        "powerpoint.notesRegionAnchors",
        ...(supportsSlideSnapshots ? ["powerpoint.slideSnapshot"] : []),
        ...(supportsShapeSnapshots ? ["powerpoint.shapeSnapshot"] : []),
        "office.setSelectedData",
      ],
    };

    const visuals: OfficeVisualSnapshot[] = [];
    for (let index = 0; index < shapeImageResults.length; index += 1) {
      const shape = selectedShapes[index];
      if (!shape) {
        continue;
      }

      visuals.push(
        await optimizeVisual(
          {
            kind: "shape",
            label: shape.name || `Selected shape ${index + 1}`,
            data: shapeImageResults[index]?.value ?? "",
            mimeType: "image/png",
            width: shape.width,
            height: shape.height,
          },
          { maxDimension: 1400 },
        ),
      );
    }

    for (let index = 0; index < slideImageResults.length; index += 1) {
      const slide = slides.items[index];
      if (!slide) {
        continue;
      }

      visuals.push(
        await optimizeVisual(
          {
            kind: "slide",
            label: `Slide ${slide.index + 1}`,
            data: slideImageResults[index]?.value ?? "",
            mimeType: "image/png",
          },
          { maxDimension: 1400 },
        ),
      );
    }

    const result: OfficeContextPayload = {
      summary: "",
      state,
      anchors: uniqueAnchors([
        ...Array.from(slideMetadataById.values()).slice(0, 30).map((slide) => ({
          kind: "slide",
          label: `Slide ${slide.index}`,
          slideId: slide.id,
          slideIndex: slide.index,
        }) as OfficeAnchor),
        ...selectedSlideNotesAnchors,
        ...masterSummaries.slice(0, 12).map((slideMaster) => ({
          kind: "slideMaster",
          id: slideMaster.id,
          label: slideMaster.name || `Master ${slideMaster.id}`,
        }) as OfficeAnchor),
        ...masterSummaries.slice(0, 12).flatMap((slideMaster) =>
          slideMaster.layouts.slice(0, 12).map((layout) => ({
            kind: "layout",
            id: layout.id,
            label: layout.name || layout.type || `Layout ${layout.id}`,
            text: slideMaster.name,
          }) as OfficeAnchor),
        ),
        ...selectedShapes.slice(0, 12).map((shape) => ({
          kind: "shape",
          label: shape.name || truncateLabel(shape.altTextTitle || shape.altTextDescription) || `Shape ${shape.id}`,
          slideId: slides.items[0]?.id,
          slideIndex: slides.items[0] ? slides.items[0].index + 1 : undefined,
          shapeId: shape.id,
          id: shape.id,
        }) as OfficeAnchor),
      ]),
      formatting: options.includeFormatting
        ? {
            pageSetup: pageSetup ? { slideWidth: formatPoints(pageSetup.slideWidth), slideHeight: formatPoints(pageSetup.slideHeight) } : undefined,
            textSelection:
              textRange && !textRange.isNullObject
                ? {
                    text: truncateText(textRange.text, 600),
                    start: textRange.start,
                    length: textRange.length,
                    fontName: textRange.font.name,
                    fontSize: textRange.font.size,
                    fontColor: textRange.font.color,
                    bold: textRange.font.bold,
                    italic: textRange.font.italic,
                    underline: textRange.font.underline,
                    horizontalAlignment: textRange.paragraphFormat.horizontalAlignment,
                    indentLevel: textRange.paragraphFormat.indentLevel,
                  }
                : undefined,
            selectedShapes: selectedShapeDescriptors.slice(0, 4),
            selectedSlides: selectedSlideSummaries.slice(0, 6).map((slide) => ({
              index: slide.index,
              layoutName: slide.layoutName,
              layoutType: slide.layoutType,
              slideMasterName: slide.slideMasterName,
            })),
          }
        : undefined,
      snippets: {
        documentStructure: {
          slides: presentationSlides.items.length,
          selectedSlides: slides.items.length,
          selectedShapes: selectedShapes.length,
          slideMasters: masterSummaries.length,
          layouts: masterSummaries.reduce((total, slideMaster) => total + slideMaster.layoutCount, 0),
        },
        slideDeck: Array.from(slideMetadataById.values()).slice(0, 30),
        selectedSlides: selectedSlideSummaries.slice(0, 6),
        selectedShapeDescriptors: selectedShapeDescriptors.slice(0, 8),
        slideMasters: masterSummaries.slice(0, 12),
        layoutUsage: layoutUsage.slice(0, 24),
      },
      visuals,
    };
    result.summary = createSummary(result);
    return result;
  });

  if (!payload.visuals?.length && maxImages > 0) {
    const fallback = await getSelectedImageAsync().catch(() => undefined);
    if (fallback) {
      payload.visuals = [await optimizeVisual({ ...fallback, kind: "slide", label: "PowerPoint selection snapshot" })];
      payload.summary = createSummary(payload);
    }
  }

  if (options.includeFormatting && supportsRequirementSet("PowerPointApi", "1.10")) {
    try {
      const packageSummary = await inspectCurrentPowerPointPresentationPackage();
      const slideIdByNumber = new Map<number, string>();
      for (const slide of getRecordArray(payload.snippets?.slideDeck)) {
        const slideNumber = toNumber(slide.index);
        const slideId = trimString(slide.id);
        if (typeof slideNumber === "number" && slideId) {
          slideIdByNumber.set(slideNumber, slideId);
        }
      }
      const selectedSlidesForNotes = getRecordArray(payload.snippets?.selectedSlides);
      payload.state.capabilities = Array.from(
        new Set([
          ...payload.state.capabilities,
          "powerpoint.presentationPackage",
          "powerpoint.slideNotes",
          "powerpoint.notesRegionAnchors",
          "powerpoint.charts",
        ]),
      );
      const slideNotesAnchors = packageSummary.notes.slice(0, 20).map((note) => {
        return {
          kind: "notesRegion",
          id: note.partName ?? `notes:${note.slideNumber}`,
          label: `Notes for Slide ${note.slideNumber}`,
          slideId: slideIdByNumber.get(note.slideNumber),
          slideIndex: note.slideNumber,
          text: truncateLabel(note.preview ?? note.text ?? (note.hasNotes ? `Speaker notes for slide ${note.slideNumber}` : undefined), 180),
        } as OfficeAnchor;
      });
      payload.anchors = uniqueAnchors([...(payload.anchors ?? []), ...slideNotesAnchors]);
      payload.formatting = {
        ...(payload.formatting ?? {}),
        presentationTheme: packageSummary.theme,
        presentationPackage: {
          partCounts: packageSummary.partCounts,
          notesSlides: packageSummary.notes.length,
          charts: packageSummary.charts.length,
        },
      };
      payload.snippets = {
        ...(payload.snippets ?? {}),
        presentationTheme: packageSummary.theme,
        selectedSlideNotes: selectedSlidesForNotes.slice(0, 12).map((slide) => {
          const slideNumber = toNumber(slide.index) ?? toNumber(slide.slideIndex) ?? 0;
          const note = packageSummary.notes.find((entry) => entry.slideNumber === slideNumber);
          return {
            slideNumber,
            slideId: trimString(slide.id) ?? slideIdByNumber.get(slideNumber),
            hasNotes: note?.hasNotes ?? false,
            preview: note?.preview,
            partName: note?.partName,
          };
        }),
        slideNotes: packageSummary.notes.slice(0, 20).map((note) => ({
          slideNumber: note.slideNumber,
          hasNotes: note.hasNotes,
          preview: note.preview,
          partName: note.partName,
        })),
        presentationCharts: packageSummary.charts.slice(0, 20).map((chart) => ({
          slideNumber: chart.slideNumber,
          chartIndex: chart.chartIndex,
          chartType: chart.chartType,
          title: chart.title,
          shapeName: chart.shapeName,
          embeddedWorkbookPartName: chart.embeddedWorkbookPartName,
          seriesCount: chart.seriesCount,
          categoryCount: chart.categoryCount,
          hasEmbeddedWorkbook: chart.hasEmbeddedWorkbook,
        })),
        presentationPackage: packageSummary.partCounts,
      };
      payload.summary = createSummary(payload);
    } catch (error) {
      payload.formatting = {
        ...(payload.formatting ?? {}),
        presentationPackageWarning:
          error instanceof Error ? error.message : "PowerPoint presentation package inspection was not available.",
      };
    }
  }

  return payload;
}

export async function waitForOfficeReady(): Promise<OfficeHost> {
  if (typeof Office === "undefined") {
    throw new Error("Office.js is not available in this taskpane.");
  }

  if (Office.context?.host) {
    return mapHost(Office.context.host);
  }

  return new Promise((resolve, reject) => {
    const timeout = window.setTimeout(() => {
      reject(new Error("Timed out waiting for the Office host. Relaunch the add-in inside Word, Excel, or PowerPoint."));
    }, 20_000);

    Office.onReady((info) => {
      window.clearTimeout(timeout);
      try {
        resolve(mapHost(info.host ?? Office.context?.host));
      } catch (error) {
        reject(error instanceof Error ? error : new Error(String(error)));
      }
    });
  });
}

export async function collectOfficeState(explicitHost?: OfficeHost): Promise<OfficeStateUpdate> {
  const host = explicitHost ?? mapHost(Office.context.host);
  const base = await buildBaseState(host);
  return getOfficeHostAdapter(host).collectState(base);
}

export async function collectOfficeContext(explicitHost?: OfficeHost, options: OfficeCaptureOptions = {}): Promise<OfficeContextPayload> {
  const host = explicitHost ?? mapHost(Office.context.host);
  const base = await buildBaseState(host);
  return getOfficeHostAdapter(host).collectContext(base, options);
}

export async function capturePromptVisuals(explicitHost?: OfficeHost, maxImages = 2): Promise<OfficeVisualSnapshot[]> {
  const payload = await collectOfficeContext(explicitHost, {
    includeFormatting: false,
    maxImages,
  });
  return payload.visuals ?? [];
}

export function buildOpenRequest(state: OfficeStateUpdate, forceNew?: boolean): OfficeSessionOpenRequest {
  return {
    host: state.host,
    documentId: state.document.id,
    documentPath: state.document.documentPath,
    documentUrl: state.document.documentUrl,
    saved: state.document.saved,
    title: state.document.title,
    selectionSummary: state.selection,
    forceNew,
  };
}

export function buildConnectorScopeContext(state: OfficeStateUpdate | undefined): ConnectorScopeContext | undefined {
  if (!state) {
    return undefined;
  }

  return {
    host: state.host,
    documentId: state.document.id,
    documentTitle: state.document.title,
    documentSaved: state.document.saved,
    documentUrl: state.document.documentUrl,
    workspaceId: state.document.workspaceDir,
  };
}

export function subscribeToOfficeChanges(listener: () => void | Promise<void>): () => void {
  const handler = () => {
    void listener();
  };

  try {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler);
  } catch {
    // Some hosts do not expose the selection handler consistently. Polling covers that case.
  }

  const interval = window.setInterval(() => {
    void listener();
  }, 2000);

  return () => {
    window.clearInterval(interval);
    try {
      Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, {
        handler,
      } as Office.RemoveHandlerOptions);
    } catch {
      // Ignore cleanup failures.
    }
  };
}

export function readOfficeTheme(): OfficeThemeSnapshot | undefined {
  try {
    const theme = Office.context?.officeTheme;
    if (!theme) {
      return undefined;
    }

    return {
      bodyBackgroundColor: theme.bodyBackgroundColor,
      bodyForegroundColor: theme.bodyForegroundColor,
      controlBackgroundColor: theme.controlBackgroundColor,
      controlForegroundColor: theme.controlForegroundColor,
      isDarkTheme: Boolean(theme.isDarkTheme),
    };
  } catch {
    return undefined;
  }
}

function matchesWordHeading(paragraph: Word.Paragraph, anchor: OfficeAnchor): boolean {
  const style = String(paragraph.styleBuiltIn || paragraph.style || "");
  const text = normalizeTextPreview(paragraph.text)?.toLowerCase() ?? "";
  const needle = normalizeTextPreview(anchor.text || anchor.label)?.toLowerCase();
  return /heading/i.test(style) && (!needle || text.includes(needle));
}

function matchesWordParagraph(paragraph: Word.Paragraph, anchor: OfficeAnchor): boolean {
  if (anchor.paragraphId && paragraph.uniqueLocalId === anchor.paragraphId) {
    return true;
  }

  const needle = normalizeTextPreview(anchor.text || anchor.label)?.toLowerCase();
  if (!needle) {
    return false;
  }

  return (normalizeTextPreview(paragraph.text)?.toLowerCase() ?? "").includes(needle);
}

async function navigateWordAnchor(anchor: OfficeAnchor): Promise<unknown> {
  return Word.run(async (context) => {
    const body = context.document.body;
    const supportsParagraphIds = supportsRequirementSet("WordApi", "1.6");
    const supportsFieldMetadata = supportsRequirementSet("WordApi", "1.5");
    const supportsContentControlSubtypes = supportsRequirementSet("WordApi", "1.3");

    if ((anchor.kind === "footnote" || anchor.kind === "endnote") && supportsRequirementSet("WordApi", "1.5")) {
      const notes = anchor.kind === "footnote" ? body.footnotes : body.endnotes;
      notes.load("items/type,items/body/text,items/reference/text");
      await context.sync();

      const indexedNote =
        parseAnchorOrdinal(anchor.id, anchor.kind) ?? parseAnchorOrdinal(anchor.label, anchor.kind);
      const noteIndex =
        typeof indexedNote === "number"
          ? indexedNote - 1
          : notes.items.findIndex(
              (entry) =>
                matchesTextQuery(entry.body.text, anchor.text || anchor.label) ||
                matchesTextQuery(entry.reference.text, anchor.text || anchor.label),
            );
      const note = noteIndex >= 0 ? notes.items[noteIndex] : undefined;
      if (note) {
        note.reference.select();
        await context.sync();
        return {
          ok: true,
          host: "word",
          anchorKind: anchor.kind,
          noteId: `${anchor.kind}:${noteIndex + 1}`,
          referenceText: truncateLabel(note.reference.text, 80),
          text: truncateLabel(note.body.text, 180),
        };
      }
    }

    if (anchor.kind === "comment" && supportsRequirementSet("WordApi", "1.4")) {
      const comments = body.getComments();
      comments.load("items/id,items/content,items/authorName,items/resolved");
      await context.sync();
      const comment = comments.items.find(
        (entry) =>
          (anchor.commentId && entry.id === anchor.commentId) ||
          matchesTextQuery(entry.content, anchor.text || anchor.label),
      );
      if (comment) {
        comment.getRange().select();
        await context.sync();
        return {
          ok: true,
          host: "word",
          anchorKind: "comment",
          commentId: comment.id,
          text: truncateLabel(comment.content, 180),
          resolved: comment.resolved,
        };
      }
    }

    if (anchor.kind === "revision" && supportsRequirementSet("WordApi", "1.6")) {
      const trackedChanges = body.getTrackedChanges();
      trackedChanges.load("items/author,items/date,items/text,items/type");
      await context.sync();
      const indexedRevision =
        parseAnchorOrdinal(anchor.revisionId, "revision") ??
        parseAnchorOrdinal(anchor.id, "revision") ??
        parseAnchorOrdinal(anchor.label, "revision");
      const changeIndex =
        typeof indexedRevision === "number"
          ? indexedRevision - 1
          : trackedChanges.items.findIndex((entry) => matchesTextQuery(entry.text, anchor.text || anchor.label));
      const change = changeIndex >= 0 ? trackedChanges.items[changeIndex] : undefined;
      if (change) {
        change.getRange().select();
        await context.sync();
        return {
          ok: true,
          host: "word",
          anchorKind: "revision",
          revisionId: `revision:${changeIndex + 1}`,
          text: truncateLabel(change.text, 180),
          type: change.type,
        };
      }
    }

    if (anchor.kind === "contentControl" && supportsRequirementSet("WordApi", "1.1")) {
      const contentControls = body.contentControls;
      const directId = parseAnchorOrdinal(anchor.id, "contentControl");
      if (typeof directId === "number") {
        const directMatch = contentControls.getByIdOrNullObject(directId);
        directMatch.load(supportsContentControlSubtypes ? "id,title,tag,text,type,subtype" : "id,title,tag,text,type");
        await context.sync();
        if (!directMatch.isNullObject) {
          directMatch.getRange().select();
          await context.sync();
          return {
            ok: true,
            host: "word",
            anchorKind: "contentControl",
            id: `contentControl:${directMatch.id}`,
            contentControlId: directMatch.id,
            title: trimString(directMatch.title),
            tag: trimString(directMatch.tag),
            text: truncateLabel(directMatch.text, 180),
            contentControlType: directMatch.type,
            contentControlSubtype: supportsContentControlSubtypes ? directMatch.subtype : undefined,
          };
        }
      }

      contentControls.load(supportsContentControlSubtypes ? "items/id,items/title,items/tag,items/text,items/type,items/subtype" : "items/id,items/title,items/tag,items/text,items/type");
      await context.sync();
      const contentControl = contentControls.items.find(
        (entry) =>
          matchesTextQuery(entry.title, anchor.text || anchor.label) ||
          matchesTextQuery(entry.tag, anchor.text || anchor.label) ||
          matchesTextQuery(entry.text, anchor.text || anchor.label),
      );
      if (contentControl) {
        contentControl.getRange().select();
        await context.sync();
        return {
          ok: true,
          host: "word",
          anchorKind: "contentControl",
          id: `contentControl:${contentControl.id}`,
          contentControlId: contentControl.id,
          title: trimString(contentControl.title),
          tag: trimString(contentControl.tag),
          text: truncateLabel(contentControl.text, 180),
          contentControlType: contentControl.type,
          contentControlSubtype: supportsContentControlSubtypes ? contentControl.subtype : undefined,
        };
      }
    }

    if (anchor.kind === "field" && supportsRequirementSet("WordApi", "1.4")) {
      const fields = body.fields;
      fields.load(supportsFieldMetadata ? "items/code,items/type,items/result/text" : "items/code,items/result/text");
      await context.sync();

      const indexedField = parseAnchorOrdinal(anchor.id, "field") ?? parseAnchorOrdinal(anchor.label, "field");
      const fieldIndex =
        typeof indexedField === "number"
          ? indexedField - 1
          : fields.items.findIndex(
              (entry) =>
                matchesTextQuery(entry.code, anchor.text || anchor.label) ||
                matchesTextQuery(entry.result.text, anchor.text || anchor.label) ||
                (supportsFieldMetadata && matchesTextQuery(String(entry.type), anchor.text || anchor.label)),
            );
      const field = fieldIndex >= 0 ? fields.items[fieldIndex] : undefined;
      if (field) {
        if (supportsFieldMetadata) {
          field.select();
        } else {
          field.result.select();
        }
        await context.sync();
        return {
          ok: true,
          host: "word",
          anchorKind: "field",
          fieldId: `field:${fieldIndex + 1}`,
          fieldIndex: fieldIndex + 1,
          fieldCode: truncateLabel(field.code, 180),
          text: truncateLabel(field.result.text, 180),
          fieldType: supportsFieldMetadata ? field.type : undefined,
        };
      }
    }

    const paragraphs = body.paragraphs;
    paragraphs.load(
      supportsParagraphIds
        ? "items/text,items/style,items/styleBuiltIn,items/uniqueLocalId"
        : "items/text,items/style,items/styleBuiltIn",
    );
    await context.sync();

    const paragraph =
      anchor.kind === "heading"
        ? paragraphs.items.find((entry) => matchesWordHeading(entry, anchor))
        : paragraphs.items.find((entry) => matchesWordParagraph(entry, anchor));
    if (paragraph) {
      paragraph.select();
      await context.sync();
      return {
        ok: true,
        host: "word",
        anchorKind: anchor.kind,
        paragraphId: supportsParagraphIds ? paragraph.uniqueLocalId : undefined,
        text: truncateLabel(paragraph.text, 180),
      };
    }

    const query = trimString(anchor.text) ?? trimString(anchor.label);
    if (query) {
      const matches = body.search(query, {
        matchCase: false,
        matchWholeWord: false,
      });
      matches.load("items");
      await context.sync();
      if (matches.items[0]) {
        matches.items[0].select();
        await context.sync();
        return { ok: true, host: "word", anchorKind: anchor.kind, query };
      }
    }

    throw new Error(`Could not find the requested Word anchor: ${anchor.label || anchor.text || anchor.kind}.`);
  });
}

async function applyWordAction(action: OfficeHostAction): Promise<unknown> {
  return Word.run(async (context) => {
    const body = context.document.body;
    const type = trimString(action.type) ?? "insertText";
    const content = action.content ?? "";
    const placement = wordInsertLocationFromPlacement(action.placement);
    const options = getActionOptions(action);
    const supportsParagraphIds = supportsRequirementSet("WordApi", "1.6");
    const supportsFieldMetadata = supportsRequirementSet("WordApi", "1.5");
    const supportsContentControlSubtypes = supportsRequirementSet("WordApi", "1.3");
    let resolvedTarget:
      | {
          range: Word.Range;
          paragraph?: Word.Paragraph | undefined;
          comment?: Word.Comment | undefined;
          revision?: { change: Word.TrackedChange; revisionId: string; revisionIndex: number } | undefined;
          contentControl?: Word.ContentControl | undefined;
          field?: { field: Word.Field; fieldIndex: number } | undefined;
        }
      | undefined;

    const resolveWordNoteTarget = async (target: OfficeAnchor) => {
      if ((target.kind !== "footnote" && target.kind !== "endnote") || !supportsRequirementSet("WordApi", "1.5")) {
        return undefined;
      }

      const notes = target.kind === "footnote" ? body.footnotes : body.endnotes;
      notes.load("items/type,items/body/text,items/reference/text");
      await context.sync();

      const indexedNote = parseAnchorOrdinal(target.id, target.kind) ?? parseAnchorOrdinal(target.label, target.kind);
      const noteIndex =
        typeof indexedNote === "number"
          ? indexedNote - 1
          : notes.items.findIndex(
              (entry) =>
                matchesTextQuery(entry.body.text, target.text || target.label) ||
                matchesTextQuery(entry.reference.text, target.text || target.label),
            );
      const note = noteIndex >= 0 ? notes.items[noteIndex] : undefined;
      return note ? { note, noteIndex: noteIndex + 1 } : undefined;
    };

    const resolveWordCommentTarget = async (target: OfficeAnchor) => {
      if (!supportsRequirementSet("WordApi", "1.4")) {
        return undefined;
      }

      const comments = body.getComments();
      comments.load("items/id,items/content,items/authorName,items/resolved");
      await context.sync();

      return comments.items.find(
        (entry) =>
          (target.commentId && entry.id === target.commentId) ||
          matchesTextQuery(entry.content, target.text || target.label),
      );
    };

    const resolveWordRevisionTarget = async (target: OfficeAnchor) => {
      if (!supportsRequirementSet("WordApi", "1.6")) {
        return undefined;
      }

      const trackedChanges = body.getTrackedChanges();
      trackedChanges.load("items/author,items/date,items/text,items/type");
      await context.sync();

      const indexedRevision =
        parseAnchorOrdinal(target.revisionId, "revision") ??
        parseAnchorOrdinal(target.id, "revision") ??
        parseAnchorOrdinal(target.label, "revision");
      const changeIndex =
        typeof indexedRevision === "number"
          ? indexedRevision - 1
          : trackedChanges.items.findIndex((entry) => matchesTextQuery(entry.text, target.text || target.label));
      const change = changeIndex >= 0 ? trackedChanges.items[changeIndex] : undefined;
      return change ? { change, revisionId: `revision:${changeIndex + 1}`, revisionIndex: changeIndex + 1 } : undefined;
    };

    const resolveWordParagraphTarget = async (target: OfficeAnchor) => {
      const paragraphs = body.paragraphs;
      paragraphs.load(
        supportsParagraphIds
          ? "items/text,items/style,items/styleBuiltIn,items/uniqueLocalId"
          : "items/text,items/style,items/styleBuiltIn",
      );
      await context.sync();

      return target.kind === "heading"
        ? paragraphs.items.find((entry) => matchesWordHeading(entry, target))
        : paragraphs.items.find((entry) => matchesWordParagraph(entry, target));
    };

    const resolveWordContentControlTarget = async (target: OfficeAnchor) => {
      if (!supportsRequirementSet("WordApi", "1.1")) {
        return undefined;
      }

      const contentControls = body.contentControls;
      const directId = parseAnchorOrdinal(target.id, "contentControl");
      if (typeof directId === "number") {
        const directMatch = contentControls.getByIdOrNullObject(directId);
        directMatch.load("id,title,tag,text,type");
        await context.sync();
        if (!directMatch.isNullObject) {
          return directMatch;
        }
      }

      contentControls.load("items/id,items/title,items/tag,items/text,items/type");
      await context.sync();

      return contentControls.items.find(
        (entry) =>
          matchesTextQuery(entry.title, target.text || target.label) ||
          matchesTextQuery(entry.tag, target.text || target.label) ||
          matchesTextQuery(entry.text, target.text || target.label),
      );
    };

    const resolveWordFieldTarget = async (target: OfficeAnchor) => {
      if (!supportsRequirementSet("WordApi", "1.4")) {
        return undefined;
      }

      const fields = body.fields;
      fields.load(supportsFieldMetadata ? "items/code,items/type,items/result/text" : "items/code,items/result/text");
      await context.sync();

      const indexedField = parseAnchorOrdinal(target.id, "field") ?? parseAnchorOrdinal(target.label, "field");
      const fieldIndex =
        typeof indexedField === "number"
          ? indexedField - 1
          : fields.items.findIndex(
              (entry) =>
                matchesTextQuery(entry.code, target.text || target.label) ||
                matchesTextQuery(entry.result.text, target.text || target.label) ||
                (supportsFieldMetadata && matchesTextQuery(String(entry.type), target.text || target.label)),
            );
      const field = fieldIndex >= 0 ? fields.items[fieldIndex] : undefined;
      return field ? { field, fieldIndex: fieldIndex + 1 } : undefined;
    };

    const resolveTargetRange = async () => {
      if (resolvedTarget) {
        return resolvedTarget;
      }

      const selection = context.document.getSelection();
      if (!action.target || action.target.kind === "selection") {
        resolvedTarget = { range: selection };
        return resolvedTarget;
      }

      if (action.target.kind === "comment") {
        const comment = await resolveWordCommentTarget(action.target);
        if (!comment) {
          throw new Error(`Could not find the requested Word comment: ${action.target.label || action.target.text || action.target.commentId || "comment"}.`);
        }
        resolvedTarget = { range: comment.getRange(), comment };
        return resolvedTarget;
      }

      if (action.target.kind === "revision") {
        const revision = await resolveWordRevisionTarget(action.target);
        if (!revision) {
          throw new Error(`Could not find the requested Word revision: ${action.target.label || action.target.text || action.target.revisionId || "revision"}.`);
        }
        resolvedTarget = { range: revision.change.getRange(), revision };
        return resolvedTarget;
      }

      if (action.target.kind === "heading" || action.target.kind === "paragraph") {
        const paragraph = await resolveWordParagraphTarget(action.target);
        if (!paragraph) {
          throw new Error(`Could not find the requested Word paragraph: ${action.target.label || action.target.text || action.target.paragraphId || action.target.kind}.`);
        }
        resolvedTarget = { range: paragraph.getRange(), paragraph };
        return resolvedTarget;
      }

      if (action.target.kind === "contentControl") {
        const contentControl = await resolveWordContentControlTarget(action.target);
        if (!contentControl) {
          throw new Error(`Could not find the requested Word content control: ${action.target.label || action.target.text || action.target.id || "content control"}.`);
        }
        resolvedTarget = { range: contentControl.getRange("Content"), contentControl };
        return resolvedTarget;
      }

      if (action.target.kind === "field") {
        const field = await resolveWordFieldTarget(action.target);
        if (!field) {
          throw new Error(`Could not find the requested Word field: ${action.target.label || action.target.text || action.target.id || "field"}.`);
        }
        resolvedTarget = { range: field.field.result, field };
        return resolvedTarget;
      }

      if (action.target.kind === "footnote" || action.target.kind === "endnote") {
        const note = await resolveWordNoteTarget(action.target);
        if (!note) {
          throw new Error(`Could not find the requested Word ${action.target.kind}: ${action.target.label || action.target.text || action.target.id || action.target.kind}.`);
        }
        resolvedTarget = { range: note.note.reference };
        return resolvedTarget;
      }

      resolvedTarget = { range: selection };
      return resolvedTarget;
    };

    const resolveCommentActionTarget = async () => {
      if (action.target?.kind === "comment") {
        const targetComment = await resolveWordCommentTarget(action.target);
        if (targetComment) {
          return targetComment;
        }
      }

      const selectionComments = context.document.getSelection().getComments();
      selectionComments.load("items/id,items/content,items/authorName,items/resolved");
      await context.sync();
      if (selectionComments.items[0]) {
        return selectionComments.items[0];
      }

      if (action.target) {
        const fallbackComment = await resolveWordCommentTarget(action.target);
        if (fallbackComment) {
          return fallbackComment;
        }
      }

      throw new Error("A Word comment target is required for this action.");
    };

    const resolveRevisionActionTarget = async () => {
      if (action.target?.kind === "revision") {
        const targetRevision = await resolveWordRevisionTarget(action.target);
        if (targetRevision) {
          return targetRevision;
        }
      }

      const selectionChanges = context.document.getSelection().getTrackedChanges();
      selectionChanges.load("items/author,items/date,items/text,items/type");
      await context.sync();
      if (selectionChanges.items[0]) {
        return {
          change: selectionChanges.items[0],
          revisionId: "revision:selection",
          revisionIndex: 1,
        };
      }

      if (action.target) {
        const fallbackRevision = await resolveWordRevisionTarget(action.target);
        if (fallbackRevision) {
          return fallbackRevision;
        }
      }

      throw new Error("A Word revision target is required for this action.");
    };

    if (type === "insertHtml") {
      const { range } = await resolveTargetRange();
      range.insertHtml(content, placement);
      await context.sync();
      return { ok: true, host: "word", action: type };
    }

    if (type === "insertText") {
      const { range } = await resolveTargetRange();
      range.insertText(content, placement);
      await context.sync();
      return { ok: true, host: "word", action: type };
    }

    if (type === "insertOoxml") {
      const { range } = await resolveTargetRange();
      range.insertOoxml(content, placement);
      await context.sync();
      return { ok: true, host: "word", action: type };
    }

    if (type === "addComment") {
      const { range } = await resolveTargetRange();
      const comment = range.insertComment(content);
      comment.load("id,content,resolved");
      await context.sync();
      return { ok: true, host: "word", action: type, commentId: comment.id, content: comment.content, resolved: comment.resolved };
    }

    if (type === "replyToComment") {
      const comment = await resolveCommentActionTarget();
      const reply = comment.reply(content);
      reply.load("id,content");
      await context.sync();
      return { ok: true, host: "word", action: type, commentId: comment.id, replyId: reply.id, content: reply.content };
    }

    if (type === "resolveComment" || type === "reopenComment" || type === "setCommentResolved") {
      const comment = await resolveCommentActionTarget();
      const resolved =
        type === "setCommentResolved"
          ? (toBoolean(options.resolved ?? action.resolved) ?? true)
          : type === "resolveComment";
      comment.resolved = resolved;
      await context.sync();
      return { ok: true, host: "word", action: type, commentId: comment.id, resolved };
    }

    if (type === "deleteComment") {
      const comment = await resolveCommentActionTarget();
      const commentId = comment.id;
      const existingContent = comment.content;
      comment.delete();
      await context.sync();
      return { ok: true, host: "word", action: type, commentId, content: existingContent };
    }

    if (type === "insertInlinePicture") {
      const { range } = await resolveTargetRange();
      const picture = range.insertInlinePictureFromBase64(content, placement);
      const altText = trimString(options.altText) ?? trimString(action.altText);
      if (altText) {
        picture.altTextTitle = altText;
      }
      picture.load("width,height,altTextTitle");
      await context.sync();
      return { ok: true, host: "word", action: type, width: picture.width, height: picture.height, altTextTitle: picture.altTextTitle };
    }

    if (type === "insertFileFromBase64") {
      const { range } = await resolveTargetRange();
      range.insertFileFromBase64(content, placement);
      await context.sync();
      return { ok: true, host: "word", action: type };
    }

    if (type === "insertTable") {
      const { range } = await resolveTargetRange();
      const values = toStringMatrix(options.values ?? action.values);
      const rowCount = toNumber(options.rowCount) ?? values?.length ?? 2;
      const columnCount = toNumber(options.columnCount) ?? values?.[0]?.length ?? 2;
      const tablePlacement = action.placement === "before" ? Word.InsertLocation.before : Word.InsertLocation.after;
      range.insertTable(rowCount, columnCount, tablePlacement, values);
      await context.sync();
      return { ok: true, host: "word", action: type, rowCount, columnCount };
    }

    if (type === "insertContentControl") {
      if (!supportsRequirementSet("WordApi", "1.1")) {
        throw new Error("Word content controls require WordApi 1.1 or newer.");
      }

      const supportsTypedContentControls = supportsRequirementSet("WordApi", "1.5");
      const { range } = await resolveTargetRange();
      const warnings: string[] = [];
      const requestedType = trimString(options.contentControlType) ?? trimString(action.contentControlType);
      let contentRange = range;

      if ((action.placement === "before" || action.placement === "after") && content) {
        contentRange = range.insertText(content, placement);
      } else if ((action.placement === "before" || action.placement === "after") && !content) {
        warnings.push("Wrapped the target range because empty content controls cannot be inserted before or after a target natively.");
      }

      if (requestedType && !supportsTypedContentControls) {
        warnings.push(`Requested content control type "${requestedType}" but this host only supports default rich text insertion.`);
      }

      const contentControl =
        requestedType && supportsTypedContentControls
          ? contentRange.insertContentControl(requestedType as never)
          : contentRange.insertContentControl();

      if (content && contentRange === range) {
        contentControl.insertText(content, Word.InsertLocation.replace);
      }

      const title = trimString(options.title) ?? trimString(action.title);
      const tag = trimString(options.tag) ?? trimString(action.tag);
      const placeholderText = trimString(options.placeholderText) ?? trimString(action.placeholderText);
      const appearance = trimString(options.appearance) ?? trimString(action.appearance);
      const color = trimString(options.color) ?? trimString(action.color);
      const cannotDelete = toBoolean(options.cannotDelete ?? action.cannotDelete);
      const cannotEdit = toBoolean(options.cannotEdit ?? action.cannotEdit);
      const removeWhenEdited = toBoolean(options.removeWhenEdited ?? action.removeWhenEdited);

      if (title) {
        contentControl.title = title;
      }
      if (tag) {
        contentControl.tag = tag;
      }
      if (placeholderText) {
        contentControl.placeholderText = placeholderText;
      }
      if (appearance) {
        contentControl.appearance = appearance as Word.ContentControlAppearance;
      }
      if (color) {
        contentControl.color = color;
      }
      if (typeof cannotDelete === "boolean") {
        contentControl.cannotDelete = cannotDelete;
      }
      if (typeof cannotEdit === "boolean") {
        contentControl.cannotEdit = cannotEdit;
      }
      if (typeof removeWhenEdited === "boolean") {
        contentControl.removeWhenEdited = removeWhenEdited;
      }

      contentControl.load(
        supportsContentControlSubtypes ? "id,title,tag,type,subtype,text" : "id,title,tag,type,text",
      );
      await context.sync();
      return {
        ok: true,
        host: "word",
        action: type,
        id: `contentControl:${contentControl.id}`,
        contentControlId: contentControl.id,
        title: trimString(contentControl.title),
        tag: trimString(contentControl.tag),
        text: truncateLabel(contentControl.text, 180),
        contentControlType: contentControl.type,
        contentControlSubtype: supportsContentControlSubtypes ? contentControl.subtype : undefined,
        warnings,
      };
    }

    if (type === "insertField") {
      if (!supportsRequirementSet("WordApi", "1.5")) {
        throw new Error("Word field insertion requires WordApi 1.5 or newer.");
      }

      const { range } = await resolveTargetRange();
      const fieldType = trimString(options.fieldType) ?? trimString(action.fieldType) ?? "Empty";
      const fieldText = trimString(options.text) ?? trimString(action.fieldText) ?? trimString(content);
      const removeFormatting = toBoolean(options.removeFormatting ?? action.removeFormatting) ?? false;
      const field = range.insertField(placement, fieldType as Word.FieldType, fieldText, removeFormatting);
      field.load("code,type,locked,result/text");
      await context.sync();

      const documentFields = body.fields;
      documentFields.load("items/code,items/type,items/result/text");
      await context.sync();

      let fieldIndex = 0;
      for (let index = documentFields.items.length - 1; index >= 0; index -= 1) {
        const entry = documentFields.items[index];
        if (!entry) {
          continue;
        }
        if (
          entry.code === field.code &&
          entry.result.text === field.result.text &&
          String(entry.type || "") === String(field.type || "")
        ) {
          fieldIndex = index + 1;
          break;
        }
      }

      return {
        ok: true,
        host: "word",
        action: type,
        fieldId: fieldIndex > 0 ? `field:${fieldIndex}` : undefined,
        fieldIndex: fieldIndex || undefined,
        fieldType: field.type,
        fieldCode: truncateLabel(field.code, 180),
        text: truncateLabel(field.result.text, 180),
        locked: field.locked,
      };
    }

    if (type === "acceptRevision" || type === "rejectRevision") {
      const revision = await resolveRevisionActionTarget();
      if (type === "acceptRevision") {
        revision.change.accept();
      } else {
        revision.change.reject();
      }
      await context.sync();
      return {
        ok: true,
        host: "word",
        action: type,
        revisionId: revision.revisionId,
        text: truncateLabel(revision.change.text, 180),
        type: revision.change.type,
      };
    }

    if (type === "acceptAllRevisions" || type === "rejectAllRevisions") {
      if (!supportsRequirementSet("WordApi", "1.6")) {
        throw new Error("Word revision actions require WordApi 1.6 or newer.");
      }

      const selectionChanges = context.document.getSelection().getTrackedChanges();
      selectionChanges.load("items/text");
      await context.sync();
      const targetChanges = selectionChanges.items.length ? selectionChanges : body.getTrackedChanges();

      if (!selectionChanges.items.length) {
        targetChanges.load("items/text");
        await context.sync();
      }

      if (type === "acceptAllRevisions") {
        targetChanges.acceptAll();
      } else {
        targetChanges.rejectAll();
      }
      await context.sync();
      return {
        ok: true,
        host: "word",
        action: type,
        scope: selectionChanges.items.length ? "selection" : "document",
        count: targetChanges.items.length,
      };
    }

    throw new Error(`Unsupported Word action: ${type}`);
  });
}

function resolveExcelWorksheet(
  context: Excel.RequestContext,
  target?: OfficeAnchor,
  allowActive = true,
): Excel.Worksheet {
  const parsedAddress = splitSheetAddress(target?.address, target?.sheetName);
  const sheetName = parsedAddress.sheetName ?? trimString(target?.sheetName);
  if (sheetName) {
    return context.workbook.worksheets.getItem(sheetName);
  }
  if (allowActive) {
    return context.workbook.worksheets.getActiveWorksheet();
  }
  throw new Error("Excel worksheet target is required for this action.");
}

function resolveExcelRange(
  context: Excel.RequestContext,
  target?: OfficeAnchor,
  allowSelected = true,
): Excel.Range {
  if (target?.namedItemName) {
    return context.workbook.names.getItem(target.namedItemName).getRange();
  }

  const parsedAddress = splitSheetAddress(target?.address, target?.sheetName);
  if (parsedAddress.address) {
    return resolveExcelWorksheet(context, target, true).getRange(parsedAddress.address);
  }

  if (allowSelected) {
    return context.workbook.getSelectedRange();
  }

  throw new Error("Excel range target is required for this action.");
}

function resolveExcelTable(context: Excel.RequestContext, target: OfficeAnchor): Excel.Table {
  if (!target.tableName) {
    throw new Error("Excel table name is required for this action.");
  }

  if (Boolean(target.sheetName || (target.address && target.address.includes("!")))) {
    return resolveExcelWorksheet(context, target, true).tables.getItem(target.tableName);
  }

  return context.workbook.tables.getItem(target.tableName);
}

function resolveExcelChart(context: Excel.RequestContext, target: OfficeAnchor): Excel.Chart {
  if (!target.chartName) {
    throw new Error("Excel chart name is required for this action.");
  }

  return resolveExcelWorksheet(context, target, true).charts.getItem(target.chartName);
}

function resolveExcelPivotTable(context: Excel.RequestContext, target: OfficeAnchor): Excel.PivotTable {
  if (!target.pivotTableName) {
    throw new Error("Excel PivotTable name is required for this action.");
  }

  if (Boolean(target.sheetName || (target.address && target.address.includes("!")))) {
    return resolveExcelWorksheet(context, target, true).pivotTables.getItem(target.pivotTableName);
  }

  return context.workbook.pivotTables.getItem(target.pivotTableName);
}

function resolveExcelTableColumn(table: Excel.Table, options: Record<string, unknown>): Excel.TableColumn {
  const columnName = trimString(options.columnName) ?? trimString(options.fieldName) ?? trimString(options.name);
  const zeroBasedIndex = toNumber(options.columnIndex);
  const oneBasedColumnNumber = toNumber(options.columnNumber);

  if (columnName) {
    return table.columns.getItem(columnName);
  }
  if (typeof zeroBasedIndex === "number" && Number.isInteger(zeroBasedIndex) && zeroBasedIndex >= 0) {
    return table.columns.getItemAt(zeroBasedIndex);
  }
  if (typeof oneBasedColumnNumber === "number" && Number.isInteger(oneBasedColumnNumber) && oneBasedColumnNumber > 0) {
    return table.columns.getItemAt(oneBasedColumnNumber - 1);
  }

  throw new Error("Excel table filter actions require columnName, fieldName, columnIndex, or columnNumber.");
}

function getExcelRuntime(): typeof Excel {
  if (typeof Excel === "undefined") {
    throw new Error("Excel runtime is unavailable.");
  }
  return Excel;
}

let excelOutlineBorderIndexes: Excel.BorderIndex[] | undefined;

function getExcelOutlineBorderIndexes(): Excel.BorderIndex[] {
  if (excelOutlineBorderIndexes) {
    return excelOutlineBorderIndexes;
  }

  const excel = getExcelRuntime();
  excelOutlineBorderIndexes = [
    excel.BorderIndex.edgeTop,
    excel.BorderIndex.edgeBottom,
    excel.BorderIndex.edgeLeft,
    excel.BorderIndex.edgeRight,
  ];
  return excelOutlineBorderIndexes;
}

let excelDefaultBorderIndexes: Excel.BorderIndex[] | undefined;

function getExcelDefaultBorderIndexes(): Excel.BorderIndex[] {
  if (excelDefaultBorderIndexes) {
    return excelDefaultBorderIndexes;
  }

  const excel = getExcelRuntime();
  excelDefaultBorderIndexes = [
    ...getExcelOutlineBorderIndexes(),
    excel.BorderIndex.insideVertical,
    excel.BorderIndex.insideHorizontal,
  ];
  return excelDefaultBorderIndexes;
}

let excelBorderIndexAliases: Record<string, Excel.BorderIndex> | undefined;

function getExcelBorderIndexAliases(): Record<string, Excel.BorderIndex> {
  if (excelBorderIndexAliases) {
    return excelBorderIndexAliases;
  }

  const excel = getExcelRuntime();
  excelBorderIndexAliases = {
    top: excel.BorderIndex.edgeTop,
    edgetop: excel.BorderIndex.edgeTop,
    bottom: excel.BorderIndex.edgeBottom,
    edgebottom: excel.BorderIndex.edgeBottom,
    left: excel.BorderIndex.edgeLeft,
    edgeleft: excel.BorderIndex.edgeLeft,
    right: excel.BorderIndex.edgeRight,
    edgeright: excel.BorderIndex.edgeRight,
    insidevertical: excel.BorderIndex.insideVertical,
    vertical: excel.BorderIndex.insideVertical,
    insidehorizontal: excel.BorderIndex.insideHorizontal,
    horizontal: excel.BorderIndex.insideHorizontal,
    diagonaldown: excel.BorderIndex.diagonalDown,
    diagdown: excel.BorderIndex.diagonalDown,
    diagonalup: excel.BorderIndex.diagonalUp,
    diagup: excel.BorderIndex.diagonalUp,
  };
  return excelBorderIndexAliases;
}

function normalizeExcelBorderIndex(value: unknown): Excel.BorderIndex | undefined {
  const normalized = trimString(value)?.replace(/[\s_-]+/g, "").toLowerCase();
  return normalized ? getExcelBorderIndexAliases()[normalized] : undefined;
}

function getExcelBorderIndexes(value: unknown): Excel.BorderIndex[] {
  const entries = Array.isArray(value) ? value : [value];
  const indexes: Excel.BorderIndex[] = [];
  for (const entry of entries) {
    const index = normalizeExcelBorderIndex(entry);
    if (index && !indexes.includes(index)) {
      indexes.push(index);
    }
  }
  return indexes;
}

function getExcelBorderUpdateOptions(value: unknown): ExcelBorderUpdateOptions | undefined {
  if (value == null) {
    return undefined;
  }

  if (typeof value === "string") {
    const normalized = value.trim();
    if (!normalized) {
      return undefined;
    }
    if (/^(none|clear)$/i.test(normalized)) {
      return { style: Excel.BorderLineStyle.none };
    }
    return { style: normalized as Excel.BorderLineStyle };
  }

  if (!isRecord(value)) {
    return undefined;
  }

  const clear = toBoolean(value.clear) ?? toBoolean(value.none);
  const color = trimString(value.color) ?? trimString(value.borderColor) ?? trimString(value.lineColor);
  const style = trimString(value.style) ?? trimString(value.borderStyle) ?? trimString(value.lineStyle);
  const weight = trimString(value.weight) ?? trimString(value.borderWeight) ?? trimString(value.lineWeight);
  const tintAndShade = toNumber(value.tintAndShade ?? value.borderTintAndShade);

  if (clear) {
    return { style: Excel.BorderLineStyle.none };
  }

  const update: ExcelBorderUpdateOptions = {};
  if (color) update.color = color;
  if (style) update.style = style as Excel.BorderLineStyle;
  if (weight) update.weight = weight as Excel.BorderWeight;
  if (typeof tintAndShade === "number") update.tintAndShade = tintAndShade;
  return Object.keys(update).length ? update : undefined;
}

function mergeExcelBorderUpdateOptions(
  base: ExcelBorderUpdateOptions | undefined,
  override: ExcelBorderUpdateOptions | undefined,
): ExcelBorderUpdateOptions | undefined {
  if (!base && !override) {
    return undefined;
  }

  return {
    ...base,
    ...override,
  };
}

function applyExcelRangeBorder(
  range: Excel.Range,
  index: Excel.BorderIndex,
  update: ExcelBorderUpdateOptions | undefined,
): boolean {
  if (!update) {
    return false;
  }

  const border = range.format.borders.getItem(index);
  if (update.color) border.color = update.color;
  if (update.style) border.style = update.style;
  if (update.weight) border.weight = update.weight;
  if (typeof update.tintAndShade === "number") border.tintAndShade = update.tintAndShade;
  return true;
}

function applyExcelRangeBorders(range: Excel.Range, options: Record<string, unknown>): boolean {
  const outlineBorderIndexes = getExcelOutlineBorderIndexes();
  const rootBorderUpdate = getExcelBorderUpdateOptions({
    color: options.borderColor ?? options.borderLineColor,
    style: options.borderStyle ?? options.lineStyle,
    weight: options.borderWeight ?? options.lineWeight,
    tintAndShade: options.borderTintAndShade,
    clear: options.clearBorders,
  });
  const sharedBorderUpdate = mergeExcelBorderUpdateOptions(
    mergeExcelBorderUpdateOptions(rootBorderUpdate, getExcelBorderUpdateOptions(options.border)),
    mergeExcelBorderUpdateOptions(getExcelBorderUpdateOptions(options.allBorders), getExcelBorderUpdateOptions(options.allBorder)),
  );
  const explicitTargetIndexes = getExcelBorderIndexes(options.borderSides ?? options.borderIndexes ?? options.sides);
  const arrayBorderIndexes =
    Array.isArray(options.borders) && options.borders.some((entry) => typeof entry === "string") ? getExcelBorderIndexes(options.borders) : [];
  const defaultTargetIndexes =
    explicitTargetIndexes.length ? explicitTargetIndexes : arrayBorderIndexes.length ? arrayBorderIndexes : getExcelDefaultBorderIndexes();
  const outlineBorderUpdate = mergeExcelBorderUpdateOptions(rootBorderUpdate, getExcelBorderUpdateOptions(options.outlineBorder));
  const sideEntries: Array<[string, unknown]> = [
    ["topBorder", options.topBorder],
    ["bottomBorder", options.bottomBorder],
    ["leftBorder", options.leftBorder],
    ["rightBorder", options.rightBorder],
    ["insideHorizontalBorder", options.insideHorizontalBorder],
    ["insideVerticalBorder", options.insideVerticalBorder],
    ["diagonalDownBorder", options.diagonalDownBorder],
    ["diagonalUpBorder", options.diagonalUpBorder],
  ];

  const usesTintAndShade =
    typeof sharedBorderUpdate?.tintAndShade === "number" ||
    typeof outlineBorderUpdate?.tintAndShade === "number" ||
    sideEntries.some(([, entry]) => typeof getExcelBorderUpdateOptions(entry)?.tintAndShade === "number") ||
    (isRecord(options.borders) &&
      Object.values(options.borders).some((entry) => typeof getExcelBorderUpdateOptions(entry)?.tintAndShade === "number")) ||
    getRecordArray(options.borders).some((entry) => typeof getExcelBorderUpdateOptions(entry)?.tintAndShade === "number");
  if (usesTintAndShade && !supportsRequirementSet("ExcelApi", "1.9")) {
    throw new Error("Excel border tintAndShade formatting requires ExcelApi 1.9.");
  }

  let applied = false;
  if (sharedBorderUpdate) {
    for (const index of defaultTargetIndexes) {
      applied = applyExcelRangeBorder(range, index, sharedBorderUpdate) || applied;
    }
  }

  if (outlineBorderUpdate) {
    for (const index of outlineBorderIndexes) {
      applied = applyExcelRangeBorder(range, index, outlineBorderUpdate) || applied;
    }
  }

  for (const [key, entry] of sideEntries) {
    const index = normalizeExcelBorderIndex(key);
    const update = mergeExcelBorderUpdateOptions(rootBorderUpdate, getExcelBorderUpdateOptions(entry));
    if (index && update) {
      applied = applyExcelRangeBorder(range, index, update) || applied;
    }
  }

  if (isRecord(options.borders)) {
    for (const [key, entry] of Object.entries(options.borders)) {
      const index = normalizeExcelBorderIndex(key);
      if (!index) {
        continue;
      }
      const update = mergeExcelBorderUpdateOptions(rootBorderUpdate, getExcelBorderUpdateOptions(entry));
      if (update) {
        applied = applyExcelRangeBorder(range, index, update) || applied;
      }
    }
  }

  for (const entry of getRecordArray(options.borders)) {
    const index =
      normalizeExcelBorderIndex(entry.side ?? entry.index ?? entry.name) ??
      (toBoolean(entry.outline) ? undefined : normalizeExcelBorderIndex(entry.borderSide ?? entry.borderIndex));
    const update = mergeExcelBorderUpdateOptions(rootBorderUpdate, getExcelBorderUpdateOptions(entry));
    if (toBoolean(entry.outline) && update) {
      for (const outlineIndex of outlineBorderIndexes) {
        applied = applyExcelRangeBorder(range, outlineIndex, update) || applied;
      }
      continue;
    }
    if (index && update) {
      applied = applyExcelRangeBorder(range, index, update) || applied;
    }
  }

  return applied;
}

function getExcelFilterValueArray(value: unknown): Array<string | Excel.FilterDatetime> {
  if (!Array.isArray(value)) {
    return [];
  }

  const result: Array<string | Excel.FilterDatetime> = [];
  for (const entry of value) {
    if (typeof entry === "string" || typeof entry === "number" || typeof entry === "boolean") {
      result.push(String(entry));
      continue;
    }
    if (isRecord(entry) && trimString(entry.date) && trimString(entry.specificity)) {
      result.push(entry as unknown as Excel.FilterDatetime);
      continue;
    }
  }
  return result;
}

function buildExcelIcon(value: unknown): Excel.Icon | undefined {
  if (!isRecord(value)) {
    return undefined;
  }

  const set = trimString(value.set) ?? trimString(value.iconSet) ?? trimString(value.customIconSet);
  const index = toNumber(value.index ?? value.iconIndex ?? value.customIconIndex);
  if (!set || typeof index !== "number") {
    return undefined;
  }

  return {
    set: set as Excel.IconSet,
    index,
  } as Excel.Icon;
}

function applyExcelTableFilter(column: Excel.TableColumn, options: Record<string, unknown>): void {
  const filter = column.filter;
  const filterType = (trimString(options.filterType) ?? trimString(options.type) ?? "").toLowerCase();

  if ((toBoolean(options.clear) ?? false) || filterType === "clear") {
    filter.clear();
    return;
  }

  if (isRecord(options.criteria)) {
    filter.apply(options.criteria as unknown as Excel.FilterCriteria);
    return;
  }

  if (filterType === "values") {
    const values = getExcelFilterValueArray(options.values);
    if (!values.length) {
      throw new Error("Excel values filter requires a non-empty values array.");
    }
    filter.applyValuesFilter(values);
    return;
  }

  if (filterType === "custom") {
    const criteria1 = trimString(options.criteria1);
    if (!criteria1) {
      throw new Error("Excel custom table filter requires criteria1.");
    }
    filter.applyCustomFilter(criteria1, trimString(options.criteria2), trimString(options.operator) as Excel.FilterOperator);
    return;
  }

  if (filterType === "dynamic") {
    const criteria = trimString(options.criteria) ?? trimString(options.dynamicCriteria);
    if (!criteria) {
      throw new Error("Excel dynamic table filter requires criteria.");
    }
    filter.applyDynamicFilter(criteria as Excel.DynamicFilterCriteria);
    return;
  }

  if (filterType === "cellcolor" || filterType === "cellColor".toLowerCase()) {
    const color = trimString(options.color) ?? trimString(options.fillColor);
    if (!color) {
      throw new Error("Excel cell color filter requires color.");
    }
    filter.applyCellColorFilter(color);
    return;
  }

  if (filterType === "fontcolor" || filterType === "fontColor".toLowerCase()) {
    const color = trimString(options.color) ?? trimString(options.fontColor);
    if (!color) {
      throw new Error("Excel font color filter requires color.");
    }
    filter.applyFontColorFilter(color);
    return;
  }

  if (filterType === "icon") {
    const icon = buildExcelIcon(options.icon ?? options.criteria);
    if (!icon) {
      throw new Error("Excel icon table filter requires icon.set and icon.index.");
    }
    filter.applyIconFilter(icon);
    return;
  }

  if (filterType === "topitems") {
    const count = toNumber(options.count);
    if (typeof count !== "number") {
      throw new Error("Excel top items filter requires count.");
    }
    filter.applyTopItemsFilter(count);
    return;
  }

  if (filterType === "toppercent") {
    const percent = toNumber(options.percent);
    if (typeof percent !== "number") {
      throw new Error("Excel top percent filter requires percent.");
    }
    filter.applyTopPercentFilter(percent);
    return;
  }

  if (filterType === "bottomitems") {
    const count = toNumber(options.count);
    if (typeof count !== "number") {
      throw new Error("Excel bottom items filter requires count.");
    }
    filter.applyBottomItemsFilter(count);
    return;
  }

  if (filterType === "bottompercent") {
    const percent = toNumber(options.percent);
    if (typeof percent !== "number") {
      throw new Error("Excel bottom percent filter requires percent.");
    }
    filter.applyBottomPercentFilter(percent);
    return;
  }

  throw new Error("Excel table filter requires criteria or a supported filterType.");
}

function resolveExcelPivotField(pivotTable: Excel.PivotTable, value: Record<string, unknown>): Excel.PivotField {
  const hierarchyName = trimString(value.hierarchyName) ?? trimString(value.fieldName) ?? trimString(value.name);
  if (!hierarchyName) {
    throw new Error("Excel PivotTable sort requires hierarchyName or fieldName.");
  }
  const fieldName = trimString(value.fieldName) ?? hierarchyName;
  return pivotTable.hierarchies.getItem(hierarchyName).fields.getItem(fieldName);
}

async function applyExcelPivotSorts(
  context: Excel.RequestContext,
  pivotTable: Excel.PivotTable,
  value: unknown,
): Promise<void> {
  const sorts = isRecord(value) ? [value] : getRecordArray(value);
  for (const sortConfig of sorts) {
    const field = resolveExcelPivotField(pivotTable, sortConfig);
    const direction = (trimString(sortConfig.sortBy) ?? trimString(sortConfig.direction) ?? "Ascending") as Excel.SortBy;
    const valuesHierarchyName =
      trimString(sortConfig.valuesHierarchy) ??
      trimString(sortConfig.valuesHierarchyName) ??
      trimString(sortConfig.dataHierarchy) ??
      trimString(sortConfig.dataHierarchyName);
    const mode = (trimString(sortConfig.mode) ?? (valuesHierarchyName ? "values" : "labels")).toLowerCase();

    if (mode === "values" || valuesHierarchyName) {
      if (!supportsRequirementSet("ExcelApi", "1.9")) {
        throw new Error("Excel PivotTable value sorting requires ExcelApi 1.9.");
      }
      if (!valuesHierarchyName) {
        throw new Error("Excel PivotTable value sorting requires valuesHierarchy or dataHierarchy.");
      }
      const valuesHierarchy = pivotTable.dataHierarchies.getItem(valuesHierarchyName);
      const scope = getStringArray(sortConfig.pivotItemScope ?? sortConfig.scopeItems);
      field.sortByValues(direction, valuesHierarchy, scope.length ? scope : undefined);
      continue;
    }

    field.sortByLabels(direction);
  }
}

function getStringArray(value: unknown): string[] {
  return Array.isArray(value) ? value.filter((entry): entry is string => typeof entry === "string" && entry.trim().length > 0) : [];
}

function getNumberArray(value: unknown): number[] {
  return Array.isArray(value) ? value.map((entry) => toNumber(entry)).filter((entry): entry is number => typeof entry === "number") : [];
}

function getRecordArray(value: unknown): Record<string, unknown>[] {
  return Array.isArray(value) ? value.filter((entry): entry is Record<string, unknown> => isRecord(entry)) : [];
}

function tryParseJsonValue(value: unknown): unknown {
  if (typeof value !== "string" || !value.trim()) {
    return value;
  }

  try {
    return JSON.parse(value);
  } catch {
    return value;
  }
}

function getChartValueArray(value: unknown): Array<string | number> {
  const candidate = tryParseJsonValue(value);
  if (!Array.isArray(candidate)) {
    return [];
  }

  return candidate
    .map((entry) => {
      const numeric = toNumber(entry);
      return typeof numeric === "number" ? numeric : trimString(entry);
    })
    .filter((entry): entry is string | number => typeof entry === "number" || typeof entry === "string");
}

function getChartSeriesInput(value: unknown): Array<{ name?: string | undefined; categories?: Array<string | number> | undefined; values: number[] }> {
  const candidate = tryParseJsonValue(value);
  if (!Array.isArray(candidate)) {
    return [];
  }

  return candidate.flatMap((entry) => {
      if (!isRecord(entry)) {
        return [];
      }

      const values = getNumberArray(entry.values);
      if (!values.length) {
        return [];
      }

      const categories = getChartValueArray(entry.categories);
      return [{
        name: trimString(entry.name),
        categories: categories.length ? categories : undefined,
        values,
      }];
    });
}

function applyExcelChartAxisOptions(axis: Excel.ChartAxis, value: unknown): void {
  if (!isRecord(value)) {
    return;
  }

  const title = trimString(value.title) ?? trimString(value.text);
  const titleVisible = toBoolean(value.titleVisible);
  const visible = toBoolean(value.visible);
  const displayUnit = trimString(value.displayUnit);
  const numberFormat = trimString(value.numberFormat);
  const minimum = toNumber(value.minimum);
  const maximum = toNumber(value.maximum);
  const majorUnit = toNumber(value.majorUnit);
  const minorUnit = toNumber(value.minorUnit);
  const majorGridlinesVisible = toBoolean(value.majorGridlinesVisible);
  const minorGridlinesVisible = toBoolean(value.minorGridlinesVisible);
  const reversePlotOrder = toBoolean(value.reversePlotOrder);
  const logBase = toNumber(value.logBase);

  if (typeof visible === "boolean") {
    axis.visible = visible;
  }
  if (title) {
    axis.title.text = title;
    axis.title.visible = true;
  } else if (typeof titleVisible === "boolean") {
    axis.title.visible = titleVisible;
  }
  if (displayUnit) {
    axis.displayUnit = displayUnit as Excel.ChartAxisDisplayUnit;
  }
  if (numberFormat) {
    axis.numberFormat = numberFormat;
    if (supportsRequirementSet("ExcelApi", "1.9")) {
      axis.linkNumberFormat = false;
    }
  }
  if (typeof minimum === "number") {
    axis.minimum = minimum;
  }
  if (typeof maximum === "number") {
    axis.maximum = maximum;
  }
  if (typeof majorUnit === "number") {
    axis.majorUnit = majorUnit;
  }
  if (typeof minorUnit === "number") {
    axis.minorUnit = minorUnit;
  }
  if (typeof majorGridlinesVisible === "boolean") {
    axis.majorGridlines.visible = majorGridlinesVisible;
  }
  if (typeof minorGridlinesVisible === "boolean") {
    axis.minorGridlines.visible = minorGridlinesVisible;
  }
  if (typeof reversePlotOrder === "boolean") {
    axis.reversePlotOrder = reversePlotOrder;
  }
  if (typeof logBase === "number") {
    axis.logBase = logBase;
  }
}

function applyExcelChartDataLabelOptions(labels: Excel.ChartDataLabels, value: unknown): void {
  if (!isRecord(value)) {
    return;
  }

  const position = trimString(value.position);
  const separator = trimString(value.separator);
  const showValue = toBoolean(value.showValue);
  const showCategoryName = toBoolean(value.showCategoryName);
  const showSeriesName = toBoolean(value.showSeriesName);
  const showLegendKey = toBoolean(value.showLegendKey);
  const showPercentage = toBoolean(value.showPercentage);
  const showBubbleSize = toBoolean(value.showBubbleSize);
  const showLeaderLines = toBoolean(value.showLeaderLines);
  const numberFormat = trimString(value.numberFormat);
  const fontColor = trimString(value.fontColor);
  const fontSize = toNumber(value.fontSize);
  const bold = toBoolean(value.bold);
  const italic = toBoolean(value.italic);
  const fillColor = trimString(value.fillColor);
  const borderColor = trimString(value.borderColor);
  const borderWeight = toNumber(value.borderWeight);

  if (position) {
    labels.position = position as Excel.ChartDataLabelPosition;
  }
  if (separator) {
    labels.separator = separator;
  }
  if (typeof showValue === "boolean") {
    labels.showValue = showValue;
  }
  if (typeof showCategoryName === "boolean") {
    labels.showCategoryName = showCategoryName;
  }
  if (typeof showSeriesName === "boolean") {
    labels.showSeriesName = showSeriesName;
  }
  if (typeof showLegendKey === "boolean") {
    labels.showLegendKey = showLegendKey;
  }
  if (typeof showPercentage === "boolean") {
    labels.showPercentage = showPercentage;
  }
  if (typeof showBubbleSize === "boolean") {
    labels.showBubbleSize = showBubbleSize;
  }
  if (typeof showLeaderLines === "boolean" && supportsRequirementSet("ExcelApi", "1.19")) {
    labels.showLeaderLines = showLeaderLines;
  }
  if (numberFormat) {
    labels.numberFormat = numberFormat;
    if (supportsRequirementSet("ExcelApi", "1.9")) {
      labels.linkNumberFormat = false;
    }
  }
  if (fontColor) {
    labels.format.font.color = fontColor;
  }
  if (typeof fontSize === "number") {
    labels.format.font.size = fontSize;
  }
  if (typeof bold === "boolean") {
    labels.format.font.bold = bold;
  }
  if (typeof italic === "boolean") {
    labels.format.font.italic = italic;
  }
  if (fillColor) {
    labels.format.fill.setSolidColor(fillColor);
  }
  if (borderColor) {
    labels.format.border.color = borderColor;
  }
  if (typeof borderWeight === "number") {
    labels.format.border.weight = borderWeight;
  }
}

async function applyExcelChartConfiguration(
  context: Excel.RequestContext,
  chart: Excel.Chart,
  value: Record<string, unknown>,
): Promise<void> {
  const source = trimString(value.source);
  if (source) {
    const parsed = splitSheetAddress(source, trimString(value.sheetName));
    const sourceWorksheet = resolveExcelWorksheet(context, { kind: "sheet", sheetName: parsed.sheetName }, true);
    chart.setData(sourceWorksheet.getRange(parsed.address ?? source), (trimString(value.seriesBy) ?? "Auto") as Excel.ChartSeriesBy);
  }

  const chartType = trimString(value.chartType);
  const title = trimString(value.title);
  const legendVisible = toBoolean(value.legendVisible);
  const legendPosition = trimString(value.legendPosition);
  const chartLeft = toNumber(value.left);
  const chartTop = toNumber(value.top);
  const chartWidth = toNumber(value.width);
  const chartHeight = toNumber(value.height);
  if (chartType) chart.chartType = chartType as Excel.ChartType;
  if (title) {
    chart.title.text = title;
    chart.title.visible = true;
  }
  if (typeof legendVisible === "boolean") {
    chart.legend.visible = legendVisible;
  }
  if (legendPosition) {
    chart.legend.position = legendPosition as Excel.ChartLegendPosition;
  }
  if (typeof chartLeft === "number") chart.left = chartLeft;
  if (typeof chartTop === "number") chart.top = chartTop;
  if (typeof chartWidth === "number") chart.width = chartWidth;
  if (typeof chartHeight === "number") chart.height = chartHeight;

  const categoryAxis: Record<string, unknown> = {};
  if (isRecord(value.categoryAxis)) Object.assign(categoryAxis, value.categoryAxis);
  if (value.categoryAxisTitle != null) categoryAxis.title = value.categoryAxisTitle;
  if (value.categoryAxisVisible != null) categoryAxis.visible = value.categoryAxisVisible;
  if (value.categoryAxisNumberFormat != null) categoryAxis.numberFormat = value.categoryAxisNumberFormat;
  if (Object.keys(categoryAxis).length) {
    applyExcelChartAxisOptions(chart.axes.categoryAxis, categoryAxis);
  }

  const valueAxis: Record<string, unknown> = {};
  if (isRecord(value.valueAxis)) Object.assign(valueAxis, value.valueAxis);
  if (value.valueAxisTitle != null) valueAxis.title = value.valueAxisTitle;
  if (value.valueAxisVisible != null) valueAxis.visible = value.valueAxisVisible;
  if (value.valueAxisDisplayUnit != null) valueAxis.displayUnit = value.valueAxisDisplayUnit;
  if (value.valueAxisMinimum != null) valueAxis.minimum = value.valueAxisMinimum;
  if (value.valueAxisMaximum != null) valueAxis.maximum = value.valueAxisMaximum;
  if (value.valueAxisMajorUnit != null) valueAxis.majorUnit = value.valueAxisMajorUnit;
  if (value.valueAxisMinorUnit != null) valueAxis.minorUnit = value.valueAxisMinorUnit;
  if (value.valueAxisMajorGridlinesVisible != null) valueAxis.majorGridlinesVisible = value.valueAxisMajorGridlinesVisible;
  if (value.valueAxisNumberFormat != null) valueAxis.numberFormat = value.valueAxisNumberFormat;
  if (Object.keys(valueAxis).length) {
    applyExcelChartAxisOptions(chart.axes.valueAxis, valueAxis);
  }

  const dataLabels: Record<string, unknown> = {};
  if (isRecord(value.dataLabels)) Object.assign(dataLabels, value.dataLabels);
  if (value.showDataLabels != null) dataLabels.visible = value.showDataLabels;
  if (value.dataLabelPosition != null) dataLabels.position = value.dataLabelPosition;
  if (value.dataLabelSeparator != null) dataLabels.separator = value.dataLabelSeparator;
  if (value.showDataLabelValue != null) dataLabels.showValue = value.showDataLabelValue;
  if (value.showDataLabelCategoryName != null) dataLabels.showCategoryName = value.showDataLabelCategoryName;
  if (value.showDataLabelSeriesName != null) dataLabels.showSeriesName = value.showDataLabelSeriesName;
  if (value.showDataLabelLegendKey != null) dataLabels.showLegendKey = value.showDataLabelLegendKey;
  if (value.showDataLabelPercentage != null) dataLabels.showPercentage = value.showDataLabelPercentage;
  if (value.showDataLabelBubbleSize != null) dataLabels.showBubbleSize = value.showDataLabelBubbleSize;
  if (value.showDataLabelLeaderLines != null) dataLabels.showLeaderLines = value.showDataLabelLeaderLines;
  if (value.dataLabelNumberFormat != null) dataLabels.numberFormat = value.dataLabelNumberFormat;
  if (value.dataLabelFontColor != null) dataLabels.fontColor = value.dataLabelFontColor;
  if (value.dataLabelFontSize != null) dataLabels.fontSize = value.dataLabelFontSize;
  if (Object.keys(dataLabels).length) {
    chart.series.load("items/name");
    await context.sync();
    const visible = toBoolean(dataLabels.visible);
    const shouldEnable = typeof visible === "boolean" ? visible : true;
    for (const series of chart.series.items) {
      series.hasDataLabels = shouldEnable;
      if (dataLabels.showLeaderLines != null && supportsRequirementSet("ExcelApi", "1.9")) {
        series.showLeaderLines = toBoolean(dataLabels.showLeaderLines) ?? false;
      }
    }
    applyExcelChartDataLabelOptions(chart.dataLabels, dataLabels);
  }
}

function getExcelPivotHierarchyCollection(
  pivotTable: Excel.PivotTable,
  axis: "row" | "column" | "filter" | "data",
):
  | Excel.RowColumnPivotHierarchyCollection
  | Excel.FilterPivotHierarchyCollection
  | Excel.DataPivotHierarchyCollection {
  if (axis === "row") {
    return pivotTable.rowHierarchies;
  }
  if (axis === "column") {
    return pivotTable.columnHierarchies;
  }
  if (axis === "filter") {
    return pivotTable.filterHierarchies;
  }
  return pivotTable.dataHierarchies;
}

async function replaceExcelPivotHierarchies(
  context: Excel.RequestContext,
  pivotTable: Excel.PivotTable,
  axis: "row" | "column" | "filter" | "data",
  names: string[],
): Promise<void> {
  const collection = getExcelPivotHierarchyCollection(pivotTable, axis) as any;
  collection.load("items/name");
  await context.sync();
  for (const item of collection.items as Array<{ name: string } & OfficeExtension.ClientObject>) {
    collection.remove(item);
  }
  for (const name of names) {
    collection.add(pivotTable.hierarchies.getItem(name));
  }
}

async function updateExcelPivotHierarchies(
  context: Excel.RequestContext,
  pivotTable: Excel.PivotTable,
  axis: "row" | "column" | "filter" | "data",
  value: unknown,
): Promise<void> {
  if (value == null) {
    return;
  }

  if (Array.isArray(value)) {
    const names = getStringArray(value);
    await replaceExcelPivotHierarchies(context, pivotTable, axis, names);
    return;
  }

  if (!isRecord(value)) {
    return;
  }

  const collection = getExcelPivotHierarchyCollection(pivotTable, axis) as any;
  const replace = getStringArray(value.replace ?? value.set ?? value.names);
  if (replace.length || Array.isArray(value.replace) || Array.isArray(value.set) || Array.isArray(value.names)) {
    await replaceExcelPivotHierarchies(context, pivotTable, axis, replace);
    return;
  }

  const removeNames = getStringArray(value.remove);
  if (removeNames.length) {
    for (const name of removeNames) {
      const existing = collection.getItemOrNullObject(name);
      existing.load("isNullObject");
      await context.sync();
      if (!existing.isNullObject) {
        collection.remove(existing);
      }
    }
  }

  for (const name of getStringArray(value.add)) {
    collection.add(pivotTable.hierarchies.getItem(name));
  }
}

async function ensureExcelPivotHierarchyAssigned(
  context: Excel.RequestContext,
  pivotTable: Excel.PivotTable,
  hierarchyName: string,
): Promise<void> {
  const rowItem = pivotTable.rowHierarchies.getItemOrNullObject(hierarchyName);
  const columnItem = pivotTable.columnHierarchies.getItemOrNullObject(hierarchyName);
  const filterItem = pivotTable.filterHierarchies.getItemOrNullObject(hierarchyName);
  const dataItem = pivotTable.dataHierarchies.getItemOrNullObject(hierarchyName);
  rowItem.load("isNullObject");
  columnItem.load("isNullObject");
  filterItem.load("isNullObject");
  dataItem.load("isNullObject");
  await context.sync();
  if (rowItem.isNullObject && columnItem.isNullObject && filterItem.isNullObject && dataItem.isNullObject) {
    pivotTable.filterHierarchies.add(pivotTable.hierarchies.getItem(hierarchyName));
  }
}

async function applyExcelPivotFilters(
  context: Excel.RequestContext,
  pivotTable: Excel.PivotTable,
  value: unknown,
): Promise<void> {
  if (!supportsRequirementSet("ExcelApi", "1.12")) {
    throw new Error("Excel PivotTable filters require ExcelApi 1.12.");
  }

  const filters = getRecordArray(value);
  for (const filterConfig of filters) {
    const hierarchyName =
      trimString(filterConfig.hierarchyName) ??
      trimString(filterConfig.fieldName) ??
      trimString(filterConfig.name);
    if (!hierarchyName) {
      throw new Error("Pivot filters require hierarchyName or fieldName.");
    }
    await ensureExcelPivotHierarchyAssigned(context, pivotTable, hierarchyName);
    const fieldName = trimString(filterConfig.fieldName) ?? hierarchyName;
    const field = pivotTable.hierarchies.getItem(hierarchyName).fields.getItem(fieldName);

    if (toBoolean(filterConfig.clearAllFilters) ?? false) {
      field.clearAllFilters();
    }
    const clearFilterType = trimString(filterConfig.clearFilterType);
    if (clearFilterType) {
      field.clearFilter(clearFilterType as Excel.PivotFilterType);
    }

    const configuredFilter =
      (isRecord(filterConfig.filter) ? filterConfig.filter : undefined) ??
      ((isRecord(filterConfig.dateFilter) || isRecord(filterConfig.labelFilter) || isRecord(filterConfig.manualFilter) || isRecord(filterConfig.valueFilter)
        ? {
            ...(isRecord(filterConfig.dateFilter) ? { dateFilter: filterConfig.dateFilter } : {}),
            ...(isRecord(filterConfig.labelFilter) ? { labelFilter: filterConfig.labelFilter } : {}),
            ...(isRecord(filterConfig.manualFilter) ? { manualFilter: filterConfig.manualFilter } : {}),
            ...(isRecord(filterConfig.valueFilter) ? { valueFilter: filterConfig.valueFilter } : {}),
          }
        : undefined) as Record<string, unknown> | undefined);

    if (configuredFilter && Object.keys(configuredFilter).length) {
      field.applyFilter(configuredFilter as unknown as Excel.PivotFilters);
    }
  }
}

async function applyExcelPivotConfiguration(
  context: Excel.RequestContext,
  pivotTable: Excel.PivotTable,
  value: Record<string, unknown>,
): Promise<void> {
  const nextName = trimString(value.pivotTableName) ?? trimString(value.name);
  if (nextName) {
    pivotTable.name = nextName;
  }

  const allowMultipleFiltersPerField = toBoolean(value.allowMultipleFiltersPerField);
  const enableDataValueEditing = toBoolean(value.enableDataValueEditing);
  const refreshOnOpen = toBoolean(value.refreshOnOpen);
  const useCustomSortLists = toBoolean(value.useCustomSortLists);
  if (typeof allowMultipleFiltersPerField === "boolean") {
    pivotTable.allowMultipleFiltersPerField = allowMultipleFiltersPerField;
  }
  if (typeof enableDataValueEditing === "boolean") {
    pivotTable.enableDataValueEditing = enableDataValueEditing;
  }
  if (typeof refreshOnOpen === "boolean" && supportsRequirementSet("ExcelApi", "1.13")) {
    pivotTable.refreshOnOpen = refreshOnOpen;
  }
  if (typeof useCustomSortLists === "boolean" && supportsRequirementSet("ExcelApi", "1.9")) {
    pivotTable.useCustomSortLists = useCustomSortLists;
  }

  await updateExcelPivotHierarchies(context, pivotTable, "row", value.rowHierarchies);
  await updateExcelPivotHierarchies(context, pivotTable, "column", value.columnHierarchies);
  await updateExcelPivotHierarchies(context, pivotTable, "filter", value.filterHierarchies);
  await updateExcelPivotHierarchies(context, pivotTable, "data", value.dataHierarchies);

  if (isRecord(value.layout)) {
    const layout = value.layout;
    const layoutType = trimString(layout.layoutType);
    const emptyCellText = trimString(layout.emptyCellText);
    const fillEmptyCells = toBoolean(layout.fillEmptyCells);
    const showColumnGrandTotals = toBoolean(layout.showColumnGrandTotals);
    const showRowGrandTotals = toBoolean(layout.showRowGrandTotals);
    const subtotalLocation = trimString(layout.subtotalLocation);
    if (layoutType) {
      pivotTable.layout.layoutType = layoutType as Excel.PivotLayoutType;
    }
    if (emptyCellText) {
      pivotTable.layout.emptyCellText = emptyCellText;
    }
    if (typeof fillEmptyCells === "boolean") {
      pivotTable.layout.fillEmptyCells = fillEmptyCells;
    }
    if (typeof showColumnGrandTotals === "boolean") {
      pivotTable.layout.showColumnGrandTotals = showColumnGrandTotals;
    }
    if (typeof showRowGrandTotals === "boolean") {
      pivotTable.layout.showRowGrandTotals = showRowGrandTotals;
    }
    if (subtotalLocation) {
      pivotTable.layout.subtotalLocation = subtotalLocation as Excel.SubtotalLocationType;
    }
  }

  if (value.filters != null || value.pivotFilters != null) {
    await applyExcelPivotFilters(context, pivotTable, value.pivotFilters ?? value.filters);
  }

  if (value.sort != null || value.pivotSorts != null || value.sorts != null) {
    await applyExcelPivotSorts(context, pivotTable, value.pivotSorts ?? value.sorts ?? value.sort);
  }

  if (toBoolean(value.refresh) ?? false) {
    pivotTable.refresh();
  }
}

function resolveExcelFormattingRange(
  context: Excel.RequestContext,
  target?: OfficeAnchor,
  allowSelected = true,
): Excel.Range {
  if (target?.tableName) {
    return resolveExcelTable(context, target).getRange();
  }

  return resolveExcelRange(context, target, allowSelected);
}

function applyExcelRangeFormatting(range: Excel.Range, action: OfficeHostAction): void {
  const options = { ...getActionOptions(action), ...action };
  const format = range.format;
  const font = format.font;
  const fill = format.fill;

  const fontName = trimString(options.fontName);
  const fontColor = trimString(options.fontColor);
  const fillColor = trimString(options.fillColor);
  const horizontalAlignment = trimString(options.horizontalAlignment);
  const verticalAlignment = trimString(options.verticalAlignment);
  const numberFormat = toStringMatrix(options.numberFormat);
  const rowHeight = toNumber(options.rowHeight);
  const columnWidth = toNumber(options.columnWidth);
  const bold = toBoolean(options.bold);
  const italic = toBoolean(options.italic);
  const underline = trimString(options.underline);
  const wrapText = toBoolean(options.wrapText);

  if (fontName) font.name = fontName;
  if (fontColor) font.color = fontColor;
  if (fillColor) fill.color = fillColor;
  if (typeof bold === "boolean") font.bold = bold;
  if (typeof italic === "boolean") font.italic = italic;
  if (underline) font.underline = underline as Excel.RangeUnderlineStyle;
  if (horizontalAlignment) format.horizontalAlignment = horizontalAlignment as Excel.HorizontalAlignment;
  if (verticalAlignment) format.verticalAlignment = verticalAlignment as Excel.VerticalAlignment;
  if (typeof wrapText === "boolean") format.wrapText = wrapText;
  if (typeof rowHeight === "number") format.rowHeight = rowHeight;
  if (typeof columnWidth === "number") format.columnWidth = columnWidth;
  if (numberFormat) range.numberFormat = numberFormat;
  if (toBoolean(options.autoFitColumns)) format.autofitColumns();
  if (toBoolean(options.autoFitRows)) format.autofitRows();
  applyExcelRangeBorders(range, options);
}

function applyExcelTableFormatting(table: Excel.Table, action: OfficeHostAction): void {
  const options = { ...getActionOptions(action), ...action };
  const style = trimString(options.tableStyle) ?? trimString(options.style);
  const showHeaders = toBoolean(options.showHeaders ?? options.headersVisible ?? options.headerRowVisible);
  const showTotals = toBoolean(options.showTotals ?? options.totalsVisible ?? options.totalRowVisible);
  const showBandedRows = toBoolean(options.showBandedRows ?? options.bandedRows ?? options.areRowsBanded);
  const showBandedColumns = toBoolean(options.showBandedColumns ?? options.bandedColumns ?? options.areColumnsBanded);
  const highlightFirstColumn = toBoolean(options.highlightFirstColumn ?? options.firstColumn ?? options.isFirstColumnHighlighted);
  const highlightLastColumn = toBoolean(options.highlightLastColumn ?? options.lastColumn ?? options.isLastColumnHighlighted);
  const showFilterButton = toBoolean(options.showFilterButton ?? options.filterButtonsVisible ?? options.filterButtonVisible);
  const requiresExcel13 = [showBandedRows, showBandedColumns, highlightFirstColumn, highlightLastColumn, showFilterButton].some(
    (value) => typeof value === "boolean",
  );

  if (requiresExcel13 && !supportsRequirementSet("ExcelApi", "1.3")) {
    throw new Error("Excel table style toggles require ExcelApi 1.3.");
  }

  if (style) {
    table.style = style;
  }
  if (typeof showHeaders === "boolean") {
    table.showHeaders = showHeaders;
  }
  if (typeof showTotals === "boolean") {
    table.showTotals = showTotals;
  }
  if (typeof showBandedRows === "boolean") {
    table.showBandedRows = showBandedRows;
  }
  if (typeof showBandedColumns === "boolean") {
    table.showBandedColumns = showBandedColumns;
  }
  if (typeof highlightFirstColumn === "boolean") {
    table.highlightFirstColumn = highlightFirstColumn;
  }
  if (typeof highlightLastColumn === "boolean") {
    table.highlightLastColumn = highlightLastColumn;
  }
  if (typeof showFilterButton === "boolean") {
    if (showFilterButton && showHeaders === false) {
      throw new Error("Excel table filter buttons require showHeaders=true.");
    }
    table.showFilterButton = showFilterButton;
  }

  applyExcelRangeFormatting(table.getRange(), action);
}

function applyConditionalRangeStyle(format: Excel.ConditionalRangeFormat, value: unknown): void {
  if (!isRecord(value)) {
    return;
  }

  const fillColor = trimString(value.fillColor);
  const fontColor = trimString(value.fontColor);
  const numberFormat = trimString(value.numberFormat);
  const bold = toBoolean(value.bold);
  const italic = toBoolean(value.italic);
  const underline = trimString(value.underline);

  if (fillColor) format.fill.color = fillColor;
  if (fontColor) format.font.color = fontColor;
  if (typeof bold === "boolean") format.font.bold = bold;
  if (typeof italic === "boolean") format.font.italic = italic;
  if (underline) format.font.underline = underline as Excel.ConditionalRangeFontUnderlineStyle;
  if (numberFormat) format.numberFormat = numberFormat;
}

function buildExcelConditionalColorScaleCriterion(
  value: unknown,
  fallback: Excel.ConditionalColorScaleCriterion,
): Excel.ConditionalColorScaleCriterion {
  if (!isRecord(value)) {
    const formula = trimString(value);
    return formula ? { ...fallback, formula } : fallback;
  }

  const criterion: Excel.ConditionalColorScaleCriterion = {
    type: (trimString(value.type) ?? trimString(value.ruleType) ?? fallback.type) as Excel.ConditionalFormatColorCriterionType,
  };
  const formula = trimString(value.formula) ?? trimString(value.value);
  const color = trimString(value.color) ?? fallback.color;
  if (formula) {
    criterion.formula = formula;
  }
  if (color) {
    criterion.color = color;
  }
  return criterion;
}

function buildExcelConditionalIconCriterion(value: Record<string, unknown>): Excel.ConditionalIconCriterion {
  const criterion: Excel.ConditionalIconCriterion = {
    type: (trimString(value.type) ?? trimString(value.ruleType) ?? "Percent") as Excel.ConditionalFormatIconRuleType,
    formula: trimString(value.formula) ?? trimString(value.value) ?? "0",
    operator: (trimString(value.operator) ?? "GreaterThanOrEqual") as Excel.ConditionalIconCriterionOperator,
  };
  const customIconSet = trimString(value.customIconSet);
  const customIconIndex = toNumber(value.customIconIndex);
  if (customIconSet || typeof customIconIndex === "number") {
    criterion.customIcon = {
      set: (customIconSet ?? "ThreeArrows") as Excel.IconSet,
      index: customIconIndex ?? 0,
    } as Excel.Icon;
  }
  return criterion;
}

async function navigateExcelAnchor(anchor: OfficeAnchor): Promise<unknown> {
  return Excel.run(async (context) => {
    if (anchor.kind === "workbook") {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      worksheet.load("name");
      worksheet.activate();
      await context.sync();
      return {
        ok: true,
        host: "excel",
        anchorKind: "workbook",
        sheetName: worksheet.name,
        label: anchor.label,
      };
    }

    if (anchor.kind === "sheet" || anchor.kind === "worksheet") {
      const worksheet = resolveExcelWorksheet(context, anchor, false);
      worksheet.activate();
      await context.sync();
      return { ok: true, host: "excel", anchorKind: anchor.kind, sheetName: anchor.sheetName };
    }

    if (anchor.kind === "namedItem") {
      const range = resolveExcelRange(context, anchor, false);
      range.load("address");
      range.worksheet.load("name");
      range.worksheet.activate();
      range.select();
      await context.sync();
      return {
        ok: true,
        host: "excel",
        anchorKind: "namedItem",
        namedItemName: anchor.namedItemName,
        address: range.address,
        sheetName: range.worksheet.name,
      };
    }

    if (anchor.kind === "table") {
      const table = resolveExcelTable(context, anchor);
      const range = table.getRange();
      const worksheet = table.worksheet;
      range.load("address");
      worksheet.load("name");
      worksheet.activate();
      range.select();
      await context.sync();
      return { ok: true, host: "excel", anchorKind: "table", tableName: anchor.tableName, address: range.address, sheetName: worksheet.name };
    }

    if (anchor.kind === "chart") {
      const chart = resolveExcelChart(context, anchor);
      chart.activate();
      chart.load("name");
      await context.sync();
      return { ok: true, host: "excel", anchorKind: "chart", chartName: chart.name };
    }

    if (anchor.kind === "pivotTable") {
      const pivotTable = resolveExcelPivotTable(context, anchor);
      const range = pivotTable.layout.getRange();
      const worksheet = pivotTable.worksheet;
      range.load("address");
      worksheet.load("name");
      worksheet.activate();
      range.select();
      await context.sync();
      return {
        ok: true,
        host: "excel",
        anchorKind: "pivotTable",
        pivotTableName: anchor.pivotTableName,
        address: range.address,
        sheetName: worksheet.name,
      };
    }

    const range = resolveExcelRange(context, anchor, true);
    const worksheet = resolveExcelWorksheet(context, anchor, true);
    range.load("address");
    worksheet.load("name");
    worksheet.activate();
    range.select();
    await context.sync();
    return { ok: true, host: "excel", anchorKind: anchor.kind, address: range.address, sheetName: worksheet.name };
  });
}

async function applyExcelAction(action: OfficeHostAction): Promise<unknown> {
  return Excel.run(async (context) => {
    const type = trimString(action.type) ?? "setRangeValues";
    const options = getActionOptions(action);

    if (type === "insertText" || type === "setRangeValues") {
      const range = resolveExcelRange(context, action.target, true);
      range.values = toValueMatrix(action.values ?? action.content, action.content);
      range.load("address");
      await context.sync();
      return { ok: true, host: "excel", action: type, address: range.address };
    }

    if (type === "setRangeFormulas") {
      const range = resolveExcelRange(context, action.target, false);
      const formulas = toStringMatrix(action.formulas ?? action.content);
      if (!formulas) {
        throw new Error("Excel formulas must be provided as a matrix.");
      }
      range.formulas = formulas;
      range.load("address");
      await context.sync();
      return { ok: true, host: "excel", action: type, address: range.address };
    }

    if (type === "setRangeNumberFormat") {
      const range = resolveExcelRange(context, action.target, false);
      const numberFormat = toStringMatrix(action.numberFormat ?? action.content);
      if (!numberFormat) {
        throw new Error("Excel number formats must be provided as a matrix.");
      }
      range.numberFormat = numberFormat;
      range.load("address");
      await context.sync();
      return { ok: true, host: "excel", action: type, address: range.address };
    }

    if (type === "formatRange") {
      const range = resolveExcelFormattingRange(context, action.target, false);
      applyExcelRangeFormatting(range, action);
      range.load("address");
      await context.sync();
      return {
        ok: true,
        host: "excel",
        action: type,
        address: range.address,
        tableName: action.target?.tableName,
      };
    }

    if (type === "insertRows") {
      const insertedRange = resolveExcelRange(context, action.target, true).getEntireRow().insert(Excel.InsertShiftDirection.down);
      insertedRange.load("address,rowCount");
      insertedRange.worksheet.load("name");
      await context.sync();
      return {
        ok: true,
        host: "excel",
        action: type,
        address: insertedRange.address,
        sheetName: insertedRange.worksheet.name,
        rowCount: insertedRange.rowCount,
      };
    }

    if (type === "insertColumns") {
      const insertedRange = resolveExcelRange(context, action.target, true).getEntireColumn().insert(Excel.InsertShiftDirection.right);
      insertedRange.load("address,columnCount");
      insertedRange.worksheet.load("name");
      await context.sync();
      return {
        ok: true,
        host: "excel",
        action: type,
        address: insertedRange.address,
        sheetName: insertedRange.worksheet.name,
        columnCount: insertedRange.columnCount,
      };
    }

    if (type === "deleteRows") {
      const rowRange = resolveExcelRange(context, action.target, true).getEntireRow();
      rowRange.load("address,rowCount");
      rowRange.worksheet.load("name");
      await context.sync();
      const address = rowRange.address;
      const sheetName = rowRange.worksheet.name;
      const rowCount = rowRange.rowCount;
      rowRange.delete(Excel.DeleteShiftDirection.up);
      await context.sync();
      return { ok: true, host: "excel", action: type, address, sheetName, rowCount };
    }

    if (type === "deleteColumns") {
      const columnRange = resolveExcelRange(context, action.target, true).getEntireColumn();
      columnRange.load("address,columnCount");
      columnRange.worksheet.load("name");
      await context.sync();
      const address = columnRange.address;
      const sheetName = columnRange.worksheet.name;
      const columnCount = columnRange.columnCount;
      columnRange.delete(Excel.DeleteShiftDirection.left);
      await context.sync();
      return { ok: true, host: "excel", action: type, address, sheetName, columnCount };
    }

    if (type === "sortRange") {
      const range = resolveExcelRange(context, action.target, false);
      const fields: Excel.SortField[] = getRecordArray(options.fields).map((field) => {
        const sortField: Excel.SortField = {
          key: toNumber(field.key) ?? 0,
          ascending: toBoolean(field.ascending) ?? true,
        };
        const sortOn = trimString(field.sortOn);
        const dataOption = trimString(field.dataOption);
        const color = trimString(field.color);
        const subField = trimString(field.subField);
        if (sortOn) sortField.sortOn = sortOn as Excel.SortOn;
        if (dataOption) sortField.dataOption = dataOption as Excel.SortDataOption;
        if (color) sortField.color = color;
        if (subField) sortField.subField = subField;
        return sortField;
      });
      if (!fields.length) {
        fields.push({
          key: toNumber(options.columnIndex) ?? 0,
          ascending: toBoolean(options.ascending) ?? true,
        });
      }
      range.sort.apply(
        fields,
        toBoolean(options.matchCase) ?? false,
        toBoolean(options.hasHeaders),
        trimString(options.orientation) as Excel.SortOrientation,
        trimString(options.method) as Excel.SortMethod,
      );
      range.load("address");
      await context.sync();
      return { ok: true, host: "excel", action: type, address: range.address };
    }

    if (type === "applyFilter") {
      const worksheet = resolveExcelWorksheet(context, action.target, true);
      const range = resolveExcelRange(context, action.target, false);
      worksheet.autoFilter.apply(
        range,
        toNumber(options.columnIndex),
        isRecord(options.criteria) ? (options.criteria as unknown as Excel.FilterCriteria) : undefined,
      );
      range.load("address");
      worksheet.load("name");
      await context.sync();
      return { ok: true, host: "excel", action: type, address: range.address, sheetName: worksheet.name };
    }

    if (type === "applyTableFilter" || type === "filterTable") {
      const table = resolveExcelTable(
        context,
        action.target ?? { kind: "table", tableName: trimString(options.tableName) ?? "", sheetName: trimString(options.sheetName) },
      );
      const column = resolveExcelTableColumn(table, options);
      table.load("name,id");
      table.worksheet.load("name");
      column.load("name,index");
      await context.sync();
      applyExcelTableFilter(column, options);
      await context.sync();
      return {
        ok: true,
        host: "excel",
        action: type,
        tableName: table.name,
        tableId: table.id,
        sheetName: table.worksheet.name,
        columnName: column.name,
        columnIndex: column.index,
      };
    }

    if (type === "clearTableFilter") {
      const table = resolveExcelTable(
        context,
        action.target ?? { kind: "table", tableName: trimString(options.tableName) ?? "", sheetName: trimString(options.sheetName) },
      );
      const column = resolveExcelTableColumn(table, options);
      table.load("name,id");
      table.worksheet.load("name");
      column.load("name,index");
      await context.sync();
      column.filter.clear();
      await context.sync();
      return {
        ok: true,
        host: "excel",
        action: type,
        tableName: table.name,
        tableId: table.id,
        sheetName: table.worksheet.name,
        columnName: column.name,
        columnIndex: column.index,
      };
    }

    if (type === "clearTableFilters") {
      const table = resolveExcelTable(
        context,
        action.target ?? { kind: "table", tableName: trimString(options.tableName) ?? "", sheetName: trimString(options.sheetName) },
      );
      table.load("name,id");
      table.worksheet.load("name");
      await context.sync();
      table.clearFilters();
      await context.sync();
      return {
        ok: true,
        host: "excel",
        action: type,
        tableName: table.name,
        tableId: table.id,
        sheetName: table.worksheet.name,
      };
    }

    if (type === "reapplyTableFilters") {
      const table = resolveExcelTable(
        context,
        action.target ?? { kind: "table", tableName: trimString(options.tableName) ?? "", sheetName: trimString(options.sheetName) },
      );
      table.load("name,id");
      table.worksheet.load("name");
      await context.sync();
      table.reapplyFilters();
      await context.sync();
      return {
        ok: true,
        host: "excel",
        action: type,
        tableName: table.name,
        tableId: table.id,
        sheetName: table.worksheet.name,
      };
    }

    if (type === "removeDuplicates") {
      if (!supportsRequirementSet("ExcelApi", "1.9")) {
        throw new Error("Excel duplicate removal requires ExcelApi 1.9.");
      }
      const range = resolveExcelRange(context, action.target, false);
      const columns = getNumberArray(options.columns ?? options.columnIndexes);
      const removeDuplicatesResult = range.removeDuplicates(columns.length ? columns : [0], toBoolean(options.includesHeader) ?? true);
      range.load("address");
      range.worksheet.load("name");
      removeDuplicatesResult.load("removed,uniqueRemaining");
      await context.sync();
      return {
        ok: true,
        host: "excel",
        action: type,
        address: range.address,
        sheetName: range.worksheet.name,
        removed: removeDuplicatesResult.removed,
        uniqueRemaining: removeDuplicatesResult.uniqueRemaining,
      };
    }

    if (type === "createWorksheet") {
      const name = trimString(options.name) ?? `Sheet_${Date.now()}`;
      const worksheet = context.workbook.worksheets.add(name);
      worksheet.load("name,position");
      await context.sync();
      worksheet.activate();
      await context.sync();
      return { ok: true, host: "excel", action: type, sheetName: worksheet.name, position: worksheet.position + 1 };
    }

    if (type === "renameWorksheet") {
      const worksheet = resolveExcelWorksheet(context, action.target, false);
      const nextName = trimString(options.name);
      if (!nextName) {
        throw new Error("Excel worksheet rename requires a name.");
      }
      worksheet.name = nextName;
      worksheet.load("name,position");
      await context.sync();
      return { ok: true, host: "excel", action: type, sheetName: worksheet.name, position: worksheet.position + 1 };
    }

    if (type === "duplicateWorksheet") {
      const worksheet = resolveExcelWorksheet(context, action.target, false);
      const relativeToName = trimString(options.relativeTo);
      const relativeTo = relativeToName ? context.workbook.worksheets.getItem(relativeToName) : undefined;
      const copy = worksheet.copy(trimString(options.positionType) as Excel.WorksheetPositionType, relativeTo);
      const nextName = trimString(options.name);
      if (nextName) {
        copy.name = nextName;
      }
      copy.load("name,position");
      await context.sync();
      return { ok: true, host: "excel", action: type, sheetName: copy.name, position: copy.position + 1 };
    }

    if (type === "deleteWorksheet") {
      const worksheet = resolveExcelWorksheet(context, action.target, false);
      worksheet.load("name,visibility");
      await context.sync();
      const sheetName = worksheet.name;
      if (worksheet.visibility === "VeryHidden") {
        worksheet.visibility = Excel.SheetVisibility.hidden;
      }
      worksheet.delete();
      await context.sync();
      return { ok: true, host: "excel", action: type, sheetName };
    }

    if (type === "setWorksheetGridlines") {
      if (!supportsRequirementSet("ExcelApi", "1.8")) {
        throw new Error("Excel worksheet gridline controls require ExcelApi 1.8.");
      }
      const worksheet = resolveExcelWorksheet(context, action.target, true);
      const visible = toBoolean(options.visible);
      if (typeof visible !== "boolean") {
        throw new Error("Excel worksheet gridline updates require visible: true or false.");
      }
      worksheet.showGridlines = visible;
      worksheet.load("name,showGridlines");
      await context.sync();
      return { ok: true, host: "excel", action: type, sheetName: worksheet.name, showGridlines: worksheet.showGridlines };
    }

    if (type === "setWorksheetHeadings") {
      if (!supportsRequirementSet("ExcelApi", "1.8")) {
        throw new Error("Excel worksheet heading controls require ExcelApi 1.8.");
      }
      const worksheet = resolveExcelWorksheet(context, action.target, true);
      const visible = toBoolean(options.visible);
      if (typeof visible !== "boolean") {
        throw new Error("Excel worksheet heading updates require visible: true or false.");
      }
      worksheet.showHeadings = visible;
      worksheet.load("name,showHeadings");
      await context.sync();
      return { ok: true, host: "excel", action: type, sheetName: worksheet.name, showHeadings: worksheet.showHeadings };
    }

    if (type === "setPrintArea") {
      if (!supportsRequirementSet("ExcelApi", "1.9")) {
        throw new Error("Excel print area controls require ExcelApi 1.9.");
      }
      const explicitPrintArea = trimString(options.address) ?? trimString(options.printArea);
      const parsedPrintArea = splitSheetAddress(explicitPrintArea, action.target?.sheetName);
      const worksheet = resolveExcelWorksheet(
        context,
        parsedPrintArea.sheetName ? ({ kind: "sheet", sheetName: parsedPrintArea.sheetName } as OfficeAnchor) : action.target,
        true,
      );
      const printAreaRange = parsedPrintArea.address
        ? worksheet.getRange(parsedPrintArea.address)
        : resolveExcelRange(context, action.target, true);
      worksheet.pageLayout.setPrintArea(printAreaRange);
      const printArea = worksheet.pageLayout.getPrintAreaOrNullObject();
      worksheet.load("name");
      printArea.load("isNullObject,address");
      await context.sync();
      return {
        ok: true,
        host: "excel",
        action: type,
        sheetName: worksheet.name,
        address: printArea.isNullObject ? undefined : printArea.address,
      };
    }

    if (type === "createTable") {
      const range = resolveExcelRange(context, action.target, false);
      const table = context.workbook.tables.add(range, toBoolean(options.hasHeaders) ?? true);
      const tableName = trimString(options.tableName) ?? trimString(action.tableName);
      if (tableName) {
        table.name = tableName;
      }
      applyExcelTableFormatting(table, action);
      table.load("name,id");
      range.load("address");
      await context.sync();
      return { ok: true, host: "excel", action: type, tableName: table.name, tableId: table.id, address: range.address };
    }

    if (type === "formatTable" || type === "updateTableStyle" || type === "configureTable" || type === "setTableStyle") {
      const table = resolveExcelTable(
        context,
        action.target ?? { kind: "table", tableName: trimString(options.tableName) ?? "", sheetName: trimString(options.sheetName) },
      );
      const range = table.getRange();
      applyExcelTableFormatting(table, action);
      table.load("name,id");
      table.worksheet.load("name");
      range.load("address");
      await context.sync();
      return {
        ok: true,
        host: "excel",
        action: type,
        tableName: table.name,
        tableId: table.id,
        sheetName: table.worksheet.name,
        address: range.address,
      };
    }

    if (type === "createChart") {
      const source = trimString(options.source) ?? withSheetName(action.target?.sheetName, action.target?.address);
      const parsedSource = splitSheetAddress(source, action.target?.sheetName);
      const sourceRange = parsedSource.address
        ? resolveExcelWorksheet(context, { kind: "sheet", sheetName: parsedSource.sheetName }, true).getRange(parsedSource.address)
        : resolveExcelRange(context, action.target, true);
      const worksheet = resolveExcelWorksheet(context, action.target, true);
      const chart = worksheet.charts.add(
        (trimString(options.chartType) ?? "ColumnClustered") as Excel.ChartType,
        sourceRange,
        (trimString(options.seriesBy) ?? "Auto") as Excel.ChartSeriesBy,
      );
      const chartName = trimString(options.chartName);
      if (chartName) {
        chart.name = chartName;
      }
      const chartTitle = trimString(options.title);
      if (chartTitle) {
        chart.title.text = chartTitle;
        chart.title.visible = true;
      }
      const chartLeft = toNumber(options.left);
      const chartTop = toNumber(options.top);
      const chartWidth = toNumber(options.width);
      const chartHeight = toNumber(options.height);
      const legendVisible = toBoolean(options.legendVisible);
      if (typeof chartLeft === "number") chart.left = chartLeft;
      if (typeof chartTop === "number") chart.top = chartTop;
      if (typeof chartWidth === "number") chart.width = chartWidth;
      if (typeof chartHeight === "number") chart.height = chartHeight;
      if (typeof legendVisible === "boolean") chart.legend.visible = legendVisible;
      await applyExcelChartConfiguration(context, chart, options);
      chart.load("name,id");
      await context.sync();
      return { ok: true, host: "excel", action: type, chartName: chart.name, chartId: chart.id };
    }

    if (type === "updateChart" || type === "formatChart" || type === "setChartAxes" || type === "setChartDataLabels") {
      const chart = resolveExcelChart(context, action.target ?? { kind: "chart", chartName: trimString(options.chartName) ?? "" });
      await applyExcelChartConfiguration(context, chart, options);
      chart.load("name,id");
      await context.sync();
      return { ok: true, host: "excel", action: type, chartName: chart.name, chartId: chart.id };
    }

    if (type === "createPivotTable") {
      const source = trimString(options.source) ?? withSheetName(action.target?.sheetName, action.target?.address);
      const destination = trimString(options.destination);
      if (!source || !destination) {
        throw new Error("Excel PivotTable creation requires source and destination.");
      }
      const name = trimString(options.pivotTableName) ?? `Pivot_${Date.now()}`;
      const pivotTable = context.workbook.pivotTables.add(name, source, destination);
      for (const hierarchyName of getStringArray(options.rowHierarchies)) {
        pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(hierarchyName));
      }
      for (const hierarchyName of getStringArray(options.columnHierarchies)) {
        pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem(hierarchyName));
      }
      for (const hierarchyName of getStringArray(options.filterHierarchies)) {
        pivotTable.filterHierarchies.add(pivotTable.hierarchies.getItem(hierarchyName));
      }
      for (const hierarchyName of getStringArray(options.dataHierarchies)) {
        pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(hierarchyName));
      }
      await applyExcelPivotConfiguration(context, pivotTable, options);
      pivotTable.load("name,id");
      await context.sync();
      return { ok: true, host: "excel", action: type, pivotTableName: pivotTable.name, pivotTableId: pivotTable.id };
    }

    if (type === "updatePivotTable" || type === "configurePivotTable" || type === "applyPivotFilter") {
      const pivotTable = resolveExcelPivotTable(
        context,
        action.target ?? { kind: "pivotTable", pivotTableName: trimString(options.pivotTableName) ?? "" },
      );
      await applyExcelPivotConfiguration(context, pivotTable, options);
      pivotTable.load("name,id");
      await context.sync();
      return { ok: true, host: "excel", action: type, pivotTableName: pivotTable.name, pivotTableId: pivotTable.id };
    }

    if (type === "sortPivotField" || type === "sortPivotByLabels" || type === "sortPivotByValues") {
      const pivotTable = resolveExcelPivotTable(
        context,
        action.target ?? { kind: "pivotTable", pivotTableName: trimString(options.pivotTableName) ?? "" },
      );
      const sortOptions = {
        ...options,
        mode:
          type === "sortPivotByValues"
            ? "values"
            : type === "sortPivotByLabels"
              ? "labels"
              : trimString(options.mode),
      };
      await applyExcelPivotSorts(context, pivotTable, sortOptions);
      pivotTable.load("name,id");
      await context.sync();
      return { ok: true, host: "excel", action: type, pivotTableName: pivotTable.name, pivotTableId: pivotTable.id };
    }

    if (type === "refreshPivotTable") {
      const pivotTable = resolveExcelPivotTable(
        context,
        action.target ?? { kind: "pivotTable", pivotTableName: trimString(options.pivotTableName) ?? "" },
      );
      pivotTable.refresh();
      pivotTable.load("name,id");
      await context.sync();
      return { ok: true, host: "excel", action: type, pivotTableName: pivotTable.name, pivotTableId: pivotTable.id };
    }

    if (type === "setDataValidation") {
      const range = resolveExcelRange(context, action.target, false);
      if (!isRecord(options.validation)) {
        throw new Error("Excel data validation requires a validation object.");
      }
      range.dataValidation.set(options.validation as Excel.Interfaces.DataValidationUpdateData);
      range.load("address");
      await context.sync();
      return { ok: true, host: "excel", action: type, address: range.address };
    }

    if (type === "clearDataValidation") {
      const range = resolveExcelRange(context, action.target, false);
      range.dataValidation.clear();
      range.load("address");
      await context.sync();
      return { ok: true, host: "excel", action: type, address: range.address };
    }

    if (type === "addConditionalFormat") {
      const range = resolveExcelRange(context, action.target, false);
      const conditionalType = (trimString(options.conditionalType) ?? trimString(options.type) ?? "Custom") as Excel.ConditionalFormatType;
      const conditionalFormat = range.conditionalFormats.add(conditionalType);
      const style = isRecord(options.style) ? options.style : undefined;

      if (conditionalType === "Custom") {
        conditionalFormat.changeRuleToCustom(trimString(options.formula) ?? "=TRUE");
        applyConditionalRangeStyle(conditionalFormat.custom.format, style);
      } else if (conditionalType === "CellValue") {
        const rule: Excel.ConditionalCellValueRule = {
          formula1: trimString(options.formula1) ?? "=0",
          operator: (trimString(options.operator) ?? "GreaterThan") as Excel.ConditionalCellValueOperator,
        };
        const formula2 = trimString(options.formula2);
        if (formula2) {
          rule.formula2 = formula2;
        }
        conditionalFormat.changeRuleToCellValue(rule);
        applyConditionalRangeStyle(conditionalFormat.cellValue.format, style);
      } else if (conditionalType === "DataBar") {
        conditionalFormat.changeRuleToDataBar();
        const dataBar = conditionalFormat.dataBar;
        const positiveFillColor = trimString(options.positiveFillColor);
        const positiveBorderColor = trimString(options.positiveBorderColor);
        const negativeFillColor = trimString(options.negativeFillColor);
        const negativeBorderColor = trimString(options.negativeBorderColor);
        const axisColor = trimString(options.axisColor);
        const axisFormat = trimString(options.axisFormat);
        const barDirection = trimString(options.barDirection);
        const showDataBarOnly = toBoolean(options.showDataBarOnly);
        if (axisColor) dataBar.axisColor = axisColor;
        if (axisFormat) dataBar.axisFormat = axisFormat as Excel.ConditionalDataBarAxisFormat;
        if (barDirection) dataBar.barDirection = barDirection as Excel.ConditionalDataBarDirection;
        if (typeof showDataBarOnly === "boolean") dataBar.showDataBarOnly = showDataBarOnly;
        if (positiveFillColor) dataBar.positiveFormat.fillColor = positiveFillColor;
        if (positiveBorderColor) dataBar.positiveFormat.borderColor = positiveBorderColor;
        if (negativeFillColor) dataBar.negativeFormat.fillColor = negativeFillColor;
        if (negativeBorderColor) dataBar.negativeFormat.borderColor = negativeBorderColor;

        const lowerBoundType = trimString(options.lowerBoundType);
        const upperBoundType = trimString(options.upperBoundType);
        const lowerBoundFormula = trimString(options.lowerBoundFormula);
        const upperBoundFormula = trimString(options.upperBoundFormula);
        if (lowerBoundType) {
          const lowerBoundRule: Excel.ConditionalDataBarRule = {
            type: lowerBoundType as Excel.ConditionalFormatRuleType,
          };
          if (lowerBoundFormula) lowerBoundRule.formula = lowerBoundFormula;
          dataBar.lowerBoundRule = lowerBoundRule;
        }
        if (upperBoundType) {
          const upperBoundRule: Excel.ConditionalDataBarRule = {
            type: upperBoundType as Excel.ConditionalFormatRuleType,
          };
          if (upperBoundFormula) upperBoundRule.formula = upperBoundFormula;
          dataBar.upperBoundRule = upperBoundRule;
        }
      } else if (conditionalType === "ColorScale") {
        conditionalFormat.changeRuleToColorScale();
        const colorScale = conditionalFormat.colorScale;
        const minimum = buildExcelConditionalColorScaleCriterion(
          isRecord(options.minimum)
            ? options.minimum
            : {
                type: options.minimumType,
                formula: options.minimumFormula,
                color: options.minimumColor,
              },
          { type: "LowestValue", color: trimString(options.minimumColor) ?? "#F8696B" },
        );
        const maximum = buildExcelConditionalColorScaleCriterion(
          isRecord(options.maximum)
            ? options.maximum
            : {
                type: options.maximumType,
                formula: options.maximumFormula,
                color: options.maximumColor,
              },
          { type: "HighestValue", color: trimString(options.maximumColor) ?? "#63BE7B" },
        );
        const includeMidpoint =
          toBoolean(options.threeColorScale) ??
          Boolean(isRecord(options.midpoint) || trimString(options.midpointType) || trimString(options.midpointFormula) || trimString(options.midpointColor));
        const criteria: Excel.ConditionalColorScaleCriteria = { minimum, maximum };
        if (includeMidpoint) {
          criteria.midpoint = buildExcelConditionalColorScaleCriterion(
            isRecord(options.midpoint)
              ? options.midpoint
              : {
                  type: options.midpointType,
                  formula: options.midpointFormula,
                  color: options.midpointColor,
                },
            {
              type: "Percentile",
              formula: "50",
              color: trimString(options.midpointColor) ?? "#FFEB84",
            },
          );
        }
        colorScale.criteria = criteria;
      } else if (conditionalType === "IconSet") {
        conditionalFormat.changeRuleToIconSet();
        const iconSet = conditionalFormat.iconSet;
        const iconSetStyle = trimString(options.iconSetStyle) ?? trimString(options.style);
        const reverseIconOrder = toBoolean(options.reverseIconOrder);
        const showIconOnly = toBoolean(options.showIconOnly);
        const criteria = getRecordArray(options.criteria).map((criterion) => buildExcelConditionalIconCriterion(criterion));
        if (iconSetStyle) {
          iconSet.style = iconSetStyle as Excel.IconSet;
        }
        if (typeof reverseIconOrder === "boolean") {
          iconSet.reverseIconOrder = reverseIconOrder;
        }
        if (typeof showIconOnly === "boolean") {
          iconSet.showIconOnly = showIconOnly;
        }
        if (criteria.length) {
          iconSet.criteria = criteria;
        }
      } else {
        throw new Error(`Unsupported Excel conditional format type: ${conditionalType}`);
      }

      conditionalFormat.load("id");
      range.load("address");
      await context.sync();
      return {
        ok: true,
        host: "excel",
        action: type,
        address: range.address,
        conditionalFormatId: conditionalFormat.id,
        conditionalType,
      };
    }

    if (type === "clearConditionalFormats") {
      const range = resolveExcelRange(context, action.target, false);
      range.conditionalFormats.clearAll();
      range.load("address");
      await context.sync();
      return { ok: true, host: "excel", action: type, address: range.address };
    }

    if (type === "insertInlinePicture") {
      const worksheet = resolveExcelWorksheet(context, action.target, true);
      const imageBase64 = action.content ?? "";
      const shape = worksheet.shapes.addImage(`data:image/png;base64,${imageBase64}`);
      shape.load("id,name,width,height");
      await context.sync();
      return { ok: true, host: "excel", action: type, shapeId: shape.id, shapeName: shape.name, width: shape.width, height: shape.height };
    }

    throw new Error(`Unsupported Excel action: ${type}`);
  });
}

async function loadPowerPointMastersWithLayouts(context: PowerPoint.RequestContext): Promise<PowerPoint.SlideMaster[]> {
  const slideMasters = context.presentation.slideMasters;
  slideMasters.load("items/id,items/name");
  await context.sync();
  for (const slideMaster of slideMasters.items) {
    slideMaster.layouts.load("items/id,items/name,items/type");
  }
  await context.sync();
  return slideMasters.items;
}

function matchesPowerPointLookup(
  value: { id?: string | undefined; name?: string | undefined; type?: string | undefined },
  lookup: { id?: string | undefined; name?: string | undefined; type?: string | undefined },
): boolean {
  if (lookup.id && value.id === lookup.id) {
    return true;
  }

  if (lookup.name && (matchesTextQuery(value.name, lookup.name) || value.id === lookup.name)) {
    return true;
  }

  if (lookup.type && (value.type ?? "").toLowerCase() === lookup.type.toLowerCase()) {
    return true;
  }

  return false;
}

function findPowerPointSlideMaster(
  slideMasters: PowerPoint.SlideMaster[],
  lookup: { id?: string | undefined; name?: string | undefined },
): PowerPoint.SlideMaster | undefined {
  if (!lookup.id && !lookup.name) {
    return undefined;
  }

  return slideMasters.find((slideMaster) => matchesPowerPointLookup({ id: slideMaster.id, name: slideMaster.name }, lookup));
}

function findPowerPointSlideLayout(
  slideMasters: PowerPoint.SlideMaster[],
  lookup: { id?: string | undefined; name?: string | undefined; type?: string | undefined },
  masterLookup?: { id?: string | undefined; name?: string | undefined },
): PowerPoint.SlideLayout | undefined {
  if (!lookup.id && !lookup.name && !lookup.type) {
    return undefined;
  }

  const scopedMasters = masterLookup ? [findPowerPointSlideMaster(slideMasters, masterLookup)].filter(Boolean) as PowerPoint.SlideMaster[] : slideMasters;
  for (const slideMaster of scopedMasters) {
    const layout = slideMaster.layouts.items.find((entry) =>
      matchesPowerPointLookup({ id: entry.id, name: entry.name, type: entry.type }, lookup),
    );
    if (layout) {
      return layout;
    }
  }

  return undefined;
}

async function resolvePowerPointLayoutSelection(
  context: PowerPoint.RequestContext,
  target: OfficeAnchor | undefined,
  options: Record<string, unknown>,
): Promise<{ slideMaster?: PowerPoint.SlideMaster | undefined; layout?: PowerPoint.SlideLayout | undefined }> {
  const layoutLookup = {
    id: trimString(options.layoutId) ?? (target?.kind === "layout" ? target.id : undefined),
    name: trimString(options.layoutName) ?? (target?.kind === "layout" ? target.label ?? target.text : undefined),
    type: trimString(options.layoutType),
  };
  const slideMasterLookup = {
    id: trimString(options.slideMasterId) ?? (target?.kind === "slideMaster" ? target.id : undefined),
    name: trimString(options.slideMasterName) ?? (target?.kind === "slideMaster" ? target.label ?? target.text : undefined),
  };

  if (!layoutLookup.id && !layoutLookup.name && !layoutLookup.type && !slideMasterLookup.id && !slideMasterLookup.name) {
    return {};
  }

  const slideMasters = await loadPowerPointMastersWithLayouts(context);
  const slideMaster = findPowerPointSlideMaster(slideMasters, slideMasterLookup);
  const layout = findPowerPointSlideLayout(slideMasters, layoutLookup, slideMasterLookup);
  return { slideMaster, layout };
}

function applyPowerPointShapeProperties(shape: PowerPoint.Shape, options: Record<string, unknown>): void {
  const left = toNumber(options.left);
  const top = toNumber(options.top);
  const width = toNumber(options.width);
  const height = toNumber(options.height);
  const rotation = toNumber(options.rotation);
  const name = trimString(options.name);
  const altTextTitle = trimString(options.altTextTitle);
  const altTextDescription = trimString(options.altTextDescription);
  const visible = toBoolean(options.visible);
  const isDecorative = toBoolean(options.isDecorative);
  const fillColor = trimString(options.fillColor);
  const fillImageBase64 = trimString(options.fillImageBase64);
  const fillTransparency = clampPercentage(options.fillTransparency);
  const clearFill = toBoolean(options.clearFill);
  const lineColor = trimString(options.lineColor);
  const lineWeight = toNumber(options.lineWeight);
  const lineTransparency = clampPercentage(options.lineTransparency);
  const lineVisible = toBoolean(options.lineVisible);
  const lineDashStyle = trimString(options.lineDashStyle);
  const lineStyle = trimString(options.lineStyle);
  const zOrder = trimString(options.zOrder);
  const hyperlinkAddress = trimString(options.hyperlinkAddress);
  const hyperlinkScreenTip = trimString(options.hyperlinkScreenTip);

  if (typeof left === "number") shape.left = left;
  if (typeof top === "number") shape.top = top;
  if (typeof width === "number") shape.width = width;
  if (typeof height === "number") shape.height = height;
  if (typeof rotation === "number") shape.rotation = rotation;
  if (name) shape.name = name;
  if (altTextTitle) shape.altTextTitle = altTextTitle;
  if (altTextDescription) shape.altTextDescription = altTextDescription;
  if (typeof visible === "boolean") shape.visible = visible;
  if (typeof isDecorative === "boolean") shape.isDecorative = isDecorative;
  if (clearFill) shape.fill.clear();
  if (fillImageBase64) shape.fill.setImage(fillImageBase64);
  if (fillColor) shape.fill.setSolidColor(fillColor);
  if (typeof fillTransparency === "number") shape.fill.transparency = fillTransparency;
  if (lineColor) shape.lineFormat.color = lineColor;
  if (typeof lineWeight === "number") shape.lineFormat.weight = lineWeight;
  if (typeof lineTransparency === "number") shape.lineFormat.transparency = lineTransparency;
  if (typeof lineVisible === "boolean") shape.lineFormat.visible = lineVisible;
  if (lineDashStyle) shape.lineFormat.dashStyle = lineDashStyle as PowerPoint.ShapeLineDashStyle;
  if (lineStyle) shape.lineFormat.style = lineStyle as PowerPoint.ShapeLineStyle;
  if (zOrder) shape.setZOrder(zOrder as PowerPoint.ShapeZOrder);
  if (hyperlinkAddress) {
    const hyperlinkOptions: PowerPoint.HyperlinkAddOptions = { address: hyperlinkAddress };
    if (hyperlinkScreenTip) {
      hyperlinkOptions.screenTip = hyperlinkScreenTip;
    }
    shape.setHyperlink(hyperlinkOptions);
  }
}

async function finalizePowerPointShapeSelection(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  shape: PowerPoint.Shape,
  action: string,
  extra: Record<string, unknown> = {},
): Promise<Record<string, unknown>> {
  shape.load("id,name");
  slide.load("id,index");
  await context.sync();
  context.presentation.setSelectedSlides([slide.id]);
  slide.setSelectedShapes([shape.id]);
  await context.sync();
  return {
    ok: true,
    host: "powerpoint",
    action,
    slideId: slide.id,
    slideIndex: slide.index + 1,
    shapeId: shape.id,
    shapeName: shape.name,
    ...extra,
  };
}

async function resolvePowerPointSlide(
  context: PowerPoint.RequestContext,
  target?: OfficeAnchor,
  allowSelected = true,
): Promise<PowerPoint.Slide> {
  if (target?.slideId) {
    return context.presentation.slides.getItem(target.slideId);
  }

  if (typeof target?.slideIndex === "number") {
    return context.presentation.slides.getItemAt(Math.max(0, target.slideIndex - 1));
  }

  if (allowSelected) {
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items/id");
    await context.sync();
    if (selectedSlides.items[0]) {
      return context.presentation.slides.getItem(selectedSlides.items[0].id);
    }
  }

  return context.presentation.slides.getItemAt(0);
}

function summarizePowerPointShape(
  slide: { id: string; index: number },
  shape: { id: string; name?: string | null; type?: unknown },
): { slideId: string; slideIndex: number; shapeId: string; shapeName?: string | undefined; shapeType?: unknown; label: string } {
  return {
    slideId: slide.id,
    slideIndex: slide.index + 1,
    shapeId: shape.id,
    shapeName: shape.name ?? undefined,
    shapeType: shape.type,
    label: shape.name || `Shape ${shape.id}`,
  };
}

async function loadSelectedPowerPointShapes(
  context: PowerPoint.RequestContext,
): Promise<{ slide?: PowerPoint.Slide | undefined; shapes: PowerPoint.Shape[] }> {
  const selectedSlides = context.presentation.getSelectedSlides();
  const selectedShapes = context.presentation.getSelectedShapes();
  selectedSlides.load("items/id,items/index");
  selectedShapes.load("items/id,items/name,items/type");
  await context.sync();
  return {
    slide: selectedSlides.items[0],
    shapes: selectedShapes.items,
  };
}

async function findPowerPointShapeInPresentation(
  context: PowerPoint.RequestContext,
  shapeId: string,
): Promise<{ slide: PowerPoint.Slide; shape: PowerPoint.Shape } | undefined> {
  const slides = context.presentation.slides;
  slides.load("items/id,items/index");
  await context.sync();

  const candidateShapes = slides.items.map((slide) => slide.shapes.getItemOrNullObject(shapeId));
  for (const shape of candidateShapes) {
    shape.load("isNullObject,id,name,type");
  }
  await context.sync();

  const matchedIndex = candidateShapes.findIndex((shape) => !shape.isNullObject);
  if (matchedIndex < 0) {
    return undefined;
  }

  const slide = slides.items[matchedIndex];
  const shape = candidateShapes[matchedIndex];
  if (!slide || !shape) {
    return undefined;
  }

  return { slide, shape };
}

async function resolvePowerPointShape(
  context: PowerPoint.RequestContext,
  target?: OfficeAnchor,
  allowSelected = true,
): Promise<{ slide: PowerPoint.Slide; shape: PowerPoint.Shape }> {
  const targetShapeId = trimString(target?.shapeId);
  if (targetShapeId && (target?.slideId || typeof target?.slideIndex === "number")) {
    const slide = await resolvePowerPointSlide(context, target, false);
    const shape = slide.shapes.getItemOrNullObject(targetShapeId);
    slide.load("id,index");
    shape.load("isNullObject,id,name,type");
    await context.sync();
    if (shape.isNullObject) {
      throw new Error(`Could not find PowerPoint shape ${targetShapeId} on the requested slide.`);
    }
    return { slide, shape };
  }

  if (targetShapeId) {
    const match = await findPowerPointShapeInPresentation(context, targetShapeId);
    if (!match) {
      throw new Error(`Could not find the requested PowerPoint shape: ${targetShapeId}.`);
    }
    return match;
  }

  if (allowSelected) {
    const selection = await loadSelectedPowerPointShapes(context);
    const shape = selection.shapes[0];
    const slide = selection.slide;
    if (shape && slide) {
      return { slide, shape };
    }
  }

  throw new Error("No PowerPoint shape is selected.");
}

async function resolvePowerPointShapes(
  context: PowerPoint.RequestContext,
  target?: OfficeAnchor,
  requestedShapeIds: string[] = [],
  allowSelected = true,
): Promise<{ slide: PowerPoint.Slide; shapes: PowerPoint.Shape[] }> {
  const shapeIds = Array.from(
    new Set([target?.shapeId, ...requestedShapeIds].map((value) => trimString(value)).filter((value): value is string => Boolean(value))),
  );

  if (!shapeIds.length) {
    if (allowSelected) {
      const selection = await loadSelectedPowerPointShapes(context);
      if (selection.slide && selection.shapes.length) {
        return { slide: selection.slide, shapes: selection.shapes };
      }
    }
    throw new Error("No PowerPoint shapes are selected.");
  }

  if (target?.slideId || typeof target?.slideIndex === "number") {
    const slide = await resolvePowerPointSlide(context, target, false);
    const shapes = shapeIds.map((shapeId) => slide.shapes.getItemOrNullObject(shapeId));
    slide.load("id,index");
    for (const shape of shapes) {
      shape.load("isNullObject,id,name,type");
    }
    await context.sync();
    const missingShapeId = shapes.find((shape) => shape.isNullObject)?.id ?? shapeIds.find((_shapeId, index) => shapes[index]?.isNullObject);
    if (missingShapeId) {
      throw new Error(`Could not find PowerPoint shape ${missingShapeId} on the requested slide.`);
    }
    return { slide, shapes };
  }

  const firstMatch = await findPowerPointShapeInPresentation(context, shapeIds[0]!);
  if (!firstMatch) {
    throw new Error(`Could not find the requested PowerPoint shape: ${shapeIds[0]}.`);
  }

  if (shapeIds.length === 1) {
    return { slide: firstMatch.slide, shapes: [firstMatch.shape] };
  }

  const additionalShapes = shapeIds.slice(1).map((shapeId) => firstMatch.slide.shapes.getItemOrNullObject(shapeId));
  for (const shape of additionalShapes) {
    shape.load("isNullObject,id,name,type");
  }
  await context.sync();
  const missingShapeId = shapeIds.slice(1).find((_shapeId, index) => additionalShapes[index]?.isNullObject);
  if (missingShapeId) {
    throw new Error(
      `Could not resolve PowerPoint shape ${missingShapeId} on slide ${firstMatch.slide.index + 1}. Pass an explicit slide target when shapes live on different slides.`,
    );
  }

  return { slide: firstMatch.slide, shapes: [firstMatch.shape, ...additionalShapes] };
}

async function selectPowerPointShapeTarget(
  target?: OfficeAnchor,
): Promise<{
  slideId: string;
  slideIndex: number;
  shapeId: string;
  shapeName?: string | undefined;
  shapeType?: string | undefined;
}> {
  return PowerPoint.run(async (context) => {
    const { slide, shape } = await resolvePowerPointShape(context, target, true);
    slide.load("id,index");
    shape.load("id,name,type");
    await context.sync();
    context.presentation.setSelectedSlides([slide.id]);
    slide.setSelectedShapes([shape.id]);
    await context.sync();
    return {
      slideId: slide.id,
      slideIndex: slide.index + 1,
      shapeId: shape.id,
      shapeName: shape.name,
      shapeType: typeof shape.type === "string" ? shape.type : undefined,
    };
  });
}

async function applyPowerPointShapeImageAction(
  action: OfficeHostAction,
  operation: string,
  image: { data: string; mimeType: string },
): Promise<unknown> {
  const options = getActionOptions(action);
  const actionOptions = {
    ...options,
    ...action,
    fillImageBase64: image.data,
  };

  try {
    return await PowerPoint.run(async (context) => {
      const { slide, shape } = await resolvePowerPointShape(context, action.target, true);
      slide.load("id,index");
      shape.load("id,name,type");
      await context.sync();
      applyPowerPointShapeProperties(shape, actionOptions);
      await context.sync();
      return finalizePowerPointShapeSelection(context, slide, shape, operation, {
        shapeType: shape.type,
        imageMimeType: image.mimeType,
        imageUpdateMode: "shape.fill.setImage",
      });
    });
  } catch (error) {
    const selection = await selectPowerPointShapeTarget(action.target);
    await setSelectedImageAsync(image.data);
    const errorInfo = serializeOfficeRuntimeError(error);
    return {
      ok: true,
      host: "powerpoint",
      action: operation,
      ...selection,
      imageMimeType: image.mimeType,
      imageUpdateMode: "setSelectedDataAsync",
      completion: "fallback",
      fallbackStrategy: "setSelectedDataAsync",
      nativeAttempted: true,
      nativeFailure: trimString(errorInfo.message),
      warning:
        typeof errorInfo.message === "string" && errorInfo.message.trim()
          ? `Native shape image update failed, so the selection-based image replacement fallback was used: ${errorInfo.message.trim()}`
          : "Used selection-based image replacement fallback.",
    };
  }
}

async function resolvePowerPointTable(
  context: PowerPoint.RequestContext,
  target?: OfficeAnchor,
  allowSelected = true,
): Promise<{ slide: PowerPoint.Slide; shape: PowerPoint.Shape; table: PowerPoint.Table }> {
  if (!supportsRequirementSet("PowerPointApi", "1.8")) {
    throw new Error("PowerPoint table actions require PowerPointApi 1.8.");
  }

  const resolved = await resolvePowerPointShape(context, target, allowSelected);
  if (!isPowerPointTableShape(resolved.shape.type)) {
    throw new Error("The requested PowerPoint shape is not a table.");
  }

  const table = resolved.shape.getTable();
  table.load("rowCount,columnCount,values");
  await context.sync();
  return { ...resolved, table };
}

function summarizePowerPointSlide(slide: { id: string; index: number }): { slideId: string; slideIndex: number; label: string } {
  return {
    slideId: slide.id,
    slideIndex: slide.index + 1,
    label: `Slide ${slide.index + 1}`,
  };
}

async function loadPowerPointSlideSummaries(
  context: PowerPoint.RequestContext,
): Promise<Array<{ slideId: string; slideIndex: number; label: string }>> {
  const slides = context.presentation.slides;
  slides.load("items/id,items/index");
  await context.sync();
  return slides.items.map((slide) => summarizePowerPointSlide(slide));
}

function findPowerPointSlideInsertionStart(
  slides: Array<{ slideId: string; slideIndex: number }>,
  targetSlideId?: string | undefined,
): number {
  if (!targetSlideId) {
    return 0;
  }

  const targetIndex = slides.findIndex((slide) => slide.slideId === targetSlideId);
  if (targetIndex < 0) {
    throw new Error(`Could not find the requested PowerPoint target slide: ${targetSlideId}.`);
  }

  return targetIndex + 1;
}

function sliceInsertedPowerPointSlides(
  beforeSlides: Array<{ slideId: string; slideIndex: number; label: string }>,
  afterSlides: Array<{ slideId: string; slideIndex: number; label: string }>,
  targetSlideId?: string | undefined,
  insertedCount?: number | undefined,
): Array<{ slideId: string; slideIndex: number; label: string }> {
  const count = insertedCount ?? Math.max(afterSlides.length - beforeSlides.length, 0);
  if (count <= 0) {
    return [];
  }

  const start = findPowerPointSlideInsertionStart(beforeSlides, targetSlideId);
  return afterSlides.slice(start, start + count);
}

async function exportTargetPowerPointSlideAsBase64(
  target?: OfficeAnchor,
): Promise<{ slideId: string; slideIndex: number; base64: string }> {
  return PowerPoint.run(async (context) => {
    const slide =
      target?.shapeId
        ? (await resolvePowerPointShape(context, target, true)).slide
        : await resolvePowerPointSlide(context, target, true);
    slide.load("id,index");
    const exportResult = slide.exportAsBase64();
    await context.sync();
    return {
      slideId: slide.id,
      slideIndex: slide.index + 1,
      base64: exportResult.value,
    };
  });
}

async function exportTargetPowerPointChartSlideAsBase64(
  target?: OfficeAnchor,
): Promise<{ slideId: string; slideIndex: number; base64: string; shapeName?: string | undefined }> {
  return PowerPoint.run(async (context) => {
    let slide: PowerPoint.Slide;
    let targetShapeName: string | undefined;

    if (target?.shapeId) {
      const resolved = await resolvePowerPointShape(context, target, true);
      slide = resolved.slide;
      resolved.shape.load("id,name,type");
      slide.load("id,index");
      const exportResult = slide.exportAsBase64();
      await context.sync();
      if (String(resolved.shape.type ?? "").toLowerCase() !== "chart") {
        throw new Error("The requested PowerPoint shape is not a chart.");
      }
      return {
        slideId: slide.id,
        slideIndex: slide.index + 1,
        base64: exportResult.value,
        shapeName: resolved.shape.name || undefined,
      };
    }

    slide = await resolvePowerPointSlide(context, target, true);
    slide.load("id,index");
    const selectedShapes = context.presentation.getSelectedShapes();
    selectedShapes.load("items/name,items/type");
    const exportResult = slide.exportAsBase64();
    await context.sync();

    const selectedChart = selectedShapes.items.find((shape) => String(shape.type ?? "").toLowerCase() === "chart");
    targetShapeName = selectedChart?.name || undefined;

    return {
      slideId: slide.id,
      slideIndex: slide.index + 1,
      base64: exportResult.value,
      shapeName: targetShapeName,
    };
  });
}

async function exportCurrentPowerPointPresentationAsBase64(): Promise<string> {
  if (!supportsRequirementSet("PowerPointApi", "1.10")) {
    throw new Error("PowerPoint presentation-package export requires PowerPointApi 1.10.");
  }

  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();
    const exportResult = slides.exportAsBase64Presentation(slides.items.map((slide) => slide.id));
    await context.sync();
    return exportResult.value;
  });
}

async function inspectCurrentPowerPointPresentationPackage(): Promise<Awaited<ReturnType<typeof inspectPowerPointPresentationBase64>>> {
  return inspectPowerPointPresentationBase64(await exportCurrentPowerPointPresentationAsBase64());
}

async function replacePowerPointSlideWithSerializedPackage(
  sourceSlide: { slideId: string; slideIndex: number },
  transformedBase64: string,
  type: string,
  extras: Record<string, unknown> = {},
): Promise<unknown> {
  const { selectShapeName: rawSelectShapeName, ...resultExtras } = extras;
  const formatting = trimString(resultExtras.formatting);
  const replaceOriginal = typeof resultExtras.replaceOriginal === "boolean" ? resultExtras.replaceOriginal : true;
  const selectShapeName = trimString(rawSelectShapeName);

  return PowerPoint.run(async (context) => {
    const beforeSlides = await loadPowerPointSlideSummaries(context);
    const insertOptions: PowerPoint.InsertSlideOptions = {
      targetSlideId: sourceSlide.slideId,
    };
    if (formatting) {
      insertOptions.formatting = formatting as PowerPoint.InsertSlideFormatting;
    }

    context.presentation.insertSlidesFromBase64(transformedBase64, insertOptions);
    await context.sync();

    const insertedAfterImport = sliceInsertedPowerPointSlides(
      beforeSlides,
      await loadPowerPointSlideSummaries(context),
      sourceSlide.slideId,
      1,
    );
    const insertedSlideIds = insertedAfterImport.map((slide) => slide.slideId);
    if (!insertedSlideIds.length) {
      throw new Error(`PowerPoint serialized ${type} did not create a replacement slide.`);
    }

    if (replaceOriginal) {
      context.presentation.slides.getItem(sourceSlide.slideId).delete();
      await context.sync();
    }

    const finalSlides = await loadPowerPointSlideSummaries(context);
    const insertedSlides = finalSlides.filter((slide) => insertedSlideIds.includes(slide.slideId));
    let selectedShapeId: string | undefined;
    let selectedShapeNameResolved: string | undefined;

    context.presentation.setSelectedSlides(insertedSlides.map((slide) => slide.slideId));
    await context.sync();

    if (selectShapeName && insertedSlides[0]) {
      const insertedSlide = context.presentation.slides.getItem(insertedSlides[0].slideId);
      insertedSlide.shapes.load("items/id,items/name");
      await context.sync();
      const normalizedTargetName = selectShapeName.toLowerCase();
      const matchingShape =
        insertedSlide.shapes.items.find((shape) => (shape.name ?? "").trim().toLowerCase() === normalizedTargetName) ??
        insertedSlide.shapes.items.find((shape) => (shape.name ?? "").trim().toLowerCase().includes(normalizedTargetName));
      if (matchingShape) {
        insertedSlide.setSelectedShapes([matchingShape.id]);
        selectedShapeId = matchingShape.id;
        selectedShapeNameResolved = matchingShape.name ?? undefined;
        await context.sync();
      }
    }

    if (!selectedShapeNameResolved) {
      selectedShapeNameResolved = trimString(resultExtras.shapeName) ?? undefined;
    }
    if (!selectedShapeId) {
      selectedShapeId = trimString(resultExtras.shapeId) ?? undefined;
    }

    return {
      ok: true,
      host: "powerpoint",
      action: type,
      slideId: insertedSlides[0]?.slideId,
      slideIndex: insertedSlides[0]?.slideIndex,
      slides: insertedSlides,
      createdSlides: insertedSlides,
      deletedSlides: replaceOriginal
        ? [{ slideId: sourceSlide.slideId, slideIndex: sourceSlide.slideIndex, label: `Slide ${sourceSlide.slideIndex}` }]
        : undefined,
      sourceSlideId: sourceSlide.slideId,
      sourceSlideIndex: sourceSlide.slideIndex,
      shapeId: selectedShapeId,
      shapeName: selectedShapeNameResolved,
      formatting,
      ...resultExtras,
    };
  });
}

async function applyPowerPointSlideNotesAction(action: OfficeHostAction, type: string): Promise<unknown> {
  const options = getActionOptions(action);
  const notesText =
    trimString(action.content) ??
    trimString(options.notesText) ??
    trimString(options.notes) ??
    trimString(options.text) ??
    "";
  const replaceOriginal = toBoolean(options.replaceOriginal) ?? true;
  const formatting = trimString(options.formatting);

  const sourceSlide = await exportTargetPowerPointSlideAsBase64(action.target);
  const transformed = await replaceSlideNotesInPowerPointPresentationBase64(sourceSlide.base64, {
    notesText,
    slideNumber: 1,
  });

  return replacePowerPointSlideWithSerializedPackage(sourceSlide, transformed.base64, type, {
    replaceOriginal,
    formatting,
    notesPartName: transformed.partName,
    notesText: truncateText(notesText, 1000),
    previousNotesText: truncateText(transformed.previousText, 1000),
    changed: transformed.changed,
    serialization: "pptx-ooxml",
  });
}

async function applyPowerPointCreateChartAction(action: OfficeHostAction, type: string): Promise<unknown> {
  const options = getActionOptions(action);
  const sourceSlide = await exportTargetPowerPointSlideAsBase64(action.target);
  const categories = getChartValueArray(options.categories ?? action.categories);
  const series = getChartSeriesInput(options.series ?? action.series);
  const transformed = await createPowerPointChartInPresentationBase64(sourceSlide.base64, {
    chartType: trimString(options.chartType) ?? trimString(action.chartType),
    shapeName: trimString(options.shapeName) ?? trimString(action.shapeName),
    title: trimString(options.title) ?? trimString(action.content),
    categories: categories.length ? categories : undefined,
    series: series.length ? series : undefined,
    left: toNumber(options.left) ?? toNumber(options.chartLeft),
    top: toNumber(options.top) ?? toNumber(options.chartTop),
    width: toNumber(options.width) ?? toNumber(options.chartWidth),
    height: toNumber(options.height) ?? toNumber(options.chartHeight),
    showLegend: toBoolean(options.showLegend) ?? undefined,
  });

  return replacePowerPointSlideWithSerializedPackage(sourceSlide, transformed.base64, type, {
    replaceOriginal: toBoolean(options.replaceOriginal) ?? true,
    formatting: trimString(options.formatting),
    selectShapeName: transformed.shapeName,
    chartIndex: transformed.chartIndex,
    chartPartName: transformed.chartPartName,
    chartType: transformed.chartType,
    chartTitle: transformed.title,
    shapeId: transformed.shapeId,
    shapeName: transformed.shapeName,
    embeddedWorkbookPartName: transformed.embeddedWorkbookPartName,
    categoryCount: transformed.categoryCount,
    seriesCount: transformed.seriesCount,
    createdCharts: [
      {
        chartIndex: transformed.chartIndex,
        chartPartName: transformed.chartPartName,
        chartType: transformed.chartType,
        title: transformed.title,
        shapeId: transformed.shapeId,
        shapeName: transformed.shapeName,
        embeddedWorkbookPartName: transformed.embeddedWorkbookPartName,
        seriesCount: transformed.seriesCount,
        categoryCount: transformed.categoryCount,
      },
    ],
    serialization: "pptx-ooxml",
  });
}

async function applyPowerPointChartAction(action: OfficeHostAction, type: string): Promise<unknown> {
  const options = getActionOptions(action);
  const sourceSlide = await exportTargetPowerPointChartSlideAsBase64(action.target);
  const categories = getChartValueArray(options.categories ?? action.categories);
  const series = getChartSeriesInput(options.series ?? action.series);
  const transformed = await updatePowerPointChartInPresentationBase64(sourceSlide.base64, {
    chartIndex: parsePositiveInteger(options.chartIndex ?? action.chartIndex),
    shapeName: trimString(options.shapeName) ?? sourceSlide.shapeName,
    title: trimString(options.title) ?? trimString(action.content),
    categories: categories.length ? categories : undefined,
    series: series.length ? series : undefined,
  });

  return replacePowerPointSlideWithSerializedPackage(sourceSlide, transformed.base64, type, {
    replaceOriginal: toBoolean(options.replaceOriginal) ?? true,
    formatting: trimString(options.formatting),
    selectShapeName: transformed.shapeName,
    chartIndex: transformed.chartIndex,
    chartPartName: transformed.chartPartName,
    chartType: transformed.chartType,
    chartTitle: transformed.title,
    shapeName: transformed.shapeName,
    embeddedWorkbookPartName: transformed.embeddedWorkbookPartName,
    categoryCount: transformed.categoryCount,
    seriesCount: transformed.seriesCount,
    serialization: "pptx-ooxml",
    warnings: transformed.warnings,
  });
}

async function resolvePowerPointSourceSlides(
  context: PowerPoint.RequestContext,
  target?: OfficeAnchor,
  requestedSlideIds: string[] = [],
): Promise<Array<{ slideId: string; slideIndex: number; label: string }>> {
  if (requestedSlideIds.length) {
    const slides = context.presentation.slides;
    const requestedSlides = requestedSlideIds.map((slideId) => slides.getItem(slideId));
    for (const slide of requestedSlides) {
      slide.load("id,index");
    }
    await context.sync();
    return requestedSlides.map((slide) => summarizePowerPointSlide(slide));
  }

  if (target?.slideId || typeof target?.slideIndex === "number") {
    const slide = await resolvePowerPointSlide(context, target, false);
    slide.load("id,index");
    await context.sync();
    return [summarizePowerPointSlide(slide)];
  }

  const selectedSlides = context.presentation.getSelectedSlides();
  selectedSlides.load("items/id,items/index");
  await context.sync();
  if (selectedSlides.items.length > 0) {
    return selectedSlides.items.map((slide) => summarizePowerPointSlide(slide));
  }

  const slide = await resolvePowerPointSlide(context, target, true);
  slide.load("id,index");
  await context.sync();
  return [summarizePowerPointSlide(slide)];
}

function getPowerPointSlideIdArray(value: unknown): string[] {
  if (Array.isArray(value)) {
    const stringValues = value.filter((entry): entry is string => typeof entry === "string" && entry.trim().length > 0);
    if (stringValues.length) {
      return stringValues;
    }

    return value
      .map((entry) => (isRecord(entry) ? trimString(entry.slideId) : undefined))
      .filter((entry): entry is string => Boolean(entry));
  }

  return [];
}

async function resolvePowerPointInsertionIndex(
  context: PowerPoint.RequestContext,
  target: OfficeAnchor | undefined,
  options: Record<string, unknown>,
): Promise<number | undefined> {
  const explicitSlideIndex = toNumber(options.slideIndex);
  if (typeof explicitSlideIndex === "number" && Number.isInteger(explicitSlideIndex)) {
    return Math.max(0, explicitSlideIndex - 1);
  }

  const relativeSlideId = trimString(options.relativeToSlideId) ?? trimString(options.targetSlideId);
  const relativeSlideIndex = toNumber(options.relativeToSlideIndex);
  const relativeTarget =
    relativeSlideId || typeof relativeSlideIndex === "number"
      ? ({
          kind: "slide",
          slideId: relativeSlideId,
          slideIndex: typeof relativeSlideIndex === "number" ? relativeSlideIndex : undefined,
        } as OfficeAnchor)
      : target?.kind === "slide"
        ? target
        : undefined;

  if (!relativeTarget) {
    return undefined;
  }

  const relativeSlide = await resolvePowerPointSlide(context, relativeTarget, false);
  relativeSlide.load("index");
  await context.sync();
  const position = (trimString(options.position) ?? trimString(options.slidePosition) ?? trimString(options.placement) ?? "after").toLowerCase();
  return position === "before" ? Math.max(0, relativeSlide.index) : relativeSlide.index + 1;
}

async function createPowerPointSlide(
  context: PowerPoint.RequestContext,
  target: OfficeAnchor | undefined,
  options: Record<string, unknown>,
): Promise<PowerPoint.Slide> {
  const slides = context.presentation.slides;
  const count = slides.getCount();
  await context.sync();

  const requestedLayoutName = trimString(options.layoutName) ?? (target?.kind === "layout" ? target.label ?? target.text : undefined);
  const requestedSlideMasterName =
    trimString(options.slideMasterName) ?? (target?.kind === "slideMaster" ? target.label ?? target.text : undefined);
  const requestedLayoutType = trimString(options.layoutType);
  const { slideMaster, layout } = await resolvePowerPointLayoutSelection(context, target, options);
  if ((requestedLayoutName || requestedLayoutType || target?.kind === "layout") && !layout) {
    throw new Error("Could not resolve the requested PowerPoint layout.");
  }
  if ((requestedSlideMasterName || target?.kind === "slideMaster") && !slideMaster && !trimString(options.slideMasterId)) {
    throw new Error("Could not resolve the requested PowerPoint slide master.");
  }

  const slideOptions: PowerPoint.AddSlideOptions = {};
  const layoutId = trimString(options.layoutId) ?? layout?.id;
  const slideMasterId = trimString(options.slideMasterId) ?? slideMaster?.id;
  if (layoutId) slideOptions.layoutId = layoutId;
  if (slideMasterId) slideOptions.slideMasterId = slideMasterId;
  slides.add(slideOptions);
  await context.sync();

  const slide = slides.getItemAt(count.value);
  slide.load("id,index");
  slide.layout.load("id,name,type");
  slide.slideMaster.load("id,name");
  await context.sync();

  const insertionIndex = await resolvePowerPointInsertionIndex(context, target, options);
  if (typeof insertionIndex === "number" && insertionIndex !== slide.index) {
    slide.moveTo(insertionIndex);
    slide.load("id,index");
    await context.sync();
  }

  return slide;
}

async function loadPowerPointSlideContentSummaries(
  context: PowerPoint.RequestContext,
  slides: Array<{ slideId: string; slideIndex: number; label: string }>,
): Promise<
  Array<{
    slideId: string;
    slideIndex: number;
    label: string;
    title?: string | undefined;
    textBlocks: string[];
    combinedText?: string | undefined;
    shapeCount: number;
  }>
> {
  if (!slides.length) {
    return [];
  }

  const sourceSlides = slides.map((slide) => context.presentation.slides.getItem(slide.slideId));
  const shapeCollections = sourceSlides.map((slide) => slide.shapes);
  for (const slide of sourceSlides) {
    slide.load("id,index");
  }
  for (const shapes of shapeCollections) {
    shapes.load("items/id,items/name,items/type,items/left,items/top");
  }
  await context.sync();

  const supportsTextFrames = supportsRequirementSet("PowerPointApi", "1.10");
  const textFrames: Array<Array<PowerPoint.TextFrame | undefined>> = shapeCollections.map((shapes) =>
    shapes.items.map((shape) => (supportsTextFrames ? shape.getTextFrameOrNullObject() : undefined)),
  );
  if (supportsTextFrames) {
    for (const slideTextFrames of textFrames) {
      for (const textFrame of slideTextFrames) {
        textFrame?.load("isNullObject,hasText,textRange/text");
      }
    }
    await context.sync();
  }

  return slides.map((slide, slideIndex) => {
    const shapes = shapeCollections[slideIndex]?.items ?? [];
    const textEntries = shapes
      .map((shape, shapeIndex) => {
        const textFrame = textFrames[slideIndex]?.[shapeIndex];
        const text =
          supportsTextFrames && textFrame && !textFrame.isNullObject && textFrame.hasText
            ? truncateText(normalizeTextPreview(textFrame.textRange.text), 500)
            : undefined;

        return {
          top: shape.top,
          left: shape.left,
          text,
          fallbackLabel: trimString(shape.name),
        };
      })
      .filter((entry) => entry.text || entry.fallbackLabel)
      .sort((left, right) => left.top - right.top || left.left - right.left);

    const title = textEntries[0]?.text ?? textEntries[0]?.fallbackLabel ?? slide.label;
    const bodyBlocks = textEntries
      .map((entry) => entry.text ?? entry.fallbackLabel)
      .filter((entry): entry is string => Boolean(entry))
      .filter((entry, index) => !(index === 0 && entry === title));

    return {
      slideId: slide.slideId,
      slideIndex: slide.slideIndex,
      label: slide.label,
      title,
      textBlocks: bodyBlocks,
      combinedText: bodyBlocks.length ? bodyBlocks.join("\n") : title,
      shapeCount: shapes.length,
    };
  });
}

function applyPowerPointTextFrameProperties(textFrame: PowerPoint.TextFrame, options: Record<string, unknown>): void {
  const autoSizeSetting = trimString(options.autoSizeSetting);
  const verticalAlignment = trimString(options.verticalAlignment);
  const wordWrap = toBoolean(options.wordWrap);
  const topMargin = toNumber(options.topMargin);
  const rightMargin = toNumber(options.rightMargin);
  const bottomMargin = toNumber(options.bottomMargin);
  const leftMargin = toNumber(options.leftMargin);

  if (autoSizeSetting) textFrame.autoSizeSetting = autoSizeSetting as PowerPoint.ShapeAutoSize;
  if (verticalAlignment) textFrame.verticalAlignment = verticalAlignment as PowerPoint.TextVerticalAlignment;
  if (typeof wordWrap === "boolean") textFrame.wordWrap = wordWrap;
  if (typeof topMargin === "number") textFrame.topMargin = topMargin;
  if (typeof rightMargin === "number") textFrame.rightMargin = rightMargin;
  if (typeof bottomMargin === "number") textFrame.bottomMargin = bottomMargin;
  if (typeof leftMargin === "number") textFrame.leftMargin = leftMargin;
}

function applyPowerPointTableCellProperties(cell: PowerPoint.TableCell, options: Record<string, unknown>): void {
  const text = trimString(options.text);
  const fillColor = trimString(options.fillColor);
  const fontColor = trimString(options.fontColor);
  const fontName = trimString(options.fontName);
  const fontSize = toNumber(options.fontSize);
  const bold = toBoolean(options.bold);
  const italic = toBoolean(options.italic);
  const underline = trimString(options.underline);
  const horizontalAlignment = trimString(options.horizontalAlignment);
  const verticalAlignment = trimString(options.verticalAlignment);
  const indentLevel = toNumber(options.indentLevel);
  const topMargin = toNumber(options.topMargin);
  const rightMargin = toNumber(options.rightMargin);
  const bottomMargin = toNumber(options.bottomMargin);
  const leftMargin = toNumber(options.leftMargin);

  if (typeof text === "string") cell.text = text;
  if (fillColor) cell.fill.setSolidColor(fillColor);
  if (fontColor) cell.font.color = fontColor;
  if (fontName) cell.font.name = fontName;
  if (typeof fontSize === "number") cell.font.size = fontSize;
  if (typeof bold === "boolean") cell.font.bold = bold;
  if (typeof italic === "boolean") cell.font.italic = italic;
  if (underline) cell.font.underline = underline as PowerPoint.ShapeFontUnderlineStyle;
  if (horizontalAlignment) cell.horizontalAlignment = horizontalAlignment as PowerPoint.ParagraphHorizontalAlignment;
  if (verticalAlignment) cell.verticalAlignment = verticalAlignment as PowerPoint.TextVerticalAlignment;
  if (typeof indentLevel === "number") cell.indentLevel = indentLevel;
  if (typeof topMargin === "number") cell.margins.top = topMargin;
  if (typeof rightMargin === "number") cell.margins.right = rightMargin;
  if (typeof bottomMargin === "number") cell.margins.bottom = bottomMargin;
  if (typeof leftMargin === "number") cell.margins.left = leftMargin;
}

async function navigatePowerPointAnchor(anchor: OfficeAnchor): Promise<unknown> {
  return PowerPoint.run(async (context) => {
    if (anchor.kind === "notesRegion") {
      const slide = await resolvePowerPointSlide(context, anchor, true);
      slide.load("id,index");
      const slideExport = supportsRequirementSet("PowerPointApi", "1.10") ? slide.exportAsBase64() : undefined;
      await context.sync();
      context.presentation.setSelectedSlides([slide.id]);
      await context.sync();

      let hasNotes = false;
      let notesPreview: string | undefined;
      let notesText: string | undefined;
      let notesPartName: string | undefined;
      if (slideExport?.value) {
        try {
          const packageSummary = await inspectPowerPointPresentationBase64(slideExport.value);
          const note = packageSummary.notes[0];
          hasNotes = note?.hasNotes ?? false;
          notesPreview = note?.preview;
          notesText = truncateText(note?.text, 800);
          notesPartName = note?.partName;
        } catch {
          // Ignore serialization failures during navigation. Slide selection still succeeds.
        }
      }

      return {
        ok: true,
        host: "powerpoint",
        anchorKind: "notesRegion",
        completion: "partial",
        fallbackStrategy: "slideSelection",
        nativeAttempted: true,
        nativeFailure: "PowerPoint does not expose direct speaker-notes pane selection in edit view.",
        slideId: slide.id,
        slideIndex: slide.index + 1,
        hasNotes,
        notesPreview,
        notesText,
        notesPartName,
        warnings: ["Selected the target slide because PowerPoint does not expose direct speaker-notes pane selection in edit view."],
      };
    }

    if (anchor.kind === "layout" || anchor.kind === "slideMaster") {
      const presentationSlides = context.presentation.slides;
      presentationSlides.load("items/id,items/index");
      await context.sync();
      for (const slide of presentationSlides.items) {
        slide.layout.load("id,name,type");
        slide.slideMaster.load("id,name");
      }
      await context.sync();

      const slide = presentationSlides.items.find((entry) =>
        anchor.kind === "layout"
          ? matchesPowerPointLookup(
              { id: entry.layout.id, name: entry.layout.name, type: entry.layout.type },
              { id: anchor.id, name: anchor.label ?? anchor.text },
            )
          : matchesPowerPointLookup({ id: entry.slideMaster.id, name: entry.slideMaster.name }, { id: anchor.id, name: anchor.label ?? anchor.text }),
      );
      if (slide) {
        context.presentation.setSelectedSlides([slide.id]);
        await context.sync();
        return {
          ok: true,
          host: "powerpoint",
          anchorKind: anchor.kind,
          completion: "partial",
          fallbackStrategy: "slideSelection",
          nativeAttempted: true,
          nativeFailure: `PowerPoint does not expose direct ${anchor.kind} selection in edit view.`,
          slideId: slide.id,
          slideIndex: slide.index + 1,
          layoutId: slide.layout.id,
          layoutName: slide.layout.name,
          slideMasterId: slide.slideMaster.id,
          slideMasterName: slide.slideMaster.name,
          warnings: [`Selected a slide using the requested ${anchor.kind} because PowerPoint does not expose direct layout/master selection in edit view.`],
        };
      }
    }

    if (anchor.kind === "shape" && anchor.shapeId) {
      const { slide, shape } = await resolvePowerPointShape(context, anchor, true);
      slide.load("id,index");
      shape.load("id,name");
      await context.sync();
      context.presentation.setSelectedSlides([slide.id]);
      slide.setSelectedShapes([shape.id]);
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        anchorKind: "shape",
        slideId: slide.id,
        slideIndex: slide.index + 1,
        shapeId: shape.id,
        shapeName: shape.name,
      };
    }

    const slide = await resolvePowerPointSlide(context, anchor, true);
    slide.load("id,index");
    await context.sync();
    context.presentation.setSelectedSlides([slide.id]);
    await context.sync();
    return { ok: true, host: "powerpoint", anchorKind: "slide", slideId: slide.id, slideIndex: slide.index + 1 };
  });
}

async function applyPowerPointAction(action: OfficeHostAction): Promise<unknown> {
  const type = trimString(action.type) ?? "insertText";
  const options = getActionOptions(action);
  const imagePayload = getActionImagePayload(action, options);
  const isShapeImageAction =
    type === "replaceShapeImage" ||
    type === "setShapeImage" ||
    type === "updateShapeImage" ||
    ((type === "setShapeProperties" || type === "updateShapeProperties") && Boolean(imagePayload));

  if (type === "insertText" && !action.target?.shapeId) {
    try {
      return await PowerPoint.run(async (context) => {
        const textRange = context.presentation.getSelectedTextRangeOrNullObject();
        textRange.load("isNullObject,text");
        await context.sync();
        if (textRange.isNullObject) {
          throw new Error("No PowerPoint text range is selected.");
        }
        textRange.text = action.content ?? "";
        await context.sync();
        return { ok: true, host: "powerpoint", action: type };
      });
    } catch (error) {
      const errorInfo = serializeOfficeRuntimeError(error);
      await setSelectedTextAsync(action.content ?? "");
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        completion: "fallback",
        fallbackStrategy: "setSelectedDataAsync",
        nativeAttempted: true,
        nativeFailure: trimString(errorInfo.message),
        warning:
          typeof errorInfo.message === "string" && errorInfo.message.trim()
            ? `Native PowerPoint text-range update failed, so the selection data fallback was used: ${errorInfo.message.trim()}`
            : "Used selection-based PowerPoint text insertion fallback.",
      };
    }
  }

  if (isShapeImageAction) {
    if (!imagePayload) {
      throw new Error("PowerPoint shape image actions require a base64 image payload.");
    }
    return applyPowerPointShapeImageAction(action, type, imagePayload);
  }

  if (type === "inspectPresentationPackage" || type === "getPresentationTheme") {
    const base64 = trimString(action.content) ?? trimString(options.base64);
    const packageSummary = base64
      ? await inspectPowerPointPresentationBase64(base64)
      : await inspectCurrentPowerPointPresentationPackage();
    return {
      ok: true,
      host: "powerpoint",
      action: type,
      serialization: "pptx-ooxml",
      theme: packageSummary.theme,
      notes: packageSummary.notes.slice(0, 20).map((note) => ({
        slideNumber: note.slideNumber,
        hasNotes: note.hasNotes,
        preview: note.preview,
        partName: note.partName,
      })),
      partCounts: packageSummary.partCounts,
    };
  }

  if (type === "getSlideNotes" || type === "readSlideNotes" || type === "inspectSlideNotes") {
    const slidePackage = await exportTargetPowerPointSlideAsBase64(action.target);
    const packageSummary = await inspectPowerPointPresentationBase64(slidePackage.base64);
    const note = packageSummary.notes[0];
    return {
      ok: true,
      host: "powerpoint",
      action: type,
      slideId: slidePackage.slideId,
      slideIndex: slidePackage.slideIndex,
      hasNotes: note?.hasNotes ?? false,
      notesText: note?.text ?? "",
      notesPreview: note?.preview,
      notesPartName: note?.partName,
      serialization: "pptx-ooxml",
    };
  }

  if (type === "getSlideCharts" || type === "inspectSlideCharts" || type === "readSlideCharts") {
    const slidePackage = await exportTargetPowerPointChartSlideAsBase64(action.target);
    const packageSummary = await inspectPowerPointPresentationBase64(slidePackage.base64);
    return {
      ok: true,
      host: "powerpoint",
      action: type,
      slideId: slidePackage.slideId,
      slideIndex: slidePackage.slideIndex,
      charts: packageSummary.charts.map((chart) => ({
        chartIndex: chart.chartIndex,
        chartPartName: chart.chartPartName,
        chartType: chart.chartType,
        title: chart.title,
        shapeName: chart.shapeName,
        embeddedWorkbookPartName: chart.embeddedWorkbookPartName,
        seriesCount: chart.seriesCount,
        categoryCount: chart.categoryCount,
        hasEmbeddedWorkbook: chart.hasEmbeddedWorkbook,
        series: chart.series.map((series) => ({
          index: series.index,
          name: series.name,
          categories: series.categories.slice(0, 12),
          values: series.values.slice(0, 12),
        })),
      })),
      serialization: "pptx-ooxml",
    };
  }

  if (type === "setSlideNotes" || type === "replaceSlideNotes") {
    return applyPowerPointSlideNotesAction(action, type);
  }

  if (type === "addSlideChart" || type === "createSlideChart" || type === "addChartToSlide" || type === "insertSlideChart") {
    return applyPowerPointCreateChartAction(action, type);
  }

  if (type === "updateSlideChart" || type === "setChartData" || type === "updateChartData" || type === "replaceChartData") {
    return applyPowerPointChartAction(action, type);
  }

  return PowerPoint.run(async (context) => {
    const actionOptions = { ...options, ...action };

    if ((type === "insertText" && action.target?.shapeId) || type === "setShapeText" || type === "clearShapeText") {
      const { slide, shape } = await resolvePowerPointShape(context, action.target, true);
      const textFrame = shape.getTextFrameOrNullObject();
      shape.load("id,name,type");
      textFrame.load("isNullObject,textRange/text");
      await context.sync();
      if (textFrame.isNullObject) {
        throw new Error("The requested PowerPoint shape does not contain text.");
      }
      if (type === "clearShapeText") {
        textFrame.deleteText();
      } else {
        const placement = trimString(action.placement) ?? trimString(options.placement);
        const nextText = trimString(action.content) ?? trimString(options.text) ?? "";
        textFrame.textRange.text = placement === "after" ? `${textFrame.textRange.text}${nextText}` : nextText;
      }
      applyPowerPointTextFrameProperties(textFrame, actionOptions);
      return finalizePowerPointShapeSelection(context, slide, shape, type);
    }

    if (type === "setShapeProperties" || type === "updateShapeProperties" || type === "moveShape" || type === "resizeShape") {
      const { slide, shape } = await resolvePowerPointShape(context, action.target, true);
      applyPowerPointShapeProperties(shape, actionOptions);
      return finalizePowerPointShapeSelection(context, slide, shape, type);
    }

    if (type === "deleteShape" || type === "deleteShapes") {
      const { slide, shapes } = await resolvePowerPointShapes(context, action.target, getStringArray(options.shapeIds), true);
      slide.load("id,index");
      for (const shape of shapes) {
        shape.load("id,name,type");
      }
      await context.sync();
      const deletedShapes = shapes.map((shape) => summarizePowerPointShape(slide, shape));
      for (const shape of shapes) {
        shape.delete();
      }
      await context.sync();
      context.presentation.setSelectedSlides([slide.id]);
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        deletedCount: deletedShapes.length,
        deletedShapeIds: deletedShapes.map((shape) => shape.shapeId),
        deletedShapes,
      };
    }

    if (type === "ungroupShape" || type === "ungroupShapes") {
      if (!supportsRequirementSet("PowerPointApi", "1.8")) {
        throw new Error("PowerPoint ungroupShape requires PowerPointApi 1.8.");
      }

      const { slide, shapes } = await resolvePowerPointShapes(context, action.target, getStringArray(options.shapeIds), true);
      slide.load("id,index");
      for (const shape of shapes) {
        shape.load("id,name,type");
      }
      await context.sync();
      const nonGroupShape = shapes.find((shape) => !isPowerPointGroupShape(shape.type));
      if (nonGroupShape) {
        throw new Error(`PowerPoint ungroupShape requires grouped shapes. ${nonGroupShape.name || nonGroupShape.id} is not a group.`);
      }

      const groups = shapes.map((shape) => shape.group);
      for (const group of groups) {
        group.shapes.load("items/id,items/name,items/type");
      }
      await context.sync();

      const deletedShapes = shapes.map((shape) => summarizePowerPointShape(slide, shape));
      const createdShapes = groups.flatMap((group) => group.shapes.items.map((shape) => summarizePowerPointShape(slide, shape)));
      for (const group of groups) {
        group.ungroup();
      }
      await context.sync();
      context.presentation.setSelectedSlides([slide.id]);
      if (createdShapes.length) {
        slide.setSelectedShapes(createdShapes.map((shape) => shape.shapeId));
      }
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        deletedShapes,
        createdShapes,
      };
    }

    if (type === "addSlide") {
      const slide = await createPowerPointSlide(context, action.target, options);
      context.presentation.setSelectedSlides([slide.id]);
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        layoutId: slide.layout.id,
        layoutName: slide.layout.name,
        layoutType: slide.layout.type,
        slideMasterId: slide.slideMaster.id,
        slideMasterName: slide.slideMaster.name,
        createdSlides: [summarizePowerPointSlide(slide)],
      };
    }

    if (type === "addAgendaSlide") {
      const requestedSourceSlideIds = getPowerPointSlideIdArray(options.sourceSlideIds ?? options.slideIds ?? options.sourceSlides);
      const sourceSlides = await resolvePowerPointSourceSlides(context, undefined, requestedSourceSlideIds);
      const sourceContent = requestedSourceSlideIds.length ? await loadPowerPointSlideContentSummaries(context, sourceSlides) : [];
      const agendaItems =
        getStringArray(options.items).length > 0
          ? getStringArray(options.items)
          : getStringArray(options.agendaItems).length > 0
            ? getStringArray(options.agendaItems)
            : sourceContent.map((slide) => slide.title ?? slide.label).filter((entry): entry is string => Boolean(entry));
      if (!agendaItems.length) {
        throw new Error("PowerPoint addAgendaSlide requires agendaItems/items or source slides with detectable titles.");
      }

      const slide = await createPowerPointSlide(context, action.target, options);
      slide.load("id,index");
      const titleText = trimString(options.title) ?? trimString(action.content) ?? "Agenda";
      const subtitleText = trimString(options.subtitle);
      const titleShape = slide.shapes.addTextBox(titleText, {
        left: toNumber(options.titleLeft) ?? 60,
        top: toNumber(options.titleTop) ?? 50,
        width: toNumber(options.titleWidth) ?? 620,
        height: toNumber(options.titleHeight) ?? 48,
      });
      const agendaShape = slide.shapes.addTextBox(
        agendaItems.map((item, index) => `${index + 1}. ${item}`).join("\n"),
        {
          left: toNumber(options.bodyLeft) ?? 80,
          top: toNumber(options.bodyTop) ?? (subtitleText ? 180 : 150),
          width: toNumber(options.bodyWidth) ?? 560,
          height: toNumber(options.bodyHeight) ?? 300,
        },
      );
      const createdShapes = [titleShape, agendaShape];
      applyPowerPointShapeProperties(titleShape, {
        name: trimString(options.titleName) ?? "Agenda Title",
        fillColor: trimString(options.titleFillColor),
        lineColor: trimString(options.titleLineColor),
      });
      applyPowerPointShapeProperties(agendaShape, {
        name: trimString(options.bodyName) ?? "Agenda Items",
        fillColor: trimString(options.bodyFillColor),
        lineColor: trimString(options.bodyLineColor),
      });

      const titleFrame = titleShape.getTextFrameOrNullObject();
      const agendaFrame = agendaShape.getTextFrameOrNullObject();
      applyPowerPointTextFrameProperties(titleFrame, {
        autoSizeSetting: trimString(options.titleAutoSizeSetting) ?? "AutoSizeShapeToFitText",
        wordWrap: toBoolean(options.titleWordWrap) ?? true,
      });
      applyPowerPointTextFrameProperties(agendaFrame, {
        autoSizeSetting: trimString(options.bodyAutoSizeSetting) ?? "AutoSizeShapeToFitText",
        wordWrap: toBoolean(options.bodyWordWrap) ?? true,
      });
      if (typeof toNumber(options.titleFontSize) === "number") {
        titleFrame.textRange.font.size = toNumber(options.titleFontSize)!;
      }
      if (trimString(options.titleFontColor)) {
        titleFrame.textRange.font.color = trimString(options.titleFontColor)!;
      }
      if (typeof toBoolean(options.titleBold) === "boolean") {
        titleFrame.textRange.font.bold = toBoolean(options.titleBold)!;
      }
      if (typeof toNumber(options.bodyFontSize) === "number") {
        agendaFrame.textRange.font.size = toNumber(options.bodyFontSize)!;
      }
      if (trimString(options.bodyFontColor)) {
        agendaFrame.textRange.font.color = trimString(options.bodyFontColor)!;
      }

      if (subtitleText) {
        const subtitleShape = slide.shapes.addTextBox(subtitleText, {
          left: toNumber(options.subtitleLeft) ?? 80,
          top: toNumber(options.subtitleTop) ?? 112,
          width: toNumber(options.subtitleWidth) ?? 560,
          height: toNumber(options.subtitleHeight) ?? 42,
        });
        createdShapes.push(subtitleShape);
        applyPowerPointShapeProperties(subtitleShape, {
          name: trimString(options.subtitleName) ?? "Agenda Subtitle",
          fillColor: trimString(options.subtitleFillColor),
          lineColor: trimString(options.subtitleLineColor),
        });
        const subtitleFrame = subtitleShape.getTextFrameOrNullObject();
        applyPowerPointTextFrameProperties(subtitleFrame, {
          autoSizeSetting: trimString(options.subtitleAutoSizeSetting) ?? "AutoSizeShapeToFitText",
          wordWrap: toBoolean(options.subtitleWordWrap) ?? true,
        });
        if (typeof toNumber(options.subtitleFontSize) === "number") {
          subtitleFrame.textRange.font.size = toNumber(options.subtitleFontSize)!;
        }
        if (trimString(options.subtitleFontColor)) {
          subtitleFrame.textRange.font.color = trimString(options.subtitleFontColor)!;
        }
      }

      for (const shape of createdShapes) {
        shape.load("id,name,type");
      }
      await context.sync();
      context.presentation.setSelectedSlides([slide.id]);
      slide.setSelectedShapes([agendaShape.id]);
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        createdSlides: [summarizePowerPointSlide(slide)],
        createdShapes: createdShapes.map((shape) => summarizePowerPointShape(slide, shape)),
        agendaItems,
        sourceSlideIds: sourceSlides.map((entry) => entry.slideId),
      };
    }

    if (type === "addTransitionSlide") {
      const slide = await createPowerPointSlide(context, action.target, options);
      slide.load("id,index");
      const titleText = trimString(options.title) ?? trimString(action.content) ?? "Section";
      const subtitleText = trimString(options.subtitle);
      const kickerText = trimString(options.kicker);
      const createdShapes: PowerPoint.Shape[] = [];

      if (kickerText) {
        const kickerShape = slide.shapes.addTextBox(kickerText, {
          left: toNumber(options.kickerLeft) ?? 72,
          top: toNumber(options.kickerTop) ?? 86,
          width: toNumber(options.kickerWidth) ?? 520,
          height: toNumber(options.kickerHeight) ?? 24,
        });
        createdShapes.push(kickerShape);
        applyPowerPointShapeProperties(kickerShape, {
          name: trimString(options.kickerName) ?? "Transition Kicker",
          fillColor: trimString(options.kickerFillColor),
          lineColor: trimString(options.kickerLineColor),
        });
        const kickerFrame = kickerShape.getTextFrameOrNullObject();
        applyPowerPointTextFrameProperties(kickerFrame, { autoSizeSetting: "AutoSizeShapeToFitText", wordWrap: false });
        if (typeof toNumber(options.kickerFontSize) === "number") {
          kickerFrame.textRange.font.size = toNumber(options.kickerFontSize)!;
        }
        if (trimString(options.kickerFontColor)) {
          kickerFrame.textRange.font.color = trimString(options.kickerFontColor)!;
        }
      }

      const titleShape = slide.shapes.addTextBox(titleText, {
        left: toNumber(options.titleLeft) ?? 72,
        top: toNumber(options.titleTop) ?? (kickerText ? 124 : 132),
        width: toNumber(options.titleWidth) ?? 580,
        height: toNumber(options.titleHeight) ?? 96,
      });
      createdShapes.push(titleShape);
      applyPowerPointShapeProperties(titleShape, {
        name: trimString(options.titleName) ?? "Transition Title",
        fillColor: trimString(options.titleFillColor),
        lineColor: trimString(options.titleLineColor),
      });
      const titleFrame = titleShape.getTextFrameOrNullObject();
      applyPowerPointTextFrameProperties(titleFrame, {
        autoSizeSetting: trimString(options.titleAutoSizeSetting) ?? "AutoSizeShapeToFitText",
        wordWrap: toBoolean(options.titleWordWrap) ?? true,
      });
      if (typeof toNumber(options.titleFontSize) === "number") {
        titleFrame.textRange.font.size = toNumber(options.titleFontSize)!;
      }
      if (trimString(options.titleFontColor)) {
        titleFrame.textRange.font.color = trimString(options.titleFontColor)!;
      }
      if (typeof toBoolean(options.titleBold) === "boolean") {
        titleFrame.textRange.font.bold = toBoolean(options.titleBold)!;
      }

      if (subtitleText) {
        const subtitleShape = slide.shapes.addTextBox(subtitleText, {
          left: toNumber(options.subtitleLeft) ?? 72,
          top: toNumber(options.subtitleTop) ?? (kickerText ? 236 : 246),
          width: toNumber(options.subtitleWidth) ?? 560,
          height: toNumber(options.subtitleHeight) ?? 58,
        });
        createdShapes.push(subtitleShape);
        applyPowerPointShapeProperties(subtitleShape, {
          name: trimString(options.subtitleName) ?? "Transition Subtitle",
          fillColor: trimString(options.subtitleFillColor),
          lineColor: trimString(options.subtitleLineColor),
        });
        const subtitleFrame = subtitleShape.getTextFrameOrNullObject();
        applyPowerPointTextFrameProperties(subtitleFrame, {
          autoSizeSetting: trimString(options.subtitleAutoSizeSetting) ?? "AutoSizeShapeToFitText",
          wordWrap: toBoolean(options.subtitleWordWrap) ?? true,
        });
        if (typeof toNumber(options.subtitleFontSize) === "number") {
          subtitleFrame.textRange.font.size = toNumber(options.subtitleFontSize)!;
        }
        if (trimString(options.subtitleFontColor)) {
          subtitleFrame.textRange.font.color = trimString(options.subtitleFontColor)!;
        }
      }

      for (const shape of createdShapes) {
        shape.load("id,name,type");
      }
      await context.sync();
      context.presentation.setSelectedSlides([slide.id]);
      slide.setSelectedShapes([titleShape.id]);
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        createdSlides: [summarizePowerPointSlide(slide)],
        createdShapes: createdShapes.map((shape) => summarizePowerPointShape(slide, shape)),
      };
    }

    if (type === "combineSlides") {
      const requestedSourceSlideIds = getPowerPointSlideIdArray(options.sourceSlideIds ?? options.slideIds ?? options.sourceSlides);
      const sourceSlides = await resolvePowerPointSourceSlides(context, undefined, requestedSourceSlideIds);
      if (!sourceSlides.length) {
        throw new Error("PowerPoint combineSlides requires at least one source slide.");
      }
      const sourceContent = await loadPowerPointSlideContentSummaries(context, sourceSlides);
      const slide = await createPowerPointSlide(context, action.target, options);
      slide.load("id,index");
      const titleText = trimString(options.title) ?? trimString(action.content) ?? "Combined Overview";
      const titleShape = slide.shapes.addTextBox(titleText, {
        left: toNumber(options.titleLeft) ?? 54,
        top: toNumber(options.titleTop) ?? 40,
        width: toNumber(options.titleWidth) ?? 620,
        height: toNumber(options.titleHeight) ?? 40,
      });
      const contentShape = slide.shapes.addTextBox(
        sourceContent
          .map((entry) => {
            const bodyPreview =
              truncateText(
                entry.textBlocks
                  .slice(0, resolvePositiveCount(options.maxBlocksPerSlide, 2))
                  .join(" "),
                toNumber(options.maxCharsPerSlide) ?? 240,
              ) ?? "";
            return bodyPreview && bodyPreview !== entry.title ? `${entry.title ?? entry.label}\n${bodyPreview}` : (entry.title ?? entry.label);
          })
          .join("\n\n"),
        {
          left: toNumber(options.bodyLeft) ?? 72,
          top: toNumber(options.bodyTop) ?? 120,
          width: toNumber(options.bodyWidth) ?? 560,
          height: toNumber(options.bodyHeight) ?? 300,
        },
      );
      const createdShapes = [titleShape, contentShape];
      applyPowerPointShapeProperties(titleShape, {
        name: trimString(options.titleName) ?? "Combined Title",
        fillColor: trimString(options.titleFillColor),
        lineColor: trimString(options.titleLineColor),
      });
      applyPowerPointShapeProperties(contentShape, {
        name: trimString(options.bodyName) ?? "Combined Content",
        fillColor: trimString(options.bodyFillColor),
        lineColor: trimString(options.bodyLineColor),
      });
      applyPowerPointTextFrameProperties(titleShape.getTextFrameOrNullObject(), {
        autoSizeSetting: trimString(options.titleAutoSizeSetting) ?? "AutoSizeShapeToFitText",
        wordWrap: toBoolean(options.titleWordWrap) ?? true,
      });
      applyPowerPointTextFrameProperties(contentShape.getTextFrameOrNullObject(), {
        autoSizeSetting: trimString(options.bodyAutoSizeSetting) ?? "AutoSizeShapeToFitText",
        wordWrap: toBoolean(options.bodyWordWrap) ?? true,
      });
      for (const shape of createdShapes) {
        shape.load("id,name,type");
      }
      await context.sync();

      const deletedSlides: Array<{ slideId: string; slideIndex: number; label: string }> = [];
      if (toBoolean(options.deleteSourceSlides ?? action.deleteSourceSlides)) {
        for (const sourceSlide of sourceSlides) {
          context.presentation.slides.getItem(sourceSlide.slideId).delete();
          deletedSlides.push(sourceSlide);
        }
        await context.sync();
        slide.load("id,index");
        await context.sync();
      }

      context.presentation.setSelectedSlides([slide.id]);
      slide.setSelectedShapes([contentShape.id]);
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        createdSlides: [summarizePowerPointSlide(slide)],
        createdShapes: createdShapes.map((shape) => summarizePowerPointShape(slide, shape)),
        sourceSlideIds: sourceSlides.map((entry) => entry.slideId),
        deletedSlides: deletedSlides.length ? deletedSlides : undefined,
      };
    }

    if (type === "reorderSlides" || type === "reorderStoryline") {
      const desiredSlideIds = Array.from(new Set(getPowerPointSlideIdArray(options.slideIds ?? options.storyline ?? options.slides)));
      if (desiredSlideIds.length < 2) {
        throw new Error(`PowerPoint ${type} requires at least two slide IDs.`);
      }

      const beforeSlides = await loadPowerPointSlideSummaries(context);
      const beforeIndexById = new Map(beforeSlides.map((slide, index) => [slide.slideId, index]));
      const missingSlideId = desiredSlideIds.find((slideId) => !beforeIndexById.has(slideId));
      if (missingSlideId) {
        throw new Error(`Could not find the requested PowerPoint slide: ${missingSlideId}.`);
      }

      let insertionIndex = await resolvePowerPointInsertionIndex(context, undefined, options);
      if (typeof insertionIndex !== "number" || Number.isNaN(insertionIndex)) {
        insertionIndex = Math.min(...desiredSlideIds.map((slideId) => beforeIndexById.get(slideId) ?? 0));
      }

      const relativeSlideId = trimString(options.relativeToSlideId) ?? trimString(options.targetSlideId);
      if (relativeSlideId && desiredSlideIds.includes(relativeSlideId)) {
        insertionIndex = Math.min(...desiredSlideIds.map((slideId) => beforeIndexById.get(slideId) ?? 0));
      }

      for (let index = 0; index < desiredSlideIds.length; index += 1) {
        context.presentation.slides.getItem(desiredSlideIds[index]!).moveTo(insertionIndex + index);
      }
      await context.sync();
      const afterSlides = await loadPowerPointSlideSummaries(context);
      const afterById = new Map(afterSlides.map((slide) => [slide.slideId, slide]));
      const reorderedSlides = desiredSlideIds
        .map((slideId) => afterById.get(slideId))
        .filter((slide): slide is { slideId: string; slideIndex: number; label: string } => Boolean(slide));
      context.presentation.setSelectedSlides(reorderedSlides.map((slide) => slide.slideId));
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideIds: reorderedSlides.map((slide) => slide.slideId),
        slides: reorderedSlides,
        startIndex: insertionIndex + 1,
      };
    }

    if (type === "applyLayout") {
      const slide = await resolvePowerPointSlide(context, action.target, true);
      const { layout } = await resolvePowerPointLayoutSelection(context, action.target, options);
      if (!layout) {
        throw new Error("PowerPoint applyLayout requires a resolvable layout.");
      }
      slide.applyLayout(layout);
      slide.load("id,index");
      slide.layout.load("id,name,type");
      slide.slideMaster.load("id,name");
      await context.sync();
      context.presentation.setSelectedSlides([slide.id]);
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        layoutId: slide.layout.id,
        layoutName: slide.layout.name,
        layoutType: slide.layout.type,
        slideMasterId: slide.slideMaster.id,
        slideMasterName: slide.slideMaster.name,
      };
    }

    if (type === "moveSlide") {
      const slide = await resolvePowerPointSlide(context, action.target, false);
      const destinationIndex = Math.max(0, (toNumber(options.slideIndex) ?? 1) - 1);
      slide.load("id,index");
      slide.moveTo(destinationIndex);
      await context.sync();
      return { ok: true, host: "powerpoint", action: type, slideId: slide.id, slideIndex: destinationIndex + 1 };
    }

    if (type === "duplicateSlide" || type === "duplicateSlides") {
      const requestedSlideIds = getStringArray(options.slideIds);
      const sourceSlides = await resolvePowerPointSourceSlides(context, action.target, requestedSlideIds);
      if (!sourceSlides.length) {
        throw new Error("PowerPoint duplication requires at least one source slide.");
      }

      const beforeSlides = await loadPowerPointSlideSummaries(context);
      const formatting = trimString(options.formatting);
      let explicitTargetSlideId = trimString(options.targetSlideId);
      if (!explicitTargetSlideId && requestedSlideIds.length > 0 && (action.target?.slideId || typeof action.target?.slideIndex === "number")) {
        const targetSlide = await resolvePowerPointSlide(context, action.target, false);
        targetSlide.load("id");
        await context.sync();
        explicitTargetSlideId = targetSlide.id;
      }
      const targetSlideId =
        explicitTargetSlideId ??
        sourceSlides
          .slice()
          .sort((left, right) => left.slideIndex - right.slideIndex)
          .at(-1)?.slideId;

      let exportResult: OfficeExtension.ClientResult<string>;
      if (sourceSlides.length === 1) {
        if (!supportsRequirementSet("PowerPointApi", "1.8")) {
          throw new Error("PowerPoint single-slide duplication requires PowerPointApi 1.8.");
        }
        const sourceSlide = sourceSlides[0];
        if (!sourceSlide) {
          throw new Error("PowerPoint duplication could not resolve the requested source slide.");
        }
        exportResult = context.presentation.slides.getItem(sourceSlide.slideId).exportAsBase64();
      } else {
        if (!supportsRequirementSet("PowerPointApi", "1.10")) {
          throw new Error("PowerPoint multi-slide duplication requires PowerPointApi 1.10.");
        }
        exportResult = context.presentation.slides.exportAsBase64Presentation(sourceSlides.map((slide) => slide.slideId));
      }
      await context.sync();

      const insertOptions: PowerPoint.InsertSlideOptions = {};
      if (formatting) {
        insertOptions.formatting = formatting as PowerPoint.InsertSlideFormatting;
      }
      if (targetSlideId) {
        insertOptions.targetSlideId = targetSlideId;
      }
      context.presentation.insertSlidesFromBase64(exportResult.value, insertOptions);
      await context.sync();

      const duplicatedSlides = sliceInsertedPowerPointSlides(beforeSlides, await loadPowerPointSlideSummaries(context), targetSlideId, sourceSlides.length);
      if (duplicatedSlides.length) {
        context.presentation.setSelectedSlides(duplicatedSlides.map((slide) => slide.slideId));
        await context.sync();
      }
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: duplicatedSlides[0]?.slideId,
        slideIndex: duplicatedSlides[0]?.slideIndex,
        slideIds: duplicatedSlides.map((slide) => slide.slideId),
        slides: duplicatedSlides,
        sourceSlideIds: sourceSlides.map((slide) => slide.slideId),
        formatting,
        targetSlideId,
      };
    }

    if (type === "deleteSlide" || type === "deleteSlides") {
      const beforeSlides = await loadPowerPointSlideSummaries(context);
      const deleteSlides = await resolvePowerPointSourceSlides(context, action.target, getStringArray(options.slideIds));
      const deleteSlideIds = Array.from(new Set(deleteSlides.map((slide) => slide.slideId)));
      if (!deleteSlideIds.length) {
        throw new Error("PowerPoint delete requires at least one slide.");
      }
      if (deleteSlideIds.length >= beforeSlides.length) {
        throw new Error("PowerPoint delete must leave at least one slide in the presentation.");
      }

      const firstDeletedIndex = beforeSlides.findIndex((slide) => deleteSlideIds.includes(slide.slideId));
      const remainingBefore = beforeSlides.filter((slide) => !deleteSlideIds.includes(slide.slideId));
      const nextSelection = remainingBefore[Math.min(Math.max(firstDeletedIndex, 0), remainingBefore.length - 1)];

      for (const slideId of deleteSlideIds) {
        context.presentation.slides.getItem(slideId).delete();
      }
      await context.sync();
      if (nextSelection?.slideId) {
        context.presentation.setSelectedSlides([nextSelection.slideId]);
        await context.sync();
      }

      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: nextSelection?.slideId,
        slideIndex: nextSelection?.slideIndex,
        deletedCount: deleteSlideIds.length,
        deletedSlideIds: deleteSlideIds,
        deletedSlides: deleteSlides,
      };
    }

    if (type === "selectSlides") {
      const slideIds = getStringArray(options.slideIds);
      if (slideIds.length) {
        context.presentation.setSelectedSlides(slideIds);
        await context.sync();
        return { ok: true, host: "powerpoint", action: type, slideIds };
      }

      const slide = await resolvePowerPointSlide(context, action.target, false);
      slide.load("id,index");
      await context.sync();
      context.presentation.setSelectedSlides([slide.id]);
      await context.sync();
      return { ok: true, host: "powerpoint", action: type, slideId: slide.id, slideIndex: slide.index + 1 };
    }

    if (type === "addTextBox") {
      const slide = await resolvePowerPointSlide(context, action.target, true);
      const shapeOptions: PowerPoint.ShapeAddOptions = {};
      const left = toNumber(options.left);
      const top = toNumber(options.top);
      const width = toNumber(options.width);
      const height = toNumber(options.height);
      if (typeof left === "number") shapeOptions.left = left;
      if (typeof top === "number") shapeOptions.top = top;
      if (typeof width === "number") shapeOptions.width = width;
      if (typeof height === "number") shapeOptions.height = height;
      const shape = slide.shapes.addTextBox(action.content ?? "", shapeOptions);
      applyPowerPointShapeProperties(shape, options);
      return finalizePowerPointShapeSelection(context, slide, shape, type);
    }

    if (type === "addGeometricShape") {
      const slide = await resolvePowerPointSlide(context, action.target, true);
      const shapeOptions: PowerPoint.ShapeAddOptions = {};
      const left = toNumber(options.left);
      const top = toNumber(options.top);
      const width = toNumber(options.width);
      const height = toNumber(options.height);
      const geometricShapeType = trimString(options.geometricShapeType) ?? trimString(options.shapeType) ?? trimString(options.type);
      if (!geometricShapeType) {
        throw new Error("PowerPoint addGeometricShape requires a geometricShapeType.");
      }
      if (typeof left === "number") shapeOptions.left = left;
      if (typeof top === "number") shapeOptions.top = top;
      if (typeof width === "number") shapeOptions.width = width;
      if (typeof height === "number") shapeOptions.height = height;
      const shape = slide.shapes.addGeometricShape(geometricShapeType as PowerPoint.GeometricShapeType, shapeOptions);
      applyPowerPointShapeProperties(shape, options);
      return finalizePowerPointShapeSelection(context, slide, shape, type, { geometricShapeType });
    }

    if (type === "groupShapes") {
      const slide = await resolvePowerPointSlide(context, action.target, true);
      const shapeIds = Array.from(
        new Set([action.target?.shapeId, ...getStringArray(options.shapeIds), ...getStringArray(options.additionalShapeIds)].filter(
          (value): value is string => Boolean(value),
        )),
      );
      if (shapeIds.length < 2) {
        throw new Error("PowerPoint groupShapes requires at least two shape IDs.");
      }
      const shape = slide.shapes.addGroup(shapeIds);
      applyPowerPointShapeProperties(shape, options);
      return finalizePowerPointShapeSelection(context, slide, shape, type, { groupedShapeIds: shapeIds });
    }

    if (type === "addTable") {
      const slide = await resolvePowerPointSlide(context, action.target, true);
      const values = toStringMatrix(options.values ?? action.values);
      const rowCount = toNumber(options.rowCount) ?? values?.length ?? 2;
      const columnCount = toNumber(options.columnCount) ?? values?.[0]?.length ?? 2;
      const tableOptions: PowerPoint.TableAddOptions = {};
      const tableLeft = toNumber(options.left);
      const tableTop = toNumber(options.top);
      const tableWidth = toNumber(options.width);
      const tableHeight = toNumber(options.height);
      const tableStyle = trimString(options.style);
      if (typeof tableLeft === "number") tableOptions.left = tableLeft;
      if (typeof tableTop === "number") tableOptions.top = tableTop;
      if (typeof tableWidth === "number") tableOptions.width = tableWidth;
      if (typeof tableHeight === "number") tableOptions.height = tableHeight;
      if (tableStyle) tableOptions.style = tableStyle as PowerPoint.TableStyle;
      if (values) tableOptions.values = values;
      const shape = slide.shapes.addTable(rowCount, columnCount, tableOptions);
      applyPowerPointShapeProperties(shape, options);
      return finalizePowerPointShapeSelection(context, slide, shape, type);
    }

    if (type === "setTableValues" || type === "updateTable") {
      const { slide, shape, table } = await resolvePowerPointTable(context, action.target, true);
      const values = toStringMatrix(options.values ?? action.values);
      if (!values?.length) {
        throw new Error("PowerPoint setTableValues requires a matrix of cell values.");
      }

      const cells: PowerPoint.TableCell[] = [];
      for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
        const row = values[rowIndex] ?? [];
        if (rowIndex >= table.rowCount) {
          throw new Error(`PowerPoint table update exceeds the available row count (${table.rowCount}).`);
        }
        for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
          if (columnIndex >= table.columnCount) {
            throw new Error(`PowerPoint table update exceeds the available column count (${table.columnCount}).`);
          }
          const cell = table.getCellOrNullObject(rowIndex, columnIndex);
          cell.load("isNullObject");
          cells.push(cell);
        }
      }
      await context.sync();
      let cellOffset = 0;
      for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
        const row = values[rowIndex] ?? [];
        for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
          const cell = cells[cellOffset];
          cellOffset += 1;
          if (!cell || cell.isNullObject) {
            throw new Error(`Could not resolve PowerPoint table cell (${rowIndex + 1}, ${columnIndex + 1}).`);
          }
          cell.text = row[columnIndex] ?? "";
        }
      }
      await context.sync();
      return finalizePowerPointShapeSelection(context, slide, shape, type, {
        rowCount: table.rowCount,
        columnCount: table.columnCount,
        values: truncateStringMatrix(values),
      });
    }

    if (type === "setTableCell" || type === "updateTableCell") {
      const { slide, shape, table } = await resolvePowerPointTable(context, action.target, true);
      if (!supportsRequirementSet("PowerPointApi", "1.9")) {
        throw new Error("PowerPoint updateTableCell requires PowerPointApi 1.9.");
      }
      const rowIndex = resolveZeroBasedIndex(actionOptions.rowIndex, actionOptions.rowNumber);
      const columnIndex = resolveZeroBasedIndex(actionOptions.columnIndex, actionOptions.columnNumber);
      if (typeof rowIndex !== "number" || typeof columnIndex !== "number") {
        throw new Error("PowerPoint updateTableCell requires rowIndex/rowNumber and columnIndex/columnNumber.");
      }
      const cell = table.getCellOrNullObject(rowIndex, columnIndex);
      cell.load("isNullObject,rowIndex,columnIndex,rowCount,columnCount,text");
      await context.sync();
      if (cell.isNullObject) {
        throw new Error(`Could not resolve PowerPoint table cell (${rowIndex + 1}, ${columnIndex + 1}).`);
      }
      applyPowerPointTableCellProperties(cell, {
        ...actionOptions,
        text: trimString(action.content) ?? trimString(actionOptions.text),
      });
      await context.sync();
      return finalizePowerPointShapeSelection(context, slide, shape, type, {
        rowIndex: cell.rowIndex,
        columnIndex: cell.columnIndex,
        rowCount: cell.rowCount,
        columnCount: cell.columnCount,
        text: cell.text,
      });
    }

    if (type === "addTableRows") {
      if (!supportsRequirementSet("PowerPointApi", "1.9")) {
        throw new Error("PowerPoint addTableRows requires PowerPointApi 1.9.");
      }
      const { slide, shape, table } = await resolvePowerPointTable(context, action.target, true);
      const insertIndex = resolveZeroBasedIndex(actionOptions.rowIndex, actionOptions.rowNumber);
      const rowCount = resolvePositiveCount(actionOptions.rowCount);
      table.rows.add(insertIndex, rowCount);
      await context.sync();
      return finalizePowerPointShapeSelection(context, slide, shape, type, {
        insertedRowIndex: insertIndex ?? table.rowCount,
        insertedRowCount: rowCount,
      });
    }

    if (type === "deleteTableRows") {
      if (!supportsRequirementSet("PowerPointApi", "1.9")) {
        throw new Error("PowerPoint deleteTableRows requires PowerPointApi 1.9.");
      }
      const { slide, shape, table } = await resolvePowerPointTable(context, action.target, true);
      const rowIndexes = Array.from(
        new Set([
          ...getNumberArray(actionOptions.rowIndexes),
          ...getNumberArray(actionOptions.rows),
          resolveZeroBasedIndex(actionOptions.rowIndex, actionOptions.rowNumber),
        ].filter((value): value is number => typeof value === "number" && value >= 0)),
      );
      if (!rowIndexes.length) {
        throw new Error("PowerPoint deleteTableRows requires at least one row index.");
      }
      const rows = rowIndexes.map((rowIndex) => table.rows.getItemAt(rowIndex));
      table.rows.deleteRows(rows);
      await context.sync();
      return finalizePowerPointShapeSelection(context, slide, shape, type, { deletedRowIndexes: rowIndexes });
    }

    if (type === "addTableColumns") {
      if (!supportsRequirementSet("PowerPointApi", "1.9")) {
        throw new Error("PowerPoint addTableColumns requires PowerPointApi 1.9.");
      }
      const { slide, shape, table } = await resolvePowerPointTable(context, action.target, true);
      const insertIndex = resolveZeroBasedIndex(actionOptions.columnIndex, actionOptions.columnNumber);
      const columnCount = resolvePositiveCount(actionOptions.columnCount);
      table.columns.add(insertIndex, columnCount);
      await context.sync();
      return finalizePowerPointShapeSelection(context, slide, shape, type, {
        insertedColumnIndex: insertIndex ?? table.columnCount,
        insertedColumnCount: columnCount,
      });
    }

    if (type === "deleteTableColumns") {
      if (!supportsRequirementSet("PowerPointApi", "1.9")) {
        throw new Error("PowerPoint deleteTableColumns requires PowerPointApi 1.9.");
      }
      const { slide, shape, table } = await resolvePowerPointTable(context, action.target, true);
      const columnIndexes = Array.from(
        new Set([
          ...getNumberArray(actionOptions.columnIndexes),
          ...getNumberArray(actionOptions.columns),
          resolveZeroBasedIndex(actionOptions.columnIndex, actionOptions.columnNumber),
        ].filter((value): value is number => typeof value === "number" && value >= 0)),
      );
      if (!columnIndexes.length) {
        throw new Error("PowerPoint deleteTableColumns requires at least one column index.");
      }
      const columns = columnIndexes.map((columnIndex) => table.columns.getItemAt(columnIndex));
      table.columns.deleteColumns(columns);
      await context.sync();
      return finalizePowerPointShapeSelection(context, slide, shape, type, { deletedColumnIndexes: columnIndexes });
    }

    if (type === "clearTable") {
      if (!supportsRequirementSet("PowerPointApi", "1.9")) {
        throw new Error("PowerPoint clearTable requires PowerPointApi 1.9.");
      }
      const { slide, shape, table } = await resolvePowerPointTable(context, action.target, true);
      table.clear();
      await context.sync();
      return finalizePowerPointShapeSelection(context, slide, shape, type);
    }

    if (type === "mergeTableCells") {
      if (!supportsRequirementSet("PowerPointApi", "1.9")) {
        throw new Error("PowerPoint mergeTableCells requires PowerPointApi 1.9.");
      }
      const { slide, shape, table } = await resolvePowerPointTable(context, action.target, true);
      const rowIndex = resolveZeroBasedIndex(actionOptions.rowIndex, actionOptions.rowNumber);
      const columnIndex = resolveZeroBasedIndex(actionOptions.columnIndex, actionOptions.columnNumber);
      if (typeof rowIndex !== "number" || typeof columnIndex !== "number") {
        throw new Error("PowerPoint mergeTableCells requires rowIndex/rowNumber and columnIndex/columnNumber.");
      }
      const rowCount = resolvePositiveCount(actionOptions.rowCount);
      const columnCount = resolvePositiveCount(actionOptions.columnCount);
      table.mergeCells(rowIndex, columnIndex, rowCount, columnCount);
      await context.sync();
      return finalizePowerPointShapeSelection(context, slide, shape, type, {
        rowIndex,
        columnIndex,
        rowCount,
        columnCount,
      });
    }

    if (type === "resizeTableCell" || type === "splitTableCell") {
      if (!supportsRequirementSet("PowerPointApi", "1.9")) {
        throw new Error(`PowerPoint ${type} requires PowerPointApi 1.9.`);
      }
      const { slide, shape, table } = await resolvePowerPointTable(context, action.target, true);
      const rowIndex = resolveZeroBasedIndex(actionOptions.rowIndex, actionOptions.rowNumber);
      const columnIndex = resolveZeroBasedIndex(actionOptions.columnIndex, actionOptions.columnNumber);
      if (typeof rowIndex !== "number" || typeof columnIndex !== "number") {
        throw new Error(`PowerPoint ${type} requires rowIndex/rowNumber and columnIndex/columnNumber.`);
      }
      const cell = table.getCellOrNullObject(rowIndex, columnIndex);
      cell.load("isNullObject,rowIndex,columnIndex");
      await context.sync();
      if (cell.isNullObject) {
        throw new Error(`Could not resolve PowerPoint table cell (${rowIndex + 1}, ${columnIndex + 1}).`);
      }
      const rowCount = resolvePositiveCount(actionOptions.rowCount);
      const columnCount = resolvePositiveCount(actionOptions.columnCount);
      if (type === "resizeTableCell") {
        cell.resize(rowCount, columnCount);
      } else {
        cell.split(rowCount, columnCount);
      }
      await context.sync();
      return finalizePowerPointShapeSelection(context, slide, shape, type, {
        rowIndex: cell.rowIndex,
        columnIndex: cell.columnIndex,
        rowCount,
        columnCount,
      });
    }

    if (type === "addLine") {
      const slide = await resolvePowerPointSlide(context, action.target, true);
      const lineOptions: PowerPoint.ShapeAddOptions = {};
      const lineLeft = toNumber(options.left);
      const lineTop = toNumber(options.top);
      const lineWidth = toNumber(options.width);
      const lineHeight = toNumber(options.height);
      if (typeof lineLeft === "number") lineOptions.left = lineLeft;
      if (typeof lineTop === "number") lineOptions.top = lineTop;
      if (typeof lineWidth === "number") lineOptions.width = lineWidth;
      if (typeof lineHeight === "number") lineOptions.height = lineHeight;
      const shape = slide.shapes.addLine(trimString(options.connectorType) as PowerPoint.ConnectorType, lineOptions);
      applyPowerPointShapeProperties(shape, options);
      return finalizePowerPointShapeSelection(context, slide, shape, type);
    }

    if (type === "addProcessFlow" || type === "addSimpleDiagram") {
      const slide = await resolvePowerPointSlide(context, action.target, true);
      slide.load("id,index");
      const stepTexts = (
        Array.isArray(actionOptions.steps)
          ? actionOptions.steps
          : typeof action.content === "string"
            ? action.content
                .split(/\r?\n|>/)
                .map((entry) => trimString(entry))
                .filter((entry): entry is string => Boolean(entry))
            : []
      )
        .map((entry) => trimString(entry))
        .filter((entry): entry is string => Boolean(entry));
      if (!stepTexts.length) {
        throw new Error("PowerPoint addProcessFlow requires steps or newline-delimited content.");
      }

      const orientation = (trimString(actionOptions.orientation) ?? "horizontal").toLowerCase();
      const stepWidth = toNumber(actionOptions.stepWidth) ?? 140;
      const stepHeight = toNumber(actionOptions.stepHeight) ?? 72;
      const spacing = toNumber(actionOptions.spacing) ?? 36;
      const startLeft = toNumber(actionOptions.left) ?? 60;
      const startTop = toNumber(actionOptions.top) ?? 140;
      const connectorType = (trimString(actionOptions.connectorType) ?? "Straight") as PowerPoint.ConnectorType;
      const geometricShapeType = (trimString(actionOptions.geometricShapeType) ?? "FlowChartProcess") as PowerPoint.GeometricShapeType;
      const stepShapes: PowerPoint.Shape[] = [];
      const connectorShapes: PowerPoint.Shape[] = [];
      const textFrames: PowerPoint.TextFrame[] = [];
      const baseName = trimString(actionOptions.name) ?? (type === "addSimpleDiagram" ? "Diagram Step" : "Process Step");

      for (let index = 0; index < stepTexts.length; index += 1) {
        const left = orientation === "vertical" ? startLeft : startLeft + index * (stepWidth + spacing);
        const top = orientation === "vertical" ? startTop + index * (stepHeight + spacing) : startTop;
        const stepShape = slide.shapes.addGeometricShape(geometricShapeType, {
          left,
          top,
          width: stepWidth,
          height: stepHeight,
        });
        applyPowerPointShapeProperties(stepShape, {
          ...actionOptions,
          left,
          top,
          width: stepWidth,
          height: stepHeight,
          name: `${baseName} ${index + 1}`,
        });
        stepShapes.push(stepShape);
        textFrames.push(stepShape.getTextFrameOrNullObject());

        if (index < stepTexts.length - 1) {
          const connectorLeft = orientation === "vertical" ? startLeft + stepWidth / 2 : left + stepWidth;
          const connectorTop = orientation === "vertical" ? top + stepHeight : startTop + stepHeight / 2;
          const connectorWidth = orientation === "vertical" ? 0 : spacing;
          const connectorHeight = orientation === "vertical" ? spacing : 0;
          const connectorShape = slide.shapes.addLine(connectorType, {
            left: connectorLeft,
            top: connectorTop,
            width: connectorWidth,
            height: connectorHeight,
          });
          applyPowerPointShapeProperties(connectorShape, {
            lineColor: trimString(actionOptions.connectorColor) ?? trimString(actionOptions.lineColor),
            lineWeight: toNumber(actionOptions.connectorWeight) ?? toNumber(actionOptions.lineWeight),
            lineTransparency: actionOptions.lineTransparency,
            lineDashStyle: trimString(actionOptions.lineDashStyle),
            lineStyle: trimString(actionOptions.lineStyle),
          });
          connectorShapes.push(connectorShape);
        }
      }

      for (const textFrame of textFrames) {
        textFrame.load("isNullObject");
      }
      await context.sync();
      for (let index = 0; index < textFrames.length; index += 1) {
        const textFrame = textFrames[index];
        if (!textFrame || textFrame.isNullObject) {
          continue;
        }
        textFrame.textRange.text = stepTexts[index] ?? "";
        applyPowerPointTextFrameProperties(textFrame, actionOptions);
      }
      await context.sync();

      for (const shape of [...stepShapes, ...connectorShapes]) {
        shape.load("id,name,type");
      }
      await context.sync();
      slide.setSelectedShapes(stepShapes.map((shape) => shape.id));
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        createdShapes: [...stepShapes, ...connectorShapes].map((shape) => summarizePowerPointShape(slide, shape)),
        stepCount: stepShapes.length,
        connectorCount: connectorShapes.length,
        orientation,
        geometricShapeType,
      };
    }

    if (type === "importSlidesFromBase64" || type === "mergePresentationFromBase64") {
      const base64 = trimString(action.content) ?? trimString(options.base64);
      if (!base64) {
        throw new Error("PowerPoint slide import requires a base64 presentation payload.");
      }
      const beforeSlides = await loadPowerPointSlideSummaries(context);
      let targetSlideId = trimString(options.targetSlideId) ?? action.target?.slideId;
      if (!targetSlideId && typeof action.target?.slideIndex === "number") {
        const targetSlide = await resolvePowerPointSlide(context, action.target, false);
        targetSlide.load("id");
        await context.sync();
        targetSlideId = targetSlide.id;
      }
      const insertOptions: PowerPoint.InsertSlideOptions = {};
      const formatting = trimString(options.formatting);
      const sourceSlideIds = getStringArray(options.sourceSlideIds);
      if (formatting) {
        insertOptions.formatting = formatting as PowerPoint.InsertSlideFormatting;
      }
      if (sourceSlideIds.length) {
        insertOptions.sourceSlideIds = sourceSlideIds;
      }
      if (targetSlideId) {
        insertOptions.targetSlideId = targetSlideId;
      }
      context.presentation.insertSlidesFromBase64(base64, insertOptions);
      await context.sync();
      const insertedSlides = sliceInsertedPowerPointSlides(
        beforeSlides,
        await loadPowerPointSlideSummaries(context),
        targetSlideId,
      );
      if (insertedSlides.length) {
        context.presentation.setSelectedSlides(insertedSlides.map((slide) => slide.slideId));
        await context.sync();
      }
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: insertedSlides[0]?.slideId,
        slideIndex: insertedSlides[0]?.slideIndex,
        slideIds: insertedSlides.map((slide) => slide.slideId),
        slides: insertedSlides,
        sourceSlideIds: sourceSlideIds.length ? sourceSlideIds : undefined,
        formatting,
        targetSlideId,
        insertedCount: insertedSlides.length,
      };
    }

    if (type === "exportSlidesAsBase64") {
      const slideIds = getStringArray(options.slideIds);
      const exportResult =
        slideIds.length > 0
          ? context.presentation.slides.exportAsBase64Presentation(slideIds)
          : (await resolvePowerPointSlide(context, action.target, true)).exportAsBase64();
      await context.sync();
      return { ok: true, host: "powerpoint", action: type, base64: exportResult.value };
    }

    if (type === "insertInlinePicture") {
      const slide = await resolvePowerPointSlide(context, action.target, true);
      const imageBase64 = action.content ?? "";
      const shapes = slide.shapes as any;
      if (typeof shapes.addImage !== "function") {
        throw new Error("PowerPoint image insertion requires PowerPointApi 1.4 or newer.");
      }
      const shape = shapes.addImage(`data:image/png;base64,${imageBase64}`) as PowerPoint.Shape;
      shape.load("id,name,width,height");
      await context.sync();
      return { ok: true, host: "powerpoint", action: type, shapeId: shape.id, shapeName: shape.name, slideId: slide.id };
    }

    throw new Error(`Unsupported PowerPoint action: ${type}`);
  });
}

const officeHostAdapters: Record<OfficeHost, OfficeHostAdapter> = {
  word: createWordOfficeHostAdapter({
    collectState: collectWordState,
    collectContext: collectWordContext,
    navigateAnchor: navigateWordAnchor,
    applyAction: applyWordAction,
  }),
  excel: createExcelOfficeHostAdapter({
    collectState: collectExcelState,
    collectContext: collectExcelContext,
    navigateAnchor: navigateExcelAnchor,
    applyAction: applyExcelAction,
  }),
  powerpoint: createPowerPointOfficeHostAdapter({
    collectState: collectPowerPointState,
    collectContext: collectPowerPointContext,
    navigateAnchor: navigatePowerPointAnchor,
    applyAction: applyPowerPointAction,
  }),
};

function getOfficeHostAdapter(host: OfficeHost): OfficeHostAdapter {
  return officeHostAdapters[host];
}

export async function navigateOfficeAnchor(host: OfficeHost, anchor: OfficeAnchor): Promise<unknown> {
  const adapter = getOfficeHostAdapter(host);
  return buildOfficeActionResult(host, `navigate:${anchor.kind}`, await adapter.navigateAnchor(anchor), {
    navigation: anchor,
    target: anchor,
  });
}

export async function applyHostAction(host: OfficeHost, action: OfficeHostAction): Promise<unknown> {
  assertDestructiveActionAllowed(host, action);
  const adapter = getOfficeHostAdapter(host);
  return buildOfficeActionResult(host, action.type, await adapter.applyAction(action), { target: action.target });
}

export async function applyHostEdit(host: OfficeHost, params: { mode: string; content: string; format?: string }): Promise<unknown> {
  const action: OfficeHostAction =
    params.format === "matrix" || params.mode === "setRangeValues"
      ? { type: "setRangeValues", values: toValueMatrix(params.content, params.content), format: params.format, content: params.content }
      : {
          type: params.format === "html" ? "insertHtml" : "insertText",
          placement: params.mode === "insertAfterSelection" ? "after" : "replace",
          format: params.format,
          content: params.content,
        };

  return applyHostAction(host, action);
}

// ---------------------------------------------------------------------------
// Document snapshot / restore for rewind
// ---------------------------------------------------------------------------

export interface DocumentSnapshotData {
  ooxml?: string;
  sheets?: Array<{
    name: string;
    usedRangeAddress: string;
    values: unknown[][];
    numberFormats: string[][];
    formulas: string[][];
  }>;
  presentationBase64?: string;
}

export async function captureDocumentSnapshot(host: OfficeHost): Promise<DocumentSnapshotData> {
  if (host === "word") {
    return Word.run(async (context) => {
      const body = context.document.body;
      const ooxml = body.getOoxml();
      await context.sync();
      return { ooxml: ooxml.value };
    });
  }

  if (host === "excel") {
    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      const sheetData: DocumentSnapshotData["sheets"] = [];
      for (const sheet of sheets.items) {
        const usedRange = sheet.getUsedRangeOrNullObject(true);
        usedRange.load("address,values,numberFormat,formulas");
        await context.sync();
        if (usedRange.isNullObject) continue;
        sheetData.push({
          name: sheet.name,
          usedRangeAddress: usedRange.address,
          values: usedRange.values as unknown[][],
          numberFormats: usedRange.numberFormat as string[][],
          formulas: usedRange.formulas as string[][],
        });
      }
      return { sheets: sheetData };
    });
  }

  if (host === "powerpoint") {
    return new Promise<DocumentSnapshotData>((resolve, reject) => {
      Office.context.document.getFileAsync(
        Office.FileType.Compressed,
        { sliceSize: 4194304 },
        (result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            reject(new Error("Failed to get PowerPoint file"));
            return;
          }
          const file = result.value;
          const sliceCount = file.sliceCount;
          const chunks: string[] = [];
          let received = 0;
          for (let i = 0; i < sliceCount; i++) {
            file.getSliceAsync(i, (sliceResult) => {
              if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                const raw = sliceResult.value.data;
                const bytes = typeof raw === "string" ? raw : btoa(String.fromCharCode(...new Uint8Array(raw)));
                chunks[i] = bytes;
              }
              received++;
              if (received === sliceCount) {
                file.closeAsync();
                resolve({ presentationBase64: chunks.join("") });
              }
            });
          }
        },
      );
    });
  }

  return {};
}

export async function restoreDocumentSnapshot(host: OfficeHost, data: DocumentSnapshotData): Promise<void> {
  if (host === "word" && data.ooxml) {
    return Word.run(async (context) => {
      context.document.body.insertOoxml(data.ooxml!, "Replace");
      await context.sync();
    });
  }

  if (host === "excel" && data.sheets?.length) {
    return Excel.run(async (context) => {
      for (const sheetData of data.sheets!) {
        let sheet: Excel.Worksheet;
        try {
          sheet = context.workbook.worksheets.getItem(sheetData.name);
        } catch {
          continue;
        }
        const addr = sheetData.usedRangeAddress.includes("!")
          ? sheetData.usedRangeAddress.split("!")[1]
          : sheetData.usedRangeAddress;
        if (!addr) continue;
        const range = sheet.getRange(addr);
        range.values = sheetData.values as (string | number | boolean)[][];
        range.numberFormat = sheetData.numberFormats as string[][];
        await context.sync();
      }
    });
  }

  if (host === "powerpoint" && data.presentationBase64) {
    return PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();
      for (const slide of slides.items) {
        slide.delete();
      }
      await context.sync();
      context.presentation.insertSlidesFromBase64(data.presentationBase64!, {
        formatting: "UseDestinationTheme" as PowerPoint.InsertSlideFormatting,
      });
      await context.sync();
    });
  }
}

// ---------------------------------------------------------------------------
// read_doc_section — paginated paragraph reading
// ---------------------------------------------------------------------------

export async function readDocumentSection(
  host: OfficeHost,
  startIndex: number,
  endIndex: number,
  includeStyles: boolean,
): Promise<unknown> {
  if (host !== "word") {
    return { error: "office_read_section is only supported for Word documents." };
  }

  return Word.run(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    const supportsParagraphIds = supportsRequirementSet("WordApi", "1.6");
    paragraphs.load(
      includeStyles
        ? supportsParagraphIds
          ? "items/text,items/uniqueLocalId,items/style,items/styleBuiltIn"
          : "items/text,items/style,items/styleBuiltIn"
        : supportsParagraphIds
          ? "items/text,items/uniqueLocalId"
          : "items/text",
    );
    await context.sync();

    const total = paragraphs.items.length;
    const safeStart = Math.max(0, Math.min(startIndex, total));
    const safeEnd = Math.max(safeStart, Math.min(endIndex, total));
    const slice = paragraphs.items.slice(safeStart, safeEnd);

    const result = slice.map((p, i) => {
      const entry: Record<string, unknown> = {
        index: safeStart + i,
        text: p.text,
      };
      if (supportsParagraphIds) entry.paragraphId = p.uniqueLocalId;
      if (includeStyles) {
        entry.style = p.styleBuiltIn || p.style;
        entry.isHeading = /heading/i.test(String(p.styleBuiltIn || p.style || ""));
      }
      return entry;
    });

    return {
      paragraphs: result,
      startIndex: safeStart,
      endIndex: safeEnd,
      totalParagraphs: total,
      hasMore: safeEnd < total,
    };
  });
}

// ---------------------------------------------------------------------------
// execute_office_js — arbitrary Office.js code execution
// ---------------------------------------------------------------------------

const BLOCKED_JS_PATTERNS = [
  /\bfetch\s*\(/i,
  /\bXMLHttpRequest\b/i,
  /\blocalStorage\b/i,
  /\bsessionStorage\b/i,
  /\bindexedDB\b/i,
  /\bWebSocket\b/i,
  /\beval\s*\(/i,
  /\bFunction\s*\(/i,
  /\bimport\s*\(/i,
  /\bnavigator\b/i,
  /\bdocument\.cookie\b/i,
];

export async function executeOfficeJs(_host: OfficeHost, code: string): Promise<unknown> {
  for (const pattern of BLOCKED_JS_PATTERNS) {
    if (pattern.test(code)) {
      return {
        error: `Code blocked: contains disallowed pattern "${pattern.source}". office_execute_js must not access network, storage, or eval.`,
      };
    }
  }

  try {
    const asyncFn = new Function("Word", "Excel", "PowerPoint", "Office", `
      "use strict";
      return (async () => {
        ${code}
      })();
    `);
    const result = await asyncFn(
      typeof Word !== "undefined" ? Word : undefined,
      typeof Excel !== "undefined" ? Excel : undefined,
      typeof PowerPoint !== "undefined" ? PowerPoint : undefined,
      typeof Office !== "undefined" ? Office : undefined,
    );
    return { ok: true, result: result ?? null };
  } catch (error) {
    return {
      ok: false,
      error: error instanceof Error ? error.message : String(error),
      stack: error instanceof Error ? error.stack : undefined,
    };
  }
}

// ---------------------------------------------------------------------------
// propose_edits — batch edit proposals for user review
// ---------------------------------------------------------------------------

export async function applyAcceptedEdits(
  edits: Array<{
    searchText: string;
    newText: string;
    kind: string;
  }>,
): Promise<{ applied: number; failed: number; errors: string[] }> {
  if (!edits.length) return { applied: 0, failed: 0, errors: [] };

  return Word.run(async (context) => {
    let applied = 0;
    let failed = 0;
    const errors: string[] = [];

    for (const edit of edits) {
      try {
        const searchText = edit.searchText;
        if (!searchText) {
          errors.push(`Edit missing searchText, skipped.`);
          failed++;
          continue;
        }

        if (searchText.length <= 255) {
          const results = context.document.body.search(searchText, { matchCase: true, matchWholeWord: false });
          results.load("items");
          await context.sync();

          if (!results.items.length) {
            errors.push(`Text not found: "${searchText.slice(0, 60)}..."`);
            failed++;
            continue;
          }

          const target = results.items[0]!;
          if (edit.kind === "delete") {
            target.delete();
          } else {
            target.insertText(edit.newText, "Replace");
          }
          await context.sync();
          applied++;
        } else {
          // Fallback for long searchText: search a short prefix, then verify full match via paragraph text
          const prefix = searchText.slice(0, 200);
          const results = context.document.body.search(prefix, { matchCase: true, matchWholeWord: false });
          results.load("items/text,items/paragraphs/items/text");
          await context.sync();

          let matched = false;
          for (const candidate of results.items) {
            const para = candidate.paragraphs.getFirst();
            para.load("text");
            await context.sync();
            if (para.text.includes(searchText)) {
              const fullRange = para.search(searchText.slice(0, 255), { matchCase: true });
              fullRange.load("items");
              await context.sync();
              if (fullRange.items.length) {
                if (edit.kind === "delete") {
                  fullRange.items[0]!.delete();
                } else {
                  fullRange.items[0]!.insertText(edit.newText, "Replace");
                }
                await context.sync();
                applied++;
                matched = true;
                break;
              }
            }
          }
          if (!matched) {
            errors.push(`Long text not found: "${searchText.slice(0, 60)}..."`);
            failed++;
          }
        }
      } catch (err) {
        failed++;
        errors.push(`Failed to apply edit: ${err instanceof Error ? err.message : String(err)}`);
      }
    }

    return { applied, failed, errors };
  });
}

export async function proposeDocumentEdits(
  host: OfficeHost,
  params: Record<string, unknown>,
): Promise<unknown> {
  if (host !== "word") {
    return { error: "office_propose_edits is only supported for Word documents." };
  }

  const edits = Array.isArray(params.edits) ? params.edits : [];
  const summary = typeof params.summary === "string" ? params.summary : "Proposed edits";

  if (!edits.length) {
    return { error: "No edits provided." };
  }

  return Word.run(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load("items/text");
    await context.sync();

    const docText = paragraphs.items.map((p) => p.text).join("\n");
    const verified: Array<Record<string, unknown>> = [];

    for (const edit of edits) {
      const record = edit as Record<string, unknown>;
      const kind = String(record.kind ?? "replace");
      const searchText = typeof record.searchText === "string" ? record.searchText : undefined;
      const oldText = typeof record.oldText === "string" ? record.oldText : searchText;
      const newText = typeof record.newText === "string" ? record.newText : "";
      const explanation = typeof record.explanation === "string" ? record.explanation : undefined;
      const id = typeof record.id === "string" ? record.id : crypto.randomUUID();

      let found = false;
      let contextPreview: string | undefined;
      if (oldText && docText.includes(oldText)) {
        found = true;
        const idx = docText.indexOf(oldText);
        const start = Math.max(0, idx - 40);
        const end = Math.min(docText.length, idx + oldText.length + 40);
        contextPreview = docText.slice(start, end);
      }

      verified.push({
        id,
        kind,
        searchText: searchText ?? oldText,
        oldText,
        newText: kind === "delete" ? "" : newText,
        explanation,
        found,
        contextPreview,
      });
    }

    return {
      proposalId: crypto.randomUUID(),
      summary,
      editCount: verified.length,
      edits: verified,
      instruction: "Present these proposed edits to the user for review. Each edit shows the text to find, the proposed change, and whether the target text was found in the document.",
    };
  });
}
