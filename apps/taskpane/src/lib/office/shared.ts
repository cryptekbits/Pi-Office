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
import type { OfficeCaptureOptions } from "../office-host-adapter-types";

export interface OfficeThemeSnapshot {
  bodyBackgroundColor: string;
  bodyForegroundColor: string;
  controlBackgroundColor: string;
  controlForegroundColor: string;
  isDarkTheme: boolean;
}

export interface ExcelCitationRecord {
  anchor: OfficeAnchor;
  label: string;
  sheetName: string;
  address: string;
  textPreview?: string | undefined;
  formula?: string | undefined;
  numberFormat?: string | undefined;
}

export interface ExcelWorksheetSnapshot {
  worksheetId: string;
  worksheetName: string;
  usedRange?: { address: string; rowCount: number; columnCount: number } | undefined;
  previewAddress: string;
  previewRows: string[];
  citedCells: ExcelCitationRecord[];
}

export interface ExcelBorderUpdateOptions {
  color?: string | undefined;
  style?: Excel.BorderLineStyle | undefined;
  weight?: Excel.BorderWeight | undefined;
  tintAndShade?: number | undefined;
}

export function mapHost(host: Office.HostType | string): OfficeHost {
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

export function basename(input: string): string {
  const parts = input.split(/[\\/]/g).filter(Boolean);
  return parts[parts.length - 1] ?? input;
}

export function toDocumentPath(documentUrl: string | undefined): string | undefined {
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

export function getUnsavedDocumentId(host: OfficeHost): string {
  const key = `pi-office-unsaved:${host}`;
  const existing = sessionStorage.getItem(key);
  if (existing) {
    return existing;
  }

  const created = `${host}-${crypto.randomUUID()}`;
  sessionStorage.setItem(key, created);
  return created;
}

export function getTaskpaneWindowId(): string {
  const key = "pi-office-window-id";
  const existing = sessionStorage.getItem(key);
  if (existing) {
    return existing;
  }

  const created = crypto.randomUUID();
  sessionStorage.setItem(key, created);
  return created;
}

export function supportsRequirementSet(name: string, version: string): boolean {
  try {
    return Office.context.requirements.isSetSupported(name, version);
  } catch {
    return false;
  }
}

export function normalizeTextPreview(value: string | undefined): string | undefined {
  const normalized = value?.replace(/\s+/g, " ").trim();
  return normalized ? normalized.slice(0, 420) : undefined;
}

export function truncateText(value: string | undefined, limit = 2400): string | undefined {
  const normalized = value?.trim();
  if (!normalized) {
    return undefined;
  }

  return normalized.length > limit ? `${normalized.slice(0, limit)}\n…` : normalized;
}

export function pluralize(count: number, singular: string, plural = `${singular}s`): string {
  return `${count} ${count === 1 ? singular : plural}`;
}

export function uniqueDetails(details: Array<string | undefined>): string[] | undefined {
  const next = details.filter((detail): detail is string => Boolean(detail?.trim())).map((detail) => detail.trim());
  return next.length ? Array.from(new Set(next)).slice(0, 6) : undefined;
}

export function containsImageMarkup(markup: string | undefined): boolean {
  return /<(?:img|w:drawing|pic:pic|a:blip|v:imagedata|v:shape)\b/i.test(markup ?? "");
}

export function stripDataUrlPrefix(input: string): { data: string; mimeType: string } {
  const match = input.match(/^data:([^;]+);base64,(.+)$/);
  if (!match) {
    return { data: input, mimeType: "image/png" };
  }

  return {
    mimeType: match[1] ?? "image/png",
    data: match[2] ?? "",
  };
}

export function formatPoints(value: number | undefined): string | undefined {
  if (typeof value !== "number" || Number.isNaN(value)) {
    return undefined;
  }

  return `${Math.round(value * 10) / 10} pt`;
}

export function formatPercent(value: number | undefined): number | undefined {
  if (typeof value !== "number" || Number.isNaN(value)) {
    return undefined;
  }

  return Math.round(value * 10) / 10;
}

export function pointsToPixels(value: number | undefined): number | undefined {
  if (typeof value !== "number" || Number.isNaN(value)) {
    return undefined;
  }

  return Math.round((value * 96) / 72);
}

export function serializeWordPages(pages: Array<{ index: number; width: number; height: number }>): Array<Record<string, unknown>> {
  return pages.map((page) => ({
    index: page.index,
    width: formatPoints(page.width),
    height: formatPoints(page.height),
  }));
}

export function buildSelectionSummary(input: {
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

export function wordStyleToMarkdownPrefix(style: string | undefined): string {
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

export function friendlyStyleName(style: string | undefined): string {
  const raw = (style ?? "Normal").replace(/([a-z])([A-Z])/g, "$1 $2");
  return raw.charAt(0).toUpperCase() + raw.slice(1);
}

export function buildStructuredSelectionPreview(
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

export function buildSelectionMeta(
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

export function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === "object" && !Array.isArray(value));
}

export function trimString(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

export function truncateLabel(value: string | undefined, limit = 96): string | undefined {
  const normalized = normalizeTextPreview(value);
  if (!normalized) {
    return undefined;
  }

  return normalized.length > limit ? `${normalized.slice(0, limit)}…` : normalized;
}

export function matchesTextQuery(value: string | undefined, query: string | undefined): boolean {
  const needle = normalizeTextPreview(query)?.toLowerCase();
  if (!needle) {
    return false;
  }

  return (normalizeTextPreview(value)?.toLowerCase() ?? "").includes(needle);
}

export function parseIndexedAnchorNumber(value: string | undefined, prefix: string): number | undefined {
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

export function parsePositiveInteger(value: unknown): number | undefined {
  const numeric = toNumber(value);
  return typeof numeric === "number" && Number.isInteger(numeric) && numeric > 0 ? numeric : undefined;
}

export function parseAnchorOrdinal(value: string | undefined, prefix: string): number | undefined {
  return parsePositiveInteger(value) ?? parseIndexedAnchorNumber(value, prefix);
}

export function clampPercentage(value: unknown): number | undefined {
  const numeric = toNumber(value);
  if (typeof numeric !== "number") {
    return undefined;
  }

  return Math.min(1, Math.max(0, numeric));
}

export function anchorKey(anchor: OfficeAnchor): string {
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

export function uniqueAnchors(anchors: OfficeAnchor[]): OfficeAnchor[] {
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

export function getActionOptions(action: OfficeHostAction): Record<string, unknown> {
  return isRecord(action.options) ? action.options : {};
}

export function toNumber(value: unknown): number | undefined {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  if (typeof value === "string" && value.trim()) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : undefined;
  }

  return undefined;
}

export function toBoolean(value: unknown): boolean | undefined {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "string") {
    if (value === "true") return true;
    if (value === "false") return false;
  }

  return undefined;
}

export function toStringMatrix(value: unknown): string[][] | undefined {
  if (!Array.isArray(value)) {
    return undefined;
  }

  return value.map((row) =>
    Array.isArray(row)
      ? row.map((cell) => (cell == null ? "" : typeof cell === "string" ? cell : String(cell)))
      : [row == null ? "" : typeof row === "string" ? row : String(row)],
  );
}

export function toCellValue(value: unknown): OfficeCellValue {
  if (value == null || typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return value as OfficeCellValue;
  }

  return JSON.stringify(value);
}

export function toValueMatrix(value: unknown, fallbackContent?: string): OfficeCellValue[][] {
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

export function splitSheetAddress(
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

export function withSheetName(sheetName: string | undefined, address: string | undefined): string | undefined {
  const normalizedAddress = trimString(address);
  if (!normalizedAddress) {
    return undefined;
  }
  if (!sheetName || normalizedAddress.includes("!")) {
    return normalizedAddress;
  }

  return `${sheetName}!${normalizedAddress}`;
}

export function isSingleCellAddress(address: string | undefined): boolean {
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

export function excelColumnName(index: number): string {
  let next = index + 1;
  let result = "";
  while (next > 0) {
    const remainder = (next - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    next = Math.floor((next - 1) / 26);
  }
  return result;
}

export function excelCellAddress(rowIndex: number, columnIndex: number): string {
  return `${excelColumnName(columnIndex)}${rowIndex + 1}`;
}

export function firstMatrixString(value: unknown): string | undefined {
  if (Array.isArray(value)) {
    const first = value[0];
    if (Array.isArray(first)) {
      return trimString(first[0]);
    }
    return trimString(first);
  }

  return trimString(value);
}

export function buildExcelCitationRecord(input: {
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

export function wordInsertLocationFromPlacement(placement: string | undefined): Word.InsertLocation {
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

export function compactStrings(values: Array<string | undefined>): string[] | undefined {
  const filtered = values.map((value) => trimString(value)).filter((value): value is string => Boolean(value));
  return filtered.length ? filtered : undefined;
}

export function toActionCompletion(value: unknown): OfficeActionCompletion | undefined {
  return value === "native" || value === "fallback" || value === "partial" ? value : undefined;
}

export function toObjectReference(host: OfficeHost, anchor: OfficeAnchor | undefined): OfficeObjectReference | undefined {
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

export function mergeObjectReferences(references: Array<OfficeObjectReference | undefined>): OfficeObjectReference[] | undefined {
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

export function extractObjectReferencesFromRecord(host: OfficeHost, record: Record<string, unknown>): OfficeObjectReference[] {
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

export function isCreationOperation(operation: string): boolean {
  return /^(add|create|insert(?!Text|Html)|duplicate|merge)/i.test(operation);
}

export function buildOfficeActionResult(
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

export function isPowerPointImageShape(type: unknown): boolean {
  return String(type ?? "").toLowerCase() === "image";
}

export function isPowerPointTableShape(type: unknown): boolean {
  return String(type ?? "").toLowerCase() === "table";
}

export function isPowerPointGroupShape(type: unknown): boolean {
  return String(type ?? "").toLowerCase() === "group";
}

export function getPowerPointShapeContentKind(type: unknown): string {
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

export function truncateStringMatrix(value: string[][] | undefined, rowLimit = 4, columnLimit = 6, cellLimit = 80): string[][] | undefined {
  if (!value?.length) {
    return undefined;
  }

  return value.slice(0, rowLimit).map((row) => row.slice(0, columnLimit).map((cell) => truncateText(cell, cellLimit) ?? ""));
}

export function resolveZeroBasedIndex(indexValue: unknown, ordinalValue: unknown): number | undefined {
  const direct = toNumber(indexValue);
  if (typeof direct === "number" && Number.isInteger(direct) && direct >= 0) {
    return direct;
  }

  const ordinal = parsePositiveInteger(ordinalValue);
  return typeof ordinal === "number" ? ordinal - 1 : undefined;
}

export function resolvePositiveCount(value: unknown, fallback = 1): number {
  return parsePositiveInteger(value) ?? fallback;
}

export function isWordPictureShape(type: unknown): boolean {
  return String(type ?? "").toLowerCase() === "picture";
}

export function countSelectedWordImages(inlinePictureCount: number, floatingImageShapeCount: number): number {
  // Word can surface a selected inline picture through both inlinePictures and shapes.
  return inlinePictureCount > 0 ? inlinePictureCount : floatingImageShapeCount;
}

export async function getSelectedTextAsync(): Promise<string> {
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

export function getSelectedMarkupAsync(coercionType: Office.CoercionType.Html | Office.CoercionType.Ooxml): Promise<string> {
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

export function setSelectedTextAsync(value: string): Promise<void> {
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

export function setSelectedImageAsync(value: string, options: Partial<Office.SetSelectedDataOptions> = {}): Promise<void> {
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

export async function buildBaseState(host: OfficeHost): Promise<OfficeStateUpdate> {
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


export function parseSelectedImageValue(value: unknown): { data: string; mimeType: string } | undefined {
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

export function getActionImagePayload(
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

export function getSelectedImageAsync(): Promise<OfficeVisualSnapshot | undefined> {
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

export function loadImageElement(sourceUrl: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("Failed to decode Office image payload."));
    image.src = sourceUrl;
  });
}

export async function optimizeVisual(
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

export async function waitForSnapshotLayout(): Promise<void> {
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

export function getWordSnapshotWidthPx(pageWidth: number | undefined, leftMargin: number | undefined, rightMargin: number | undefined): number {
  const contentWidth = typeof pageWidth === "number" ? pageWidth - (leftMargin ?? 0) - (rightMargin ?? 0) : undefined;
  return Math.max(520, Math.min(960, pointsToPixels(contentWidth) ?? 760));
}

export async function renderHtmlSelectionSnapshot(params: {
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

export function createSummary(payload: OfficeContextPayload): string {
  const lines = [
    `${HOST_LABELS[payload.state.host]} context for "${payload.state.document.title}"`,
    `Selection: ${payload.state.selection.label}`,
    payload.state.selection.textPreview ? `Preview: ${payload.state.selection.textPreview}` : undefined,
    payload.visuals?.length ? `${payload.visuals.length} visual snapshot(s) attached.` : undefined,
  ];
  return lines.filter(Boolean).join("\n");
}

export function serializeOfficeRuntimeError(error: unknown): Record<string, unknown> {
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
    windowId: getTaskpaneWindowId(),
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


export function getStringArray(value: unknown): string[] {
  return Array.isArray(value) ? value.filter((entry): entry is string => typeof entry === "string" && entry.trim().length > 0) : [];
}

export function getNumberArray(value: unknown): number[] {
  return Array.isArray(value) ? value.map((entry) => toNumber(entry)).filter((entry): entry is number => typeof entry === "number") : [];
}

export function getRecordArray(value: unknown): Record<string, unknown>[] {
  return Array.isArray(value) ? value.filter((entry): entry is Record<string, unknown> => isRecord(entry)) : [];
}

export function tryParseJsonValue(value: unknown): unknown {
  if (typeof value !== "string" || !value.trim()) {
    return value;
  }

  try {
    return JSON.parse(value);
  } catch {
    return value;
  }
}

export function getChartValueArray(value: unknown): Array<string | number> {
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

export function getChartSeriesInput(value: unknown): Array<{ name?: string | undefined; categories?: Array<string | number> | undefined; values: number[] }> {
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

