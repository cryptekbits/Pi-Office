import JSZip from "jszip";

const DRAWING_ML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

const CHART_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const PACKAGE_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
const WORKBOOK_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const WORKSHEET_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

const CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const WORKBOOK_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const WORKSHEET_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

const EMUS_PER_POINT = 12700;
const DEFAULT_CHART_LEFT = 72;
const DEFAULT_CHART_TOP = 90;
const DEFAULT_CHART_WIDTH = 432;
const DEFAULT_CHART_HEIGHT = 240;

const THEME_COLOR_LABELS: Record<string, string> = {
  dk1: "Dark 1",
  lt1: "Light 1",
  dk2: "Dark 2",
  lt2: "Light 2",
  accent1: "Accent 1",
  accent2: "Accent 2",
  accent3: "Accent 3",
  accent4: "Accent 4",
  accent5: "Accent 5",
  accent6: "Accent 6",
  hlink: "Hyperlink",
  folHlink: "Followed Hyperlink",
};

export interface PowerPointThemeColor {
  key: string;
  label: string;
  value?: string | undefined;
  source?: string | undefined;
}

export interface PowerPointThemeMetadata {
  partName: string;
  name?: string | undefined;
  colorSchemeName?: string | undefined;
  fontSchemeName?: string | undefined;
  colors: PowerPointThemeColor[];
  fonts: {
    majorLatin?: string | undefined;
    minorLatin?: string | undefined;
    majorEastAsian?: string | undefined;
    minorEastAsian?: string | undefined;
    majorComplexScript?: string | undefined;
    minorComplexScript?: string | undefined;
  };
}

export interface PowerPointSlideNotesSummary {
  slideNumber: number;
  partName: string;
  hasNotes: boolean;
  text: string;
  preview?: string | undefined;
}

export type PowerPointChartValue = string | number;

export interface PowerPointChartSeriesSummary {
  index: number;
  name?: string | undefined;
  categories: PowerPointChartValue[];
  values: number[];
}

export interface PowerPointChartSummary {
  slideNumber: number;
  slidePartName: string;
  chartIndex: number;
  chartPartName: string;
  chartType?: string | undefined;
  title?: string | undefined;
  shapeId?: string | undefined;
  shapeName?: string | undefined;
  embeddedWorkbookPartName?: string | undefined;
  hasEmbeddedWorkbook: boolean;
  seriesCount: number;
  categoryCount: number;
  series: PowerPointChartSeriesSummary[];
}

export interface PowerPointPresentationPackageSummary {
  theme?: PowerPointThemeMetadata | undefined;
  notes: PowerPointSlideNotesSummary[];
  charts: PowerPointChartSummary[];
  partCounts: {
    slides: number;
    slideMasters: number;
    layouts: number;
    themes: number;
    notesSlides: number;
    charts: number;
    media: number;
    customXmlParts: number;
  };
}

export interface PowerPointSlideNotesMutationResult {
  base64: string;
  slideNumber: number;
  partName: string;
  changed: boolean;
  previousText: string;
  nextText: string;
}

export interface PowerPointChartSeriesInput {
  name?: string | undefined;
  categories?: PowerPointChartValue[] | undefined;
  values: number[];
}

export interface PowerPointChartMutationInput {
  chartIndex?: number | undefined;
  shapeName?: string | undefined;
  chartType?: string | undefined;
  title?: string | undefined;
  categories?: PowerPointChartValue[] | undefined;
  series?: PowerPointChartSeriesInput[] | undefined;
}

export interface PowerPointChartMutationResult {
  base64: string;
  slideNumber: number;
  chartIndex: number;
  chartPartName: string;
  chartType?: string | undefined;
  title?: string | undefined;
  shapeName?: string | undefined;
  embeddedWorkbookPartName?: string | undefined;
  seriesCount: number;
  categoryCount: number;
  hadEmbeddedWorkbook: boolean;
  warnings?: string[] | undefined;
}

export interface PowerPointChartCreationInput extends PowerPointChartMutationInput {
  left?: number | undefined;
  top?: number | undefined;
  width?: number | undefined;
  height?: number | undefined;
  showLegend?: boolean | undefined;
}

export interface PowerPointChartCreationResult {
  base64: string;
  slideNumber: number;
  chartIndex: number;
  chartPartName: string;
  chartType: string;
  title?: string | undefined;
  shapeId: string;
  shapeName: string;
  embeddedWorkbookPartName: string;
  seriesCount: number;
  categoryCount: number;
}

interface NormalizedPowerPointChartSeries {
  name: string;
  categories: PowerPointChartValue[];
  values: number[];
}

interface NormalizedPowerPointChartData {
  title?: string | undefined;
  chartType: ResolvedPowerPointChartType;
  categories: PowerPointChartValue[];
  series: NormalizedPowerPointChartSeries[];
  showLegend: boolean;
}

interface ResolvedPowerPointChartType {
  key: "column" | "bar" | "line" | "area" | "pie" | "doughnut";
  plotElementName: "barChart" | "lineChart" | "areaChart" | "pieChart" | "doughnutChart";
  barDirection?: "col" | "bar" | undefined;
  grouping?: "clustered" | "stacked" | "percentStacked" | "standard" | undefined;
  requiresAxes: boolean;
  supportsMultipleSeries: boolean;
}

interface PowerPointChartWorkbookLayout {
  sheetName: string;
  dimensionRef: string;
  categoriesFormula: string;
  series: Array<{
    nameFormula: string;
    valuesFormula: string;
  }>;
}

function normalizeText(value: string | null | undefined): string {
  return (value ?? "").replace(/\r\n?/g, "\n");
}

function trimText(value: string | null | undefined): string {
  return normalizeText(value).trim();
}

function truncatePreview(value: string, limit = 280): string | undefined {
  const normalized = trimText(value).replace(/\s+/g, " ");
  if (!normalized) {
    return undefined;
  }
  return normalized.length > limit ? `${normalized.slice(0, limit)}…` : normalized;
}

function parseXml(xml: string, partName: string): XMLDocument {
  const document = new DOMParser().parseFromString(xml, "application/xml");
  const parserError = document.getElementsByTagName("parsererror")[0];
  if (parserError) {
    throw new Error(`Failed to parse ${partName} as XML.`);
  }
  return document;
}

function listElementsByLocalName(parent: Document | Element, localName: string): Element[] {
  return Array.from(parent.getElementsByTagNameNS("*", localName));
}

function getFirstElementByLocalName(parent: Document | Element, localName: string): Element | undefined {
  return listElementsByLocalName(parent, localName)[0];
}

function getDirectChildrenByLocalName(parent: Element | undefined, localName: string): Element[] {
  return parent ? Array.from(parent.children).filter((child) => child.localName === localName) : [];
}

function getDirectChildByLocalName(parent: Element | undefined, localName: string): Element | undefined {
  return getDirectChildrenByLocalName(parent, localName)[0];
}

function getAttributeNsOrName(element: Element, namespaceUri: string, localName: string, fallbackName: string): string | undefined {
  return trimText(element.getAttributeNS(namespaceUri, localName)) || trimText(element.getAttribute(fallbackName)) || undefined;
}

function getPartDirectory(partName: string): string {
  const normalized = partName.replace(/\\/g, "/");
  const index = normalized.lastIndexOf("/");
  return index >= 0 ? normalized.slice(0, index) : "";
}

function getPartBasename(partName: string): string {
  const normalized = partName.replace(/\\/g, "/");
  const index = normalized.lastIndexOf("/");
  return index >= 0 ? normalized.slice(index + 1) : normalized;
}

function resolvePackagePartName(sourcePartName: string, target: string): string {
  if (!target) {
    return sourcePartName;
  }

  const normalizedTarget = target.replace(/\\/g, "/");
  if (/^[a-z]+:/i.test(normalizedTarget)) {
    throw new Error(`External package targets are not supported in PPTX transforms: ${normalizedTarget}`);
  }

  const sourceDirectory = getPartDirectory(sourcePartName);
  const segments = normalizedTarget.startsWith("/")
    ? normalizedTarget.replace(/^\/+/, "").split("/")
    : [...(sourceDirectory ? sourceDirectory.split("/") : []), ...normalizedTarget.split("/")];
  const resolved: string[] = [];
  for (const segment of segments) {
    if (!segment || segment === ".") {
      continue;
    }
    if (segment === "..") {
      resolved.pop();
      continue;
    }
    resolved.push(segment);
  }

  return resolved.join("/");
}

function getRelationshipsPartName(sourcePartName: string): string {
  const directory = getPartDirectory(sourcePartName);
  const basename = getPartBasename(sourcePartName);
  return `${directory ? `${directory}/` : ""}_rels/${basename}.rels`;
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function toSpreadsheetColumnName(index: number): string {
  let current = Math.max(1, Math.floor(index));
  let result = "";
  while (current > 0) {
    const remainder = (current - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    current = Math.floor((current - 1) / 26);
  }
  return result || "A";
}

function toSpreadsheetCellReference(columnIndex: number, rowIndex: number): string {
  return `${toSpreadsheetColumnName(columnIndex)}${Math.max(1, Math.floor(rowIndex))}`;
}

function buildSpreadsheetRangeFormula(sheetName: string, columnIndex: number, startRow: number, endRow: number): string {
  const start = `$${toSpreadsheetCellReference(columnIndex, startRow).replace(/(\D+)(\d+)/, "$1$$2")}`;
  const end = `$${toSpreadsheetCellReference(columnIndex, endRow).replace(/(\D+)(\d+)/, "$1$$2")}`;
  return `${sheetName}!${start}:${end}`;
}

function buildSpreadsheetCellFormula(sheetName: string, columnIndex: number, rowIndex: number): string {
  const cell = toSpreadsheetCellReference(columnIndex, rowIndex).replace(/(\D+)(\d+)/, "$1$$2");
  return `${sheetName}!$${cell}`;
}

function toEmu(value: number | undefined, fallback: number): number {
  return Math.round((typeof value === "number" ? value : fallback) * EMUS_PER_POINT);
}

function getRelativePackageTarget(sourcePartName: string, targetPartName: string): string {
  const fromDirectory = getPartDirectory(sourcePartName);
  const fromSegments = fromDirectory ? fromDirectory.split("/") : [];
  const toSegments = targetPartName.replace(/\\/g, "/").split("/");
  let commonIndex = 0;
  while (commonIndex < fromSegments.length && commonIndex < toSegments.length && fromSegments[commonIndex] === toSegments[commonIndex]) {
    commonIndex += 1;
  }

  const upSegments = new Array(fromSegments.length - commonIndex).fill("..");
  const downSegments = toSegments.slice(commonIndex);
  return [...upSegments, ...downSegments].join("/");
}

function createRelationshipsDocument(): XMLDocument {
  return parseXml(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_REL_NS}"></Relationships>`,
    "relationships",
  );
}

async function loadOrCreateRelationshipsDocument(
  zip: JSZip,
  sourcePartName: string,
): Promise<{ partName: string; document: XMLDocument; root: Element }> {
  const partName = getRelationshipsPartName(sourcePartName);
  const xml = await readZipText(zip, partName);
  const document = xml ? parseXml(xml, partName) : createRelationshipsDocument();
  return { partName, document, root: document.documentElement };
}

function getNextRelationshipId(document: XMLDocument): string {
  const maxId = listElementsByLocalName(document, "Relationship").reduce((current, relationship) => {
    const match = trimText(relationship.getAttribute("Id")).match(/^rId(\d+)$/i);
    const value = match ? Number(match[1]) : 0;
    return Math.max(current, Number.isFinite(value) ? value : 0);
  }, 0);
  return `rId${maxId + 1}`;
}

function ensureRelationship(
  document: XMLDocument,
  sourcePartName: string,
  relationshipType: string,
  targetPartName: string,
): string {
  const root = document.documentElement;
  const relativeTarget = getRelativePackageTarget(sourcePartName, targetPartName);
  const existing = listElementsByLocalName(document, "Relationship").find((relationship) => {
    const type = trimText(relationship.getAttribute("Type"));
    const target = trimText(relationship.getAttribute("Target"));
    return type === relationshipType && resolvePackagePartName(sourcePartName, target) === targetPartName;
  });
  if (existing) {
    return trimText(existing.getAttribute("Id")) || getNextRelationshipId(document);
  }

  const relationship = document.createElementNS(PACKAGE_REL_NS, "Relationship");
  const relationshipId = getNextRelationshipId(document);
  relationship.setAttribute("Id", relationshipId);
  relationship.setAttribute("Type", relationshipType);
  relationship.setAttribute("Target", relativeTarget);
  root.appendChild(relationship);
  return relationshipId;
}

async function loadContentTypesDocument(zip: JSZip): Promise<XMLDocument> {
  const xml = await readZipText(zip, "[Content_Types].xml");
  if (!xml) {
    throw new Error("The exported PowerPoint package is missing [Content_Types].xml.");
  }
  return parseXml(xml, "[Content_Types].xml");
}

function ensureContentTypeDefault(document: XMLDocument, extension: string, contentType: string): void {
  const root = document.documentElement;
  const existing = getDirectChildrenByLocalName(root, "Default").find(
    (entry) => trimText(entry.getAttribute("Extension")).toLowerCase() === extension.toLowerCase(),
  );
  if (existing) {
    existing.setAttribute("ContentType", contentType);
    return;
  }

  const entry = document.createElementNS(CONTENT_TYPES_NS, "Default");
  entry.setAttribute("Extension", extension);
  entry.setAttribute("ContentType", contentType);
  root.appendChild(entry);
}

function ensureContentTypeOverride(document: XMLDocument, partName: string, contentType: string): void {
  const root = document.documentElement;
  const normalizedPartName = `/${partName.replace(/^\/+/, "")}`;
  const existing = getDirectChildrenByLocalName(root, "Override").find(
    (entry) => trimText(entry.getAttribute("PartName")) === normalizedPartName,
  );
  if (existing) {
    existing.setAttribute("ContentType", contentType);
    return;
  }

  const entry = document.createElementNS(CONTENT_TYPES_NS, "Override");
  entry.setAttribute("PartName", normalizedPartName);
  entry.setAttribute("ContentType", contentType);
  root.appendChild(entry);
}

function getNextPartNumber(zip: JSZip, pattern: RegExp): number {
  return (
    Object.keys(zip.files).reduce((current, name) => {
      const match = name.match(pattern);
      const value = match ? Number(match[1]) : 0;
      return Math.max(current, Number.isFinite(value) ? value : 0);
    }, 0) + 1
  );
}

function readThemeColorValue(container: Element): { value?: string | undefined; source?: string | undefined } {
  const srgb = getFirstElementByLocalName(container, "srgbClr");
  if (srgb) {
    const value = trimText(srgb.getAttribute("val"));
    return value ? { value: `#${value.toUpperCase()}`, source: "srgbClr" } : {};
  }

  const system = getFirstElementByLocalName(container, "sysClr");
  if (system) {
    const fallback = trimText(system.getAttribute("lastClr"));
    return fallback ? { value: `#${fallback.toUpperCase()}`, source: "sysClr" } : { source: "sysClr" };
  }

  return {};
}

function readFontTypeface(container: Element | undefined): string | undefined {
  const typeface = trimText(container?.getAttribute("typeface"));
  return typeface || undefined;
}

async function loadZipFromBase64(base64: string): Promise<JSZip> {
  return JSZip.loadAsync(base64, { base64: true });
}

async function readZipText(zip: JSZip, partName: string): Promise<string | undefined> {
  return zip.file(partName)?.async("string");
}

async function inspectPowerPointTheme(zip: JSZip): Promise<PowerPointThemeMetadata | undefined> {
  const themePartName = Object.keys(zip.files)
    .filter((name) => /^ppt\/theme\/theme\d+\.xml$/i.test(name))
    .sort((left, right) => left.localeCompare(right))[0];

  if (!themePartName) {
    return undefined;
  }

  const xml = await readZipText(zip, themePartName);
  if (!xml) {
    return undefined;
  }

  const document = parseXml(xml, themePartName);
  const theme = document.documentElement;
  const colorScheme = getFirstElementByLocalName(theme, "clrScheme");
  const fontScheme = getFirstElementByLocalName(theme, "fontScheme");
  const majorFont = fontScheme ? getFirstElementByLocalName(fontScheme, "majorFont") : undefined;
  const minorFont = fontScheme ? getFirstElementByLocalName(fontScheme, "minorFont") : undefined;

  return {
    partName: themePartName,
    name: trimText(theme.getAttribute("name")) || undefined,
    colorSchemeName: trimText(colorScheme?.getAttribute("name")) || undefined,
    fontSchemeName: trimText(fontScheme?.getAttribute("name")) || undefined,
    colors: colorScheme
      ? Array.from(colorScheme.children).map((entry) => ({
          key: entry.localName,
          label: THEME_COLOR_LABELS[entry.localName] ?? entry.localName,
          ...readThemeColorValue(entry),
        }))
      : [],
    fonts: {
      majorLatin: readFontTypeface(majorFont ? getFirstElementByLocalName(majorFont, "latin") : undefined),
      minorLatin: readFontTypeface(minorFont ? getFirstElementByLocalName(minorFont, "latin") : undefined),
      majorEastAsian: readFontTypeface(majorFont ? getFirstElementByLocalName(majorFont, "ea") : undefined),
      minorEastAsian: readFontTypeface(minorFont ? getFirstElementByLocalName(minorFont, "ea") : undefined),
      majorComplexScript: readFontTypeface(majorFont ? getFirstElementByLocalName(majorFont, "cs") : undefined),
      minorComplexScript: readFontTypeface(minorFont ? getFirstElementByLocalName(minorFont, "cs") : undefined),
    },
  };
}

async function readRelationshipMap(zip: JSZip, sourcePartName: string): Promise<Map<string, string>> {
  const relationshipPartName = getRelationshipsPartName(sourcePartName);
  const xml = await readZipText(zip, relationshipPartName);
  if (!xml) {
    return new Map();
  }

  const document = parseXml(xml, relationshipPartName);
  const relationships = listElementsByLocalName(document, "Relationship");
  const map = new Map<string, string>();
  for (const relationship of relationships) {
    const id = trimText(relationship.getAttribute("Id"));
    const target = trimText(relationship.getAttribute("Target"));
    if (!id || !target) {
      continue;
    }
    map.set(id, resolvePackagePartName(sourcePartName, target));
  }
  return map;
}

function getSortedSlidePartNames(zip: JSZip): string[] {
  return Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
    .sort((left, right) => {
      const leftNumber = Number(left.match(/slide(\d+)\.xml/i)?.[1] ?? "0");
      const rightNumber = Number(right.match(/slide(\d+)\.xml/i)?.[1] ?? "0");
      return leftNumber - rightNumber;
    });
}

function readChartPointValues(container: Element | undefined): PowerPointChartValue[] {
  if (!container) {
    return [];
  }

  const holder =
    getDirectChildByLocalName(container, "strLit") ??
    getDirectChildByLocalName(container, "numLit") ??
    getDirectChildByLocalName(getDirectChildByLocalName(container, "strRef") ?? container, "strCache") ??
    getDirectChildByLocalName(getDirectChildByLocalName(container, "numRef") ?? container, "numCache") ??
    getDirectChildByLocalName(getDirectChildByLocalName(container, "multiLvlStrRef") ?? container, "multiLvlStrCache");

  if (!holder) {
    const directValue = trimText(getDirectChildByLocalName(container, "v")?.textContent);
    return directValue ? [directValue] : [];
  }

  if (holder.localName === "multiLvlStrCache") {
    return getDirectChildrenByLocalName(holder, "lvl").flatMap((level) =>
      getDirectChildrenByLocalName(level, "pt")
        .sort((left, right) => Number(left.getAttribute("idx") ?? "0") - Number(right.getAttribute("idx") ?? "0"))
        .map((point) => trimText(getFirstElementByLocalName(point, "v")?.textContent) ?? ""),
    );
  }

  return getDirectChildrenByLocalName(holder, "pt")
    .sort((left, right) => Number(left.getAttribute("idx") ?? "0") - Number(right.getAttribute("idx") ?? "0"))
    .map((point) => trimText(getFirstElementByLocalName(point, "v")?.textContent) ?? "")
    .map((value) => {
      const numeric = Number(value);
      return Number.isFinite(numeric) && holder.localName.startsWith("num") ? numeric : value;
    });
}

function parseChartSeriesName(seriesElement: Element): string | undefined {
  const textElement = getDirectChildByLocalName(seriesElement, "tx");
  if (!textElement) {
    return undefined;
  }

  const directValue = trimText(getDirectChildByLocalName(textElement, "v")?.textContent);
  if (directValue) {
    return directValue;
  }

  const cachedValues = readChartPointValues(textElement);
  const firstValue = cachedValues[0];
  return typeof firstValue === "number" ? String(firstValue) : firstValue;
}

function parseChartSeries(seriesElement: Element): PowerPointChartSeriesSummary {
  const categories = readChartPointValues(
    getDirectChildByLocalName(seriesElement, "cat") ?? getDirectChildByLocalName(seriesElement, "xVal"),
  );
  const values = readChartPointValues(
    getDirectChildByLocalName(seriesElement, "val") ??
      getDirectChildByLocalName(seriesElement, "yVal") ??
      getDirectChildByLocalName(seriesElement, "bubbleSize"),
  )
    .map((value) => (typeof value === "number" ? value : Number(value)))
    .filter((value): value is number => Number.isFinite(value));

  return {
    index: Number(getDirectChildByLocalName(seriesElement, "idx")?.getAttribute("val") ?? "0") + 1,
    name: parseChartSeriesName(seriesElement),
    categories,
    values,
  };
}

function buildDrawingParagraph(document: XMLDocument, text: string): Element {
  const paragraph = document.createElementNS(DRAWING_ML_NS, "a:p");
  const run = document.createElementNS(DRAWING_ML_NS, "a:r");
  const runProps = document.createElementNS(DRAWING_ML_NS, "a:rPr");
  runProps.setAttribute("lang", "en-US");
  runProps.setAttribute("dirty", "0");
  run.appendChild(runProps);

  const textNode = document.createElementNS(DRAWING_ML_NS, "a:t");
  if (/^\s|\s$/.test(text)) {
    textNode.setAttribute("xml:space", "preserve");
  }
  textNode.textContent = text;
  run.appendChild(textNode);
  paragraph.appendChild(run);

  const endParagraphProps = document.createElementNS(DRAWING_ML_NS, "a:endParaRPr");
  endParagraphProps.setAttribute("lang", "en-US");
  endParagraphProps.setAttribute("dirty", "0");
  paragraph.appendChild(endParagraphProps);

  return paragraph;
}

function replaceRichTextBody(document: XMLDocument, richElement: Element, text: string): void {
  const bodyPr = getDirectChildByLocalName(richElement, "bodyPr") ?? document.createElementNS(DRAWING_ML_NS, "a:bodyPr");
  const listStyle = getDirectChildByLocalName(richElement, "lstStyle") ?? document.createElementNS(DRAWING_ML_NS, "a:lstStyle");

  while (richElement.firstChild) {
    richElement.removeChild(richElement.firstChild);
  }

  richElement.appendChild(bodyPr);
  richElement.appendChild(listStyle);
  richElement.appendChild(buildDrawingParagraph(document, text));
}

async function inspectChartsInSlidePart(
  zip: JSZip,
  slidePartName: string,
  slideNumber: number,
): Promise<PowerPointChartSummary[]> {
  const slideXml = await readZipText(zip, slidePartName);
  if (!slideXml) {
    return [];
  }

  const slideDocument = parseXml(slideXml, slidePartName);
  const slideRelationships = await readRelationshipMap(zip, slidePartName);
  const graphicFrames = listElementsByLocalName(slideDocument, "graphicFrame");
  const charts: PowerPointChartSummary[] = [];

  for (const frame of graphicFrames) {
    const chartReference = getFirstElementByLocalName(frame, "chart");
    if (!chartReference) {
      continue;
    }

    const relationshipId = getAttributeNsOrName(chartReference, OFFICE_REL_NS, "id", "r:id");
    const chartPartName = relationshipId ? slideRelationships.get(relationshipId) : undefined;
    if (!chartPartName) {
      continue;
    }

    const chartXml = await readZipText(zip, chartPartName);
    if (!chartXml) {
      continue;
    }

    const chartDocument = parseXml(chartXml, chartPartName);
    const chartRoot = chartDocument.documentElement;
    const chartElement = getFirstElementByLocalName(chartRoot, "chart");
    const plotArea = chartElement ? getDirectChildByLocalName(chartElement, "plotArea") : undefined;
    const plotTypes = plotArea
      ? Array.from(plotArea.children).filter((child) => child.localName.endsWith("Chart"))
      : [];
    const series = plotTypes.flatMap((plotType) => getDirectChildrenByLocalName(plotType, "ser").map((entry) => parseChartSeries(entry)));
    const titleElement = chartElement ? getDirectChildByLocalName(chartElement, "title") : undefined;
    const title = titleElement ? extractTextRuns(titleElement) : undefined;
    const chartRelationships = await readRelationshipMap(zip, chartPartName);
    const embeddedWorkbookPartName = Array.from(chartRelationships.values()).find((partName) => /^ppt\/embeddings\//i.test(partName));
    const nonVisualProps = getFirstElementByLocalName(frame, "cNvPr");

    charts.push({
      slideNumber,
      slidePartName,
      chartIndex: charts.length + 1,
      chartPartName,
      chartType: plotTypes.map((entry) => entry.localName).filter(Boolean).join(", ") || undefined,
      title: trimText(title) || undefined,
      shapeId: trimText(nonVisualProps?.getAttribute("id")) || undefined,
      shapeName: trimText(nonVisualProps?.getAttribute("name")) || undefined,
      embeddedWorkbookPartName,
      hasEmbeddedWorkbook: Boolean(embeddedWorkbookPartName),
      seriesCount: series.length,
      categoryCount: Math.max(...series.map((entry) => entry.categories.length), 0),
      series,
    });
  }

  return charts;
}

function extractTextRuns(root: Document | Element): string {
  return listElementsByLocalName(root, "t")
    .map((node) => node.textContent ?? "")
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function findSpeakerNotesShape(document: XMLDocument): Element | undefined {
  const shapes = listElementsByLocalName(document, "sp");
  const preferredShape = shapes.find((shape) => {
    const placeholder = getFirstElementByLocalName(shape, "ph");
    const type = trimText(placeholder?.getAttribute("type")).toLowerCase();
    return type === "body" || type === "obj";
  });
  if (preferredShape) {
    return preferredShape;
  }

  return shapes.find((shape) => Boolean(getFirstElementByLocalName(shape, "txBody")));
}

function extractNotesText(document: XMLDocument): string {
  const speakerNotesShape = findSpeakerNotesShape(document);
  const textBody = speakerNotesShape ? getFirstElementByLocalName(speakerNotesShape, "txBody") : undefined;
  return textBody ? extractTextRuns(textBody) : extractTextRuns(document);
}

function buildNotesParagraph(document: XMLDocument, text: string): Element {
  return buildDrawingParagraph(document, text);
}

function replaceNotesText(document: XMLDocument, nextText: string): string {
  const speakerNotesShape = findSpeakerNotesShape(document);
  if (!speakerNotesShape) {
    throw new Error("Could not resolve the speaker-notes text shape in the exported slide package.");
  }

  let textBody = getFirstElementByLocalName(speakerNotesShape, "txBody");
  if (!textBody) {
    textBody = document.createElementNS(speakerNotesShape.namespaceURI, "p:txBody");
    textBody.appendChild(document.createElementNS(DRAWING_ML_NS, "a:bodyPr"));
    textBody.appendChild(document.createElementNS(DRAWING_ML_NS, "a:lstStyle"));
    speakerNotesShape.appendChild(textBody);
  }

  const previousText = extractNotesText(document);
  const bodyPr = getDirectChildrenByLocalName(textBody, "bodyPr")[0] ?? document.createElementNS(DRAWING_ML_NS, "a:bodyPr");
  const listStyle = getDirectChildrenByLocalName(textBody, "lstStyle")[0] ?? document.createElementNS(DRAWING_ML_NS, "a:lstStyle");

  while (textBody.firstChild) {
    textBody.removeChild(textBody.firstChild);
  }
  textBody.appendChild(bodyPr);
  textBody.appendChild(listStyle);

  const paragraphs = normalizeText(nextText).split("\n");
  if (!paragraphs.length) {
    textBody.appendChild(buildNotesParagraph(document, ""));
  } else {
    for (const paragraphText of paragraphs) {
      textBody.appendChild(buildNotesParagraph(document, paragraphText));
    }
  }

  return previousText;
}

function getSortedNotesPartNames(zip: JSZip): string[] {
  return Object.keys(zip.files)
    .filter((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(name))
    .sort((left, right) => {
      const leftNumber = Number(left.match(/notesSlide(\d+)\.xml/i)?.[1] ?? "0");
      const rightNumber = Number(right.match(/notesSlide(\d+)\.xml/i)?.[1] ?? "0");
      return leftNumber - rightNumber;
    });
}

function setChartCachePoints(document: XMLDocument, cacheElement: Element, values: PowerPointChartValue[], numeric: boolean): void {
  for (const child of getDirectChildrenByLocalName(cacheElement, "ptCount")) {
    cacheElement.removeChild(child);
  }
  for (const child of getDirectChildrenByLocalName(cacheElement, "pt")) {
    cacheElement.removeChild(child);
  }
  for (const child of getDirectChildrenByLocalName(cacheElement, "lvl")) {
    cacheElement.removeChild(child);
  }

  const pointCount = document.createElementNS(CHART_NS, "c:ptCount");
  pointCount.setAttribute("val", String(values.length));
  cacheElement.appendChild(pointCount);

  values.forEach((value, index) => {
    const point = document.createElementNS(CHART_NS, "c:pt");
    point.setAttribute("idx", String(index));
    const valueNode = document.createElementNS(CHART_NS, "c:v");
    const serialized = numeric ? String(Number(value)) : String(value);
    valueNode.textContent = serialized;
    point.appendChild(valueNode);
    cacheElement.appendChild(point);
  });
}

function removeDirectChildrenByLocalName(parent: Element, localNames: string[]): void {
  for (const localName of localNames) {
    for (const child of getDirectChildrenByLocalName(parent, localName)) {
      parent.removeChild(child);
    }
  }
}

function resolveOrCreateChartCache(
  document: XMLDocument,
  parent: Element,
  requestedKind: "string" | "number",
): { cacheElement: Element; kind: "string" | "number"; referenceType: "ref" | "literal" | "direct" } {
  const existingStringRef = getDirectChildByLocalName(parent, "strRef");
  const existingNumRef = getDirectChildByLocalName(parent, "numRef");
  const existingStringLiteral = getDirectChildByLocalName(parent, "strLit");
  const existingNumLiteral = getDirectChildByLocalName(parent, "numLit");

  if (existingStringRef) {
    let cache = getDirectChildByLocalName(existingStringRef, "strCache");
    if (!cache) {
      cache = document.createElementNS(CHART_NS, "c:strCache");
      existingStringRef.appendChild(cache);
    }
    return { cacheElement: cache, kind: "string", referenceType: "ref" };
  }

  if (existingNumRef) {
    let cache = getDirectChildByLocalName(existingNumRef, "numCache");
    if (!cache) {
      cache = document.createElementNS(CHART_NS, "c:numCache");
      existingNumRef.appendChild(cache);
    }
    return { cacheElement: cache, kind: "number", referenceType: "ref" };
  }

  if (existingStringLiteral) {
    return { cacheElement: existingStringLiteral, kind: "string", referenceType: "literal" };
  }

  if (existingNumLiteral) {
    return { cacheElement: existingNumLiteral, kind: "number", referenceType: "literal" };
  }

  const literalName = requestedKind === "number" ? "numLit" : "strLit";
  const cacheElement = document.createElementNS(CHART_NS, `c:${literalName}`);
  parent.appendChild(cacheElement);
  return { cacheElement, kind: requestedKind, referenceType: "literal" };
}

function resolveOrCreateChartReference(
  document: XMLDocument,
  parent: Element,
  requestedKind: "string" | "number",
): { referenceElement: Element; formulaElement: Element; cacheElement: Element; kind: "string" | "number" } {
  const referenceElementName = requestedKind === "number" ? "numRef" : "strRef";
  const cacheElementName = requestedKind === "number" ? "numCache" : "strCache";
  let referenceElement = getDirectChildByLocalName(parent, referenceElementName);

  if (!referenceElement) {
    removeDirectChildrenByLocalName(parent, ["strRef", "numRef", "strLit", "numLit", "multiLvlStrRef", "v"]);
    referenceElement = document.createElementNS(CHART_NS, `c:${referenceElementName}`);
    parent.appendChild(referenceElement);
  }

  let formulaElement = getDirectChildByLocalName(referenceElement, "f");
  if (!formulaElement) {
    formulaElement = document.createElementNS(CHART_NS, "c:f");
    referenceElement.insertBefore(formulaElement, referenceElement.firstChild);
  }

  let cacheElement = getDirectChildByLocalName(referenceElement, cacheElementName);
  if (!cacheElement) {
    cacheElement = document.createElementNS(CHART_NS, `c:${cacheElementName}`);
    referenceElement.appendChild(cacheElement);
  }

  return { referenceElement, formulaElement, cacheElement, kind: requestedKind };
}

function setChartSeriesName(document: XMLDocument, seriesElement: Element, name: string, formula?: string | undefined): void {
  let textElement = getDirectChildByLocalName(seriesElement, "tx");
  if (!textElement) {
    textElement = document.createElementNS(CHART_NS, "c:tx");
    const orderReference =
      getDirectChildByLocalName(seriesElement, "idx")?.nextSibling ?? getDirectChildByLocalName(seriesElement, "idx");
    seriesElement.insertBefore(textElement, orderReference ?? seriesElement.firstChild);
  }

  if (formula) {
    const reference = resolveOrCreateChartReference(document, textElement, "string");
    reference.formulaElement.textContent = formula;
    setChartCachePoints(document, reference.cacheElement, [name], false);
    return;
  }

  const directValue = getDirectChildByLocalName(textElement, "v");
  if (directValue) {
    directValue.textContent = name;
    return;
  }

  const cache = resolveOrCreateChartCache(document, textElement, "string");
  setChartCachePoints(document, cache.cacheElement, [name], false);
}

function setChartSeriesValues(
  document: XMLDocument,
  seriesElement: Element,
  input: {
    categories: PowerPointChartValue[];
    values: number[];
    categoriesFormula?: string | undefined;
    valuesFormula?: string | undefined;
  },
): void {
  let categoryElement = getDirectChildByLocalName(seriesElement, "cat") ?? getDirectChildByLocalName(seriesElement, "xVal");
  if (!categoryElement && input.categories.length) {
    categoryElement = document.createElementNS(CHART_NS, "c:cat");
    const valueReference =
      getDirectChildByLocalName(seriesElement, "val") ??
      getDirectChildByLocalName(seriesElement, "yVal") ??
      getDirectChildByLocalName(seriesElement, "bubbleSize");
    seriesElement.insertBefore(categoryElement, valueReference ?? null);
  }
  if (categoryElement && input.categories.length) {
    const categoryKind = typeof input.categories[0] === "number" ? "number" : "string";
    if (categoryKind === "number" && input.categories.some((value) => typeof value !== "number" && !Number.isFinite(Number(value)))) {
      throw new Error("Chart categories must be numeric for a numeric cache.");
    }
    if (input.categoriesFormula) {
      const reference = resolveOrCreateChartReference(document, categoryElement, categoryKind);
      reference.formulaElement.textContent = input.categoriesFormula;
      setChartCachePoints(document, reference.cacheElement, input.categories, reference.kind === "number");
    } else {
      const cache = resolveOrCreateChartCache(document, categoryElement, categoryKind);
      setChartCachePoints(document, cache.cacheElement, input.categories, cache.kind === "number");
    }
  }

  let valueElement =
    getDirectChildByLocalName(seriesElement, "val") ??
    getDirectChildByLocalName(seriesElement, "yVal") ??
    getDirectChildByLocalName(seriesElement, "bubbleSize");
  if (!valueElement) {
    valueElement = document.createElementNS(CHART_NS, "c:val");
    seriesElement.appendChild(valueElement);
  }
  if (valueElement) {
    if (input.valuesFormula) {
      const reference = resolveOrCreateChartReference(document, valueElement, "number");
      reference.formulaElement.textContent = input.valuesFormula;
      setChartCachePoints(document, reference.cacheElement, input.values, true);
    } else {
      const cache = resolveOrCreateChartCache(document, valueElement, "number");
      setChartCachePoints(document, cache.cacheElement, input.values, true);
    }
  }
}

function setChartTitle(document: XMLDocument, chartElement: Element, title: string): void {
  let titleElement = getDirectChildByLocalName(chartElement, "title");
  if (!titleElement) {
    titleElement = document.createElementNS(CHART_NS, "c:title");
    const plotArea = getDirectChildByLocalName(chartElement, "plotArea");
    chartElement.insertBefore(titleElement, plotArea ?? chartElement.firstChild);
  }

  let textElement = getDirectChildByLocalName(titleElement, "tx");
  if (!textElement) {
    textElement = document.createElementNS(CHART_NS, "c:tx");
    titleElement.insertBefore(textElement, titleElement.firstChild);
  }

  const richElement = getDirectChildByLocalName(textElement, "rich");
  if (richElement) {
    replaceRichTextBody(document, richElement, title);
    return;
  }

  const directValue = getDirectChildByLocalName(textElement, "v");
  if (directValue) {
    directValue.textContent = title;
    return;
  }

  const cache = resolveOrCreateChartCache(document, textElement, "string");
  setChartCachePoints(document, cache.cacheElement, [title], false);
}

function getChartValuesKey(values: PowerPointChartValue[]): string {
  return values.map((value) => `${typeof value === "number" ? "n" : "s"}:${String(value)}`).join("|");
}

function normalizePowerPointChartType(value?: string | undefined, fallbackPlotElement?: Element | undefined): ResolvedPowerPointChartType {
  const normalized = trimText(value).toLowerCase().replace(/[\s_-]+/g, "");
  const existingPlotName = fallbackPlotElement?.localName ?? "";
  const existingBarDirection = trimText(getDirectChildByLocalName(fallbackPlotElement, "barDir")?.getAttribute("val")).toLowerCase();
  const existingGrouping = trimText(getDirectChildByLocalName(fallbackPlotElement, "grouping")?.getAttribute("val")).toLowerCase();

  const wantsPercentStacked = normalized.includes("percent") || normalized.includes("100");
  const wantsStacked = wantsPercentStacked || normalized.includes("stacked");
  const isGenericExistingPlotReference =
    !normalized ||
    normalized === "barchart" ||
    normalized === "linechart" ||
    normalized === "areachart" ||
    normalized === "piechart" ||
    normalized === "doughnutchart";

  if (fallbackPlotElement && isGenericExistingPlotReference) {
    if (existingPlotName === "doughnutChart") {
      return {
        key: "doughnut",
        plotElementName: "doughnutChart",
        requiresAxes: false,
        supportsMultipleSeries: false,
      };
    }

    if (existingPlotName === "pieChart") {
      return {
        key: "pie",
        plotElementName: "pieChart",
        requiresAxes: false,
        supportsMultipleSeries: false,
      };
    }

    if (existingPlotName === "areaChart") {
      return {
        key: "area",
        plotElementName: "areaChart",
        grouping:
          existingGrouping === "percentStacked" || existingGrouping === "percentstacked"
            ? "percentStacked"
            : existingGrouping === "stacked"
              ? "stacked"
              : "standard",
        requiresAxes: true,
        supportsMultipleSeries: true,
      };
    }

    if (existingPlotName === "lineChart") {
      return {
        key: "line",
        plotElementName: "lineChart",
        grouping: "standard",
        requiresAxes: true,
        supportsMultipleSeries: true,
      };
    }

    if (existingPlotName === "barChart") {
      return {
        key: existingBarDirection === "bar" ? "bar" : "column",
        plotElementName: "barChart",
        barDirection: existingBarDirection === "bar" ? "bar" : "col",
        grouping:
          existingGrouping === "percentStacked" || existingGrouping === "percentstacked"
            ? "percentStacked"
            : existingGrouping === "stacked"
              ? "stacked"
              : "clustered",
        requiresAxes: true,
        supportsMultipleSeries: true,
      };
    }
  }

  if (normalized.includes("doughnut") || normalized.includes("donut") || existingPlotName === "doughnutChart") {
    return {
      key: "doughnut",
      plotElementName: "doughnutChart",
      requiresAxes: false,
      supportsMultipleSeries: false,
    };
  }

  if (normalized.includes("pie") || existingPlotName === "pieChart") {
    return {
      key: "pie",
      plotElementName: "pieChart",
      requiresAxes: false,
      supportsMultipleSeries: false,
    };
  }

  if (normalized.includes("area") || existingPlotName === "areaChart") {
    return {
      key: "area",
      plotElementName: "areaChart",
      grouping: wantsPercentStacked ? "percentStacked" : wantsStacked ? "stacked" : "standard",
      requiresAxes: true,
      supportsMultipleSeries: true,
    };
  }

  if (normalized.includes("line") || existingPlotName === "lineChart") {
    return {
      key: "line",
      plotElementName: "lineChart",
      grouping: "standard",
      requiresAxes: true,
      supportsMultipleSeries: true,
    };
  }

  if (normalized.includes("bar") || (existingPlotName === "barChart" && existingBarDirection === "bar")) {
    return {
      key: "bar",
      plotElementName: "barChart",
      barDirection: "bar",
      grouping: wantsPercentStacked ? "percentStacked" : wantsStacked ? "stacked" : "clustered",
      requiresAxes: true,
      supportsMultipleSeries: true,
    };
  }

  return {
    key: "column",
    plotElementName: "barChart",
    barDirection: "col",
    grouping: wantsPercentStacked ? "percentStacked" : wantsStacked ? "stacked" : "clustered",
    requiresAxes: true,
    supportsMultipleSeries: true,
  };
}

function normalizePowerPointChartData(
  chartType: ResolvedPowerPointChartType,
  input: PowerPointChartMutationInput | PowerPointChartCreationInput,
  fallback: {
    title?: string | undefined;
    series?: PowerPointChartSeriesSummary[] | undefined;
  } = {},
): NormalizedPowerPointChartData {
  const requestedSeries =
    input.series?.length
      ? input.series.map((series, index) => ({
          name: trimText(series.name) || `Series ${index + 1}`,
          categories: [...(series.categories ?? [])],
          values: [...series.values],
        }))
      : fallback.series?.map((series, index) => ({
          name: trimText(series.name) || `Series ${index + 1}`,
          categories: [...series.categories],
          values: [...series.values],
        })) ?? [];

  if (!requestedSeries.length) {
    throw new Error("PowerPoint chart actions require at least one series.");
  }

  if (!chartType.supportsMultipleSeries && requestedSeries.length > 1) {
    throw new Error(`${chartType.key} charts currently support exactly one series.`);
  }

  const explicitCategories = input.categories?.length ? [...input.categories] : undefined;
  const maxValueCount = Math.max(...requestedSeries.map((series) => series.values.length), 0);
  const defaultCategories =
    explicitCategories ??
    requestedSeries.find((series) => series.categories.length)?.categories ??
    Array.from({ length: maxValueCount }, (_entry, index) => index + 1);

  if (!defaultCategories.length) {
    throw new Error("PowerPoint chart actions require at least one category.");
  }

  const categoryKey = getChartValuesKey(defaultCategories);
  const normalizedSeries = requestedSeries.map((series, index) => {
    const categories = series.categories.length ? [...series.categories] : [...defaultCategories];
    if (categories.length !== series.values.length) {
      throw new Error(
        `Chart series ${index + 1} has ${series.values.length} values but ${categories.length} categories. Each series must align with the category axis.`,
      );
    }

    if (getChartValuesKey(categories) !== categoryKey) {
      throw new Error("All PowerPoint chart series must share the same category labels for this chart model.");
    }

    return {
      name: trimText(series.name) || `Series ${index + 1}`,
      categories,
      values: series.values.map((value) => {
        const numeric = Number(value);
        if (!Number.isFinite(numeric)) {
          throw new Error(`Chart series ${index + 1} contains a non-numeric value that cannot be written to the embedded workbook.`);
        }
        return numeric;
      }),
    } satisfies NormalizedPowerPointChartSeries;
  });

  return {
    title: trimText(input.title) || trimText(fallback.title) || undefined,
    chartType,
    categories: [...defaultCategories],
    series: normalizedSeries,
    showLegend: typeof (input as PowerPointChartCreationInput).showLegend === "boolean"
      ? Boolean((input as PowerPointChartCreationInput).showLegend)
      : normalizedSeries.length > 1,
  };
}

function buildPowerPointChartWorkbookLayout(data: NormalizedPowerPointChartData): PowerPointChartWorkbookLayout {
  const sheetName = "Sheet1";
  const lastColumnIndex = data.series.length + 1;
  const lastRowIndex = data.categories.length + 1;
  return {
    sheetName,
    dimensionRef: `A1:${toSpreadsheetCellReference(lastColumnIndex, lastRowIndex)}`,
    categoriesFormula: buildSpreadsheetRangeFormula(sheetName, 1, 2, lastRowIndex),
    series: data.series.map((_series, index) => ({
      nameFormula: buildSpreadsheetCellFormula(sheetName, index + 2, 1),
      valuesFormula: buildSpreadsheetRangeFormula(sheetName, index + 2, 2, lastRowIndex),
    })),
  };
}

function buildWorksheetCellXml(reference: string, value: string | number): string {
  if (typeof value === "number") {
    return `<c r="${reference}"><v>${value}</v></c>`;
  }

  const preserveSpace = /^\s|\s$/.test(value);
  return `<c r="${reference}" t="inlineStr"><is><t${preserveSpace ? ' xml:space="preserve"' : ""}>${escapeXml(value)}</t></is></c>`;
}

async function buildPowerPointChartWorkbookBase64(
  data: NormalizedPowerPointChartData,
  layout: PowerPointChartWorkbookLayout,
): Promise<string> {
  const workbookZip = new JSZip();
  const lastColumnIndex = data.series.length + 1;
  const lastRowIndex = data.categories.length + 1;
  const rowXml: string[] = [];

  rowXml.push(
    `<row r="1" spans="1:${lastColumnIndex}">${data.series
      .map((series, index) => buildWorksheetCellXml(toSpreadsheetCellReference(index + 2, 1), series.name))
      .join("")}</row>`,
  );

  data.categories.forEach((category, categoryIndex) => {
    const rowNumber = categoryIndex + 2;
    const cells = [buildWorksheetCellXml(toSpreadsheetCellReference(1, rowNumber), category)];
    data.series.forEach((series, seriesIndex) => {
      cells.push(buildWorksheetCellXml(toSpreadsheetCellReference(seriesIndex + 2, rowNumber), series.values[categoryIndex] ?? 0));
    });
    rowXml.push(`<row r="${rowNumber}" spans="1:${lastColumnIndex}">${cells.join("")}</row>`);
  });

  workbookZip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${CONTENT_TYPES_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="${WORKBOOK_CONTENT_TYPE}"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="${WORKSHEET_CONTENT_TYPE}"/>
</Types>`,
  );
  workbookZip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${WORKBOOK_RELATIONSHIP_TYPE}" Target="xl/workbook.xml"/>
</Relationships>`,
  );
  workbookZip.file(
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${OFFICE_REL_NS}">
  <bookViews><workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/></bookViews>
  <sheets><sheet name="${layout.sheetName}" sheetId="1" r:id="rId1"/></sheets>
  <calcPr calcId="191029" fullCalcOnLoad="1"/>
</workbook>`,
  );
  workbookZip.file(
    "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${WORKSHEET_RELATIONSHIP_TYPE}" Target="worksheets/sheet1.xml"/>
</Relationships>`,
  );
  workbookZip.file(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${SPREADSHEET_NS}">
  <dimension ref="${layout.dimensionRef}"/>
  <sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>${rowXml.join("")}</sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`,
  );

  return workbookZip.generateAsync({ type: "base64", compression: "DEFLATE", compressionOptions: { level: 6 } });
}

function buildChartSeriesElement(
  document: XMLDocument,
  data: NormalizedPowerPointChartData,
  layout: PowerPointChartWorkbookLayout,
  series: NormalizedPowerPointChartSeries,
  index: number,
): Element {
  const seriesElement = document.createElementNS(CHART_NS, "c:ser");
  const idx = document.createElementNS(CHART_NS, "c:idx");
  idx.setAttribute("val", String(index));
  seriesElement.appendChild(idx);

  const order = document.createElementNS(CHART_NS, "c:order");
  order.setAttribute("val", String(index));
  seriesElement.appendChild(order);

  setChartSeriesName(document, seriesElement, series.name, layout.series[index]?.nameFormula);

  if (data.chartType.key === "line") {
    const marker = document.createElementNS(CHART_NS, "c:marker");
    const symbol = document.createElementNS(CHART_NS, "c:symbol");
    symbol.setAttribute("val", "none");
    marker.appendChild(symbol);
    seriesElement.appendChild(marker);
  }

  setChartSeriesValues(document, seriesElement, {
    categories: data.categories,
    values: series.values,
    categoriesFormula: layout.categoriesFormula,
    valuesFormula: layout.series[index]?.valuesFormula,
  });

  if (data.chartType.key === "line") {
    const smooth = document.createElementNS(CHART_NS, "c:smooth");
    smooth.setAttribute("val", "0");
    seriesElement.appendChild(smooth);
  }

  return seriesElement;
}

function buildChartDataLabelsElement(document: XMLDocument, includeLeaderLines = false): Element {
  const labels = document.createElementNS(CHART_NS, "c:dLbls");
  const entries: Array<[string, string]> = [
    ["showLegendKey", "0"],
    ["showVal", "0"],
    ["showCatName", "0"],
    ["showSerName", "0"],
    ["showPercent", "0"],
    ["showBubbleSize", "0"],
  ];
  if (includeLeaderLines) {
    entries.push(["showLeaderLines", "1"]);
  }
  for (const [localName, value] of entries) {
    const entry = document.createElementNS(CHART_NS, `c:${localName}`);
    entry.setAttribute("val", value);
    labels.appendChild(entry);
  }
  return labels;
}

function buildCategoryAxisElement(document: XMLDocument, categoryAxisId: string, valueAxisId: string): Element {
  const axis = document.createElementNS(CHART_NS, "c:catAx");
  const parts: Array<[string, Record<string, string> | undefined]> = [
    ["axId", { val: categoryAxisId }],
    ["delete", { val: "0" }],
    ["axPos", { val: "b" }],
    ["majorTickMark", { val: "out" }],
    ["minorTickMark", { val: "none" }],
    ["tickLblPos", { val: "nextTo" }],
    ["crossAx", { val: valueAxisId }],
    ["crosses", { val: "autoZero" }],
    ["auto", { val: "1" }],
    ["lblAlgn", { val: "ctr" }],
    ["lblOffset", { val: "100" }],
    ["noMultiLvlLbl", { val: "0" }],
  ];

  const scaling = document.createElementNS(CHART_NS, "c:scaling");
  const orientation = document.createElementNS(CHART_NS, "c:orientation");
  orientation.setAttribute("val", "minMax");
  scaling.appendChild(orientation);
  const axisId = document.createElementNS(CHART_NS, "c:axId");
  axisId.setAttribute("val", categoryAxisId);
  axis.appendChild(axisId);
  axis.appendChild(scaling);

  for (const [localName, attributes] of parts.slice(1)) {
    const entry = document.createElementNS(CHART_NS, `c:${localName}`);
    for (const [attributeName, attributeValue] of Object.entries(attributes ?? {})) {
      entry.setAttribute(attributeName, attributeValue);
    }
    axis.appendChild(entry);
  }

  return axis;
}

function buildValueAxisElement(document: XMLDocument, categoryAxisId: string, valueAxisId: string): Element {
  const axis = document.createElementNS(CHART_NS, "c:valAx");
  const axisId = document.createElementNS(CHART_NS, "c:axId");
  axisId.setAttribute("val", valueAxisId);
  axis.appendChild(axisId);

  axis.appendChild(document.createElementNS(CHART_NS, "c:scaling"));

  const entries: Array<[string, Record<string, string> | undefined]> = [
    ["delete", { val: "0" }],
    ["axPos", { val: "l" }],
    ["majorGridlines", undefined],
    ["majorTickMark", { val: "out" }],
    ["minorTickMark", { val: "none" }],
    ["tickLblPos", { val: "nextTo" }],
    ["crossAx", { val: categoryAxisId }],
    ["crosses", { val: "autoZero" }],
  ];
  for (const [localName, attributes] of entries) {
    const entry = document.createElementNS(CHART_NS, `c:${localName}`);
    for (const [attributeName, attributeValue] of Object.entries(attributes ?? {})) {
      entry.setAttribute(attributeName, attributeValue);
    }
    axis.appendChild(entry);
  }
  return axis;
}

function buildChartPlotElement(
  document: XMLDocument,
  data: NormalizedPowerPointChartData,
  layout: PowerPointChartWorkbookLayout,
  axisIds: { categoryAxisId: string; valueAxisId: string },
): Element {
  const plotElement = document.createElementNS(CHART_NS, `c:${data.chartType.plotElementName}`);

  if (data.chartType.barDirection) {
    const barDirection = document.createElementNS(CHART_NS, "c:barDir");
    barDirection.setAttribute("val", data.chartType.barDirection);
    plotElement.appendChild(barDirection);
  }

  if (data.chartType.grouping) {
    const grouping = document.createElementNS(CHART_NS, "c:grouping");
    grouping.setAttribute("val", data.chartType.grouping);
    plotElement.appendChild(grouping);
  }

  if (data.chartType.key === "line" || data.chartType.key === "area" || data.chartType.key === "pie" || data.chartType.key === "doughnut") {
    const varyColors = document.createElementNS(CHART_NS, "c:varyColors");
    varyColors.setAttribute("val", data.chartType.key === "pie" || data.chartType.key === "doughnut" ? "1" : "0");
    plotElement.appendChild(varyColors);
  }

  data.series.forEach((series, index) => {
    plotElement.appendChild(buildChartSeriesElement(document, data, layout, series, index));
  });

  if (data.chartType.key === "line") {
    const marker = document.createElementNS(CHART_NS, "c:marker");
    marker.setAttribute("val", "1");
    plotElement.appendChild(marker);
    const smooth = document.createElementNS(CHART_NS, "c:smooth");
    smooth.setAttribute("val", "0");
    plotElement.appendChild(smooth);
  }

  if (data.chartType.key === "area") {
    plotElement.appendChild(buildChartDataLabelsElement(document));
  }

  if (data.chartType.key === "doughnut") {
    plotElement.appendChild(buildChartDataLabelsElement(document, true));
    const firstSliceAngle = document.createElementNS(CHART_NS, "c:firstSliceAng");
    firstSliceAngle.setAttribute("val", "0");
    plotElement.appendChild(firstSliceAngle);
    const holeSize = document.createElementNS(CHART_NS, "c:holeSize");
    holeSize.setAttribute("val", "50");
    plotElement.appendChild(holeSize);
  }

  if (data.chartType.requiresAxes) {
    const categoryAxisId = document.createElementNS(CHART_NS, "c:axId");
    categoryAxisId.setAttribute("val", axisIds.categoryAxisId);
    plotElement.appendChild(categoryAxisId);
    const valueAxisId = document.createElementNS(CHART_NS, "c:axId");
    valueAxisId.setAttribute("val", axisIds.valueAxisId);
    plotElement.appendChild(valueAxisId);
  }

  return plotElement;
}

function rebuildPowerPointChartDocument(
  document: XMLDocument,
  chartElement: Element,
  data: NormalizedPowerPointChartData,
  layout: PowerPointChartWorkbookLayout,
): void {
  let autoTitleDeleted = getDirectChildByLocalName(chartElement, "autoTitleDeleted");
  if (!autoTitleDeleted) {
    autoTitleDeleted = document.createElementNS(CHART_NS, "c:autoTitleDeleted");
    chartElement.insertBefore(autoTitleDeleted, chartElement.firstChild);
  }
  autoTitleDeleted.setAttribute("val", data.title ? "0" : "1");

  if (data.title) {
    setChartTitle(document, chartElement, data.title);
  } else {
    removeDirectChildrenByLocalName(chartElement, ["title"]);
  }

  let plotArea = getDirectChildByLocalName(chartElement, "plotArea");
  if (!plotArea) {
    plotArea = document.createElementNS(CHART_NS, "c:plotArea");
    chartElement.appendChild(plotArea);
  }
  while (plotArea.firstChild) {
    plotArea.removeChild(plotArea.firstChild);
  }

  const axisSeed = 1200000000 + data.categories.length * 100 + data.series.length * 10;
  const axisIds = {
    categoryAxisId: String(axisSeed),
    valueAxisId: String(axisSeed + 137),
  };
  plotArea.appendChild(buildChartPlotElement(document, data, layout, axisIds));
  if (data.chartType.requiresAxes) {
    plotArea.appendChild(buildCategoryAxisElement(document, axisIds.categoryAxisId, axisIds.valueAxisId));
    plotArea.appendChild(buildValueAxisElement(document, axisIds.categoryAxisId, axisIds.valueAxisId));
  }

  if (data.showLegend && data.series.length > 1) {
    let legend = getDirectChildByLocalName(chartElement, "legend");
    if (!legend) {
      legend = document.createElementNS(CHART_NS, "c:legend");
      const firstTailEntry =
        getDirectChildByLocalName(chartElement, "plotVisOnly") ??
        getDirectChildByLocalName(chartElement, "dispBlanksAs") ??
        getDirectChildByLocalName(chartElement, "showDLblsOverMax");
      chartElement.insertBefore(legend, firstTailEntry ?? null);
    }
    while (legend.firstChild) {
      legend.removeChild(legend.firstChild);
    }
    const legendPos = document.createElementNS(CHART_NS, "c:legendPos");
    legendPos.setAttribute("val", "r");
    legend.appendChild(legendPos);
    legend.appendChild(document.createElementNS(CHART_NS, "c:layout"));
    const overlay = document.createElementNS(CHART_NS, "c:overlay");
    overlay.setAttribute("val", "0");
    legend.appendChild(overlay);
  } else {
    removeDirectChildrenByLocalName(chartElement, ["legend"]);
  }

  let dispBlanksAs = getDirectChildByLocalName(chartElement, "dispBlanksAs");
  if (!dispBlanksAs) {
    dispBlanksAs = document.createElementNS(CHART_NS, "c:dispBlanksAs");
    chartElement.appendChild(dispBlanksAs);
  }
  dispBlanksAs.setAttribute("val", "gap");
}

function createPowerPointChartDocument(
  data: NormalizedPowerPointChartData,
  layout: PowerPointChartWorkbookLayout,
): XMLDocument {
  const document = parseXml(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${DRAWING_ML_NS}" xmlns:r="${OFFICE_REL_NS}">
  <c:date1904 val="0"/>
  <c:chart>
    <c:autoTitleDeleted val="1"/>
    <c:plotArea/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
  <c:txPr>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p><a:pPr><a:defRPr sz="1800"/></a:pPr><a:endParaRPr lang="en-US"/></a:p>
  </c:txPr>
</c:chartSpace>`,
    "chart",
  );
  const chartElement = getFirstElementByLocalName(document.documentElement, "chart");
  if (!chartElement) {
    throw new Error("Could not initialize the PowerPoint chart OOXML skeleton.");
  }
  rebuildPowerPointChartDocument(document, chartElement, data, layout);
  return document;
}

function ensureChartExternalDataBinding(document: XMLDocument, relationshipId: string): void {
  const root = document.documentElement;
  let externalData = getDirectChildByLocalName(root, "externalData");
  if (!externalData) {
    externalData = document.createElementNS(CHART_NS, "c:externalData");
    root.appendChild(externalData);
  }
  externalData.setAttributeNS(OFFICE_REL_NS, "r:id", relationshipId);
  externalData.setAttribute("r:id", relationshipId);

  let autoUpdate = getDirectChildByLocalName(externalData, "autoUpdate");
  if (!autoUpdate) {
    autoUpdate = document.createElementNS(CHART_NS, "c:autoUpdate");
    externalData.appendChild(autoUpdate);
  }
  autoUpdate.setAttribute("val", "0");
}

async function writeEmbeddedWorkbookForChart(
  zip: JSZip,
  workbookPartName: string,
  data: NormalizedPowerPointChartData,
  layout: PowerPointChartWorkbookLayout,
): Promise<void> {
  zip.file(workbookPartName, await buildPowerPointChartWorkbookBase64(data, layout), { base64: true });
}

async function ensureEmbeddedWorkbookBindingForChart(
  zip: JSZip,
  chartPartName: string,
  chartDocument: XMLDocument,
  workbookPartName: string,
  data: NormalizedPowerPointChartData,
  layout: PowerPointChartWorkbookLayout,
): Promise<void> {
  const chartRelationships = await loadOrCreateRelationshipsDocument(zip, chartPartName);
  const relationshipId = ensureRelationship(chartRelationships.document, chartPartName, PACKAGE_RELATIONSHIP_TYPE, workbookPartName);
  ensureChartExternalDataBinding(chartDocument, relationshipId);
  zip.file(chartRelationships.partName, new XMLSerializer().serializeToString(chartRelationships.document));
  await writeEmbeddedWorkbookForChart(zip, workbookPartName, data, layout);
}

function getNextPowerPointShapeId(document: XMLDocument): number {
  return (
    listElementsByLocalName(document, "cNvPr").reduce((current, entry) => {
      const id = Number(entry.getAttribute("id") ?? "0");
      return Math.max(current, Number.isFinite(id) ? id : 0);
    }, 0) + 1
  );
}

function appendChildBeforeExtensionList(parent: Element, child: Element): void {
  const extensionList = getDirectChildByLocalName(parent, "extLst");
  parent.insertBefore(child, extensionList ?? null);
}

function buildPowerPointChartFrame(
  document: XMLDocument,
  input: {
    relationshipId: string;
    shapeId: number;
    shapeName: string;
    left?: number | undefined;
    top?: number | undefined;
    width?: number | undefined;
    height?: number | undefined;
  },
): Element {
  const frame = document.createElementNS(PRESENTATION_NS, "p:graphicFrame");

  const nvGraphicFrame = document.createElementNS(PRESENTATION_NS, "p:nvGraphicFramePr");
  const nonVisualProps = document.createElementNS(PRESENTATION_NS, "p:cNvPr");
  nonVisualProps.setAttribute("id", String(input.shapeId));
  nonVisualProps.setAttribute("name", input.shapeName);
  nvGraphicFrame.appendChild(nonVisualProps);

  const graphicFrameProps = document.createElementNS(PRESENTATION_NS, "p:cNvGraphicFramePr");
  const graphicFrameLocks = document.createElementNS(DRAWING_ML_NS, "a:graphicFrameLocks");
  graphicFrameLocks.setAttribute("noGrp", "1");
  graphicFrameProps.appendChild(graphicFrameLocks);
  nvGraphicFrame.appendChild(graphicFrameProps);
  nvGraphicFrame.appendChild(document.createElementNS(PRESENTATION_NS, "p:nvPr"));
  frame.appendChild(nvGraphicFrame);

  const transform = document.createElementNS(PRESENTATION_NS, "p:xfrm");
  const offset = document.createElementNS(DRAWING_ML_NS, "a:off");
  offset.setAttribute("x", String(toEmu(input.left, DEFAULT_CHART_LEFT)));
  offset.setAttribute("y", String(toEmu(input.top, DEFAULT_CHART_TOP)));
  transform.appendChild(offset);
  const extent = document.createElementNS(DRAWING_ML_NS, "a:ext");
  extent.setAttribute("cx", String(toEmu(input.width, DEFAULT_CHART_WIDTH)));
  extent.setAttribute("cy", String(toEmu(input.height, DEFAULT_CHART_HEIGHT)));
  transform.appendChild(extent);
  frame.appendChild(transform);

  const graphic = document.createElementNS(DRAWING_ML_NS, "a:graphic");
  const graphicData = document.createElementNS(DRAWING_ML_NS, "a:graphicData");
  graphicData.setAttribute("uri", CHART_NS);
  const chartReference = document.createElementNS(CHART_NS, "c:chart");
  chartReference.setAttributeNS(OFFICE_REL_NS, "r:id", input.relationshipId);
  chartReference.setAttribute("r:id", input.relationshipId);
  graphicData.appendChild(chartReference);
  graphic.appendChild(graphicData);
  frame.appendChild(graphic);

  return frame;
}

function selectChartSummary(charts: PowerPointChartSummary[], input: PowerPointChartMutationInput): PowerPointChartSummary {
  const requestedShapeName = trimText(input.shapeName)?.toLowerCase();
  if (requestedShapeName) {
    const byShapeName = charts.find((chart) => chart.shapeName?.trim().toLowerCase() === requestedShapeName);
    if (byShapeName) {
      return byShapeName;
    }

    const fuzzyShapeMatch = charts.find((chart) => chart.shapeName?.trim().toLowerCase().includes(requestedShapeName));
    if (fuzzyShapeMatch) {
      return fuzzyShapeMatch;
    }
  }

  if (input.chartIndex && charts[input.chartIndex - 1]) {
    return charts[input.chartIndex - 1]!;
  }

  if (charts.length === 1) {
    return charts[0]!;
  }

  throw new Error("Multiple charts were found in the exported slide package. Pass chartIndex or shapeName to disambiguate.");
}

export async function inspectPowerPointPresentationBase64(base64: string): Promise<PowerPointPresentationPackageSummary> {
  const zip = await loadZipFromBase64(base64);
  const slidePartNames = getSortedSlidePartNames(zip);
  const notesPartNames = getSortedNotesPartNames(zip);
  const notes = await Promise.all(
    notesPartNames.map(async (partName) => {
      const xml = await readZipText(zip, partName);
      const text = xml ? extractNotesText(parseXml(xml, partName)) : "";
      const slideNumber = Number(partName.match(/notesSlide(\d+)\.xml/i)?.[1] ?? "0");
      return {
        slideNumber,
        partName,
        hasNotes: Boolean(text),
        text,
        preview: truncatePreview(text),
      } satisfies PowerPointSlideNotesSummary;
    }),
  );
  const charts = (
    await Promise.all(
      slidePartNames.map((partName, index) => inspectChartsInSlidePart(zip, partName, index + 1)),
    )
  ).flat();

  return {
    theme: await inspectPowerPointTheme(zip),
    notes,
    charts,
    partCounts: {
      slides: slidePartNames.length,
      slideMasters: Object.keys(zip.files).filter((name) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(name)).length,
      layouts: Object.keys(zip.files).filter((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(name)).length,
      themes: Object.keys(zip.files).filter((name) => /^ppt\/theme\/theme\d+\.xml$/i.test(name)).length,
      notesSlides: notesPartNames.length,
      charts: Object.keys(zip.files).filter((name) => /^ppt\/charts\/chart\d+\.xml$/i.test(name)).length,
      media: Object.keys(zip.files).filter((name) => /^ppt\/media\//i.test(name)).length,
      customXmlParts: Object.keys(zip.files).filter((name) => /^customXml\//i.test(name)).length,
    },
  };
}

export async function replaceSlideNotesInPowerPointPresentationBase64(
  base64: string,
  input: { notesText: string; slideNumber?: number | undefined },
): Promise<PowerPointSlideNotesMutationResult> {
  const zip = await loadZipFromBase64(base64);
  const notesPartNames = getSortedNotesPartNames(zip);
  if (!notesPartNames.length) {
    throw new Error("The exported PowerPoint package does not contain a notes slide part to update.");
  }

  const requestedSlideNumber = input.slideNumber ?? 1;
  const partName =
    notesPartNames.find((name) => Number(name.match(/notesSlide(\d+)\.xml/i)?.[1] ?? "0") === requestedSlideNumber) ?? notesPartNames[0];
  if (!partName) {
    throw new Error("Could not resolve the notes slide part to update.");
  }

  const xml = await readZipText(zip, partName);
  if (!xml) {
    throw new Error(`Could not load ${partName} from the exported PowerPoint package.`);
  }

  const document = parseXml(xml, partName);
  const previousText = replaceNotesText(document, input.notesText);
  zip.file(partName, new XMLSerializer().serializeToString(document));

  return {
    base64: await zip.generateAsync({ type: "base64", compression: "DEFLATE", compressionOptions: { level: 6 } }),
    slideNumber: requestedSlideNumber,
    partName,
    changed: normalizeText(previousText) !== normalizeText(input.notesText),
    previousText,
    nextText: input.notesText,
  };
}

export async function createPowerPointChartInPresentationBase64(
  base64: string,
  input: PowerPointChartCreationInput,
): Promise<PowerPointChartCreationResult> {
  const zip = await loadZipFromBase64(base64);
  const slidePartNames = getSortedSlidePartNames(zip);
  if (!slidePartNames.length) {
    throw new Error("The exported PowerPoint package does not contain a slide to update.");
  }

  const slidePartName = slidePartNames[0]!;
  const slideXml = await readZipText(zip, slidePartName);
  if (!slideXml) {
    throw new Error(`Could not load ${slidePartName} from the exported PowerPoint package.`);
  }

  const slideDocument = parseXml(slideXml, slidePartName);
  const shapeTree = getFirstElementByLocalName(slideDocument, "spTree");
  if (!shapeTree) {
    throw new Error(`Could not resolve the slide shape tree inside ${slidePartName}.`);
  }

  const chartType = normalizePowerPointChartType(input.chartType);
  const data = normalizePowerPointChartData(chartType, input);
  const layout = buildPowerPointChartWorkbookLayout(data);
  const chartDocument = createPowerPointChartDocument(data, layout);
  const existingCharts = await inspectChartsInSlidePart(zip, slidePartName, 1);

  const chartPartNumber = getNextPartNumber(zip, /^ppt\/charts\/chart(\d+)\.xml$/i);
  const workbookPartNumber = getNextPartNumber(zip, /^ppt\/embeddings\/Microsoft_Excel_Sheet(\d+)\.xlsx$/i);
  const chartPartName = `ppt/charts/chart${chartPartNumber}.xml`;
  const embeddedWorkbookPartName = `ppt/embeddings/Microsoft_Excel_Sheet${workbookPartNumber}.xlsx`;

  const slideRelationships = await loadOrCreateRelationshipsDocument(zip, slidePartName);
  const chartRelationshipId = ensureRelationship(slideRelationships.document, slidePartName, CHART_RELATIONSHIP_TYPE, chartPartName);
  const shapeId = getNextPowerPointShapeId(slideDocument);
  const shapeName = trimText(input.shapeName) || `Chart ${shapeId}`;
  appendChildBeforeExtensionList(
    shapeTree,
    buildPowerPointChartFrame(slideDocument, {
      relationshipId: chartRelationshipId,
      shapeId,
      shapeName,
      left: input.left,
      top: input.top,
      width: input.width,
      height: input.height,
    }),
  );

  const contentTypesDocument = await loadContentTypesDocument(zip);
  ensureContentTypeDefault(contentTypesDocument, "xlsx", XLSX_CONTENT_TYPE);
  ensureContentTypeOverride(contentTypesDocument, chartPartName, CHART_CONTENT_TYPE);
  await ensureEmbeddedWorkbookBindingForChart(zip, chartPartName, chartDocument, embeddedWorkbookPartName, data, layout);

  zip.file(slidePartName, new XMLSerializer().serializeToString(slideDocument));
  zip.file(slideRelationships.partName, new XMLSerializer().serializeToString(slideRelationships.document));
  zip.file(chartPartName, new XMLSerializer().serializeToString(chartDocument));
  zip.file("[Content_Types].xml", new XMLSerializer().serializeToString(contentTypesDocument));

  return {
    base64: await zip.generateAsync({ type: "base64", compression: "DEFLATE", compressionOptions: { level: 6 } }),
    slideNumber: 1,
    chartIndex: existingCharts.length + 1,
    chartPartName,
    chartType: data.chartType.key,
    title: data.title,
    shapeId: String(shapeId),
    shapeName,
    embeddedWorkbookPartName,
    seriesCount: data.series.length,
    categoryCount: data.categories.length,
  };
}

export async function updatePowerPointChartInPresentationBase64(
  base64: string,
  input: PowerPointChartMutationInput,
): Promise<PowerPointChartMutationResult> {
  const zip = await loadZipFromBase64(base64);
  const slidePartNames = getSortedSlidePartNames(zip);
  if (!slidePartNames.length) {
    throw new Error("The exported PowerPoint package does not contain a slide to update.");
  }

  const slidePartName = slidePartNames[0]!;
  const charts = await inspectChartsInSlidePart(zip, slidePartName, 1);
  if (!charts.length) {
    throw new Error("The exported PowerPoint slide package does not contain a chart to update.");
  }

  const selectedChart = selectChartSummary(charts, input);
  const chartXml = await readZipText(zip, selectedChart.chartPartName);
  if (!chartXml) {
    throw new Error(`Could not load ${selectedChart.chartPartName} from the exported PowerPoint package.`);
  }

  const chartDocument = parseXml(chartXml, selectedChart.chartPartName);
  const chartRoot = chartDocument.documentElement;
  const chartElement = getFirstElementByLocalName(chartRoot, "chart");
  if (!chartElement) {
    throw new Error(`Could not resolve the chart root inside ${selectedChart.chartPartName}.`);
  }

  const existingPlotElement =
    getDirectChildByLocalName(chartElement, "plotArea")
      ? Array.from(getDirectChildByLocalName(chartElement, "plotArea")!.children).find((child) => child.localName.endsWith("Chart"))
      : undefined;
  const chartType = normalizePowerPointChartType(input.chartType ?? selectedChart.chartType, existingPlotElement);
  const data = normalizePowerPointChartData(chartType, input, {
    title: selectedChart.title,
    series: selectedChart.series,
  });
  const workbookLayout = buildPowerPointChartWorkbookLayout(data);
  rebuildPowerPointChartDocument(chartDocument, chartElement, data, workbookLayout);

  const contentTypesDocument = await loadContentTypesDocument(zip);
  ensureContentTypeDefault(contentTypesDocument, "xlsx", XLSX_CONTENT_TYPE);
  const embeddedWorkbookPartName =
    selectedChart.embeddedWorkbookPartName ?? `ppt/embeddings/Microsoft_Excel_Sheet${getNextPartNumber(zip, /^ppt\/embeddings\/Microsoft_Excel_Sheet(\d+)\.xlsx$/i)}.xlsx`;
  await ensureEmbeddedWorkbookBindingForChart(zip, selectedChart.chartPartName, chartDocument, embeddedWorkbookPartName, data, workbookLayout);

  zip.file(selectedChart.chartPartName, new XMLSerializer().serializeToString(chartDocument));
  zip.file("[Content_Types].xml", new XMLSerializer().serializeToString(contentTypesDocument));

  return {
    base64: await zip.generateAsync({ type: "base64", compression: "DEFLATE", compressionOptions: { level: 6 } }),
    slideNumber: 1,
    chartIndex: selectedChart.chartIndex,
    chartPartName: selectedChart.chartPartName,
    chartType: data.chartType.key,
    title: data.title,
    shapeName: selectedChart.shapeName,
    embeddedWorkbookPartName,
    seriesCount: data.series.length,
    categoryCount: data.categories.length,
    hadEmbeddedWorkbook: selectedChart.hasEmbeddedWorkbook,
  };
}
