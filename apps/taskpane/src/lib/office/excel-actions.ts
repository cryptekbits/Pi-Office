import type { OfficeAnchor, OfficeContextPayload, OfficeHostAction, OfficeStateUpdate, OfficeSelectionMeta } from "@pi-office/pi-office-pack/protocol";
import type { OfficeCaptureOptions } from "../office-host-adapter-types";
import type { ExcelBorderUpdateOptions, ExcelCitationRecord, ExcelWorksheetSnapshot } from "./shared";
import {
  supportsRequirementSet,
  isRecord,
  trimString,
  uniqueAnchors,
  getActionOptions,
  toNumber,
  toBoolean,
  toStringMatrix,
  toValueMatrix,
  splitSheetAddress,
  withSheetName,
  isSingleCellAddress,
  excelCellAddress,
  firstMatrixString,
  buildExcelCitationRecord,
  getSelectedImageAsync,
  optimizeVisual,
  createSummary,
} from "./shared";
import {
  resolveExcelWorksheet,
  resolveExcelRange,
  resolveExcelTable,
  resolveExcelChart,
  resolveExcelPivotTable,
  resolveExcelTableColumn,
} from "./excel-targets";

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

function toExcelClearApplyTo(value: unknown): Excel.ClearApplyTo {
  const normalized = trimString(value)?.replace(/[\s_-]/g, "").toLowerCase();
  switch (normalized) {
    case "contents":
    case "content":
      return Excel.ClearApplyTo.contents;
    case "formats":
    case "format":
      return Excel.ClearApplyTo.formats;
    case "hyperlinks":
    case "hyperlink":
      return Excel.ClearApplyTo.hyperlinks;
    case "removehyperlinks":
    case "removehyperlink":
      return Excel.ClearApplyTo.removeHyperlinks;
    case "all":
    default:
      return Excel.ClearApplyTo.all;
  }
}


export async function applyExcelAction(action: OfficeHostAction): Promise<unknown> {
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

    if (type === "getRangeValues") {
      const range = resolveExcelRange(context, action.target, true);
      const includeValues = toBoolean(options.includeValues) ?? true;
      const includeText = toBoolean(options.includeText) ?? true;
      const includeFormulas = toBoolean(options.includeFormulas) ?? true;
      const includeNumberFormat = toBoolean(options.includeNumberFormat) ?? true;
      range.load(["address", "rowCount", "columnCount", "values", "text", "formulas", "numberFormat"]);
      range.worksheet.load("name");
      await context.sync();
      return {
        ok: true,
        host: "excel",
        action: type,
        sheetName: range.worksheet.name,
        address: range.address,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        ...(includeValues ? { values: range.values } : {}),
        ...(includeText ? { text: range.text } : {}),
        ...(includeFormulas ? { formulas: range.formulas } : {}),
        ...(includeNumberFormat ? { numberFormat: range.numberFormat } : {}),
      };
    }

    if (type === "clearRange") {
      const range = resolveExcelRange(context, action.target, false);
      range.load("address,rowCount,columnCount");
      range.worksheet.load("name");
      await context.sync();
      const applyTo = toExcelClearApplyTo(options.applyTo ?? options.clearApplyTo ?? options.clearType);
      range.clear(applyTo);
      await context.sync();
      return {
        ok: true,
        host: "excel",
        action: type,
        sheetName: range.worksheet.name,
        address: range.address,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        applyTo,
      };
    }

    if (type === "resizeRange") {
      const sourceRange = resolveExcelRange(context, action.target, false);
      sourceRange.load("address,rowCount,columnCount");
      sourceRange.worksheet.load("name");
      await context.sync();

      const requestedRowCount = toNumber(options.rowCount ?? options.rows ?? options.targetRowCount);
      const requestedColumnCount = toNumber(options.columnCount ?? options.columns ?? options.targetColumnCount);
      const requestedRowDelta = toNumber(options.rowDelta ?? options.rowsDelta ?? options.deltaRows);
      const requestedColumnDelta = toNumber(options.columnDelta ?? options.columnsDelta ?? options.deltaColumns);

      const rowDelta =
        typeof requestedRowCount === "number"
          ? Math.trunc(requestedRowCount) - sourceRange.rowCount
          : Math.trunc(requestedRowDelta ?? 0);
      const columnDelta =
        typeof requestedColumnCount === "number"
          ? Math.trunc(requestedColumnCount) - sourceRange.columnCount
          : Math.trunc(requestedColumnDelta ?? 0);

      if (sourceRange.rowCount + rowDelta < 1 || sourceRange.columnCount + columnDelta < 1) {
        throw new Error("Excel resizeRange cannot produce a range smaller than 1x1.");
      }

      const resizedRange = sourceRange.getResizedRange(rowDelta, columnDelta);
      resizedRange.load("address,rowCount,columnCount");
      resizedRange.worksheet.load("name");
      if (toBoolean(options.activate ?? options.select) ?? false) {
        resizedRange.select();
      }
      await context.sync();

      return {
        ok: true,
        host: "excel",
        action: type,
        sourceSheetName: sourceRange.worksheet.name,
        sourceAddress: sourceRange.address,
        sheetName: resizedRange.worksheet.name,
        address: resizedRange.address,
        rowCount: resizedRange.rowCount,
        columnCount: resizedRange.columnCount,
      };
    }

    if (type === "copyRange") {
      const sourceRange = resolveExcelRange(context, action.target, true);
      sourceRange.load("address");
      sourceRange.worksheet.load("name");

      const destinationAddressInput =
        trimString(options.destinationAddress) ??
        trimString(options.targetAddress) ??
        trimString(options.address);
      if (!destinationAddressInput) {
        throw new Error("Excel copyRange requires destinationAddress.");
      }
      const destinationSheetInput =
        trimString(options.destinationSheetName) ??
        trimString(options.targetSheetName) ??
        trimString(options.sheetName);
      const parsedDestination = splitSheetAddress(destinationAddressInput, destinationSheetInput);
      const destinationWorksheet = parsedDestination.sheetName
        ? resolveExcelWorksheet(context, { kind: "sheet", sheetName: parsedDestination.sheetName }, false)
        : sourceRange.worksheet;
      const destinationAddress = parsedDestination.address ?? destinationAddressInput;
      const destinationRange = destinationWorksheet.getRange(destinationAddress);
      destinationRange.load("address");
      destinationWorksheet.load("name");
      await context.sync();

      const copyType = trimString(options.copyType) as Excel.RangeCopyType | undefined;
      const skipBlanks = toBoolean(options.skipBlanks) ?? false;
      const transpose = toBoolean(options.transpose) ?? false;
      destinationRange.copyFrom(sourceRange, copyType, skipBlanks, transpose);
      if (toBoolean(options.selectDestination) ?? false) {
        destinationRange.select();
      }
      await context.sync();

      return {
        ok: true,
        host: "excel",
        action: type,
        sourceSheetName: sourceRange.worksheet.name,
        sourceAddress: sourceRange.address,
        destinationSheetName: destinationWorksheet.name,
        destinationAddress: destinationRange.address,
        copyType: copyType ?? "All",
        skipBlanks,
        transpose,
      };
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


