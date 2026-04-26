import type { ExcelBorderUpdateOptions } from "../shared";
import {
  supportsRequirementSet,
  getRecordArray,
  isRecord,
  trimString,
  toNumber,
  toBoolean,
} from "../shared";

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

export function applyExcelRangeBorders(range: Excel.Range, options: Record<string, unknown>): boolean {
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
