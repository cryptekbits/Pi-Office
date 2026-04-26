import {
  isRecord,
  trimString,
  toNumber,
  toBoolean,
} from "../shared";

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

export function applyExcelTableFilter(column: Excel.TableColumn, options: Record<string, unknown>): void {
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
