import type { OfficeHostAction } from "@pi-office/pi-office-pack/protocol";
import {
  supportsRequirementSet,
  isRecord,
  trimString,
  toBoolean,
  withSheetName,
  getStringArray,
  getRecordArray,
} from "../shared";
import { resolveExcelPivotTable } from "../excel-targets";

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

function getExcelPivotHierarchyCollection(
  pivotTable: Excel.PivotTable,
  axis: "row" | "column" | "filter" | "data",
):
  | Excel.RowColumnPivotHierarchyCollection
  | Excel.FilterPivotHierarchyCollection
  | Excel.DataPivotHierarchyCollection {
  if (axis === "row") return pivotTable.rowHierarchies;
  if (axis === "column") return pivotTable.columnHierarchies;
  if (axis === "filter") return pivotTable.filterHierarchies;
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
    await replaceExcelPivotHierarchies(context, pivotTable, axis, getStringArray(value));
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
  if (typeof allowMultipleFiltersPerField === "boolean") pivotTable.allowMultipleFiltersPerField = allowMultipleFiltersPerField;
  if (typeof enableDataValueEditing === "boolean") pivotTable.enableDataValueEditing = enableDataValueEditing;
  if (typeof refreshOnOpen === "boolean" && supportsRequirementSet("ExcelApi", "1.13")) pivotTable.refreshOnOpen = refreshOnOpen;
  if (typeof useCustomSortLists === "boolean" && supportsRequirementSet("ExcelApi", "1.9")) pivotTable.useCustomSortLists = useCustomSortLists;

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
    if (layoutType) pivotTable.layout.layoutType = layoutType as Excel.PivotLayoutType;
    if (emptyCellText) pivotTable.layout.emptyCellText = emptyCellText;
    if (typeof fillEmptyCells === "boolean") pivotTable.layout.fillEmptyCells = fillEmptyCells;
    if (typeof showColumnGrandTotals === "boolean") pivotTable.layout.showColumnGrandTotals = showColumnGrandTotals;
    if (typeof showRowGrandTotals === "boolean") pivotTable.layout.showRowGrandTotals = showRowGrandTotals;
    if (subtotalLocation) pivotTable.layout.subtotalLocation = subtotalLocation as Excel.SubtotalLocationType;
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

export async function applyExcelCreatePivotTableAction(
  context: Excel.RequestContext,
  action: OfficeHostAction,
  type: string,
  options: Record<string, unknown>,
): Promise<unknown> {
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

export async function applyExcelUpdatePivotTableAction(
  context: Excel.RequestContext,
  action: OfficeHostAction,
  type: string,
  options: Record<string, unknown>,
): Promise<unknown> {
  const pivotTable = resolveExcelPivotTable(
    context,
    action.target ?? { kind: "pivotTable", pivotTableName: trimString(options.pivotTableName) ?? "" },
  );
  await applyExcelPivotConfiguration(context, pivotTable, options);
  pivotTable.load("name,id");
  await context.sync();
  return { ok: true, host: "excel", action: type, pivotTableName: pivotTable.name, pivotTableId: pivotTable.id };
}

export async function applyExcelPivotSortAction(
  context: Excel.RequestContext,
  action: OfficeHostAction,
  type: string,
  options: Record<string, unknown>,
): Promise<unknown> {
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

export async function applyExcelRefreshPivotTableAction(
  context: Excel.RequestContext,
  action: OfficeHostAction,
  type: string,
  options: Record<string, unknown>,
): Promise<unknown> {
  const pivotTable = resolveExcelPivotTable(
    context,
    action.target ?? { kind: "pivotTable", pivotTableName: trimString(options.pivotTableName) ?? "" },
  );
  pivotTable.refresh();
  pivotTable.load("name,id");
  await context.sync();
  return { ok: true, host: "excel", action: type, pivotTableName: pivotTable.name, pivotTableId: pivotTable.id };
}
