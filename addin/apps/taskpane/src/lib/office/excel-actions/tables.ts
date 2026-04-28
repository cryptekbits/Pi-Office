import type { OfficeHostAction } from "@pi-office/pi-office-pack/protocol";
import { getActionOptions, supportsRequirementSet, toBoolean, trimString } from "../shared";
import { resolveExcelRange, resolveExcelTable, resolveExcelTableColumn } from "../excel-targets";
import { applyExcelTableFilter } from "./table-filters";
import { applyExcelRangeFormatting } from "./ranges";

export function isExcelTableAction(type: string): boolean {
  return [
    "createTable",
    "formatTable",
    "updateTableStyle",
    "configureTable",
    "setTableStyle",
    "applyTableFilter",
    "filterTable",
    "clearTableFilter",
    "clearTableFilters",
    "reapplyTableFilters",
  ].includes(type);
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

  if (style) table.style = style;
  if (typeof showHeaders === "boolean") table.showHeaders = showHeaders;
  if (typeof showTotals === "boolean") table.showTotals = showTotals;
  if (typeof showBandedRows === "boolean") table.showBandedRows = showBandedRows;
  if (typeof showBandedColumns === "boolean") table.showBandedColumns = showBandedColumns;
  if (typeof highlightFirstColumn === "boolean") table.highlightFirstColumn = highlightFirstColumn;
  if (typeof highlightLastColumn === "boolean") table.highlightLastColumn = highlightLastColumn;
  if (typeof showFilterButton === "boolean") {
    if (showFilterButton && showHeaders === false) {
      throw new Error("Excel table filter buttons require showHeaders=true.");
    }
    table.showFilterButton = showFilterButton;
  }

  applyExcelRangeFormatting(table.getRange(), action);
}

export async function applyExcelTableAction(
  context: Excel.RequestContext,
  action: OfficeHostAction,
  type: string,
  options: Record<string, unknown>,
): Promise<unknown> {
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

  throw new Error(`Unsupported Excel table action: ${type}`);
}
