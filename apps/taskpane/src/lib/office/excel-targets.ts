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

export function resolveExcelWorksheet(
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

export function resolveExcelRange(
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

export function resolveExcelTable(context: Excel.RequestContext, target: OfficeAnchor): Excel.Table {
  if (!target.tableName) {
    throw new Error("Excel table name is required for this action.");
  }

  if (Boolean(target.sheetName || (target.address && target.address.includes("!")))) {
    return resolveExcelWorksheet(context, target, true).tables.getItem(target.tableName);
  }

  return context.workbook.tables.getItem(target.tableName);
}

export function resolveExcelChart(context: Excel.RequestContext, target: OfficeAnchor): Excel.Chart {
  if (!target.chartName) {
    throw new Error("Excel chart name is required for this action.");
  }

  return resolveExcelWorksheet(context, target, true).charts.getItem(target.chartName);
}

export function resolveExcelPivotTable(context: Excel.RequestContext, target: OfficeAnchor): Excel.PivotTable {
  if (!target.pivotTableName) {
    throw new Error("Excel PivotTable name is required for this action.");
  }

  if (Boolean(target.sheetName || (target.address && target.address.includes("!")))) {
    return resolveExcelWorksheet(context, target, true).pivotTables.getItem(target.pivotTableName);
  }

  return context.workbook.pivotTables.getItem(target.pivotTableName);
}

export function resolveExcelTableColumn(table: Excel.Table, options: Record<string, unknown>): Excel.TableColumn {
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

