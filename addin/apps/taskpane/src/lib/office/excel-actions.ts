import type { OfficeAnchor, OfficeContextPayload, OfficeHostAction, OfficeStateUpdate, OfficeSelectionMeta } from "@pi-office/pi-office-pack/protocol";
import type { OfficeCaptureOptions } from "../office-host-adapter-types";
import type { ExcelCitationRecord, ExcelWorksheetSnapshot } from "./shared";
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
  getNumberArray,
  getRecordArray,
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
  resolveExcelTableColumn,
} from "./excel-targets";
import {
  applyExcelCreateChartAction,
  applyExcelExtractChartXmlAction,
  applyExcelUpdateChartAction,
} from "./excel-actions/charts";
import { applyExcelRangeBorders } from "./excel-actions/formatting";
import {
  applyExcelCreatePivotTableAction,
  applyExcelPivotSortAction,
  applyExcelRefreshPivotTableAction,
  applyExcelUpdatePivotTableAction,
} from "./excel-actions/pivots";
import { applyExcelTableFilter } from "./excel-actions/table-filters";

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
      return applyExcelCreateChartAction(context, action, type, options);
    }

    if (type === "updateChart" || type === "formatChart" || type === "setChartAxes" || type === "setChartDataLabels") {
      return applyExcelUpdateChartAction(context, action, type, options);
    }

    if (type === "extractChartXml") {
      return applyExcelExtractChartXmlAction(context, action, type, options);
    }

    if (type === "createPivotTable") {
      return applyExcelCreatePivotTableAction(context, action, type, options);
    }

    if (type === "updatePivotTable" || type === "configurePivotTable" || type === "applyPivotFilter") {
      return applyExcelUpdatePivotTableAction(context, action, type, options);
    }

    if (type === "sortPivotField" || type === "sortPivotByLabels" || type === "sortPivotByValues") {
      return applyExcelPivotSortAction(context, action, type, options);
    }

    if (type === "refreshPivotTable") {
      return applyExcelRefreshPivotTableAction(context, action, type, options);
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


