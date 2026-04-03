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

export async function navigateExcelAnchor(anchor: OfficeAnchor): Promise<unknown> {
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

