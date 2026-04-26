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

export async function collectExcelState(base: OfficeStateUpdate): Promise<OfficeStateUpdate> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["address", "rowCount", "columnCount", "text"]);
    range.worksheet.load("name");
    await context.sync();

    const preview = range.text.flat().join(" | ").trim();
    const previewRows = range.text.slice(0, 8).map((row) => row.join("\t"));
    const excelStructuredPreview = previewRows.length > 0 ? previewRows.join("\n") : undefined;
    const excelMeta: OfficeSelectionMeta = {
      paragraphCount: range.rowCount,
      firstParagraphStyle: range.rowCount === 1 && range.columnCount === 1 ? "cell" : "range",
      styleHistogram: { [`${range.rowCount}x${range.columnCount}`]: 1 },
    };
    return {
      ...base,
      selection: {
        label: `${range.worksheet.name}!${range.address}`,
        kind: preview ? "text" : "empty",
        imageCount: 0,
        objectCount: 0,
        textPreview: preview || undefined,
        structuredPreview: excelStructuredPreview,
        selectionMeta: excelMeta,
        details: [`${range.rowCount} rows`, `${range.columnCount} columns`],
      },
      capabilities: ["excel.range", "excel.values", "excel.formulas", "excel.format"],
    };
  });
}


export async function collectExcelContext(base: OfficeStateUpdate, options: OfficeCaptureOptions): Promise<OfficeContextPayload> {
  const payload = await Excel.run(async (context) => {
    const normalizedScope = trimString(options.scope)?.toLowerCase();
    const supportsWorksheetView = supportsRequirementSet("ExcelApi", "1.8");
    const supportsPageLayout = supportsRequirementSet("ExcelApi", "1.9");
    const range = context.workbook.getSelectedRange();
    const workbook = context.workbook;
    const worksheets = workbook.worksheets;
    const namedItems = workbook.names;
    const activeWorksheet = range.worksheet;
    const format = range.format;
    const font = format.font;
    const fill = format.fill;
    const tables = activeWorksheet.tables;
    const charts = activeWorksheet.charts;
    const pivotTables = activeWorksheet.pivotTables;
    const activePrintArea = supportsPageLayout ? activeWorksheet.pageLayout.getPrintAreaOrNullObject() : undefined;

    range.load(["address", "rowCount", "columnCount", "text", "formulas", "numberFormat"]);
    activeWorksheet.load("name,id");
    if (supportsWorksheetView) {
      activeWorksheet.load("showGridlines,showHeadings");
    }
    worksheets.load("items/name,items/id,items/position,items/visibility");
    namedItems.load("items/name,items/type");
    tables.load("items/name,items/id");
    charts.load("items/name,items/id");
    pivotTables.load("items/name,items/id");
    format.load(["horizontalAlignment", "verticalAlignment", "wrapText", "rowHeight", "columnWidth"]);
    font.load(["name", "size", "color", "bold", "italic", "underline"]);
    fill.load("color");
    activePrintArea?.load("isNullObject,address");
    await context.sync();

    const usedRanges = worksheets.items.map((worksheet) => worksheet.getUsedRangeOrNullObject(true));
    const worksheetTables = worksheets.items.map((worksheet) => worksheet.tables);
    const worksheetCharts = worksheets.items.map((worksheet) => worksheet.charts);
    const worksheetPivotTables = worksheets.items.map((worksheet) => worksheet.pivotTables);
    const worksheetPrintAreas = supportsPageLayout ? worksheets.items.map((worksheet) => worksheet.pageLayout.getPrintAreaOrNullObject()) : [];
    for (const usedRange of usedRanges) {
      usedRange.load("isNullObject,address,rowCount,columnCount,rowIndex,columnIndex");
    }
    for (const worksheet of worksheets.items) {
      if (supportsWorksheetView) {
        worksheet.load("showGridlines,showHeadings");
      }
    }
    for (const tableCollection of worksheetTables) {
      tableCollection.load("items/name,items/id");
    }
    for (const chartCollection of worksheetCharts) {
      chartCollection.load("items/name,items/id");
    }
    for (const pivotCollection of worksheetPivotTables) {
      pivotCollection.load("items/name,items/id");
    }
    for (const printArea of worksheetPrintAreas) {
      printArea.load("isNullObject,address");
    }
    await context.sync();

    const selectionCellRanges: Excel.Range[] = [];
    const maxSelectionCells = normalizedScope === "selection" ? 16 : 12;
    let selectionCellBudget = maxSelectionCells;
    for (let rowIndex = 0; rowIndex < range.rowCount && selectionCellBudget > 0; rowIndex += 1) {
      for (let columnIndex = 0; columnIndex < range.columnCount && selectionCellBudget > 0; columnIndex += 1) {
        const cell = range.getCell(rowIndex, columnIndex);
        cell.load(["address", "text", "formulas", "numberFormat"]);
        selectionCellRanges.push(cell);
        selectionCellBudget -= 1;
      }
    }

    const activeWorksheetIndex = worksheets.items.findIndex((worksheet) => worksheet.id === activeWorksheet.id || worksheet.name === activeWorksheet.name);
    const snapshotIndexes = new Set<number>();
    if (activeWorksheetIndex >= 0) {
      snapshotIndexes.add(activeWorksheetIndex);
    }
    if (normalizedScope !== "selection" && normalizedScope !== "worksheet") {
      for (let index = 0; index < worksheets.items.length && snapshotIndexes.size < 6; index += 1) {
        const usedRange = usedRanges[index];
        if (index !== activeWorksheetIndex && usedRange && !usedRange.isNullObject && usedRange.address) {
          snapshotIndexes.add(index);
        }
      }
    }

    const worksheetSnapshotRanges: Array<{
      worksheetIndex: number;
      previewRange: Excel.Range;
      startRowIndex: number;
      startColumnIndex: number;
      rowCount: number;
      columnCount: number;
    }> = [];
    for (const worksheetIndex of snapshotIndexes) {
      const usedRange = usedRanges[worksheetIndex];
      const worksheet = worksheets.items[worksheetIndex];
      if (!worksheet || !usedRange || usedRange.isNullObject) {
        continue;
      }

      const previewRowCount = Math.min(usedRange.rowCount, 3);
      const previewColumnCount = Math.min(usedRange.columnCount, 4);
      if (previewRowCount < 1 || previewColumnCount < 1) {
        continue;
      }

      const previewRange = worksheet.getRangeByIndexes(
        usedRange.rowIndex,
        usedRange.columnIndex,
        previewRowCount,
        previewColumnCount,
      );
      previewRange.load(["address", "text", "formulas", "numberFormat"]);
      worksheetSnapshotRanges.push({
        worksheetIndex,
        previewRange,
        startRowIndex: usedRange.rowIndex,
        startColumnIndex: usedRange.columnIndex,
        rowCount: previewRowCount,
        columnCount: previewColumnCount,
      });
    }
    await context.sync();

    const worksheetSummaries = worksheets.items.slice(0, 24).map((worksheet, index) => ({
      id: worksheet.id,
      name: worksheet.name,
      position: worksheet.position + 1,
      visibility: worksheet.visibility,
      usedRange:
        usedRanges[index] && !usedRanges[index].isNullObject
          ? {
              address: usedRanges[index].address,
              rowCount: usedRanges[index].rowCount,
              columnCount: usedRanges[index].columnCount,
            }
          : undefined,
      showGridlines: supportsWorksheetView ? worksheet.showGridlines : undefined,
      showHeadings: supportsWorksheetView ? worksheet.showHeadings : undefined,
      printArea:
        supportsPageLayout && worksheetPrintAreas[index] && !worksheetPrintAreas[index].isNullObject
          ? worksheetPrintAreas[index].address
          : undefined,
      tableCount: worksheetTables[index]?.items.length ?? 0,
      chartCount: worksheetCharts[index]?.items.length ?? 0,
      pivotTableCount: worksheetPivotTables[index]?.items.length ?? 0,
      tables: worksheetTables[index]?.items.slice(0, 6).map((table) => table.name) ?? [],
      charts: worksheetCharts[index]?.items.slice(0, 6).map((chart) => chart.name) ?? [],
      pivotTables: worksheetPivotTables[index]?.items.slice(0, 6).map((pivotTable) => pivotTable.name) ?? [],
    }));
    const workbookTables = worksheets.items.flatMap((worksheet, index) =>
      (worksheetTables[index]?.items ?? []).map((table) => ({
        sheetName: worksheet.name,
        name: table.name,
        id: table.id,
      })),
    );
    const workbookCharts = worksheets.items.flatMap((worksheet, index) =>
      (worksheetCharts[index]?.items ?? []).map((chart) => ({
        sheetName: worksheet.name,
        name: chart.name,
        id: chart.id,
      })),
    );
    const workbookPivotTables = worksheets.items.flatMap((worksheet, index) =>
      (worksheetPivotTables[index]?.items ?? []).map((pivotTable) => ({
        sheetName: worksheet.name,
        name: pivotTable.name,
        id: pivotTable.id,
      })),
    );

    const preview = range.text.flat().join(" | ").trim();
    const selectionRangeCitation = buildExcelCitationRecord({
      kind: isSingleCellAddress(range.address) ? "cell" : "range",
      sheetName: activeWorksheet.name,
      address: range.address,
      text: preview || undefined,
      formula: range.rowCount === 1 && range.columnCount === 1 ? trimString(range.formulas[0]?.[0]) : undefined,
      numberFormat: range.rowCount === 1 && range.columnCount === 1 ? trimString(range.numberFormat[0]?.[0]) : undefined,
    });
    const selectionCellCitations = selectionCellRanges.map((cell) =>
      buildExcelCitationRecord({
        kind: "cell",
        sheetName: activeWorksheet.name,
        address: cell.address,
        text: firstMatrixString(cell.text),
        formula: firstMatrixString(cell.formulas),
        numberFormat: firstMatrixString(cell.numberFormat),
      }),
    );
    const workbookSheetSnapshots: ExcelWorksheetSnapshot[] = worksheetSnapshotRanges.flatMap((snapshot) => {
      const worksheetSummary = worksheetSummaries[snapshot.worksheetIndex];
      if (!worksheetSummary) {
        return [];
      }
      const previewRows = snapshot.previewRange.text
        .slice(0, 3)
        .map((row) => row.map((value) => trimString(value) ?? "").join(" | "))
        .filter(Boolean);
      const citedCells: ExcelCitationRecord[] = [];
      for (let rowOffset = 0; rowOffset < snapshot.rowCount && citedCells.length < 6; rowOffset += 1) {
        for (let columnOffset = 0; columnOffset < snapshot.columnCount && citedCells.length < 6; columnOffset += 1) {
          citedCells.push(
            buildExcelCitationRecord({
              kind: "cell",
              sheetName: worksheetSummary.name,
              address: excelCellAddress(snapshot.startRowIndex + rowOffset, snapshot.startColumnIndex + columnOffset),
              text: snapshot.previewRange.text[rowOffset]?.[columnOffset],
              formula: trimString(snapshot.previewRange.formulas[rowOffset]?.[columnOffset]),
              numberFormat: trimString(snapshot.previewRange.numberFormat[rowOffset]?.[columnOffset]),
            }),
          );
        }
      }
      return [{
        worksheetId: worksheetSummary.id,
        worksheetName: worksheetSummary.name,
        usedRange: worksheetSummary.usedRange,
        previewAddress: snapshot.previewRange.address,
        previewRows,
        citedCells,
      }];
    });
    const activeWorksheetSnapshot =
      workbookSheetSnapshots.find((snapshot) => snapshot.worksheetId === activeWorksheet.id) ?? workbookSheetSnapshots[0];

    const state: OfficeStateUpdate = {
      ...base,
      selection: {
        label: `${activeWorksheet.name}!${range.address}`,
        kind: preview ? "text" : "empty",
        imageCount: 0,
        objectCount: 0,
        textPreview: preview || undefined,
        details: [`${range.rowCount} rows`, `${range.columnCount} columns`],
      },
      capabilities: [
        "excel.workbook",
        "excel.range",
        "excel.values",
        "excel.formulas",
        "excel.format",
        "excel.citations",
        "excel.sheetSnapshots",
        ...(supportsWorksheetView ? ["excel.worksheetView"] : []),
        ...(supportsPageLayout ? ["excel.pageLayout"] : []),
      ],
    };

    const result: OfficeContextPayload = {
      summary: "",
      state,
      anchors: uniqueAnchors([
        {
          kind: "workbook",
          label: base.document.title || "Workbook",
        } as OfficeAnchor,
        {
          kind: selectionRangeCitation.anchor.kind,
          label: selectionRangeCitation.label,
          sheetName: activeWorksheet.name,
          address: range.address,
        } as OfficeAnchor,
        ...selectionCellCitations.map((citation) => citation.anchor as OfficeAnchor),
        ...worksheets.items.slice(0, 24).map((worksheet) => ({
          kind: "sheet",
          label: worksheet.name,
          sheetName: worksheet.name,
        }) as OfficeAnchor),
        ...worksheetSummaries
          .filter((worksheet) => Boolean(worksheet.usedRange?.address))
          .map((worksheet) => ({
            kind: "range",
            label: `${worksheet.name}!${worksheet.usedRange!.address}`,
            sheetName: worksheet.name,
            address: worksheet.usedRange!.address,
          }) as OfficeAnchor),
        ...namedItems.items.slice(0, 24).map((namedItem) => ({
          kind: "namedItem",
          label: namedItem.name,
          namedItemName: namedItem.name,
        }) as OfficeAnchor),
        ...workbookTables.slice(0, 24).map((table) => ({
          kind: "table",
          label: table.name,
          sheetName: table.sheetName,
          tableName: table.name,
          id: table.id,
        }) as OfficeAnchor),
        ...workbookCharts.slice(0, 24).map((chart) => ({
          kind: "chart",
          label: chart.name,
          sheetName: chart.sheetName,
          chartName: chart.name,
          id: chart.id,
        }) as OfficeAnchor),
        ...workbookPivotTables.slice(0, 24).map((pivotTable) => ({
          kind: "pivotTable",
          label: pivotTable.name,
          sheetName: pivotTable.sheetName,
          pivotTableName: pivotTable.name,
          id: pivotTable.id,
        }) as OfficeAnchor),
        ...workbookSheetSnapshots.flatMap((snapshot) => snapshot.citedCells.slice(0, 4).map((citation) => citation.anchor as OfficeAnchor)).slice(0, 24),
      ]),
      formatting: options.includeFormatting
        ? {
            selectionFormat: {
              fontName: font.name,
              fontSize: font.size,
              fontColor: font.color,
              bold: font.bold,
              italic: font.italic,
              underline: font.underline,
              fillColor: fill.color,
              horizontalAlignment: format.horizontalAlignment,
              verticalAlignment: format.verticalAlignment,
              wrapText: format.wrapText,
              rowHeight: format.rowHeight,
              columnWidth: format.columnWidth,
            },
            worksheetView: {
              worksheetId: activeWorksheet.id,
              showGridlines: supportsWorksheetView ? activeWorksheet.showGridlines : undefined,
              showHeadings: supportsWorksheetView ? activeWorksheet.showHeadings : undefined,
              printArea: supportsPageLayout && activePrintArea && !activePrintArea.isNullObject ? activePrintArea.address : undefined,
            },
          }
        : undefined,
      snippets: {
        documentStructure: {
          worksheets: worksheets.items.length,
          namedItems: namedItems.items.length,
          workbookTables: workbookTables.length,
          workbookCharts: workbookCharts.length,
          workbookPivotTables: workbookPivotTables.length,
          activeSheetTables: tables.items.length,
          activeSheetCharts: charts.items.length,
          activeSheetPivotTables: pivotTables.items.length,
        },
        workbookSheets: worksheetSummaries,
        workbookSheetSnapshots,
        activeWorksheetSnapshot,
        namedItems: namedItems.items.slice(0, 24).map((namedItem) => ({
          name: namedItem.name,
          type: namedItem.type,
        })),
        selectionRangeCitation,
        selectionCellCitations,
        textPreviewRows: range.text.slice(0, 4).map((row) => row.join(" | ")),
        formulas: range.formulas.slice(0, 4),
        numberFormats: range.numberFormat.slice(0, 4),
        workbookObjects: {
          tables: workbookTables.slice(0, 24),
          charts: workbookCharts.slice(0, 24),
          pivotTables: workbookPivotTables.slice(0, 24),
        },
        activeWorksheetObjects: {
          worksheetId: activeWorksheet.id,
          worksheetName: activeWorksheet.name,
          tables: tables.items.map((table) => table.name).slice(0, 12),
          charts: charts.items.map((chart) => chart.name).slice(0, 12),
          pivotTables: pivotTables.items.map((pivotTable) => pivotTable.name).slice(0, 12),
        },
        activeWorksheetView: {
          showGridlines: supportsWorksheetView ? activeWorksheet.showGridlines : undefined,
          showHeadings: supportsWorksheetView ? activeWorksheet.showHeadings : undefined,
          printArea: supportsPageLayout && activePrintArea && !activePrintArea.isNullObject ? activePrintArea.address : undefined,
        },
      },
    };
    result.summary = createSummary(result);
    return result;
  });

  if ((options.maxImages ?? 0) > 0) {
    const fallback = await getSelectedImageAsync().catch(() => undefined);
    if (fallback) {
      payload.visuals = [await optimizeVisual({ ...fallback, kind: "worksheet", label: "Worksheet selection snapshot" })];
      payload.summary = createSummary(payload);
    }
  }

  return payload;
}


