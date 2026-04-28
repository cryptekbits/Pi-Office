import type { OfficeHostAction } from "@pi-office/pi-office-pack/protocol";
import {
  getNumberArray,
  resolvePositiveCount,
  resolveZeroBasedIndex,
  supportsRequirementSet,
  toNumber,
  toStringMatrix,
  trimString,
  truncateStringMatrix,
} from "../shared";
import {
  applyPowerPointShapeProperties,
  applyPowerPointTableCellProperties,
  finalizePowerPointShapeSelection,
  resolvePowerPointSlide,
  resolvePowerPointTable,
} from "../powerpoint-helpers";

const POWERPOINT_TABLE_ACTIONS = new Set([
  "addTable",
  "setTableValues",
  "updateTable",
  "setTableCell",
  "updateTableCell",
  "addTableRows",
  "deleteTableRows",
  "addTableColumns",
  "deleteTableColumns",
  "clearTable",
  "mergeTableCells",
  "resizeTableCell",
  "splitTableCell",
]);

export function isPowerPointTableAction(type: string): boolean {
  return POWERPOINT_TABLE_ACTIONS.has(type);
}

export async function applyPowerPointTableAction(
  context: PowerPoint.RequestContext,
  action: OfficeHostAction,
  type: string,
  options: Record<string, unknown>,
  actionOptions: Record<string, unknown>,
): Promise<unknown> {
  if (type === "addTable") {
    const slide = await resolvePowerPointSlide(context, action.target, true);
    const values = toStringMatrix(options.values ?? action.values);
    const rowCount = toNumber(options.rowCount) ?? values?.length ?? 2;
    const columnCount = toNumber(options.columnCount) ?? values?.[0]?.length ?? 2;
    const tableOptions: PowerPoint.TableAddOptions = {};
    const tableLeft = toNumber(options.left);
    const tableTop = toNumber(options.top);
    const tableWidth = toNumber(options.width);
    const tableHeight = toNumber(options.height);
    const tableStyle = trimString(options.style);
    if (typeof tableLeft === "number") tableOptions.left = tableLeft;
    if (typeof tableTop === "number") tableOptions.top = tableTop;
    if (typeof tableWidth === "number") tableOptions.width = tableWidth;
    if (typeof tableHeight === "number") tableOptions.height = tableHeight;
    if (tableStyle) tableOptions.style = tableStyle as PowerPoint.TableStyle;
    if (values) tableOptions.values = values;
    const shape = slide.shapes.addTable(rowCount, columnCount, tableOptions);
    applyPowerPointShapeProperties(shape, options);
    return finalizePowerPointShapeSelection(context, slide, shape, type);
  }

  if (type === "setTableValues" || type === "updateTable") {
    const { slide, shape, table } = await resolvePowerPointTable(context, action.target, true);
    const values = toStringMatrix(options.values ?? action.values);
    if (!values?.length) {
      throw new Error("PowerPoint setTableValues requires a matrix of cell values.");
    }

    const cells: PowerPoint.TableCell[] = [];
    for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
      const row = values[rowIndex] ?? [];
      if (rowIndex >= table.rowCount) {
        throw new Error(`PowerPoint table update exceeds the available row count (${table.rowCount}).`);
      }
      for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
        if (columnIndex >= table.columnCount) {
          throw new Error(`PowerPoint table update exceeds the available column count (${table.columnCount}).`);
        }
        const cell = table.getCellOrNullObject(rowIndex, columnIndex);
        cell.load("isNullObject");
        cells.push(cell);
      }
    }
    await context.sync();
    let cellOffset = 0;
    for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
      const row = values[rowIndex] ?? [];
      for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
        const cell = cells[cellOffset];
        cellOffset += 1;
        if (!cell || cell.isNullObject) {
          throw new Error(`Could not resolve PowerPoint table cell (${rowIndex + 1}, ${columnIndex + 1}).`);
        }
        cell.text = row[columnIndex] ?? "";
      }
    }
    await context.sync();
    return finalizePowerPointShapeSelection(context, slide, shape, type, {
      rowCount: table.rowCount,
      columnCount: table.columnCount,
      values: truncateStringMatrix(values),
    });
  }

  if (type === "setTableCell" || type === "updateTableCell") {
    const { slide, shape, table } = await resolvePowerPointTable(context, action.target, true);
    if (!supportsRequirementSet("PowerPointApi", "1.9")) {
      throw new Error("PowerPoint updateTableCell requires PowerPointApi 1.9.");
    }
    const rowIndex = resolveZeroBasedIndex(actionOptions.rowIndex, actionOptions.rowNumber);
    const columnIndex = resolveZeroBasedIndex(actionOptions.columnIndex, actionOptions.columnNumber);
    if (typeof rowIndex !== "number" || typeof columnIndex !== "number") {
      throw new Error("PowerPoint updateTableCell requires rowIndex/rowNumber and columnIndex/columnNumber.");
    }
    const cell = table.getCellOrNullObject(rowIndex, columnIndex);
    cell.load("isNullObject,rowIndex,columnIndex,rowCount,columnCount,text");
    await context.sync();
    if (cell.isNullObject) {
      throw new Error(`Could not resolve PowerPoint table cell (${rowIndex + 1}, ${columnIndex + 1}).`);
    }
    applyPowerPointTableCellProperties(cell, {
      ...actionOptions,
      text: trimString(action.content) ?? trimString(actionOptions.text),
    });
    await context.sync();
    return finalizePowerPointShapeSelection(context, slide, shape, type, {
      rowIndex: cell.rowIndex,
      columnIndex: cell.columnIndex,
      rowCount: cell.rowCount,
      columnCount: cell.columnCount,
      text: cell.text,
    });
  }

  if (type === "addTableRows") {
    if (!supportsRequirementSet("PowerPointApi", "1.9")) {
      throw new Error("PowerPoint addTableRows requires PowerPointApi 1.9.");
    }
    const { slide, shape, table } = await resolvePowerPointTable(context, action.target, true);
    const insertIndex = resolveZeroBasedIndex(actionOptions.rowIndex, actionOptions.rowNumber);
    const rowCount = resolvePositiveCount(actionOptions.rowCount);
    table.rows.add(insertIndex, rowCount);
    await context.sync();
    return finalizePowerPointShapeSelection(context, slide, shape, type, {
      insertedRowIndex: insertIndex ?? table.rowCount,
      insertedRowCount: rowCount,
    });
  }

  if (type === "deleteTableRows") {
    if (!supportsRequirementSet("PowerPointApi", "1.9")) {
      throw new Error("PowerPoint deleteTableRows requires PowerPointApi 1.9.");
    }
    const { slide, shape, table } = await resolvePowerPointTable(context, action.target, true);
    const rowIndexes = Array.from(
      new Set([
        ...getNumberArray(actionOptions.rowIndexes),
        ...getNumberArray(actionOptions.rows),
        resolveZeroBasedIndex(actionOptions.rowIndex, actionOptions.rowNumber),
      ].filter((value): value is number => typeof value === "number" && value >= 0)),
    );
    if (!rowIndexes.length) {
      throw new Error("PowerPoint deleteTableRows requires at least one row index.");
    }
    const rows = rowIndexes.map((rowIndex) => table.rows.getItemAt(rowIndex));
    table.rows.deleteRows(rows);
    await context.sync();
    return finalizePowerPointShapeSelection(context, slide, shape, type, { deletedRowIndexes: rowIndexes });
  }

  if (type === "addTableColumns") {
    if (!supportsRequirementSet("PowerPointApi", "1.9")) {
      throw new Error("PowerPoint addTableColumns requires PowerPointApi 1.9.");
    }
    const { slide, shape, table } = await resolvePowerPointTable(context, action.target, true);
    const insertIndex = resolveZeroBasedIndex(actionOptions.columnIndex, actionOptions.columnNumber);
    const columnCount = resolvePositiveCount(actionOptions.columnCount);
    table.columns.add(insertIndex, columnCount);
    await context.sync();
    return finalizePowerPointShapeSelection(context, slide, shape, type, {
      insertedColumnIndex: insertIndex ?? table.columnCount,
      insertedColumnCount: columnCount,
    });
  }

  if (type === "deleteTableColumns") {
    if (!supportsRequirementSet("PowerPointApi", "1.9")) {
      throw new Error("PowerPoint deleteTableColumns requires PowerPointApi 1.9.");
    }
    const { slide, shape, table } = await resolvePowerPointTable(context, action.target, true);
    const columnIndexes = Array.from(
      new Set([
        ...getNumberArray(actionOptions.columnIndexes),
        ...getNumberArray(actionOptions.columns),
        resolveZeroBasedIndex(actionOptions.columnIndex, actionOptions.columnNumber),
      ].filter((value): value is number => typeof value === "number" && value >= 0)),
    );
    if (!columnIndexes.length) {
      throw new Error("PowerPoint deleteTableColumns requires at least one column index.");
    }
    const columns = columnIndexes.map((columnIndex) => table.columns.getItemAt(columnIndex));
    table.columns.deleteColumns(columns);
    await context.sync();
    return finalizePowerPointShapeSelection(context, slide, shape, type, { deletedColumnIndexes: columnIndexes });
  }

  if (type === "clearTable") {
    if (!supportsRequirementSet("PowerPointApi", "1.9")) {
      throw new Error("PowerPoint clearTable requires PowerPointApi 1.9.");
    }
    const { slide, shape, table } = await resolvePowerPointTable(context, action.target, true);
    table.clear();
    await context.sync();
    return finalizePowerPointShapeSelection(context, slide, shape, type);
  }

  if (type === "mergeTableCells") {
    if (!supportsRequirementSet("PowerPointApi", "1.9")) {
      throw new Error("PowerPoint mergeTableCells requires PowerPointApi 1.9.");
    }
    const { slide, shape, table } = await resolvePowerPointTable(context, action.target, true);
    const rowIndex = resolveZeroBasedIndex(actionOptions.rowIndex, actionOptions.rowNumber);
    const columnIndex = resolveZeroBasedIndex(actionOptions.columnIndex, actionOptions.columnNumber);
    if (typeof rowIndex !== "number" || typeof columnIndex !== "number") {
      throw new Error("PowerPoint mergeTableCells requires rowIndex/rowNumber and columnIndex/columnNumber.");
    }
    const rowCount = resolvePositiveCount(actionOptions.rowCount);
    const columnCount = resolvePositiveCount(actionOptions.columnCount);
    table.mergeCells(rowIndex, columnIndex, rowCount, columnCount);
    await context.sync();
    return finalizePowerPointShapeSelection(context, slide, shape, type, {
      rowIndex,
      columnIndex,
      rowCount,
      columnCount,
    });
  }

  if (type === "resizeTableCell" || type === "splitTableCell") {
    if (!supportsRequirementSet("PowerPointApi", "1.9")) {
      throw new Error(`PowerPoint ${type} requires PowerPointApi 1.9.`);
    }
    const { slide, shape, table } = await resolvePowerPointTable(context, action.target, true);
    const rowIndex = resolveZeroBasedIndex(actionOptions.rowIndex, actionOptions.rowNumber);
    const columnIndex = resolveZeroBasedIndex(actionOptions.columnIndex, actionOptions.columnNumber);
    if (typeof rowIndex !== "number" || typeof columnIndex !== "number") {
      throw new Error(`PowerPoint ${type} requires rowIndex/rowNumber and columnIndex/columnNumber.`);
    }
    const cell = table.getCellOrNullObject(rowIndex, columnIndex);
    cell.load("isNullObject,rowIndex,columnIndex");
    await context.sync();
    if (cell.isNullObject) {
      throw new Error(`Could not resolve PowerPoint table cell (${rowIndex + 1}, ${columnIndex + 1}).`);
    }
    const rowCount = resolvePositiveCount(actionOptions.rowCount);
    const columnCount = resolvePositiveCount(actionOptions.columnCount);
    if (type === "resizeTableCell") {
      cell.resize(rowCount, columnCount);
    } else {
      cell.split(rowCount, columnCount);
    }
    await context.sync();
    return finalizePowerPointShapeSelection(context, slide, shape, type, {
      rowIndex: cell.rowIndex,
      columnIndex: cell.columnIndex,
      rowCount,
      columnCount,
    });
  }

  throw new Error(`Unsupported PowerPoint table action: ${type}`);
}
