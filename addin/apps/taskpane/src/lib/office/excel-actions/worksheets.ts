import type { OfficeAnchor, OfficeHostAction } from "@pi-office/pi-office-pack/protocol";
import { getActionOptions, splitSheetAddress, supportsRequirementSet, toBoolean, trimString } from "../shared";
import { resolveExcelRange, resolveExcelWorksheet } from "../excel-targets";

export function isExcelWorksheetAction(type: string): boolean {
  return [
    "createWorksheet",
    "renameWorksheet",
    "duplicateWorksheet",
    "deleteWorksheet",
    "setWorksheetGridlines",
    "setWorksheetHeadings",
    "setPrintArea",
  ].includes(type);
}

export async function applyExcelWorksheetAction(
  context: Excel.RequestContext,
  action: OfficeHostAction,
  type: string,
  options = getActionOptions(action),
): Promise<unknown> {
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

  throw new Error(`Unsupported Excel worksheet action: ${type}`);
}
