import type { OfficeHostAction } from "@pi-office/pi-office-pack/protocol";
import { getActionOptions, trimString } from "./shared";
import {
  applyExcelCreateChartAction,
  applyExcelExtractChartXmlAction,
  applyExcelUpdateChartAction,
} from "./excel-actions/charts";
import { applyExcelMediaAction, isExcelMediaAction } from "./excel-actions/media";
import {
  applyExcelCreatePivotTableAction,
  applyExcelPivotSortAction,
  applyExcelRefreshPivotTableAction,
  applyExcelUpdatePivotTableAction,
} from "./excel-actions/pivots";
import { applyExcelRangeAction, isExcelRangeAction } from "./excel-actions/ranges";
import { applyExcelTableAction, isExcelTableAction } from "./excel-actions/tables";
import { applyExcelWorksheetAction, isExcelWorksheetAction } from "./excel-actions/worksheets";

export async function applyExcelAction(action: OfficeHostAction): Promise<unknown> {
  return Excel.run(async (context) => {
    const type = trimString(action.type) ?? "setRangeValues";
    const options = getActionOptions(action);

    if (isExcelRangeAction(type)) {
      return applyExcelRangeAction(context, action, type, options);
    }

    if (isExcelWorksheetAction(type)) {
      return applyExcelWorksheetAction(context, action, type, options);
    }

    if (isExcelTableAction(type)) {
      return applyExcelTableAction(context, action, type, options);
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

    if (isExcelMediaAction(type)) {
      return applyExcelMediaAction(context, action, type);
    }

    throw new Error(`Unsupported Excel action: ${type}`);
  });
}
