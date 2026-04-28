import type { OfficeHostAction } from "@pi-office/pi-office-pack/protocol";
import { inspectPowerPointPresentationBase64 } from "../../powerpoint-transform";
import { trimString } from "../shared";
import {
  applyPowerPointChartAction,
  applyPowerPointCreateChartAction,
  applyPowerPointSlideNotesAction,
  exportTargetPowerPointChartSlideAsBase64,
  exportTargetPowerPointSlideAsBase64,
  inspectCurrentPowerPointPresentationPackage,
} from "../powerpoint-helpers";

const POWERPOINT_PACKAGE_ACTIONS = new Set([
  "inspectPresentationPackage",
  "getPresentationTheme",
  "getSlideNotes",
  "readSlideNotes",
  "inspectSlideNotes",
  "getSlideCharts",
  "inspectSlideCharts",
  "readSlideCharts",
  "setSlideNotes",
  "replaceSlideNotes",
  "addSlideChart",
  "createSlideChart",
  "addChartToSlide",
  "insertSlideChart",
  "updateSlideChart",
  "setChartData",
  "updateChartData",
  "replaceChartData",
]);

export function isPowerPointPackageAction(type: string): boolean {
  return POWERPOINT_PACKAGE_ACTIONS.has(type);
}

export async function applyPowerPointPackageAction(
  action: OfficeHostAction,
  type: string,
  options: Record<string, unknown>,
): Promise<unknown> {
  if (type === "inspectPresentationPackage" || type === "getPresentationTheme") {
    const base64 = trimString(action.content) ?? trimString(options.base64);
    const packageSummary = base64
      ? await inspectPowerPointPresentationBase64(base64)
      : await inspectCurrentPowerPointPresentationPackage();
    return {
      ok: true,
      host: "powerpoint",
      action: type,
      serialization: "pptx-ooxml",
      theme: packageSummary.theme,
      notes: packageSummary.notes.slice(0, 20).map((note) => ({
        slideNumber: note.slideNumber,
        hasNotes: note.hasNotes,
        preview: note.preview,
        partName: note.partName,
      })),
      partCounts: packageSummary.partCounts,
    };
  }

  if (type === "getSlideNotes" || type === "readSlideNotes" || type === "inspectSlideNotes") {
    const slidePackage = await exportTargetPowerPointSlideAsBase64(action.target);
    const packageSummary = await inspectPowerPointPresentationBase64(slidePackage.base64);
    const note = packageSummary.notes[0];
    return {
      ok: true,
      host: "powerpoint",
      action: type,
      slideId: slidePackage.slideId,
      slideIndex: slidePackage.slideIndex,
      hasNotes: note?.hasNotes ?? false,
      notesText: note?.text ?? "",
      notesPreview: note?.preview,
      notesPartName: note?.partName,
      serialization: "pptx-ooxml",
    };
  }

  if (type === "getSlideCharts" || type === "inspectSlideCharts" || type === "readSlideCharts") {
    const slidePackage = await exportTargetPowerPointChartSlideAsBase64(action.target);
    const packageSummary = await inspectPowerPointPresentationBase64(slidePackage.base64);
    return {
      ok: true,
      host: "powerpoint",
      action: type,
      slideId: slidePackage.slideId,
      slideIndex: slidePackage.slideIndex,
      charts: packageSummary.charts.map((chart) => ({
        chartIndex: chart.chartIndex,
        chartPartName: chart.chartPartName,
        chartType: chart.chartType,
        title: chart.title,
        shapeName: chart.shapeName,
        embeddedWorkbookPartName: chart.embeddedWorkbookPartName,
        seriesCount: chart.seriesCount,
        categoryCount: chart.categoryCount,
        hasEmbeddedWorkbook: chart.hasEmbeddedWorkbook,
        series: chart.series.map((series) => ({
          index: series.index,
          name: series.name,
          categories: series.categories.slice(0, 12),
          values: series.values.slice(0, 12),
        })),
      })),
      serialization: "pptx-ooxml",
    };
  }

  if (type === "setSlideNotes" || type === "replaceSlideNotes") {
    return applyPowerPointSlideNotesAction(action, type);
  }

  if (type === "addSlideChart" || type === "createSlideChart" || type === "addChartToSlide" || type === "insertSlideChart") {
    return applyPowerPointCreateChartAction(action, type);
  }

  if (type === "updateSlideChart" || type === "setChartData" || type === "updateChartData" || type === "replaceChartData") {
    return applyPowerPointChartAction(action, type);
  }

  throw new Error(`Unsupported PowerPoint package action: ${type}`);
}
