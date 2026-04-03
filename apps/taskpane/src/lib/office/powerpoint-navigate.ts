import type { OfficeAnchor, OfficeContextPayload, OfficeHostAction, OfficeStateUpdate, OfficeSelectionMeta, OfficeVisualSnapshot } from "@pi-office/pi-office-pack/protocol";
import type { OfficeCaptureOptions } from "../office-host-adapter-types";
import {
  createPowerPointChartInPresentationBase64,
  inspectPowerPointPresentationBase64,
  replaceSlideNotesInPowerPointPresentationBase64,
  updatePowerPointChartInPresentationBase64,
} from "../powerpoint-transform";
import {
  supportsRequirementSet,
  normalizeTextPreview,
  truncateText,
  pluralize,
  uniqueDetails,
  formatPoints,
  buildSelectionSummary,
  isRecord,
  trimString,
  truncateLabel,
  matchesTextQuery,
  parsePositiveInteger,
  clampPercentage,
  uniqueAnchors,
  getActionOptions,
  toNumber,
  toBoolean,
  toStringMatrix,
  isPowerPointImageShape,
  isPowerPointTableShape,
  isPowerPointGroupShape,
  getPowerPointShapeContentKind,
  truncateStringMatrix,
  resolveZeroBasedIndex,
  resolvePositiveCount,
  getSelectedTextAsync,
  setSelectedTextAsync,
  setSelectedImageAsync,
  getActionImagePayload,
  getSelectedImageAsync,
  optimizeVisual,
  createSummary,
  serializeOfficeRuntimeError,
  getStringArray,
  getNumberArray,
  getRecordArray,
  getChartValueArray,
  getChartSeriesInput,
} from "./shared";
import {
  loadPowerPointMastersWithLayouts,
  matchesPowerPointLookup,
  findPowerPointSlideMaster,
  findPowerPointSlideLayout,
  resolvePowerPointLayoutSelection,
  applyPowerPointShapeProperties,
  finalizePowerPointShapeSelection,
  resolvePowerPointSlide,
  summarizePowerPointShape,
  loadSelectedPowerPointShapes,
  findPowerPointShapeInPresentation,
  resolvePowerPointShape,
  resolvePowerPointShapes,
  selectPowerPointShapeTarget,
  applyPowerPointShapeImageAction,
  resolvePowerPointTable,
  summarizePowerPointSlide,
  loadPowerPointSlideSummaries,
  findPowerPointSlideInsertionStart,
  sliceInsertedPowerPointSlides,
  exportTargetPowerPointSlideAsBase64,
  exportTargetPowerPointChartSlideAsBase64,
  exportCurrentPowerPointPresentationAsBase64,
  inspectCurrentPowerPointPresentationPackage,
  replacePowerPointSlideWithSerializedPackage,
  applyPowerPointSlideNotesAction,
  applyPowerPointCreateChartAction,
  applyPowerPointChartAction,
  resolvePowerPointSourceSlides,
  getPowerPointSlideIdArray,
  resolvePowerPointInsertionIndex,
  createPowerPointSlide,
  loadPowerPointSlideContentSummaries,
  applyPowerPointTextFrameProperties,
  applyPowerPointTableCellProperties,
} from "./powerpoint-helpers";

export async function navigatePowerPointAnchor(anchor: OfficeAnchor): Promise<unknown> {
  return PowerPoint.run(async (context) => {
    if (anchor.kind === "notesRegion") {
      const slide = await resolvePowerPointSlide(context, anchor, true);
      slide.load("id,index");
      const slideExport = supportsRequirementSet("PowerPointApi", "1.10") ? slide.exportAsBase64() : undefined;
      await context.sync();
      context.presentation.setSelectedSlides([slide.id]);
      await context.sync();

      let hasNotes = false;
      let notesPreview: string | undefined;
      let notesText: string | undefined;
      let notesPartName: string | undefined;
      if (slideExport?.value) {
        try {
          const packageSummary = await inspectPowerPointPresentationBase64(slideExport.value);
          const note = packageSummary.notes[0];
          hasNotes = note?.hasNotes ?? false;
          notesPreview = note?.preview;
          notesText = truncateText(note?.text, 800);
          notesPartName = note?.partName;
        } catch {
          // Ignore serialization failures during navigation. Slide selection still succeeds.
        }
      }

      return {
        ok: true,
        host: "powerpoint",
        anchorKind: "notesRegion",
        completion: "partial",
        fallbackStrategy: "slideSelection",
        nativeAttempted: true,
        nativeFailure: "PowerPoint does not expose direct speaker-notes pane selection in edit view.",
        slideId: slide.id,
        slideIndex: slide.index + 1,
        hasNotes,
        notesPreview,
        notesText,
        notesPartName,
        warnings: ["Selected the target slide because PowerPoint does not expose direct speaker-notes pane selection in edit view."],
      };
    }

    if (anchor.kind === "layout" || anchor.kind === "slideMaster") {
      const presentationSlides = context.presentation.slides;
      presentationSlides.load("items/id,items/index");
      await context.sync();
      for (const slide of presentationSlides.items) {
        slide.layout.load("id,name,type");
        slide.slideMaster.load("id,name");
      }
      await context.sync();

      const slide = presentationSlides.items.find((entry) =>
        anchor.kind === "layout"
          ? matchesPowerPointLookup(
              { id: entry.layout.id, name: entry.layout.name, type: entry.layout.type },
              { id: anchor.id, name: anchor.label ?? anchor.text },
            )
          : matchesPowerPointLookup({ id: entry.slideMaster.id, name: entry.slideMaster.name }, { id: anchor.id, name: anchor.label ?? anchor.text }),
      );
      if (slide) {
        context.presentation.setSelectedSlides([slide.id]);
        await context.sync();
        return {
          ok: true,
          host: "powerpoint",
          anchorKind: anchor.kind,
          completion: "partial",
          fallbackStrategy: "slideSelection",
          nativeAttempted: true,
          nativeFailure: `PowerPoint does not expose direct ${anchor.kind} selection in edit view.`,
          slideId: slide.id,
          slideIndex: slide.index + 1,
          layoutId: slide.layout.id,
          layoutName: slide.layout.name,
          slideMasterId: slide.slideMaster.id,
          slideMasterName: slide.slideMaster.name,
          warnings: [`Selected a slide using the requested ${anchor.kind} because PowerPoint does not expose direct layout/master selection in edit view.`],
        };
      }
    }

    if (anchor.kind === "shape" && anchor.shapeId) {
      const { slide, shape } = await resolvePowerPointShape(context, anchor, true);
      slide.load("id,index");
      shape.load("id,name");
      await context.sync();
      context.presentation.setSelectedSlides([slide.id]);
      slide.setSelectedShapes([shape.id]);
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        anchorKind: "shape",
        slideId: slide.id,
        slideIndex: slide.index + 1,
        shapeId: shape.id,
        shapeName: shape.name,
      };
    }

    const slide = await resolvePowerPointSlide(context, anchor, true);
    slide.load("id,index");
    await context.sync();
    context.presentation.setSelectedSlides([slide.id]);
    await context.sync();
    return { ok: true, host: "powerpoint", anchorKind: "slide", slideId: slide.id, slideIndex: slide.index + 1 };
  });
}

