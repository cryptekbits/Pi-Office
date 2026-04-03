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

export async function loadPowerPointMastersWithLayouts(context: PowerPoint.RequestContext): Promise<PowerPoint.SlideMaster[]> {
  const slideMasters = context.presentation.slideMasters;
  slideMasters.load("items/id,items/name");
  await context.sync();
  for (const slideMaster of slideMasters.items) {
    slideMaster.layouts.load("items/id,items/name,items/type");
  }
  await context.sync();
  return slideMasters.items;
}

export function matchesPowerPointLookup(
  value: { id?: string | undefined; name?: string | undefined; type?: string | undefined },
  lookup: { id?: string | undefined; name?: string | undefined; type?: string | undefined },
): boolean {
  if (lookup.id && value.id === lookup.id) {
    return true;
  }

  if (lookup.name && (matchesTextQuery(value.name, lookup.name) || value.id === lookup.name)) {
    return true;
  }

  if (lookup.type && (value.type ?? "").toLowerCase() === lookup.type.toLowerCase()) {
    return true;
  }

  return false;
}

export function findPowerPointSlideMaster(
  slideMasters: PowerPoint.SlideMaster[],
  lookup: { id?: string | undefined; name?: string | undefined },
): PowerPoint.SlideMaster | undefined {
  if (!lookup.id && !lookup.name) {
    return undefined;
  }

  return slideMasters.find((slideMaster) => matchesPowerPointLookup({ id: slideMaster.id, name: slideMaster.name }, lookup));
}

export function findPowerPointSlideLayout(
  slideMasters: PowerPoint.SlideMaster[],
  lookup: { id?: string | undefined; name?: string | undefined; type?: string | undefined },
  masterLookup?: { id?: string | undefined; name?: string | undefined },
): PowerPoint.SlideLayout | undefined {
  if (!lookup.id && !lookup.name && !lookup.type) {
    return undefined;
  }

  const scopedMasters = masterLookup ? [findPowerPointSlideMaster(slideMasters, masterLookup)].filter(Boolean) as PowerPoint.SlideMaster[] : slideMasters;
  for (const slideMaster of scopedMasters) {
    const layout = slideMaster.layouts.items.find((entry) =>
      matchesPowerPointLookup({ id: entry.id, name: entry.name, type: entry.type }, lookup),
    );
    if (layout) {
      return layout;
    }
  }

  return undefined;
}

export async function resolvePowerPointLayoutSelection(
  context: PowerPoint.RequestContext,
  target: OfficeAnchor | undefined,
  options: Record<string, unknown>,
): Promise<{ slideMaster?: PowerPoint.SlideMaster | undefined; layout?: PowerPoint.SlideLayout | undefined }> {
  const layoutLookup = {
    id: trimString(options.layoutId) ?? (target?.kind === "layout" ? target.id : undefined),
    name: trimString(options.layoutName) ?? (target?.kind === "layout" ? target.label ?? target.text : undefined),
    type: trimString(options.layoutType),
  };
  const slideMasterLookup = {
    id: trimString(options.slideMasterId) ?? (target?.kind === "slideMaster" ? target.id : undefined),
    name: trimString(options.slideMasterName) ?? (target?.kind === "slideMaster" ? target.label ?? target.text : undefined),
  };

  if (!layoutLookup.id && !layoutLookup.name && !layoutLookup.type && !slideMasterLookup.id && !slideMasterLookup.name) {
    return {};
  }

  const slideMasters = await loadPowerPointMastersWithLayouts(context);
  const slideMaster = findPowerPointSlideMaster(slideMasters, slideMasterLookup);
  const layout = findPowerPointSlideLayout(slideMasters, layoutLookup, slideMasterLookup);
  return { slideMaster, layout };
}

export function applyPowerPointShapeProperties(shape: PowerPoint.Shape, options: Record<string, unknown>): void {
  const left = toNumber(options.left);
  const top = toNumber(options.top);
  const width = toNumber(options.width);
  const height = toNumber(options.height);
  const rotation = toNumber(options.rotation);
  const name = trimString(options.name);
  const altTextTitle = trimString(options.altTextTitle);
  const altTextDescription = trimString(options.altTextDescription);
  const visible = toBoolean(options.visible);
  const isDecorative = toBoolean(options.isDecorative);
  const fillColor = trimString(options.fillColor);
  const fillImageBase64 = trimString(options.fillImageBase64);
  const fillTransparency = clampPercentage(options.fillTransparency);
  const clearFill = toBoolean(options.clearFill);
  const lineColor = trimString(options.lineColor);
  const lineWeight = toNumber(options.lineWeight);
  const lineTransparency = clampPercentage(options.lineTransparency);
  const lineVisible = toBoolean(options.lineVisible);
  const lineDashStyle = trimString(options.lineDashStyle);
  const lineStyle = trimString(options.lineStyle);
  const zOrder = trimString(options.zOrder);
  const hyperlinkAddress = trimString(options.hyperlinkAddress);
  const hyperlinkScreenTip = trimString(options.hyperlinkScreenTip);

  if (typeof left === "number") shape.left = left;
  if (typeof top === "number") shape.top = top;
  if (typeof width === "number") shape.width = width;
  if (typeof height === "number") shape.height = height;
  if (typeof rotation === "number") shape.rotation = rotation;
  if (name) shape.name = name;
  if (altTextTitle) shape.altTextTitle = altTextTitle;
  if (altTextDescription) shape.altTextDescription = altTextDescription;
  if (typeof visible === "boolean") shape.visible = visible;
  if (typeof isDecorative === "boolean") shape.isDecorative = isDecorative;
  if (clearFill) shape.fill.clear();
  if (fillImageBase64) shape.fill.setImage(fillImageBase64);
  if (fillColor) shape.fill.setSolidColor(fillColor);
  if (typeof fillTransparency === "number") shape.fill.transparency = fillTransparency;
  if (lineColor) shape.lineFormat.color = lineColor;
  if (typeof lineWeight === "number") shape.lineFormat.weight = lineWeight;
  if (typeof lineTransparency === "number") shape.lineFormat.transparency = lineTransparency;
  if (typeof lineVisible === "boolean") shape.lineFormat.visible = lineVisible;
  if (lineDashStyle) shape.lineFormat.dashStyle = lineDashStyle as PowerPoint.ShapeLineDashStyle;
  if (lineStyle) shape.lineFormat.style = lineStyle as PowerPoint.ShapeLineStyle;
  if (zOrder) shape.setZOrder(zOrder as PowerPoint.ShapeZOrder);
  if (hyperlinkAddress) {
    const hyperlinkOptions: PowerPoint.HyperlinkAddOptions = { address: hyperlinkAddress };
    if (hyperlinkScreenTip) {
      hyperlinkOptions.screenTip = hyperlinkScreenTip;
    }
    shape.setHyperlink(hyperlinkOptions);
  }
}

export async function finalizePowerPointShapeSelection(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  shape: PowerPoint.Shape,
  action: string,
  extra: Record<string, unknown> = {},
): Promise<Record<string, unknown>> {
  shape.load("id,name");
  slide.load("id,index");
  await context.sync();
  context.presentation.setSelectedSlides([slide.id]);
  slide.setSelectedShapes([shape.id]);
  await context.sync();
  return {
    ok: true,
    host: "powerpoint",
    action,
    slideId: slide.id,
    slideIndex: slide.index + 1,
    shapeId: shape.id,
    shapeName: shape.name,
    ...extra,
  };
}

export async function resolvePowerPointSlide(
  context: PowerPoint.RequestContext,
  target?: OfficeAnchor,
  allowSelected = true,
): Promise<PowerPoint.Slide> {
  if (target?.slideId) {
    return context.presentation.slides.getItem(target.slideId);
  }

  if (typeof target?.slideIndex === "number") {
    return context.presentation.slides.getItemAt(Math.max(0, target.slideIndex - 1));
  }

  if (allowSelected) {
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items/id");
    await context.sync();
    if (selectedSlides.items[0]) {
      return context.presentation.slides.getItem(selectedSlides.items[0].id);
    }
  }

  return context.presentation.slides.getItemAt(0);
}

export function summarizePowerPointShape(
  slide: { id: string; index: number },
  shape: { id: string; name?: string | null; type?: unknown },
): { slideId: string; slideIndex: number; shapeId: string; shapeName?: string | undefined; shapeType?: unknown; label: string } {
  return {
    slideId: slide.id,
    slideIndex: slide.index + 1,
    shapeId: shape.id,
    shapeName: shape.name ?? undefined,
    shapeType: shape.type,
    label: shape.name || `Shape ${shape.id}`,
  };
}

export async function loadSelectedPowerPointShapes(
  context: PowerPoint.RequestContext,
): Promise<{ slide?: PowerPoint.Slide | undefined; shapes: PowerPoint.Shape[] }> {
  const selectedSlides = context.presentation.getSelectedSlides();
  const selectedShapes = context.presentation.getSelectedShapes();
  selectedSlides.load("items/id,items/index");
  selectedShapes.load("items/id,items/name,items/type");
  await context.sync();
  return {
    slide: selectedSlides.items[0],
    shapes: selectedShapes.items,
  };
}

export async function findPowerPointShapeInPresentation(
  context: PowerPoint.RequestContext,
  shapeId: string,
): Promise<{ slide: PowerPoint.Slide; shape: PowerPoint.Shape } | undefined> {
  const slides = context.presentation.slides;
  slides.load("items/id,items/index");
  await context.sync();

  const candidateShapes = slides.items.map((slide) => slide.shapes.getItemOrNullObject(shapeId));
  for (const shape of candidateShapes) {
    shape.load("isNullObject,id,name,type");
  }
  await context.sync();

  const matchedIndex = candidateShapes.findIndex((shape) => !shape.isNullObject);
  if (matchedIndex < 0) {
    return undefined;
  }

  const slide = slides.items[matchedIndex];
  const shape = candidateShapes[matchedIndex];
  if (!slide || !shape) {
    return undefined;
  }

  return { slide, shape };
}

export async function resolvePowerPointShape(
  context: PowerPoint.RequestContext,
  target?: OfficeAnchor,
  allowSelected = true,
): Promise<{ slide: PowerPoint.Slide; shape: PowerPoint.Shape }> {
  const targetShapeId = trimString(target?.shapeId);
  if (targetShapeId && (target?.slideId || typeof target?.slideIndex === "number")) {
    const slide = await resolvePowerPointSlide(context, target, false);
    const shape = slide.shapes.getItemOrNullObject(targetShapeId);
    slide.load("id,index");
    shape.load("isNullObject,id,name,type");
    await context.sync();
    if (shape.isNullObject) {
      throw new Error(`Could not find PowerPoint shape ${targetShapeId} on the requested slide.`);
    }
    return { slide, shape };
  }

  if (targetShapeId) {
    const match = await findPowerPointShapeInPresentation(context, targetShapeId);
    if (!match) {
      throw new Error(`Could not find the requested PowerPoint shape: ${targetShapeId}.`);
    }
    return match;
  }

  if (allowSelected) {
    const selection = await loadSelectedPowerPointShapes(context);
    const shape = selection.shapes[0];
    const slide = selection.slide;
    if (shape && slide) {
      return { slide, shape };
    }
  }

  throw new Error("No PowerPoint shape is selected.");
}

export async function resolvePowerPointShapes(
  context: PowerPoint.RequestContext,
  target?: OfficeAnchor,
  requestedShapeIds: string[] = [],
  allowSelected = true,
): Promise<{ slide: PowerPoint.Slide; shapes: PowerPoint.Shape[] }> {
  const shapeIds = Array.from(
    new Set([target?.shapeId, ...requestedShapeIds].map((value) => trimString(value)).filter((value): value is string => Boolean(value))),
  );

  if (!shapeIds.length) {
    if (allowSelected) {
      const selection = await loadSelectedPowerPointShapes(context);
      if (selection.slide && selection.shapes.length) {
        return { slide: selection.slide, shapes: selection.shapes };
      }
    }
    throw new Error("No PowerPoint shapes are selected.");
  }

  if (target?.slideId || typeof target?.slideIndex === "number") {
    const slide = await resolvePowerPointSlide(context, target, false);
    const shapes = shapeIds.map((shapeId) => slide.shapes.getItemOrNullObject(shapeId));
    slide.load("id,index");
    for (const shape of shapes) {
      shape.load("isNullObject,id,name,type");
    }
    await context.sync();
    const missingShapeId = shapes.find((shape) => shape.isNullObject)?.id ?? shapeIds.find((_shapeId, index) => shapes[index]?.isNullObject);
    if (missingShapeId) {
      throw new Error(`Could not find PowerPoint shape ${missingShapeId} on the requested slide.`);
    }
    return { slide, shapes };
  }

  const firstMatch = await findPowerPointShapeInPresentation(context, shapeIds[0]!);
  if (!firstMatch) {
    throw new Error(`Could not find the requested PowerPoint shape: ${shapeIds[0]}.`);
  }

  if (shapeIds.length === 1) {
    return { slide: firstMatch.slide, shapes: [firstMatch.shape] };
  }

  const additionalShapes = shapeIds.slice(1).map((shapeId) => firstMatch.slide.shapes.getItemOrNullObject(shapeId));
  for (const shape of additionalShapes) {
    shape.load("isNullObject,id,name,type");
  }
  await context.sync();
  const missingShapeId = shapeIds.slice(1).find((_shapeId, index) => additionalShapes[index]?.isNullObject);
  if (missingShapeId) {
    throw new Error(
      `Could not resolve PowerPoint shape ${missingShapeId} on slide ${firstMatch.slide.index + 1}. Pass an explicit slide target when shapes live on different slides.`,
    );
  }

  return { slide: firstMatch.slide, shapes: [firstMatch.shape, ...additionalShapes] };
}

export async function selectPowerPointShapeTarget(
  target?: OfficeAnchor,
): Promise<{
  slideId: string;
  slideIndex: number;
  shapeId: string;
  shapeName?: string | undefined;
  shapeType?: string | undefined;
}> {
  return PowerPoint.run(async (context) => {
    const { slide, shape } = await resolvePowerPointShape(context, target, true);
    slide.load("id,index");
    shape.load("id,name,type");
    await context.sync();
    context.presentation.setSelectedSlides([slide.id]);
    slide.setSelectedShapes([shape.id]);
    await context.sync();
    return {
      slideId: slide.id,
      slideIndex: slide.index + 1,
      shapeId: shape.id,
      shapeName: shape.name,
      shapeType: typeof shape.type === "string" ? shape.type : undefined,
    };
  });
}

export async function applyPowerPointShapeImageAction(
  action: OfficeHostAction,
  operation: string,
  image: { data: string; mimeType: string },
): Promise<unknown> {
  const options = getActionOptions(action);
  const actionOptions = {
    ...options,
    ...action,
    fillImageBase64: image.data,
  };

  try {
    return await PowerPoint.run(async (context) => {
      const { slide, shape } = await resolvePowerPointShape(context, action.target, true);
      slide.load("id,index");
      shape.load("id,name,type");
      await context.sync();
      applyPowerPointShapeProperties(shape, actionOptions);
      await context.sync();
      return finalizePowerPointShapeSelection(context, slide, shape, operation, {
        shapeType: shape.type,
        imageMimeType: image.mimeType,
        imageUpdateMode: "shape.fill.setImage",
      });
    });
  } catch (error) {
    const selection = await selectPowerPointShapeTarget(action.target);
    await setSelectedImageAsync(image.data);
    const errorInfo = serializeOfficeRuntimeError(error);
    return {
      ok: true,
      host: "powerpoint",
      action: operation,
      ...selection,
      imageMimeType: image.mimeType,
      imageUpdateMode: "setSelectedDataAsync",
      completion: "fallback",
      fallbackStrategy: "setSelectedDataAsync",
      nativeAttempted: true,
      nativeFailure: trimString(errorInfo.message),
      warning:
        typeof errorInfo.message === "string" && errorInfo.message.trim()
          ? `Native shape image update failed, so the selection-based image replacement fallback was used: ${errorInfo.message.trim()}`
          : "Used selection-based image replacement fallback.",
    };
  }
}

export async function resolvePowerPointTable(
  context: PowerPoint.RequestContext,
  target?: OfficeAnchor,
  allowSelected = true,
): Promise<{ slide: PowerPoint.Slide; shape: PowerPoint.Shape; table: PowerPoint.Table }> {
  if (!supportsRequirementSet("PowerPointApi", "1.8")) {
    throw new Error("PowerPoint table actions require PowerPointApi 1.8.");
  }

  const resolved = await resolvePowerPointShape(context, target, allowSelected);
  if (!isPowerPointTableShape(resolved.shape.type)) {
    throw new Error("The requested PowerPoint shape is not a table.");
  }

  const table = resolved.shape.getTable();
  table.load("rowCount,columnCount,values");
  await context.sync();
  return { ...resolved, table };
}

export function summarizePowerPointSlide(slide: { id: string; index: number }): { slideId: string; slideIndex: number; label: string } {
  return {
    slideId: slide.id,
    slideIndex: slide.index + 1,
    label: `Slide ${slide.index + 1}`,
  };
}

export async function loadPowerPointSlideSummaries(
  context: PowerPoint.RequestContext,
): Promise<Array<{ slideId: string; slideIndex: number; label: string }>> {
  const slides = context.presentation.slides;
  slides.load("items/id,items/index");
  await context.sync();
  return slides.items.map((slide) => summarizePowerPointSlide(slide));
}

export function findPowerPointSlideInsertionStart(
  slides: Array<{ slideId: string; slideIndex: number }>,
  targetSlideId?: string | undefined,
): number {
  if (!targetSlideId) {
    return 0;
  }

  const targetIndex = slides.findIndex((slide) => slide.slideId === targetSlideId);
  if (targetIndex < 0) {
    throw new Error(`Could not find the requested PowerPoint target slide: ${targetSlideId}.`);
  }

  return targetIndex + 1;
}

export function sliceInsertedPowerPointSlides(
  beforeSlides: Array<{ slideId: string; slideIndex: number; label: string }>,
  afterSlides: Array<{ slideId: string; slideIndex: number; label: string }>,
  targetSlideId?: string | undefined,
  insertedCount?: number | undefined,
): Array<{ slideId: string; slideIndex: number; label: string }> {
  const count = insertedCount ?? Math.max(afterSlides.length - beforeSlides.length, 0);
  if (count <= 0) {
    return [];
  }

  const start = findPowerPointSlideInsertionStart(beforeSlides, targetSlideId);
  return afterSlides.slice(start, start + count);
}

export async function exportTargetPowerPointSlideAsBase64(
  target?: OfficeAnchor,
): Promise<{ slideId: string; slideIndex: number; base64: string }> {
  return PowerPoint.run(async (context) => {
    const slide =
      target?.shapeId
        ? (await resolvePowerPointShape(context, target, true)).slide
        : await resolvePowerPointSlide(context, target, true);
    slide.load("id,index");
    const exportResult = slide.exportAsBase64();
    await context.sync();
    return {
      slideId: slide.id,
      slideIndex: slide.index + 1,
      base64: exportResult.value,
    };
  });
}

export async function exportTargetPowerPointChartSlideAsBase64(
  target?: OfficeAnchor,
): Promise<{ slideId: string; slideIndex: number; base64: string; shapeName?: string | undefined }> {
  return PowerPoint.run(async (context) => {
    let slide: PowerPoint.Slide;
    let targetShapeName: string | undefined;

    if (target?.shapeId) {
      const resolved = await resolvePowerPointShape(context, target, true);
      slide = resolved.slide;
      resolved.shape.load("id,name,type");
      slide.load("id,index");
      const exportResult = slide.exportAsBase64();
      await context.sync();
      if (String(resolved.shape.type ?? "").toLowerCase() !== "chart") {
        throw new Error("The requested PowerPoint shape is not a chart.");
      }
      return {
        slideId: slide.id,
        slideIndex: slide.index + 1,
        base64: exportResult.value,
        shapeName: resolved.shape.name || undefined,
      };
    }

    slide = await resolvePowerPointSlide(context, target, true);
    slide.load("id,index");
    const selectedShapes = context.presentation.getSelectedShapes();
    selectedShapes.load("items/name,items/type");
    const exportResult = slide.exportAsBase64();
    await context.sync();

    const selectedChart = selectedShapes.items.find((shape) => String(shape.type ?? "").toLowerCase() === "chart");
    targetShapeName = selectedChart?.name || undefined;

    return {
      slideId: slide.id,
      slideIndex: slide.index + 1,
      base64: exportResult.value,
      shapeName: targetShapeName,
    };
  });
}

export async function exportCurrentPowerPointPresentationAsBase64(): Promise<string> {
  if (!supportsRequirementSet("PowerPointApi", "1.10")) {
    throw new Error("PowerPoint presentation-package export requires PowerPointApi 1.10.");
  }

  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();
    const exportResult = slides.exportAsBase64Presentation(slides.items.map((slide) => slide.id));
    await context.sync();
    return exportResult.value;
  });
}

export async function inspectCurrentPowerPointPresentationPackage(): Promise<Awaited<ReturnType<typeof inspectPowerPointPresentationBase64>>> {
  return inspectPowerPointPresentationBase64(await exportCurrentPowerPointPresentationAsBase64());
}

export async function replacePowerPointSlideWithSerializedPackage(
  sourceSlide: { slideId: string; slideIndex: number },
  transformedBase64: string,
  type: string,
  extras: Record<string, unknown> = {},
): Promise<unknown> {
  const { selectShapeName: rawSelectShapeName, ...resultExtras } = extras;
  const formatting = trimString(resultExtras.formatting);
  const replaceOriginal = typeof resultExtras.replaceOriginal === "boolean" ? resultExtras.replaceOriginal : true;
  const selectShapeName = trimString(rawSelectShapeName);

  return PowerPoint.run(async (context) => {
    const beforeSlides = await loadPowerPointSlideSummaries(context);
    const insertOptions: PowerPoint.InsertSlideOptions = {
      targetSlideId: sourceSlide.slideId,
    };
    if (formatting) {
      insertOptions.formatting = formatting as PowerPoint.InsertSlideFormatting;
    }

    context.presentation.insertSlidesFromBase64(transformedBase64, insertOptions);
    await context.sync();

    const insertedAfterImport = sliceInsertedPowerPointSlides(
      beforeSlides,
      await loadPowerPointSlideSummaries(context),
      sourceSlide.slideId,
      1,
    );
    const insertedSlideIds = insertedAfterImport.map((slide) => slide.slideId);
    if (!insertedSlideIds.length) {
      throw new Error(`PowerPoint serialized ${type} did not create a replacement slide.`);
    }

    if (replaceOriginal) {
      context.presentation.slides.getItem(sourceSlide.slideId).delete();
      await context.sync();
    }

    const finalSlides = await loadPowerPointSlideSummaries(context);
    const insertedSlides = finalSlides.filter((slide) => insertedSlideIds.includes(slide.slideId));
    let selectedShapeId: string | undefined;
    let selectedShapeNameResolved: string | undefined;

    context.presentation.setSelectedSlides(insertedSlides.map((slide) => slide.slideId));
    await context.sync();

    if (selectShapeName && insertedSlides[0]) {
      const insertedSlide = context.presentation.slides.getItem(insertedSlides[0].slideId);
      insertedSlide.shapes.load("items/id,items/name");
      await context.sync();
      const normalizedTargetName = selectShapeName.toLowerCase();
      const matchingShape =
        insertedSlide.shapes.items.find((shape) => (shape.name ?? "").trim().toLowerCase() === normalizedTargetName) ??
        insertedSlide.shapes.items.find((shape) => (shape.name ?? "").trim().toLowerCase().includes(normalizedTargetName));
      if (matchingShape) {
        insertedSlide.setSelectedShapes([matchingShape.id]);
        selectedShapeId = matchingShape.id;
        selectedShapeNameResolved = matchingShape.name ?? undefined;
        await context.sync();
      }
    }

    if (!selectedShapeNameResolved) {
      selectedShapeNameResolved = trimString(resultExtras.shapeName) ?? undefined;
    }
    if (!selectedShapeId) {
      selectedShapeId = trimString(resultExtras.shapeId) ?? undefined;
    }

    return {
      ok: true,
      host: "powerpoint",
      action: type,
      slideId: insertedSlides[0]?.slideId,
      slideIndex: insertedSlides[0]?.slideIndex,
      slides: insertedSlides,
      createdSlides: insertedSlides,
      deletedSlides: replaceOriginal
        ? [{ slideId: sourceSlide.slideId, slideIndex: sourceSlide.slideIndex, label: `Slide ${sourceSlide.slideIndex}` }]
        : undefined,
      sourceSlideId: sourceSlide.slideId,
      sourceSlideIndex: sourceSlide.slideIndex,
      shapeId: selectedShapeId,
      shapeName: selectedShapeNameResolved,
      formatting,
      ...resultExtras,
    };
  });
}

export async function applyPowerPointSlideNotesAction(action: OfficeHostAction, type: string): Promise<unknown> {
  const options = getActionOptions(action);
  const notesText =
    trimString(action.content) ??
    trimString(options.notesText) ??
    trimString(options.notes) ??
    trimString(options.text) ??
    "";
  const replaceOriginal = toBoolean(options.replaceOriginal) ?? true;
  const formatting = trimString(options.formatting);

  const sourceSlide = await exportTargetPowerPointSlideAsBase64(action.target);
  const transformed = await replaceSlideNotesInPowerPointPresentationBase64(sourceSlide.base64, {
    notesText,
    slideNumber: 1,
  });

  return replacePowerPointSlideWithSerializedPackage(sourceSlide, transformed.base64, type, {
    replaceOriginal,
    formatting,
    notesPartName: transformed.partName,
    notesText: truncateText(notesText, 1000),
    previousNotesText: truncateText(transformed.previousText, 1000),
    changed: transformed.changed,
    serialization: "pptx-ooxml",
  });
}

export async function applyPowerPointCreateChartAction(action: OfficeHostAction, type: string): Promise<unknown> {
  const options = getActionOptions(action);
  const sourceSlide = await exportTargetPowerPointSlideAsBase64(action.target);
  const categories = getChartValueArray(options.categories ?? action.categories);
  const series = getChartSeriesInput(options.series ?? action.series);
  const transformed = await createPowerPointChartInPresentationBase64(sourceSlide.base64, {
    chartType: trimString(options.chartType) ?? trimString(action.chartType),
    shapeName: trimString(options.shapeName) ?? trimString(action.shapeName),
    title: trimString(options.title) ?? trimString(action.content),
    categories: categories.length ? categories : undefined,
    series: series.length ? series : undefined,
    left: toNumber(options.left) ?? toNumber(options.chartLeft),
    top: toNumber(options.top) ?? toNumber(options.chartTop),
    width: toNumber(options.width) ?? toNumber(options.chartWidth),
    height: toNumber(options.height) ?? toNumber(options.chartHeight),
    showLegend: toBoolean(options.showLegend) ?? undefined,
  });

  return replacePowerPointSlideWithSerializedPackage(sourceSlide, transformed.base64, type, {
    replaceOriginal: toBoolean(options.replaceOriginal) ?? true,
    formatting: trimString(options.formatting),
    selectShapeName: transformed.shapeName,
    chartIndex: transformed.chartIndex,
    chartPartName: transformed.chartPartName,
    chartType: transformed.chartType,
    chartTitle: transformed.title,
    shapeId: transformed.shapeId,
    shapeName: transformed.shapeName,
    embeddedWorkbookPartName: transformed.embeddedWorkbookPartName,
    categoryCount: transformed.categoryCount,
    seriesCount: transformed.seriesCount,
    createdCharts: [
      {
        chartIndex: transformed.chartIndex,
        chartPartName: transformed.chartPartName,
        chartType: transformed.chartType,
        title: transformed.title,
        shapeId: transformed.shapeId,
        shapeName: transformed.shapeName,
        embeddedWorkbookPartName: transformed.embeddedWorkbookPartName,
        seriesCount: transformed.seriesCount,
        categoryCount: transformed.categoryCount,
      },
    ],
    serialization: "pptx-ooxml",
  });
}

export async function applyPowerPointChartAction(action: OfficeHostAction, type: string): Promise<unknown> {
  const options = getActionOptions(action);
  const sourceSlide = await exportTargetPowerPointChartSlideAsBase64(action.target);
  const categories = getChartValueArray(options.categories ?? action.categories);
  const series = getChartSeriesInput(options.series ?? action.series);
  const transformed = await updatePowerPointChartInPresentationBase64(sourceSlide.base64, {
    chartIndex: parsePositiveInteger(options.chartIndex ?? action.chartIndex),
    shapeName: trimString(options.shapeName) ?? sourceSlide.shapeName,
    title: trimString(options.title) ?? trimString(action.content),
    categories: categories.length ? categories : undefined,
    series: series.length ? series : undefined,
  });

  return replacePowerPointSlideWithSerializedPackage(sourceSlide, transformed.base64, type, {
    replaceOriginal: toBoolean(options.replaceOriginal) ?? true,
    formatting: trimString(options.formatting),
    selectShapeName: transformed.shapeName,
    chartIndex: transformed.chartIndex,
    chartPartName: transformed.chartPartName,
    chartType: transformed.chartType,
    chartTitle: transformed.title,
    shapeName: transformed.shapeName,
    embeddedWorkbookPartName: transformed.embeddedWorkbookPartName,
    categoryCount: transformed.categoryCount,
    seriesCount: transformed.seriesCount,
    serialization: "pptx-ooxml",
    warnings: transformed.warnings,
  });
}

export async function resolvePowerPointSourceSlides(
  context: PowerPoint.RequestContext,
  target?: OfficeAnchor,
  requestedSlideIds: string[] = [],
): Promise<Array<{ slideId: string; slideIndex: number; label: string }>> {
  if (requestedSlideIds.length) {
    const slides = context.presentation.slides;
    const requestedSlides = requestedSlideIds.map((slideId) => slides.getItem(slideId));
    for (const slide of requestedSlides) {
      slide.load("id,index");
    }
    await context.sync();
    return requestedSlides.map((slide) => summarizePowerPointSlide(slide));
  }

  if (target?.slideId || typeof target?.slideIndex === "number") {
    const slide = await resolvePowerPointSlide(context, target, false);
    slide.load("id,index");
    await context.sync();
    return [summarizePowerPointSlide(slide)];
  }

  const selectedSlides = context.presentation.getSelectedSlides();
  selectedSlides.load("items/id,items/index");
  await context.sync();
  if (selectedSlides.items.length > 0) {
    return selectedSlides.items.map((slide) => summarizePowerPointSlide(slide));
  }

  const slide = await resolvePowerPointSlide(context, target, true);
  slide.load("id,index");
  await context.sync();
  return [summarizePowerPointSlide(slide)];
}

export function getPowerPointSlideIdArray(value: unknown): string[] {
  if (Array.isArray(value)) {
    const stringValues = value.filter((entry): entry is string => typeof entry === "string" && entry.trim().length > 0);
    if (stringValues.length) {
      return stringValues;
    }

    return value
      .map((entry) => (isRecord(entry) ? trimString(entry.slideId) : undefined))
      .filter((entry): entry is string => Boolean(entry));
  }

  return [];
}

export async function resolvePowerPointInsertionIndex(
  context: PowerPoint.RequestContext,
  target: OfficeAnchor | undefined,
  options: Record<string, unknown>,
): Promise<number | undefined> {
  const explicitSlideIndex = toNumber(options.slideIndex);
  if (typeof explicitSlideIndex === "number" && Number.isInteger(explicitSlideIndex)) {
    return Math.max(0, explicitSlideIndex - 1);
  }

  const relativeSlideId = trimString(options.relativeToSlideId) ?? trimString(options.targetSlideId);
  const relativeSlideIndex = toNumber(options.relativeToSlideIndex);
  const relativeTarget =
    relativeSlideId || typeof relativeSlideIndex === "number"
      ? ({
          kind: "slide",
          slideId: relativeSlideId,
          slideIndex: typeof relativeSlideIndex === "number" ? relativeSlideIndex : undefined,
        } as OfficeAnchor)
      : target?.kind === "slide"
        ? target
        : undefined;

  if (!relativeTarget) {
    return undefined;
  }

  const relativeSlide = await resolvePowerPointSlide(context, relativeTarget, false);
  relativeSlide.load("index");
  await context.sync();
  const position = (trimString(options.position) ?? trimString(options.slidePosition) ?? trimString(options.placement) ?? "after").toLowerCase();
  return position === "before" ? Math.max(0, relativeSlide.index) : relativeSlide.index + 1;
}

export async function createPowerPointSlide(
  context: PowerPoint.RequestContext,
  target: OfficeAnchor | undefined,
  options: Record<string, unknown>,
): Promise<PowerPoint.Slide> {
  const slides = context.presentation.slides;
  const count = slides.getCount();
  await context.sync();

  const requestedLayoutName = trimString(options.layoutName) ?? (target?.kind === "layout" ? target.label ?? target.text : undefined);
  const requestedSlideMasterName =
    trimString(options.slideMasterName) ?? (target?.kind === "slideMaster" ? target.label ?? target.text : undefined);
  const requestedLayoutType = trimString(options.layoutType);
  const { slideMaster, layout } = await resolvePowerPointLayoutSelection(context, target, options);
  if ((requestedLayoutName || requestedLayoutType || target?.kind === "layout") && !layout) {
    throw new Error("Could not resolve the requested PowerPoint layout.");
  }
  if ((requestedSlideMasterName || target?.kind === "slideMaster") && !slideMaster && !trimString(options.slideMasterId)) {
    throw new Error("Could not resolve the requested PowerPoint slide master.");
  }

  const slideOptions: PowerPoint.AddSlideOptions = {};
  const layoutId = trimString(options.layoutId) ?? layout?.id;
  const slideMasterId = trimString(options.slideMasterId) ?? slideMaster?.id;
  if (layoutId) slideOptions.layoutId = layoutId;
  if (slideMasterId) slideOptions.slideMasterId = slideMasterId;
  slides.add(slideOptions);
  await context.sync();

  const slide = slides.getItemAt(count.value);
  slide.load("id,index");
  slide.layout.load("id,name,type");
  slide.slideMaster.load("id,name");
  await context.sync();

  const insertionIndex = await resolvePowerPointInsertionIndex(context, target, options);
  if (typeof insertionIndex === "number" && insertionIndex !== slide.index) {
    slide.moveTo(insertionIndex);
    slide.load("id,index");
    await context.sync();
  }

  return slide;
}

export async function loadPowerPointSlideContentSummaries(
  context: PowerPoint.RequestContext,
  slides: Array<{ slideId: string; slideIndex: number; label: string }>,
): Promise<
  Array<{
    slideId: string;
    slideIndex: number;
    label: string;
    title?: string | undefined;
    textBlocks: string[];
    combinedText?: string | undefined;
    shapeCount: number;
  }>
> {
  if (!slides.length) {
    return [];
  }

  const sourceSlides = slides.map((slide) => context.presentation.slides.getItem(slide.slideId));
  const shapeCollections = sourceSlides.map((slide) => slide.shapes);
  for (const slide of sourceSlides) {
    slide.load("id,index");
  }
  for (const shapes of shapeCollections) {
    shapes.load("items/id,items/name,items/type,items/left,items/top");
  }
  await context.sync();

  const supportsTextFrames = supportsRequirementSet("PowerPointApi", "1.10");
  const textFrames: Array<Array<PowerPoint.TextFrame | undefined>> = shapeCollections.map((shapes) =>
    shapes.items.map((shape) => (supportsTextFrames ? shape.getTextFrameOrNullObject() : undefined)),
  );
  if (supportsTextFrames) {
    for (const slideTextFrames of textFrames) {
      for (const textFrame of slideTextFrames) {
        textFrame?.load("isNullObject,hasText,textRange/text");
      }
    }
    await context.sync();
  }

  return slides.map((slide, slideIndex) => {
    const shapes = shapeCollections[slideIndex]?.items ?? [];
    const textEntries = shapes
      .map((shape, shapeIndex) => {
        const textFrame = textFrames[slideIndex]?.[shapeIndex];
        const text =
          supportsTextFrames && textFrame && !textFrame.isNullObject && textFrame.hasText
            ? truncateText(normalizeTextPreview(textFrame.textRange.text), 500)
            : undefined;

        return {
          top: shape.top,
          left: shape.left,
          text,
          fallbackLabel: trimString(shape.name),
        };
      })
      .filter((entry) => entry.text || entry.fallbackLabel)
      .sort((left, right) => left.top - right.top || left.left - right.left);

    const title = textEntries[0]?.text ?? textEntries[0]?.fallbackLabel ?? slide.label;
    const bodyBlocks = textEntries
      .map((entry) => entry.text ?? entry.fallbackLabel)
      .filter((entry): entry is string => Boolean(entry))
      .filter((entry, index) => !(index === 0 && entry === title));

    return {
      slideId: slide.slideId,
      slideIndex: slide.slideIndex,
      label: slide.label,
      title,
      textBlocks: bodyBlocks,
      combinedText: bodyBlocks.length ? bodyBlocks.join("\n") : title,
      shapeCount: shapes.length,
    };
  });
}

export function applyPowerPointTextFrameProperties(textFrame: PowerPoint.TextFrame, options: Record<string, unknown>): void {
  const autoSizeSetting = trimString(options.autoSizeSetting);
  const verticalAlignment = trimString(options.verticalAlignment);
  const wordWrap = toBoolean(options.wordWrap);
  const topMargin = toNumber(options.topMargin);
  const rightMargin = toNumber(options.rightMargin);
  const bottomMargin = toNumber(options.bottomMargin);
  const leftMargin = toNumber(options.leftMargin);

  if (autoSizeSetting) textFrame.autoSizeSetting = autoSizeSetting as PowerPoint.ShapeAutoSize;
  if (verticalAlignment) textFrame.verticalAlignment = verticalAlignment as PowerPoint.TextVerticalAlignment;
  if (typeof wordWrap === "boolean") textFrame.wordWrap = wordWrap;
  if (typeof topMargin === "number") textFrame.topMargin = topMargin;
  if (typeof rightMargin === "number") textFrame.rightMargin = rightMargin;
  if (typeof bottomMargin === "number") textFrame.bottomMargin = bottomMargin;
  if (typeof leftMargin === "number") textFrame.leftMargin = leftMargin;
}

export function applyPowerPointTableCellProperties(cell: PowerPoint.TableCell, options: Record<string, unknown>): void {
  const text = trimString(options.text);
  const fillColor = trimString(options.fillColor);
  const fontColor = trimString(options.fontColor);
  const fontName = trimString(options.fontName);
  const fontSize = toNumber(options.fontSize);
  const bold = toBoolean(options.bold);
  const italic = toBoolean(options.italic);
  const underline = trimString(options.underline);
  const horizontalAlignment = trimString(options.horizontalAlignment);
  const verticalAlignment = trimString(options.verticalAlignment);
  const indentLevel = toNumber(options.indentLevel);
  const topMargin = toNumber(options.topMargin);
  const rightMargin = toNumber(options.rightMargin);
  const bottomMargin = toNumber(options.bottomMargin);
  const leftMargin = toNumber(options.leftMargin);

  if (typeof text === "string") cell.text = text;
  if (fillColor) cell.fill.setSolidColor(fillColor);
  if (fontColor) cell.font.color = fontColor;
  if (fontName) cell.font.name = fontName;
  if (typeof fontSize === "number") cell.font.size = fontSize;
  if (typeof bold === "boolean") cell.font.bold = bold;
  if (typeof italic === "boolean") cell.font.italic = italic;
  if (underline) cell.font.underline = underline as PowerPoint.ShapeFontUnderlineStyle;
  if (horizontalAlignment) cell.horizontalAlignment = horizontalAlignment as PowerPoint.ParagraphHorizontalAlignment;
  if (verticalAlignment) cell.verticalAlignment = verticalAlignment as PowerPoint.TextVerticalAlignment;
  if (typeof indentLevel === "number") cell.indentLevel = indentLevel;
  if (typeof topMargin === "number") cell.margins.top = topMargin;
  if (typeof rightMargin === "number") cell.margins.right = rightMargin;
  if (typeof bottomMargin === "number") cell.margins.bottom = bottomMargin;
  if (typeof leftMargin === "number") cell.margins.left = leftMargin;
}

