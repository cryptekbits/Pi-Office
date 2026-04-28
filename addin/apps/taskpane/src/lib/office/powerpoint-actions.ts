import type { OfficeContextPayload, OfficeHostAction, OfficeStateUpdate, OfficeSelectionMeta, OfficeVisualSnapshot } from "@pi-office/pi-office-pack/protocol";
import type { OfficeCaptureOptions } from "../office-host-adapter-types";
import { inspectPowerPointPresentationBase64 } from "../powerpoint-transform";
import {
  supportsRequirementSet,
  normalizeTextPreview,
  truncateText,
  pluralize,
  uniqueDetails,
  formatPoints,
  buildSelectionSummary,
  trimString,
  truncateLabel,
  matchesTextQuery,
  clampPercentage,
  uniqueAnchors,
  getActionOptions,
  toNumber,
  isPowerPointImageShape,
  isPowerPointTableShape,
  isPowerPointGroupShape,
  getPowerPointShapeContentKind,
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
} from "./shared";
import {
  loadPowerPointMastersWithLayouts,
  resolvePowerPointLayoutSelection,
  applyPowerPointShapeProperties,
  finalizePowerPointShapeSelection,
  resolvePowerPointSlide,
  summarizePowerPointShape,
  loadSelectedPowerPointShapes,
  findPowerPointShapeInPresentation,
  resolvePowerPointShape,
  resolvePowerPointShapes,
  applyPowerPointShapeImageAction,
  summarizePowerPointSlide,
  loadPowerPointSlideSummaries,
  findPowerPointSlideInsertionStart,
  sliceInsertedPowerPointSlides,
  exportTargetPowerPointSlideAsBase64,
  exportTargetPowerPointChartSlideAsBase64,
  inspectCurrentPowerPointPresentationPackage,
  applyPowerPointSlideNotesAction,
  applyPowerPointCreateChartAction,
  applyPowerPointChartAction,
  loadPowerPointSlideContentSummaries,
  applyPowerPointTextFrameProperties,
} from "./powerpoint-helpers";
import { applyPowerPointMediaAction, isPowerPointMediaAction } from "./powerpoint-actions/media";
import { applyPowerPointSlideStructureAction, isPowerPointSlideStructureAction } from "./powerpoint-actions/slide-structure";
import { applyPowerPointTableAction, isPowerPointTableAction } from "./powerpoint-actions/tables";

export async function applyPowerPointAction(action: OfficeHostAction): Promise<unknown> {
  const type = trimString(action.type) ?? "insertText";
  const options = getActionOptions(action);
  const imagePayload = getActionImagePayload(action, options);
  const isShapeImageAction =
    type === "replaceShapeImage" ||
    type === "setShapeImage" ||
    type === "updateShapeImage" ||
    ((type === "setShapeProperties" || type === "updateShapeProperties") && Boolean(imagePayload));

  if (type === "insertText" && !action.target?.shapeId) {
    try {
      return await PowerPoint.run(async (context) => {
        const textRange = context.presentation.getSelectedTextRangeOrNullObject();
        textRange.load("isNullObject,text");
        await context.sync();
        if (textRange.isNullObject) {
          throw new Error("No PowerPoint text range is selected.");
        }
        textRange.text = action.content ?? "";
        await context.sync();
        return { ok: true, host: "powerpoint", action: type };
      });
    } catch (error) {
      const errorInfo = serializeOfficeRuntimeError(error);
      await setSelectedTextAsync(action.content ?? "");
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        completion: "fallback",
        fallbackStrategy: "setSelectedDataAsync",
        nativeAttempted: true,
        nativeFailure: trimString(errorInfo.message),
        warning:
          typeof errorInfo.message === "string" && errorInfo.message.trim()
            ? `Native PowerPoint text-range update failed, so the selection data fallback was used: ${errorInfo.message.trim()}`
            : "Used selection-based PowerPoint text insertion fallback.",
      };
    }
  }

  if (isShapeImageAction) {
    if (!imagePayload) {
      throw new Error("PowerPoint shape image actions require a base64 image payload.");
    }
    return applyPowerPointShapeImageAction(action, type, imagePayload);
  }

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

  if (isPowerPointMediaAction(type)) {
    return applyPowerPointMediaAction(action, type, options);
  }

  return PowerPoint.run(async (context) => {
    const actionOptions = { ...options, ...action };

    if (type === "getPresentationStructure" || type === "readPresentationStructure") {
      const maxSlides = resolvePositiveCount(actionOptions.maxSlides, 20);
      const includeSlideText = actionOptions.includeSlideText !== false;
      const slides = await loadPowerPointSlideSummaries(context);
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load("items/id,items/index");
      await context.sync();

      const slideContent = await loadPowerPointSlideContentSummaries(context, slides.slice(0, maxSlides));
      const supportsLayoutMetadata = supportsRequirementSet("PowerPointApi", "1.3");
      const slideMasters = supportsLayoutMetadata ? await loadPowerPointMastersWithLayouts(context) : [];

      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideCount: slides.length,
        selectedSlideCount: selectedSlides.items.length,
        selectedSlideIds: selectedSlides.items.map((slide) => slide.id),
        slides,
        slidePreviews: slideContent.map((slide) => ({
          slideId: slide.slideId,
          slideIndex: slide.slideIndex,
          label: slide.label,
          title: slide.title,
          shapeCount: slide.shapeCount,
          textPreview: includeSlideText ? truncateText(slide.combinedText, 400) : undefined,
        })),
        slideMasters: slideMasters.map((slideMaster) => ({
          id: slideMaster.id,
          name: slideMaster.name,
          layoutCount: slideMaster.layouts.items.length,
          layouts: slideMaster.layouts.items.map((layout) => ({
            id: layout.id,
            name: layout.name,
            type: layout.type,
          })),
        })),
      };
    }

    if (type === "getSlide" || type === "readSlide") {
      const includeShapes = actionOptions.includeShapes !== false;
      const includeSlideText = actionOptions.includeSlideText !== false;
      const slide = await resolvePowerPointSlide(context, action.target, true);
      slide.load("id,index");
      const supportsLayoutMetadata = supportsRequirementSet("PowerPointApi", "1.3");
      if (supportsLayoutMetadata) {
        slide.layout.load("id,name,type");
        slide.slideMaster.load("id,name");
      }
      const shapes = slide.shapes;
      if (includeShapes) {
        shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height,items/rotation");
      }
      await context.sync();

      const slideSummary = summarizePowerPointSlide(slide);
      const contentSummary = includeSlideText
        ? (await loadPowerPointSlideContentSummaries(context, [slideSummary]))[0]
        : undefined;
      const listedShapes = includeShapes
        ? shapes.items.map((shape) => ({
            ...summarizePowerPointShape(slide, shape),
            contentKind: getPowerPointShapeContentKind(String(shape.type ?? "shape")),
            left: formatPoints(shape.left),
            top: formatPoints(shape.top),
            width: formatPoints(shape.width),
            height: formatPoints(shape.height),
            rotation: shape.rotation,
          }))
        : undefined;

      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        label: slideSummary.label,
        layoutId: supportsLayoutMetadata ? slide.layout.id : undefined,
        layoutName: supportsLayoutMetadata ? slide.layout.name : undefined,
        layoutType: supportsLayoutMetadata ? slide.layout.type : undefined,
        slideMasterId: supportsLayoutMetadata ? slide.slideMaster.id : undefined,
        slideMasterName: supportsLayoutMetadata ? slide.slideMaster.name : undefined,
        title: contentSummary?.title,
        textPreview: includeSlideText ? truncateText(contentSummary?.combinedText, 500) : undefined,
        shapeCount: includeShapes ? shapes.items.length : contentSummary?.shapeCount,
        shapes: listedShapes,
      };
    }

    if (type === "listSlideShapes" || type === "getSlideShapes") {
      const maxShapes = resolvePositiveCount(actionOptions.maxShapes, 200);
      const slide = await resolvePowerPointSlide(context, action.target, true);
      const shapes = slide.shapes;
      slide.load("id,index");
      shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height,items/rotation");
      await context.sync();

      const listedShapes = shapes.items.slice(0, maxShapes).map((shape) => ({
        ...summarizePowerPointShape(slide, shape),
        contentKind: getPowerPointShapeContentKind(String(shape.type ?? "shape")),
        left: formatPoints(shape.left),
        top: formatPoints(shape.top),
        width: formatPoints(shape.width),
        height: formatPoints(shape.height),
        rotation: shape.rotation,
      }));

      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        shapeCount: shapes.items.length,
        returnedShapeCount: listedShapes.length,
        maxShapes,
        shapes: listedShapes,
      };
    }

    if ((type === "insertText" && action.target?.shapeId) || type === "setShapeText" || type === "clearShapeText") {
      const { slide, shape } = await resolvePowerPointShape(context, action.target, true);
      const textFrame = shape.getTextFrameOrNullObject();
      shape.load("id,name,type");
      textFrame.load("isNullObject,textRange/text");
      await context.sync();
      if (textFrame.isNullObject) {
        throw new Error("The requested PowerPoint shape does not contain text.");
      }
      if (type === "clearShapeText") {
        textFrame.deleteText();
      } else {
        const placement = trimString(action.placement) ?? trimString(options.placement);
        const nextText = trimString(action.content) ?? trimString(options.text) ?? "";
        textFrame.textRange.text = placement === "after" ? `${textFrame.textRange.text}${nextText}` : nextText;
      }
      applyPowerPointTextFrameProperties(textFrame, actionOptions);
      return finalizePowerPointShapeSelection(context, slide, shape, type);
    }

    if (type === "setShapeProperties" || type === "updateShapeProperties" || type === "moveShape" || type === "resizeShape") {
      const { slide, shape } = await resolvePowerPointShape(context, action.target, true);
      applyPowerPointShapeProperties(shape, actionOptions);
      return finalizePowerPointShapeSelection(context, slide, shape, type);
    }

    if (type === "deleteShape" || type === "deleteShapes") {
      const { slide, shapes } = await resolvePowerPointShapes(context, action.target, getStringArray(options.shapeIds), true);
      slide.load("id,index");
      for (const shape of shapes) {
        shape.load("id,name,type");
      }
      await context.sync();
      const deletedShapes = shapes.map((shape) => summarizePowerPointShape(slide, shape));
      for (const shape of shapes) {
        shape.delete();
      }
      await context.sync();
      context.presentation.setSelectedSlides([slide.id]);
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        deletedCount: deletedShapes.length,
        deletedShapeIds: deletedShapes.map((shape) => shape.shapeId),
        deletedShapes,
      };
    }

    if (type === "ungroupShape" || type === "ungroupShapes") {
      if (!supportsRequirementSet("PowerPointApi", "1.8")) {
        throw new Error("PowerPoint ungroupShape requires PowerPointApi 1.8.");
      }

      const { slide, shapes } = await resolvePowerPointShapes(context, action.target, getStringArray(options.shapeIds), true);
      slide.load("id,index");
      for (const shape of shapes) {
        shape.load("id,name,type");
      }
      await context.sync();
      const nonGroupShape = shapes.find((shape) => !isPowerPointGroupShape(shape.type));
      if (nonGroupShape) {
        throw new Error(`PowerPoint ungroupShape requires grouped shapes. ${nonGroupShape.name || nonGroupShape.id} is not a group.`);
      }

      const groups = shapes.map((shape) => shape.group);
      for (const group of groups) {
        group.shapes.load("items/id,items/name,items/type");
      }
      await context.sync();

      const deletedShapes = shapes.map((shape) => summarizePowerPointShape(slide, shape));
      const createdShapes = groups.flatMap((group) => group.shapes.items.map((shape) => summarizePowerPointShape(slide, shape)));
      for (const group of groups) {
        group.ungroup();
      }
      await context.sync();
      context.presentation.setSelectedSlides([slide.id]);
      if (createdShapes.length) {
        slide.setSelectedShapes(createdShapes.map((shape) => shape.shapeId));
      }
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        deletedShapes,
        createdShapes,
      };
    }

    if (isPowerPointSlideStructureAction(type)) {
      return applyPowerPointSlideStructureAction(context, action, type, options);
    }

    if (type === "addTextBox") {
      const slide = await resolvePowerPointSlide(context, action.target, true);
      const shapeOptions: PowerPoint.ShapeAddOptions = {};
      const left = toNumber(options.left);
      const top = toNumber(options.top);
      const width = toNumber(options.width);
      const height = toNumber(options.height);
      if (typeof left === "number") shapeOptions.left = left;
      if (typeof top === "number") shapeOptions.top = top;
      if (typeof width === "number") shapeOptions.width = width;
      if (typeof height === "number") shapeOptions.height = height;
      const shape = slide.shapes.addTextBox(action.content ?? "", shapeOptions);
      applyPowerPointShapeProperties(shape, options);
      return finalizePowerPointShapeSelection(context, slide, shape, type);
    }

    if (type === "addGeometricShape") {
      const slide = await resolvePowerPointSlide(context, action.target, true);
      const shapeOptions: PowerPoint.ShapeAddOptions = {};
      const left = toNumber(options.left);
      const top = toNumber(options.top);
      const width = toNumber(options.width);
      const height = toNumber(options.height);
      const geometricShapeType = trimString(options.geometricShapeType) ?? trimString(options.shapeType) ?? trimString(options.type);
      if (!geometricShapeType) {
        throw new Error("PowerPoint addGeometricShape requires a geometricShapeType.");
      }
      if (typeof left === "number") shapeOptions.left = left;
      if (typeof top === "number") shapeOptions.top = top;
      if (typeof width === "number") shapeOptions.width = width;
      if (typeof height === "number") shapeOptions.height = height;
      const shape = slide.shapes.addGeometricShape(geometricShapeType as PowerPoint.GeometricShapeType, shapeOptions);
      applyPowerPointShapeProperties(shape, options);
      return finalizePowerPointShapeSelection(context, slide, shape, type, { geometricShapeType });
    }

    if (type === "groupShapes") {
      const slide = await resolvePowerPointSlide(context, action.target, true);
      const shapeIds = Array.from(
        new Set([action.target?.shapeId, ...getStringArray(options.shapeIds), ...getStringArray(options.additionalShapeIds)].filter(
          (value): value is string => Boolean(value),
        )),
      );
      if (shapeIds.length < 2) {
        throw new Error("PowerPoint groupShapes requires at least two shape IDs.");
      }
      const shape = slide.shapes.addGroup(shapeIds);
      applyPowerPointShapeProperties(shape, options);
      return finalizePowerPointShapeSelection(context, slide, shape, type, { groupedShapeIds: shapeIds });
    }

    if (isPowerPointTableAction(type)) {
      return applyPowerPointTableAction(context, action, type, options, actionOptions);
    }

    if (type === "addLine") {
      const slide = await resolvePowerPointSlide(context, action.target, true);
      const lineOptions: PowerPoint.ShapeAddOptions = {};
      const lineLeft = toNumber(options.left);
      const lineTop = toNumber(options.top);
      const lineWidth = toNumber(options.width);
      const lineHeight = toNumber(options.height);
      if (typeof lineLeft === "number") lineOptions.left = lineLeft;
      if (typeof lineTop === "number") lineOptions.top = lineTop;
      if (typeof lineWidth === "number") lineOptions.width = lineWidth;
      if (typeof lineHeight === "number") lineOptions.height = lineHeight;
      const shape = slide.shapes.addLine(trimString(options.connectorType) as PowerPoint.ConnectorType, lineOptions);
      applyPowerPointShapeProperties(shape, options);
      return finalizePowerPointShapeSelection(context, slide, shape, type);
    }

    if (type === "addProcessFlow" || type === "addSimpleDiagram") {
      const slide = await resolvePowerPointSlide(context, action.target, true);
      slide.load("id,index");
      const stepTexts = (
        Array.isArray(actionOptions.steps)
          ? actionOptions.steps
          : typeof action.content === "string"
            ? action.content
                .split(/\r?\n|>/)
                .map((entry) => trimString(entry))
                .filter((entry): entry is string => Boolean(entry))
            : []
      )
        .map((entry) => trimString(entry))
        .filter((entry): entry is string => Boolean(entry));
      if (!stepTexts.length) {
        throw new Error("PowerPoint addProcessFlow requires steps or newline-delimited content.");
      }

      const orientation = (trimString(actionOptions.orientation) ?? "horizontal").toLowerCase();
      const stepWidth = toNumber(actionOptions.stepWidth) ?? 140;
      const stepHeight = toNumber(actionOptions.stepHeight) ?? 72;
      const spacing = toNumber(actionOptions.spacing) ?? 36;
      const startLeft = toNumber(actionOptions.left) ?? 60;
      const startTop = toNumber(actionOptions.top) ?? 140;
      const connectorType = (trimString(actionOptions.connectorType) ?? "Straight") as PowerPoint.ConnectorType;
      const geometricShapeType = (trimString(actionOptions.geometricShapeType) ?? "FlowChartProcess") as PowerPoint.GeometricShapeType;
      const stepShapes: PowerPoint.Shape[] = [];
      const connectorShapes: PowerPoint.Shape[] = [];
      const textFrames: PowerPoint.TextFrame[] = [];
      const baseName = trimString(actionOptions.name) ?? (type === "addSimpleDiagram" ? "Diagram Step" : "Process Step");

      for (let index = 0; index < stepTexts.length; index += 1) {
        const left = orientation === "vertical" ? startLeft : startLeft + index * (stepWidth + spacing);
        const top = orientation === "vertical" ? startTop + index * (stepHeight + spacing) : startTop;
        const stepShape = slide.shapes.addGeometricShape(geometricShapeType, {
          left,
          top,
          width: stepWidth,
          height: stepHeight,
        });
        applyPowerPointShapeProperties(stepShape, {
          ...actionOptions,
          left,
          top,
          width: stepWidth,
          height: stepHeight,
          name: `${baseName} ${index + 1}`,
        });
        stepShapes.push(stepShape);
        textFrames.push(stepShape.getTextFrameOrNullObject());

        if (index < stepTexts.length - 1) {
          const connectorLeft = orientation === "vertical" ? startLeft + stepWidth / 2 : left + stepWidth;
          const connectorTop = orientation === "vertical" ? top + stepHeight : startTop + stepHeight / 2;
          const connectorWidth = orientation === "vertical" ? 0 : spacing;
          const connectorHeight = orientation === "vertical" ? spacing : 0;
          const connectorShape = slide.shapes.addLine(connectorType, {
            left: connectorLeft,
            top: connectorTop,
            width: connectorWidth,
            height: connectorHeight,
          });
          applyPowerPointShapeProperties(connectorShape, {
            lineColor: trimString(actionOptions.connectorColor) ?? trimString(actionOptions.lineColor),
            lineWeight: toNumber(actionOptions.connectorWeight) ?? toNumber(actionOptions.lineWeight),
            lineTransparency: actionOptions.lineTransparency,
            lineDashStyle: trimString(actionOptions.lineDashStyle),
            lineStyle: trimString(actionOptions.lineStyle),
          });
          connectorShapes.push(connectorShape);
        }
      }

      for (const textFrame of textFrames) {
        textFrame.load("isNullObject");
      }
      await context.sync();
      for (let index = 0; index < textFrames.length; index += 1) {
        const textFrame = textFrames[index];
        if (!textFrame || textFrame.isNullObject) {
          continue;
        }
        textFrame.textRange.text = stepTexts[index] ?? "";
        applyPowerPointTextFrameProperties(textFrame, actionOptions);
      }
      await context.sync();

      for (const shape of [...stepShapes, ...connectorShapes]) {
        shape.load("id,name,type");
      }
      await context.sync();
      slide.setSelectedShapes(stepShapes.map((shape) => shape.id));
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        createdShapes: [...stepShapes, ...connectorShapes].map((shape) => summarizePowerPointShape(slide, shape)),
        stepCount: stepShapes.length,
        connectorCount: connectorShapes.length,
        orientation,
        geometricShapeType,
      };
    }

    if (type === "importSlidesFromBase64" || type === "mergePresentationFromBase64") {
      const base64 = trimString(action.content) ?? trimString(options.base64);
      if (!base64) {
        throw new Error("PowerPoint slide import requires a base64 presentation payload.");
      }
      const beforeSlides = await loadPowerPointSlideSummaries(context);
      let targetSlideId = trimString(options.targetSlideId) ?? action.target?.slideId;
      if (!targetSlideId && typeof action.target?.slideIndex === "number") {
        const targetSlide = await resolvePowerPointSlide(context, action.target, false);
        targetSlide.load("id");
        await context.sync();
        targetSlideId = targetSlide.id;
      }
      const insertOptions: PowerPoint.InsertSlideOptions = {};
      const formatting = trimString(options.formatting);
      const sourceSlideIds = getStringArray(options.sourceSlideIds);
      if (formatting) {
        insertOptions.formatting = formatting as PowerPoint.InsertSlideFormatting;
      }
      if (sourceSlideIds.length) {
        insertOptions.sourceSlideIds = sourceSlideIds;
      }
      if (targetSlideId) {
        insertOptions.targetSlideId = targetSlideId;
      }
      context.presentation.insertSlidesFromBase64(base64, insertOptions);
      await context.sync();
      const insertedSlides = sliceInsertedPowerPointSlides(
        beforeSlides,
        await loadPowerPointSlideSummaries(context),
        targetSlideId,
      );
      if (insertedSlides.length) {
        context.presentation.setSelectedSlides(insertedSlides.map((slide) => slide.slideId));
        await context.sync();
      }
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: insertedSlides[0]?.slideId,
        slideIndex: insertedSlides[0]?.slideIndex,
        slideIds: insertedSlides.map((slide) => slide.slideId),
        slides: insertedSlides,
        sourceSlideIds: sourceSlideIds.length ? sourceSlideIds : undefined,
        formatting,
        targetSlideId,
        insertedCount: insertedSlides.length,
      };
    }

    if (type === "exportSlidesAsBase64") {
      const slideIds = getStringArray(options.slideIds);
      const exportResult =
        slideIds.length > 0
          ? context.presentation.slides.exportAsBase64Presentation(slideIds)
          : (await resolvePowerPointSlide(context, action.target, true)).exportAsBase64();
      await context.sync();
      return { ok: true, host: "powerpoint", action: type, base64: exportResult.value };
    }

    throw new Error(`Unsupported PowerPoint action: ${type}`);
  });
}

