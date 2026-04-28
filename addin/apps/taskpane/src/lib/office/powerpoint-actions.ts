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
} from "./powerpoint-helpers";
import { applyPowerPointTableAction, isPowerPointTableAction } from "./powerpoint-actions/tables";
import { resolvePowerPointIcon, searchPowerPointIcons } from "./powerpoint-icons";

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

  if (type === "searchIcons") {
    const query =
      trimString(action.content) ??
      trimString(options.query) ??
      trimString(options.search) ??
      "";
    if (!query) {
      throw new Error("PowerPoint searchIcons requires a non-empty query.");
    }

    const maxResults = resolvePositiveCount(options.maxResults, 12);
    const icons = searchPowerPointIcons(query, maxResults).map((icon) => ({
      id: icon.id,
      name: icon.name,
      keywords: icon.keywords,
      glyph: icon.glyph,
    }));

    return {
      ok: true,
      host: "powerpoint",
      action: type,
      query,
      maxResults,
      totalMatches: icons.length,
      icons,
      catalog: "taskpane-runtime-icon-catalog",
      note: "Use insertIcon with iconId to place a selected icon on the target slide.",
    };
  }

  if (type === "insertIcon") {
    const iconQuery =
      trimString(action.content) ??
      trimString(options.iconId) ??
      trimString(options.iconName) ??
      trimString(options.query) ??
      "";
    if (!iconQuery) {
      throw new Error("PowerPoint insertIcon requires iconId, iconName, query, or content.");
    }

    const icon = resolvePowerPointIcon(iconQuery);
    if (!icon) {
      throw new Error(`Could not find an icon for "${iconQuery}". Run search_icons first to discover supported icon IDs.`);
    }

    const providedBase64 = trimString(options.iconBase64) ?? trimString(options.base64);
    if (providedBase64) {
      if (action.target?.shapeId) {
        const replaced = await applyPowerPointShapeImageAction(action, type, { data: providedBase64, mimeType: "image/png" });
        return isRecord(replaced)
          ? {
              ...replaced,
              iconId: icon.id,
              iconName: icon.name,
              iconGlyph: icon.glyph,
              iconKeywords: icon.keywords,
              catalog: "taskpane-runtime-icon-catalog",
              insertionMode: "shape-image-replace",
            }
          : replaced;
      }

      return PowerPoint.run(async (context) => {
        const slide = await resolvePowerPointSlide(context, action.target, true);
        slide.load("id,index");
        const shapes = slide.shapes as any;
        if (typeof shapes.addImage !== "function") {
          throw new Error("PowerPoint insertIcon image mode requires PowerPointApi 1.4 or newer.");
        }
        const shape = shapes.addImage(`data:image/png;base64,${providedBase64}`) as PowerPoint.Shape;
        applyPowerPointShapeProperties(shape, options);
        shape.load("id,name,type");
        await context.sync();
        context.presentation.setSelectedSlides([slide.id]);
        slide.setSelectedShapes([shape.id]);
        await context.sync();
        return {
          ok: true,
          host: "powerpoint",
          action: type,
          slideId: slide.id,
          slideIndex: slide.index + 1,
          shapeId: shape.id,
          shapeName: shape.name,
          shapeType: shape.type,
          iconId: icon.id,
          iconName: icon.name,
          iconGlyph: icon.glyph,
          iconKeywords: icon.keywords,
          catalog: "taskpane-runtime-icon-catalog",
          insertionMode: "image-shape",
        };
      });
    }

    return PowerPoint.run(async (context) => {
      const slide = await resolvePowerPointSlide(context, action.target, true);
      slide.load("id,index");
      const textOptions: PowerPoint.ShapeAddOptions = {};
      const left = toNumber(options.left);
      const top = toNumber(options.top);
      const width = toNumber(options.width);
      const height = toNumber(options.height);
      if (typeof left === "number") textOptions.left = left;
      if (typeof top === "number") textOptions.top = top;
      if (typeof width === "number") textOptions.width = width;
      if (typeof height === "number") textOptions.height = height;
      const shape = slide.shapes.addTextBox(icon.glyph, textOptions);
      applyPowerPointShapeProperties(shape, {
        ...options,
        name: trimString(options.name) ?? `Icon ${icon.name}`,
      });
      const textFrame = shape.getTextFrameOrNullObject();
      textFrame.load("isNullObject");
      shape.load("id,name,type");
      await context.sync();

      if (!textFrame.isNullObject) {
        const fontSize = toNumber(options.fontSize) ?? 28;
        textFrame.textRange.font.size = fontSize;
        const fontColor = trimString(options.fontColor) ?? trimString(options.fillColor);
        if (fontColor) {
          textFrame.textRange.font.color = fontColor;
        }
        textFrame.wordWrap = false;
      }
      await context.sync();

      context.presentation.setSelectedSlides([slide.id]);
      slide.setSelectedShapes([shape.id]);
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        shapeId: shape.id,
        shapeName: shape.name,
        shapeType: shape.type,
        iconId: icon.id,
        iconName: icon.name,
        iconGlyph: icon.glyph,
        iconKeywords: icon.keywords,
        catalog: "taskpane-runtime-icon-catalog",
        insertionMode: "glyph-textbox",
      };
    });
  }

  if (type === "copyImageBetweenSlides") {
    const sourceSlideIndex =
      parsePositiveInteger(options.sourceSlideIndex) ??
      parsePositiveInteger(options.fromSlideIndex);
    const sourceShapeId = trimString(options.sourceShapeId) ?? trimString(options.fromShapeId);
    const sourceTarget: OfficeAnchor | undefined =
      sourceShapeId || trimString(options.sourceSlideId) || typeof sourceSlideIndex === "number"
        ? {
            kind: "shape",
            slideId: trimString(options.sourceSlideId) ?? trimString(options.fromSlideId),
            slideIndex: sourceSlideIndex,
            shapeId: sourceShapeId,
          }
        : undefined;

    const directImageBase64 =
      trimString(action.content) ??
      trimString(options.sourceImageBase64) ??
      trimString(options.base64);

    let imageBase64 = directImageBase64;
    let sourceSummary: {
      sourceSlideId?: string;
      sourceSlideIndex?: number;
      sourceShapeId?: string;
      sourceShapeName?: string;
    } = {};

    if (!imageBase64) {
      if (!supportsRequirementSet("PowerPointApi", "1.10")) {
        throw new Error("PowerPoint copyImageBetweenSlides without sourceImageBase64 requires PowerPointApi 1.10 for shape snapshot export.");
      }

      const extracted = await PowerPoint.run(async (context) => {
        const resolved = await resolvePowerPointShape(context, sourceTarget, true);
        resolved.slide.load("id,index");
        resolved.shape.load("id,name");
        const exportResult = resolved.shape.getImageAsBase64({ format: "Png", width: 1400 });
        await context.sync();
        return {
          imageBase64: exportResult.value,
          sourceSlideId: resolved.slide.id,
          sourceSlideIndex: resolved.slide.index + 1,
          sourceShapeId: resolved.shape.id,
          sourceShapeName: resolved.shape.name,
        };
      });
      imageBase64 = extracted.imageBase64;
      sourceSummary = {
        sourceSlideId: extracted.sourceSlideId,
        sourceSlideIndex: extracted.sourceSlideIndex,
        sourceShapeId: extracted.sourceShapeId,
        sourceShapeName: extracted.sourceShapeName,
      };
    }

    if (!imageBase64) {
      throw new Error("PowerPoint copyImageBetweenSlides could not resolve an image payload.");
    }

    if (action.target?.shapeId) {
      const replaced = await applyPowerPointShapeImageAction(action, type, { data: imageBase64, mimeType: "image/png" });
      return isRecord(replaced)
        ? {
            ...replaced,
            ...sourceSummary,
            copiedImageMimeType: "image/png",
            insertionMode: "shape-image-replace",
          }
        : replaced;
    }

    return PowerPoint.run(async (context) => {
      const slide = await resolvePowerPointSlide(context, action.target, true);
      slide.load("id,index");
      const shapes = slide.shapes as any;
      if (typeof shapes.addImage !== "function") {
        throw new Error("PowerPoint copyImageBetweenSlides insertion requires PowerPointApi 1.4 or newer.");
      }
      const shape = shapes.addImage(`data:image/png;base64,${imageBase64}`) as PowerPoint.Shape;
      applyPowerPointShapeProperties(shape, options);
      shape.load("id,name,type");
      await context.sync();
      context.presentation.setSelectedSlides([slide.id]);
      slide.setSelectedShapes([shape.id]);
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        shapeId: shape.id,
        shapeName: shape.name,
        shapeType: shape.type,
        ...sourceSummary,
        copiedImageMimeType: "image/png",
        insertionMode: "image-shape",
      };
    });
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

    if (type === "addSlide") {
      const slide = await createPowerPointSlide(context, action.target, options);
      context.presentation.setSelectedSlides([slide.id]);
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        layoutId: slide.layout.id,
        layoutName: slide.layout.name,
        layoutType: slide.layout.type,
        slideMasterId: slide.slideMaster.id,
        slideMasterName: slide.slideMaster.name,
        createdSlides: [summarizePowerPointSlide(slide)],
      };
    }

    if (type === "addAgendaSlide") {
      const requestedSourceSlideIds = getPowerPointSlideIdArray(options.sourceSlideIds ?? options.slideIds ?? options.sourceSlides);
      const sourceSlides = await resolvePowerPointSourceSlides(context, undefined, requestedSourceSlideIds);
      const sourceContent = requestedSourceSlideIds.length ? await loadPowerPointSlideContentSummaries(context, sourceSlides) : [];
      const agendaItems =
        getStringArray(options.items).length > 0
          ? getStringArray(options.items)
          : getStringArray(options.agendaItems).length > 0
            ? getStringArray(options.agendaItems)
            : sourceContent.map((slide) => slide.title ?? slide.label).filter((entry): entry is string => Boolean(entry));
      if (!agendaItems.length) {
        throw new Error("PowerPoint addAgendaSlide requires agendaItems/items or source slides with detectable titles.");
      }

      const slide = await createPowerPointSlide(context, action.target, options);
      slide.load("id,index");
      const titleText = trimString(options.title) ?? trimString(action.content) ?? "Agenda";
      const subtitleText = trimString(options.subtitle);
      const titleShape = slide.shapes.addTextBox(titleText, {
        left: toNumber(options.titleLeft) ?? 60,
        top: toNumber(options.titleTop) ?? 50,
        width: toNumber(options.titleWidth) ?? 620,
        height: toNumber(options.titleHeight) ?? 48,
      });
      const agendaShape = slide.shapes.addTextBox(
        agendaItems.map((item, index) => `${index + 1}. ${item}`).join("\n"),
        {
          left: toNumber(options.bodyLeft) ?? 80,
          top: toNumber(options.bodyTop) ?? (subtitleText ? 180 : 150),
          width: toNumber(options.bodyWidth) ?? 560,
          height: toNumber(options.bodyHeight) ?? 300,
        },
      );
      const createdShapes = [titleShape, agendaShape];
      applyPowerPointShapeProperties(titleShape, {
        name: trimString(options.titleName) ?? "Agenda Title",
        fillColor: trimString(options.titleFillColor),
        lineColor: trimString(options.titleLineColor),
      });
      applyPowerPointShapeProperties(agendaShape, {
        name: trimString(options.bodyName) ?? "Agenda Items",
        fillColor: trimString(options.bodyFillColor),
        lineColor: trimString(options.bodyLineColor),
      });

      const titleFrame = titleShape.getTextFrameOrNullObject();
      const agendaFrame = agendaShape.getTextFrameOrNullObject();
      applyPowerPointTextFrameProperties(titleFrame, {
        autoSizeSetting: trimString(options.titleAutoSizeSetting) ?? "AutoSizeShapeToFitText",
        wordWrap: toBoolean(options.titleWordWrap) ?? true,
      });
      applyPowerPointTextFrameProperties(agendaFrame, {
        autoSizeSetting: trimString(options.bodyAutoSizeSetting) ?? "AutoSizeShapeToFitText",
        wordWrap: toBoolean(options.bodyWordWrap) ?? true,
      });
      if (typeof toNumber(options.titleFontSize) === "number") {
        titleFrame.textRange.font.size = toNumber(options.titleFontSize)!;
      }
      if (trimString(options.titleFontColor)) {
        titleFrame.textRange.font.color = trimString(options.titleFontColor)!;
      }
      if (typeof toBoolean(options.titleBold) === "boolean") {
        titleFrame.textRange.font.bold = toBoolean(options.titleBold)!;
      }
      if (typeof toNumber(options.bodyFontSize) === "number") {
        agendaFrame.textRange.font.size = toNumber(options.bodyFontSize)!;
      }
      if (trimString(options.bodyFontColor)) {
        agendaFrame.textRange.font.color = trimString(options.bodyFontColor)!;
      }

      if (subtitleText) {
        const subtitleShape = slide.shapes.addTextBox(subtitleText, {
          left: toNumber(options.subtitleLeft) ?? 80,
          top: toNumber(options.subtitleTop) ?? 112,
          width: toNumber(options.subtitleWidth) ?? 560,
          height: toNumber(options.subtitleHeight) ?? 42,
        });
        createdShapes.push(subtitleShape);
        applyPowerPointShapeProperties(subtitleShape, {
          name: trimString(options.subtitleName) ?? "Agenda Subtitle",
          fillColor: trimString(options.subtitleFillColor),
          lineColor: trimString(options.subtitleLineColor),
        });
        const subtitleFrame = subtitleShape.getTextFrameOrNullObject();
        applyPowerPointTextFrameProperties(subtitleFrame, {
          autoSizeSetting: trimString(options.subtitleAutoSizeSetting) ?? "AutoSizeShapeToFitText",
          wordWrap: toBoolean(options.subtitleWordWrap) ?? true,
        });
        if (typeof toNumber(options.subtitleFontSize) === "number") {
          subtitleFrame.textRange.font.size = toNumber(options.subtitleFontSize)!;
        }
        if (trimString(options.subtitleFontColor)) {
          subtitleFrame.textRange.font.color = trimString(options.subtitleFontColor)!;
        }
      }

      for (const shape of createdShapes) {
        shape.load("id,name,type");
      }
      await context.sync();
      context.presentation.setSelectedSlides([slide.id]);
      slide.setSelectedShapes([agendaShape.id]);
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        createdSlides: [summarizePowerPointSlide(slide)],
        createdShapes: createdShapes.map((shape) => summarizePowerPointShape(slide, shape)),
        agendaItems,
        sourceSlideIds: sourceSlides.map((entry) => entry.slideId),
      };
    }

    if (type === "addTransitionSlide") {
      const slide = await createPowerPointSlide(context, action.target, options);
      slide.load("id,index");
      const titleText = trimString(options.title) ?? trimString(action.content) ?? "Section";
      const subtitleText = trimString(options.subtitle);
      const kickerText = trimString(options.kicker);
      const createdShapes: PowerPoint.Shape[] = [];

      if (kickerText) {
        const kickerShape = slide.shapes.addTextBox(kickerText, {
          left: toNumber(options.kickerLeft) ?? 72,
          top: toNumber(options.kickerTop) ?? 86,
          width: toNumber(options.kickerWidth) ?? 520,
          height: toNumber(options.kickerHeight) ?? 24,
        });
        createdShapes.push(kickerShape);
        applyPowerPointShapeProperties(kickerShape, {
          name: trimString(options.kickerName) ?? "Transition Kicker",
          fillColor: trimString(options.kickerFillColor),
          lineColor: trimString(options.kickerLineColor),
        });
        const kickerFrame = kickerShape.getTextFrameOrNullObject();
        applyPowerPointTextFrameProperties(kickerFrame, { autoSizeSetting: "AutoSizeShapeToFitText", wordWrap: false });
        if (typeof toNumber(options.kickerFontSize) === "number") {
          kickerFrame.textRange.font.size = toNumber(options.kickerFontSize)!;
        }
        if (trimString(options.kickerFontColor)) {
          kickerFrame.textRange.font.color = trimString(options.kickerFontColor)!;
        }
      }

      const titleShape = slide.shapes.addTextBox(titleText, {
        left: toNumber(options.titleLeft) ?? 72,
        top: toNumber(options.titleTop) ?? (kickerText ? 124 : 132),
        width: toNumber(options.titleWidth) ?? 580,
        height: toNumber(options.titleHeight) ?? 96,
      });
      createdShapes.push(titleShape);
      applyPowerPointShapeProperties(titleShape, {
        name: trimString(options.titleName) ?? "Transition Title",
        fillColor: trimString(options.titleFillColor),
        lineColor: trimString(options.titleLineColor),
      });
      const titleFrame = titleShape.getTextFrameOrNullObject();
      applyPowerPointTextFrameProperties(titleFrame, {
        autoSizeSetting: trimString(options.titleAutoSizeSetting) ?? "AutoSizeShapeToFitText",
        wordWrap: toBoolean(options.titleWordWrap) ?? true,
      });
      if (typeof toNumber(options.titleFontSize) === "number") {
        titleFrame.textRange.font.size = toNumber(options.titleFontSize)!;
      }
      if (trimString(options.titleFontColor)) {
        titleFrame.textRange.font.color = trimString(options.titleFontColor)!;
      }
      if (typeof toBoolean(options.titleBold) === "boolean") {
        titleFrame.textRange.font.bold = toBoolean(options.titleBold)!;
      }

      if (subtitleText) {
        const subtitleShape = slide.shapes.addTextBox(subtitleText, {
          left: toNumber(options.subtitleLeft) ?? 72,
          top: toNumber(options.subtitleTop) ?? (kickerText ? 236 : 246),
          width: toNumber(options.subtitleWidth) ?? 560,
          height: toNumber(options.subtitleHeight) ?? 58,
        });
        createdShapes.push(subtitleShape);
        applyPowerPointShapeProperties(subtitleShape, {
          name: trimString(options.subtitleName) ?? "Transition Subtitle",
          fillColor: trimString(options.subtitleFillColor),
          lineColor: trimString(options.subtitleLineColor),
        });
        const subtitleFrame = subtitleShape.getTextFrameOrNullObject();
        applyPowerPointTextFrameProperties(subtitleFrame, {
          autoSizeSetting: trimString(options.subtitleAutoSizeSetting) ?? "AutoSizeShapeToFitText",
          wordWrap: toBoolean(options.subtitleWordWrap) ?? true,
        });
        if (typeof toNumber(options.subtitleFontSize) === "number") {
          subtitleFrame.textRange.font.size = toNumber(options.subtitleFontSize)!;
        }
        if (trimString(options.subtitleFontColor)) {
          subtitleFrame.textRange.font.color = trimString(options.subtitleFontColor)!;
        }
      }

      for (const shape of createdShapes) {
        shape.load("id,name,type");
      }
      await context.sync();
      context.presentation.setSelectedSlides([slide.id]);
      slide.setSelectedShapes([titleShape.id]);
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        createdSlides: [summarizePowerPointSlide(slide)],
        createdShapes: createdShapes.map((shape) => summarizePowerPointShape(slide, shape)),
      };
    }

    if (type === "combineSlides") {
      const requestedSourceSlideIds = getPowerPointSlideIdArray(options.sourceSlideIds ?? options.slideIds ?? options.sourceSlides);
      const sourceSlides = await resolvePowerPointSourceSlides(context, undefined, requestedSourceSlideIds);
      if (!sourceSlides.length) {
        throw new Error("PowerPoint combineSlides requires at least one source slide.");
      }
      const sourceContent = await loadPowerPointSlideContentSummaries(context, sourceSlides);
      const slide = await createPowerPointSlide(context, action.target, options);
      slide.load("id,index");
      const titleText = trimString(options.title) ?? trimString(action.content) ?? "Combined Overview";
      const titleShape = slide.shapes.addTextBox(titleText, {
        left: toNumber(options.titleLeft) ?? 54,
        top: toNumber(options.titleTop) ?? 40,
        width: toNumber(options.titleWidth) ?? 620,
        height: toNumber(options.titleHeight) ?? 40,
      });
      const contentShape = slide.shapes.addTextBox(
        sourceContent
          .map((entry) => {
            const bodyPreview =
              truncateText(
                entry.textBlocks
                  .slice(0, resolvePositiveCount(options.maxBlocksPerSlide, 2))
                  .join(" "),
                toNumber(options.maxCharsPerSlide) ?? 240,
              ) ?? "";
            return bodyPreview && bodyPreview !== entry.title ? `${entry.title ?? entry.label}\n${bodyPreview}` : (entry.title ?? entry.label);
          })
          .join("\n\n"),
        {
          left: toNumber(options.bodyLeft) ?? 72,
          top: toNumber(options.bodyTop) ?? 120,
          width: toNumber(options.bodyWidth) ?? 560,
          height: toNumber(options.bodyHeight) ?? 300,
        },
      );
      const createdShapes = [titleShape, contentShape];
      applyPowerPointShapeProperties(titleShape, {
        name: trimString(options.titleName) ?? "Combined Title",
        fillColor: trimString(options.titleFillColor),
        lineColor: trimString(options.titleLineColor),
      });
      applyPowerPointShapeProperties(contentShape, {
        name: trimString(options.bodyName) ?? "Combined Content",
        fillColor: trimString(options.bodyFillColor),
        lineColor: trimString(options.bodyLineColor),
      });
      applyPowerPointTextFrameProperties(titleShape.getTextFrameOrNullObject(), {
        autoSizeSetting: trimString(options.titleAutoSizeSetting) ?? "AutoSizeShapeToFitText",
        wordWrap: toBoolean(options.titleWordWrap) ?? true,
      });
      applyPowerPointTextFrameProperties(contentShape.getTextFrameOrNullObject(), {
        autoSizeSetting: trimString(options.bodyAutoSizeSetting) ?? "AutoSizeShapeToFitText",
        wordWrap: toBoolean(options.bodyWordWrap) ?? true,
      });
      for (const shape of createdShapes) {
        shape.load("id,name,type");
      }
      await context.sync();

      const deletedSlides: Array<{ slideId: string; slideIndex: number; label: string }> = [];
      if (toBoolean(options.deleteSourceSlides ?? action.deleteSourceSlides)) {
        for (const sourceSlide of sourceSlides) {
          context.presentation.slides.getItem(sourceSlide.slideId).delete();
          deletedSlides.push(sourceSlide);
        }
        await context.sync();
        slide.load("id,index");
        await context.sync();
      }

      context.presentation.setSelectedSlides([slide.id]);
      slide.setSelectedShapes([contentShape.id]);
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        createdSlides: [summarizePowerPointSlide(slide)],
        createdShapes: createdShapes.map((shape) => summarizePowerPointShape(slide, shape)),
        sourceSlideIds: sourceSlides.map((entry) => entry.slideId),
        deletedSlides: deletedSlides.length ? deletedSlides : undefined,
      };
    }

    if (type === "reorderSlides" || type === "reorderStoryline") {
      const desiredSlideIds = Array.from(new Set(getPowerPointSlideIdArray(options.slideIds ?? options.storyline ?? options.slides)));
      if (desiredSlideIds.length < 2) {
        throw new Error(`PowerPoint ${type} requires at least two slide IDs.`);
      }

      const beforeSlides = await loadPowerPointSlideSummaries(context);
      const beforeIndexById = new Map(beforeSlides.map((slide, index) => [slide.slideId, index]));
      const missingSlideId = desiredSlideIds.find((slideId) => !beforeIndexById.has(slideId));
      if (missingSlideId) {
        throw new Error(`Could not find the requested PowerPoint slide: ${missingSlideId}.`);
      }

      let insertionIndex = await resolvePowerPointInsertionIndex(context, undefined, options);
      if (typeof insertionIndex !== "number" || Number.isNaN(insertionIndex)) {
        insertionIndex = Math.min(...desiredSlideIds.map((slideId) => beforeIndexById.get(slideId) ?? 0));
      }

      const relativeSlideId = trimString(options.relativeToSlideId) ?? trimString(options.targetSlideId);
      if (relativeSlideId && desiredSlideIds.includes(relativeSlideId)) {
        insertionIndex = Math.min(...desiredSlideIds.map((slideId) => beforeIndexById.get(slideId) ?? 0));
      }

      for (let index = 0; index < desiredSlideIds.length; index += 1) {
        context.presentation.slides.getItem(desiredSlideIds[index]!).moveTo(insertionIndex + index);
      }
      await context.sync();
      const afterSlides = await loadPowerPointSlideSummaries(context);
      const afterById = new Map(afterSlides.map((slide) => [slide.slideId, slide]));
      const reorderedSlides = desiredSlideIds
        .map((slideId) => afterById.get(slideId))
        .filter((slide): slide is { slideId: string; slideIndex: number; label: string } => Boolean(slide));
      context.presentation.setSelectedSlides(reorderedSlides.map((slide) => slide.slideId));
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideIds: reorderedSlides.map((slide) => slide.slideId),
        slides: reorderedSlides,
        startIndex: insertionIndex + 1,
      };
    }

    if (type === "applyLayout") {
      const slide = await resolvePowerPointSlide(context, action.target, true);
      const { layout } = await resolvePowerPointLayoutSelection(context, action.target, options);
      if (!layout) {
        throw new Error("PowerPoint applyLayout requires a resolvable layout.");
      }
      slide.applyLayout(layout);
      slide.load("id,index");
      slide.layout.load("id,name,type");
      slide.slideMaster.load("id,name");
      await context.sync();
      context.presentation.setSelectedSlides([slide.id]);
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: slide.id,
        slideIndex: slide.index + 1,
        layoutId: slide.layout.id,
        layoutName: slide.layout.name,
        layoutType: slide.layout.type,
        slideMasterId: slide.slideMaster.id,
        slideMasterName: slide.slideMaster.name,
      };
    }

    if (type === "moveSlide") {
      const slide = await resolvePowerPointSlide(context, action.target, false);
      const destinationIndex = Math.max(0, (toNumber(options.slideIndex) ?? 1) - 1);
      slide.load("id,index");
      slide.moveTo(destinationIndex);
      await context.sync();
      return { ok: true, host: "powerpoint", action: type, slideId: slide.id, slideIndex: destinationIndex + 1 };
    }

    if (type === "duplicateSlide" || type === "duplicateSlides") {
      const requestedSlideIds = getStringArray(options.slideIds);
      const sourceSlides = await resolvePowerPointSourceSlides(context, action.target, requestedSlideIds);
      if (!sourceSlides.length) {
        throw new Error("PowerPoint duplication requires at least one source slide.");
      }

      const beforeSlides = await loadPowerPointSlideSummaries(context);
      const formatting = trimString(options.formatting);
      let explicitTargetSlideId = trimString(options.targetSlideId);
      if (!explicitTargetSlideId && requestedSlideIds.length > 0 && (action.target?.slideId || typeof action.target?.slideIndex === "number")) {
        const targetSlide = await resolvePowerPointSlide(context, action.target, false);
        targetSlide.load("id");
        await context.sync();
        explicitTargetSlideId = targetSlide.id;
      }
      const targetSlideId =
        explicitTargetSlideId ??
        sourceSlides
          .slice()
          .sort((left, right) => left.slideIndex - right.slideIndex)
          .at(-1)?.slideId;

      let exportResult: OfficeExtension.ClientResult<string>;
      if (sourceSlides.length === 1) {
        if (!supportsRequirementSet("PowerPointApi", "1.8")) {
          throw new Error("PowerPoint single-slide duplication requires PowerPointApi 1.8.");
        }
        const sourceSlide = sourceSlides[0];
        if (!sourceSlide) {
          throw new Error("PowerPoint duplication could not resolve the requested source slide.");
        }
        exportResult = context.presentation.slides.getItem(sourceSlide.slideId).exportAsBase64();
      } else {
        if (!supportsRequirementSet("PowerPointApi", "1.10")) {
          throw new Error("PowerPoint multi-slide duplication requires PowerPointApi 1.10.");
        }
        exportResult = context.presentation.slides.exportAsBase64Presentation(sourceSlides.map((slide) => slide.slideId));
      }
      await context.sync();

      const insertOptions: PowerPoint.InsertSlideOptions = {};
      if (formatting) {
        insertOptions.formatting = formatting as PowerPoint.InsertSlideFormatting;
      }
      if (targetSlideId) {
        insertOptions.targetSlideId = targetSlideId;
      }
      context.presentation.insertSlidesFromBase64(exportResult.value, insertOptions);
      await context.sync();

      const duplicatedSlides = sliceInsertedPowerPointSlides(beforeSlides, await loadPowerPointSlideSummaries(context), targetSlideId, sourceSlides.length);
      if (duplicatedSlides.length) {
        context.presentation.setSelectedSlides(duplicatedSlides.map((slide) => slide.slideId));
        await context.sync();
      }
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: duplicatedSlides[0]?.slideId,
        slideIndex: duplicatedSlides[0]?.slideIndex,
        slideIds: duplicatedSlides.map((slide) => slide.slideId),
        slides: duplicatedSlides,
        sourceSlideIds: sourceSlides.map((slide) => slide.slideId),
        formatting,
        targetSlideId,
      };
    }

    if (type === "deleteSlide" || type === "deleteSlides") {
      const beforeSlides = await loadPowerPointSlideSummaries(context);
      const deleteSlides = await resolvePowerPointSourceSlides(context, action.target, getStringArray(options.slideIds));
      const deleteSlideIds = Array.from(new Set(deleteSlides.map((slide) => slide.slideId)));
      if (!deleteSlideIds.length) {
        throw new Error("PowerPoint delete requires at least one slide.");
      }
      if (deleteSlideIds.length >= beforeSlides.length) {
        throw new Error("PowerPoint delete must leave at least one slide in the presentation.");
      }

      const firstDeletedIndex = beforeSlides.findIndex((slide) => deleteSlideIds.includes(slide.slideId));
      const remainingBefore = beforeSlides.filter((slide) => !deleteSlideIds.includes(slide.slideId));
      const nextSelection = remainingBefore[Math.min(Math.max(firstDeletedIndex, 0), remainingBefore.length - 1)];

      for (const slideId of deleteSlideIds) {
        context.presentation.slides.getItem(slideId).delete();
      }
      await context.sync();
      if (nextSelection?.slideId) {
        context.presentation.setSelectedSlides([nextSelection.slideId]);
        await context.sync();
      }

      return {
        ok: true,
        host: "powerpoint",
        action: type,
        slideId: nextSelection?.slideId,
        slideIndex: nextSelection?.slideIndex,
        deletedCount: deleteSlideIds.length,
        deletedSlideIds: deleteSlideIds,
        deletedSlides: deleteSlides,
      };
    }

    if (type === "selectSlides") {
      const slideIds = getStringArray(options.slideIds);
      if (slideIds.length) {
        context.presentation.setSelectedSlides(slideIds);
        await context.sync();
        return { ok: true, host: "powerpoint", action: type, slideIds };
      }

      const slide = await resolvePowerPointSlide(context, action.target, false);
      slide.load("id,index");
      await context.sync();
      context.presentation.setSelectedSlides([slide.id]);
      await context.sync();
      return { ok: true, host: "powerpoint", action: type, slideId: slide.id, slideIndex: slide.index + 1 };
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

    if (type === "insertInlinePicture") {
      const slide = await resolvePowerPointSlide(context, action.target, true);
      const imageBase64 = action.content ?? "";
      const shapes = slide.shapes as any;
      if (typeof shapes.addImage !== "function") {
        throw new Error("PowerPoint image insertion requires PowerPointApi 1.4 or newer.");
      }
      const shape = shapes.addImage(`data:image/png;base64,${imageBase64}`) as PowerPoint.Shape;
      shape.load("id,name,width,height");
      await context.sync();
      return { ok: true, host: "powerpoint", action: type, shapeId: shape.id, shapeName: shape.name, slideId: slide.id };
    }

    throw new Error(`Unsupported PowerPoint action: ${type}`);
  });
}

