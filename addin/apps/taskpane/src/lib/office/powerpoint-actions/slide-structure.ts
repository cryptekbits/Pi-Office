import type { OfficeHostAction } from "@pi-office/pi-office-pack/protocol";
import {
  getStringArray,
  resolvePositiveCount,
  supportsRequirementSet,
  toBoolean,
  toNumber,
  trimString,
  truncateText,
} from "../shared";
import {
  applyPowerPointShapeProperties,
  applyPowerPointTextFrameProperties,
  createPowerPointSlide,
  getPowerPointSlideIdArray,
  loadPowerPointSlideContentSummaries,
  loadPowerPointSlideSummaries,
  resolvePowerPointInsertionIndex,
  resolvePowerPointLayoutSelection,
  resolvePowerPointSlide,
  resolvePowerPointSourceSlides,
  sliceInsertedPowerPointSlides,
  summarizePowerPointShape,
  summarizePowerPointSlide,
} from "../powerpoint-helpers";

const POWERPOINT_SLIDE_STRUCTURE_ACTIONS = new Set([
  "addSlide",
  "addAgendaSlide",
  "addTransitionSlide",
  "combineSlides",
  "reorderSlides",
  "reorderStoryline",
  "applyLayout",
  "moveSlide",
  "duplicateSlide",
  "duplicateSlides",
  "deleteSlide",
  "deleteSlides",
  "selectSlides",
]);

export function isPowerPointSlideStructureAction(type: string): boolean {
  return POWERPOINT_SLIDE_STRUCTURE_ACTIONS.has(type);
}

export async function applyPowerPointSlideStructureAction(
  context: PowerPoint.RequestContext,
  action: OfficeHostAction,
  type: string,
  options: Record<string, unknown>,
): Promise<unknown> {
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

  throw new Error(`Unsupported PowerPoint slide-structure action: ${type}`);
}
