import type { OfficeAnchor, OfficeHostAction } from "@pi-office/pi-office-pack/protocol";
import {
  getActionImagePayload,
  isRecord,
  parsePositiveInteger,
  resolvePositiveCount,
  supportsRequirementSet,
  toNumber,
  trimString,
} from "../shared";
import {
  applyPowerPointShapeImageAction,
  applyPowerPointShapeProperties,
  finalizePowerPointShapeSelection,
  resolvePowerPointShape,
  resolvePowerPointSlide,
} from "../powerpoint-helpers";
import { resolvePowerPointIcon, searchPowerPointIcons } from "../powerpoint-icons";

const POWERPOINT_MEDIA_ACTIONS = new Set([
  "searchIcons",
  "insertIcon",
  "copyImageBetweenSlides",
  "insertInlinePicture",
]);

export function isPowerPointMediaAction(type: string): boolean {
  return POWERPOINT_MEDIA_ACTIONS.has(type);
}

export async function applyPowerPointMediaAction(
  action: OfficeHostAction,
  type: string,
  options: Record<string, unknown>,
): Promise<unknown> {
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
    return copyImageBetweenSlides(action, type, options);
  }

  if (type === "insertInlinePicture") {
    return PowerPoint.run(async (context) => {
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
    });
  }

  throw new Error(`Unsupported PowerPoint media action: ${type}`);
}

async function copyImageBetweenSlides(
  action: OfficeHostAction,
  type: string,
  options: Record<string, unknown>,
): Promise<unknown> {
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
