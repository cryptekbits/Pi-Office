import type { OfficeAnchor, OfficeHostAction } from "@pi-office/pi-office-pack/protocol";
import {
  getActionImagePayload,
  isRecord,
  loadImageElement,
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
import {
  buildPowerPointIconSvg,
  resolvePowerPointIcon,
  searchPowerPointIcons,
  type PowerPointIconCatalogEntry,
} from "../powerpoint-icons";
import { getPowerPointAssetSourceCatalog } from "../powerpoint-assets";

const POWERPOINT_MEDIA_ACTIONS = new Set([
  "searchIcons",
  "insertIcon",
  "copyImageBetweenSlides",
  "insertInlinePicture",
]);

export function isPowerPointMediaAction(type: string): boolean {
  return POWERPOINT_MEDIA_ACTIONS.has(type);
}

function stripImageDataUrl(dataUrl: string): { data: string; mimeType: string } {
  const commaIndex = dataUrl.indexOf(",");
  const metadata = commaIndex >= 0 ? dataUrl.slice(5, commaIndex) : "";
  return {
    data: commaIndex >= 0 ? dataUrl.slice(commaIndex + 1) : dataUrl,
    mimeType: metadata.split(";")[0] || "image/png",
  };
}

function encodeUtf8Base64(value: string): string {
  if (typeof btoa !== "function") {
    const bufferCtor = (
      globalThis as unknown as {
        Buffer?: { from: (input: string, encoding: "utf8") => { toString: (encoding: "base64") => string } };
      }
    ).Buffer;
    if (bufferCtor) {
      return bufferCtor.from(value, "utf8").toString("base64");
    }
    throw new Error("Base64 encoding is unavailable for PowerPoint icon insertion.");
  }

  const bytes = new TextEncoder().encode(value);
  let binary = "";
  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }
  return btoa(binary);
}

async function renderPowerPointIconAsset(
  icon: PowerPointIconCatalogEntry,
  options: Record<string, unknown>,
): Promise<{
  data: string;
  mimeType: string;
  width: number;
  height: number;
  sourceFormat: "svg-rasterized-png" | "svg";
}> {
  const color =
    trimString(options.iconColor) ??
    trimString(options.strokeColor) ??
    trimString(options.fillColor) ??
    trimString(options.fontColor);
  const svg = buildPowerPointIconSvg(icon, { color });
  const width = 320;
  const height = 320;

  if (
    typeof document === "undefined" ||
    typeof Blob === "undefined" ||
    typeof URL === "undefined" ||
    typeof btoa !== "function"
  ) {
    return {
      data: encodeUtf8Base64(svg),
      mimeType: "image/svg+xml",
      width: 160,
      height: 160,
      sourceFormat: "svg",
    };
  }

  const svgUrl = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml;charset=utf-8" }));
  try {
    const image = await loadImageElement(svgUrl);
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d");
    if (!ctx) {
      throw new Error("Canvas rendering is unavailable for PowerPoint icon insertion.");
    }
    ctx.clearRect(0, 0, width, height);
    ctx.drawImage(image, 0, 0, width, height);
    const parsed = stripImageDataUrl(canvas.toDataURL("image/png"));
    return {
      data: parsed.data,
      mimeType: parsed.mimeType,
      width,
      height,
      sourceFormat: "svg-rasterized-png",
    };
  } finally {
    URL.revokeObjectURL(svgUrl);
  }
}

function withDefaultIconShapeOptions(icon: PowerPointIconCatalogEntry, options: Record<string, unknown>): Record<string, unknown> {
  const shapeOptions = { ...options };
  delete shapeOptions.fillColor;
  delete shapeOptions.fontColor;
  delete shapeOptions.iconColor;
  delete shapeOptions.strokeColor;
  return {
    width: 40,
    height: 40,
    ...shapeOptions,
    name: trimString(options.name) ?? `Icon ${icon.name}`,
    altTextTitle: trimString(options.altTextTitle) ?? icon.name,
    altTextDescription:
      trimString(options.altTextDescription) ??
      `Pi-Office built-in ${icon.name} icon from the SVG icon catalog.`,
  };
}

function truncatePowerPointAssetText(value: string | undefined, maxLength: number): string | undefined {
  if (!value) {
    return undefined;
  }
  return value.length > maxLength ? `${value.slice(0, Math.max(0, maxLength - 1)).trimEnd()}...` : value;
}

function normalizePowerPointImageAssetSourceId(options: Record<string, unknown>): string {
  const sourceId =
    trimString(options.assetSourceId) ??
    trimString(options.sourceAssetId) ??
    trimString(options.imageAssetSourceId);
  switch (sourceId) {
    case "generated-image-base64":
    case "provided-image-base64":
    case "powerpoint-shape-snapshot":
      return sourceId;
    default:
      return "provided-image-base64";
  }
}

function withDefaultImageShapeOptions(options: Record<string, unknown>, assetSourceId: string): Record<string, unknown> {
  const altText =
    trimString(options.altText) ??
    trimString(options.generatedImagePrompt) ??
    trimString(options.prompt);
  if (assetSourceId !== "generated-image-base64") {
    return {
      ...options,
      altTextTitle: trimString(options.altTextTitle) ?? truncatePowerPointAssetText(altText, 120),
      altTextDescription: trimString(options.altTextDescription) ?? truncatePowerPointAssetText(altText, 240),
    };
  }

  const prompt = truncatePowerPointAssetText(trimString(options.generatedImagePrompt) ?? altText, 180);
  const model = trimString(options.generatedImageModel);
  const descriptionParts = ["Pi-Office generated image from generate_image."];
  if (model) {
    descriptionParts.push(`Model: ${model}.`);
  }
  if (prompt) {
    descriptionParts.push(`Prompt: ${prompt}`);
  }

  return {
    ...options,
    name: trimString(options.name) ?? "Generated Image",
    altTextTitle: trimString(options.altTextTitle) ?? truncatePowerPointAssetText(altText ?? "Generated Image", 120),
    altTextDescription: trimString(options.altTextDescription) ?? descriptionParts.join(" "),
  };
}

function shouldAllowGlyphFallback(options: Record<string, unknown>): boolean {
  return options.allowGlyphFallback === true || options.fallback === "glyph-textbox";
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
      assetFormat: "svg",
      preferredInsertionFormat: "svg-rasterized-png",
      source: "pi-office-built-in-svg-icon-catalog",
    }));
    const { getPowerPointReusableSlideComponentCatalog } = await import("../powerpoint-components");

    return {
      ok: true,
      host: "powerpoint",
      action: type,
      query,
      maxResults,
      totalMatches: icons.length,
      icons,
      catalog: "pi-office-built-in-svg-icon-catalog",
      supportedAssetSources: getPowerPointAssetSourceCatalog(),
      reusableComponents: getPowerPointReusableSlideComponentCatalog(),
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
              catalog: "pi-office-built-in-svg-icon-catalog",
              iconAssetSourceId: "provided-image-base64",
              iconAssetSource: "provided-base64",
              iconAssetMimeType: "image/png",
              iconAssetFormat: "png",
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
        applyPowerPointShapeProperties(shape, withDefaultIconShapeOptions(icon, options));
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
          catalog: "pi-office-built-in-svg-icon-catalog",
          iconAssetSourceId: "provided-image-base64",
          iconAssetSource: "provided-base64",
          iconAssetMimeType: "image/png",
          iconAssetFormat: "png",
          insertionMode: "image-shape",
        };
      });
    }

    return PowerPoint.run(async (context) => {
      const slide = await resolvePowerPointSlide(context, action.target, true);
      slide.load("id,index");
      const shapes = slide.shapes as any;
      if (typeof shapes.addImage !== "function") {
        if (!shouldAllowGlyphFallback(options)) {
          throw new Error(
            "PowerPoint insertIcon requires image insertion support. Set allowGlyphFallback=true only if a plain text glyph fallback is acceptable.",
          );
        }
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
        applyPowerPointShapeProperties(shape, withDefaultIconShapeOptions(icon, options));
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
          catalog: "pi-office-built-in-svg-icon-catalog",
          iconAssetSourceId: "built-in-icon-svg",
          completion: "fallback",
          fallbackStrategy: "explicit-glyph-textbox",
          insertionMode: "glyph-textbox",
        };
      }

      const asset = await renderPowerPointIconAsset(icon, options);
      const shape = shapes.addImage(`data:${asset.mimeType};base64,${asset.data}`) as PowerPoint.Shape;
      applyPowerPointShapeProperties(shape, withDefaultIconShapeOptions(icon, options));
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
        catalog: "pi-office-built-in-svg-icon-catalog",
        iconAssetSourceId: "built-in-icon-svg",
        iconAssetSource: "pi-office-built-in-svg-icon-catalog",
        iconAssetMimeType: asset.mimeType,
        iconAssetFormat: asset.sourceFormat,
        iconAssetWidth: asset.width,
        iconAssetHeight: asset.height,
        insertionMode: "icon-image-shape",
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
      const assetSourceId = normalizePowerPointImageAssetSourceId(options);
      const assetMimeType = trimString(options.assetMimeType) ?? trimString(options.mimeType) ?? "image/png";
      const shapes = slide.shapes as any;
      if (typeof shapes.addImage !== "function") {
        throw new Error("PowerPoint image insertion requires PowerPointApi 1.4 or newer.");
      }
      const shape = shapes.addImage(`data:${assetMimeType};base64,${imageBase64}`) as PowerPoint.Shape;
      applyPowerPointShapeProperties(shape, withDefaultImageShapeOptions(options, assetSourceId));
      shape.load("id,name,width,height,type");
      await context.sync();
      context.presentation.setSelectedSlides([slide.id]);
      slide.setSelectedShapes([shape.id]);
      await context.sync();
      return {
        ok: true,
        host: "powerpoint",
        action: type,
        shapeId: shape.id,
        shapeName: shape.name,
        shapeType: shape.type,
        slideId: slide.id,
        assetSourceId,
        assetMimeType,
        insertionMode: "image-shape",
      };
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
    assetSourceId?: string;
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
      assetSourceId: "powerpoint-shape-snapshot",
    };
  } else {
    const assetSourceId = normalizePowerPointImageAssetSourceId(options);
    sourceSummary = {
      assetSourceId,
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
    applyPowerPointShapeProperties(shape, withDefaultImageShapeOptions(options, sourceSummary.assetSourceId ?? "provided-image-base64"));
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
