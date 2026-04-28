export type PowerPointAssetPreviewKind = "icon" | "image" | "component";

export interface PowerPointAssetPreview {
  kind: PowerPointAssetPreviewKind;
  assetSourceId: string;
  label: string;
  insertionMode: string;
  verificationHint: string;
  format?: string | undefined;
  mimeType?: string | undefined;
  width?: number | undefined;
  height?: number | undefined;
  shapeId?: string | undefined;
  shapeName?: string | undefined;
  slideId?: string | undefined;
  slideIndex?: number | undefined;
  iconId?: string | undefined;
  iconGlyph?: string | undefined;
  componentId?: string | undefined;
  createdShapeCount?: number | undefined;
  componentTextPreview?: string | undefined;
  source?: string | undefined;
  sourceShapeId?: string | undefined;
  sourceShapeName?: string | undefined;
  sourceSlideId?: string | undefined;
  sourceSlideIndex?: number | undefined;
  fallbackStrategy?: string | undefined;
}

interface SharedPreviewArgs {
  assetSourceId: string;
  label?: string | undefined;
  insertionMode: string;
  format?: string | undefined;
  mimeType?: string | undefined;
  width?: number | undefined;
  height?: number | undefined;
  shapeId?: string | undefined;
  shapeName?: string | undefined;
  slideId?: string | undefined;
  slideIndex?: number | undefined;
  source?: string | undefined;
  fallbackStrategy?: string | undefined;
}

export interface PowerPointIconAssetPreviewArgs extends SharedPreviewArgs {
  iconId: string;
  iconName: string;
  iconGlyph?: string | undefined;
}

export interface PowerPointImageAssetPreviewArgs extends SharedPreviewArgs {
  sourceShapeId?: string | undefined;
  sourceShapeName?: string | undefined;
  sourceSlideId?: string | undefined;
  sourceSlideIndex?: number | undefined;
}

export interface PowerPointComponentAssetPreviewArgs extends SharedPreviewArgs {
  componentId: string;
  componentLabel: string;
  createdShapeCount?: number | undefined;
  componentPreview?: Record<string, unknown> | undefined;
}

function cleanText(value: string | undefined): string | undefined {
  const trimmed = typeof value === "string" ? value.trim() : "";
  return trimmed ? trimmed : undefined;
}

function finiteNumber(value: number | undefined): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function truncateText(value: string | undefined, maxLength: number): string | undefined {
  const text = cleanText(value);
  if (!text) {
    return undefined;
  }
  return text.length > maxLength ? `${text.slice(0, Math.max(0, maxLength - 1)).trimEnd()}...` : text;
}

function sourceDefaultLabel(assetSourceId: string): string {
  switch (assetSourceId) {
    case "built-in-icon-svg":
      return "Built-in PowerPoint icon";
    case "generated-image-base64":
      return "Generated image";
    case "powerpoint-shape-snapshot":
      return "PowerPoint shape snapshot";
    case "reusable-slide-component":
      return "Reusable slide component";
    default:
      return "PowerPoint image";
  }
}

function createSharedPreview(
  kind: PowerPointAssetPreviewKind,
  args: SharedPreviewArgs,
  verificationHint: string,
): PowerPointAssetPreview {
  return {
    kind,
    assetSourceId: args.assetSourceId,
    label: truncateText(args.label, 120) ?? sourceDefaultLabel(args.assetSourceId),
    insertionMode: args.insertionMode,
    verificationHint,
    format: cleanText(args.format),
    mimeType: cleanText(args.mimeType),
    width: finiteNumber(args.width),
    height: finiteNumber(args.height),
    shapeId: cleanText(args.shapeId),
    shapeName: cleanText(args.shapeName),
    slideId: cleanText(args.slideId),
    slideIndex: finiteNumber(args.slideIndex),
    source: cleanText(args.source),
    fallbackStrategy: cleanText(args.fallbackStrategy),
  };
}

export function summarizePowerPointComponentPreview(preview: Record<string, unknown>): string | undefined {
  const parts = Object.entries(preview)
    .map(([key, value]) => {
      if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
        const text = String(value).trim();
        return text ? `${key}: ${text}` : undefined;
      }
      return undefined;
    })
    .filter((value): value is string => Boolean(value));
  return truncateText(parts.join(" | "), 180);
}

export function createPowerPointIconAssetPreview(args: PowerPointIconAssetPreviewArgs): PowerPointAssetPreview {
  const preview = createSharedPreview("icon", {
    ...args,
    label: args.label ?? args.iconName,
  }, "Select the inserted icon shape and run verify_slide_visual to confirm icon asset metadata.");
  preview.iconId = cleanText(args.iconId);
  preview.iconGlyph = cleanText(args.iconGlyph);
  return preview;
}

export function createPowerPointImageAssetPreview(args: PowerPointImageAssetPreviewArgs): PowerPointAssetPreview {
  const preview = createSharedPreview(
    "image",
    {
      ...args,
      label: args.label ?? args.sourceShapeName,
    },
    "Select the inserted image shape and run verify_slide_visual to inspect image-shape metadata.",
  );
  preview.sourceShapeId = cleanText(args.sourceShapeId);
  preview.sourceShapeName = cleanText(args.sourceShapeName);
  preview.sourceSlideId = cleanText(args.sourceSlideId);
  preview.sourceSlideIndex = finiteNumber(args.sourceSlideIndex);
  return preview;
}

export function createPowerPointComponentAssetPreview(args: PowerPointComponentAssetPreviewArgs): PowerPointAssetPreview {
  const preview = createSharedPreview(
    "component",
    {
      ...args,
      assetSourceId: "reusable-slide-component",
      label: args.label ?? args.componentLabel,
      insertionMode: args.insertionMode || "native-shape-component",
    },
    "Select the created component shapes and run verify_slide_visual to classify reusable component metadata.",
  );
  preview.componentId = cleanText(args.componentId);
  preview.createdShapeCount = finiteNumber(args.createdShapeCount);
  preview.componentTextPreview = args.componentPreview ? summarizePowerPointComponentPreview(args.componentPreview) : undefined;
  return preview;
}
