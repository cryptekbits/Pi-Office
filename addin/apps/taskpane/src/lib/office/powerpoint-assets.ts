import { isRecord, trimString } from "./shared";

export type PowerPointAssetSourceStatus = "available" | "available-through-composition" | "planned";

export interface PowerPointAssetSourceDefinition {
  id: string;
  label: string;
  status: PowerPointAssetSourceStatus;
  inputFormats: string[];
  insertionPath: string;
  verificationPath: string;
  fallback?: string | undefined;
}

export interface PowerPointVerifiedAssetDescriptor {
  shapeId?: string | undefined;
  shapeName?: string | undefined;
  slideId?: string | undefined;
  slideIndex?: number | undefined;
  contentKind?: string | undefined;
  sourceId: string;
  sourceLabel: string;
  assetName?: string | undefined;
  componentId?: string | undefined;
  componentPart?: string | undefined;
  expectedInsertionMode: string;
  verificationStatus: "verified-image-shape" | "selected-image-shape" | "verified-component-shape" | "fallback-or-unknown";
  fallbackStrategy?: string | undefined;
}

const GENERATED_IMAGE_ALT_PREFIX = "Pi-Office generated image from generate_image.";

const POWERPOINT_REUSABLE_COMPONENT_IDS = ["metric-card", "quote-callout", "section-divider"];
const POWERPOINT_REUSABLE_COMPONENT_LABELS: Record<string, string> = {
  "metric-card": "Metric card",
  "quote-callout": "Quote callout",
  "section-divider": "Section divider",
};

export const POWERPOINT_ASSET_SOURCE_CATALOG: readonly PowerPointAssetSourceDefinition[] = [
  {
    id: "built-in-icon-svg",
    label: "Built-in SVG icon catalog",
    status: "available",
    inputFormats: ["original SVG geometry"],
    insertionPath: "Rasterized to a PowerPoint image shape with addImage where supported.",
    verificationPath: "verify_slide_visual can confirm the selected inserted shape is an image and preserve Pi-Office icon alt text.",
    fallback: "Glyph text boxes are only used when allowGlyphFallback=true or fallback is glyph-textbox.",
  },
  {
    id: "provided-image-base64",
    label: "Provided image payload",
    status: "available",
    inputFormats: ["PNG base64"],
    insertionPath: "Inserted or used as a target shape image through PowerPoint image APIs.",
    verificationPath: "verify_slide_visual can return the selected image shape snapshot and shape metadata.",
  },
  {
    id: "powerpoint-shape-snapshot",
    label: "Existing PowerPoint shape snapshot",
    status: "available",
    inputFormats: ["PowerPoint shape PNG snapshot"],
    insertionPath: "Copied from an existing shape snapshot into a destination slide or image placeholder.",
    verificationPath: "verify_slide_visual can inspect the copied image shape on the destination slide.",
  },
  {
    id: "selected-image-shape",
    label: "Selected PowerPoint image shape",
    status: "available",
    inputFormats: ["existing selected image shape"],
    insertionPath: "Existing presentation content; source provenance cannot be inferred from Office.js shape metadata alone.",
    verificationPath: "verify_slide_visual can confirm the selected shape is an image and return its shape metadata.",
  },
  {
    id: "generated-image-base64",
    label: "Generated image handoff",
    status: "available-through-composition",
    inputFormats: ["OpenAI image generation result as PNG base64"],
    insertionPath: "Use generate_image first, then insert the returned base64 through the PowerPoint image insertion path.",
    verificationPath: "verify_slide_visual can inspect the inserted image shape; provenance depends on the generating tool result.",
  },
  {
    id: "reusable-slide-component",
    label: "Reusable slide component",
    status: "available",
    inputFormats: [...POWERPOINT_REUSABLE_COMPONENT_IDS],
    insertionPath: "Native shape components inserted through insert_slide_element operation add_reusable_component.",
    verificationPath: "verify_slide_visual can classify selected component shapes from Pi-Office alt text metadata.",
  },
];

export function getPowerPointAssetSourceCatalog(): PowerPointAssetSourceDefinition[] {
  return POWERPOINT_ASSET_SOURCE_CATALOG.map((source) => ({ ...source, inputFormats: [...source.inputFormats] }));
}

function toNumber(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function sourceLabel(sourceId: string): string {
  return POWERPOINT_ASSET_SOURCE_CATALOG.find((source) => source.id === sourceId)?.label ?? sourceId;
}

export function inferPowerPointAssetDescriptor(shape: unknown): PowerPointVerifiedAssetDescriptor | undefined {
  if (!isRecord(shape)) {
    return undefined;
  }

  const contentKind = trimString(shape.contentKind);
  const shapeId = trimString(shape.id) ?? trimString(shape.shapeId);
  const shapeName = trimString(shape.name) ?? trimString(shape.shapeName);
  const altTextTitle = trimString(shape.altTextTitle);
  const altTextDescription = trimString(shape.altTextDescription);
  const slideId = trimString(shape.slideId);
  const slideIndex = toNumber(shape.slideIndex);
  const iconMatch = altTextDescription?.match(/^Pi-Office built-in (.+) icon from the SVG icon catalog\.$/i);
  const iconNameMatch = shapeName?.match(/^Icon\s+(.+)$/i);
  const generatedImageMatch = altTextDescription?.startsWith(GENERATED_IMAGE_ALT_PREFIX);
  const reusableComponentMatch = altTextDescription?.match(/^Pi-Office reusable slide component ([a-z0-9-]+)(?:: (.+))?\.$/i);

  if (iconMatch) {
    const assetName = iconMatch[1] ?? altTextTitle ?? iconNameMatch?.[1];
    return {
      shapeId,
      shapeName,
      slideId,
      slideIndex,
      contentKind,
      sourceId: "built-in-icon-svg",
      sourceLabel: sourceLabel("built-in-icon-svg"),
      assetName,
      expectedInsertionMode: contentKind === "image" ? "icon-image-shape" : "glyph-textbox",
      verificationStatus: contentKind === "image" ? "verified-image-shape" : "fallback-or-unknown",
      fallbackStrategy: contentKind === "image" ? undefined : "explicit-glyph-textbox-or-legacy-shape",
    };
  }

  if (iconNameMatch) {
    return {
      shapeId,
      shapeName,
      slideId,
      slideIndex,
      contentKind,
      sourceId: "selected-image-shape",
      sourceLabel: sourceLabel("selected-image-shape"),
      assetName: altTextTitle ?? iconNameMatch[1],
      expectedInsertionMode: contentKind === "image" ? "image-shape" : "glyph-textbox",
      verificationStatus: contentKind === "image" ? "selected-image-shape" : "fallback-or-unknown",
      fallbackStrategy: "name-only-icon-metadata",
    };
  }

  if (generatedImageMatch) {
    return {
      shapeId,
      shapeName,
      slideId,
      slideIndex,
      contentKind,
      sourceId: "generated-image-base64",
      sourceLabel: sourceLabel("generated-image-base64"),
      assetName: altTextTitle ?? shapeName ?? "Generated Image",
      expectedInsertionMode: "image-shape",
      verificationStatus: contentKind === "image" ? "verified-image-shape" : "fallback-or-unknown",
      fallbackStrategy: contentKind === "image" ? undefined : "generated-image-metadata-on-non-image-shape",
    };
  }

  if (reusableComponentMatch) {
    const componentId = reusableComponentMatch[1] ?? "";
    const componentPart = reusableComponentMatch[2];
    return {
      shapeId,
      shapeName,
      slideId,
      slideIndex,
      contentKind,
      sourceId: "reusable-slide-component",
      sourceLabel: sourceLabel("reusable-slide-component"),
      assetName: POWERPOINT_REUSABLE_COMPONENT_LABELS[componentId] ?? altTextTitle ?? shapeName,
      componentId,
      componentPart,
      expectedInsertionMode: "native-shape-component",
      verificationStatus: "verified-component-shape",
    };
  }

  if (contentKind === "image") {
    return {
      shapeId,
      shapeName,
      slideId,
      slideIndex,
      contentKind,
      sourceId: "selected-image-shape",
      sourceLabel: sourceLabel("selected-image-shape"),
      assetName: altTextTitle ?? shapeName,
      expectedInsertionMode: "image-shape",
      verificationStatus: "selected-image-shape",
    };
  }

  return undefined;
}
