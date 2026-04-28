import type { OfficeHostAction } from "@pi-office/pi-office-pack/protocol";
import {
  resolvePowerPointReusableSlideComponent,
  type PowerPointReusableSlideComponentDefinition,
} from "../powerpoint-components";
import {
  applyPowerPointShapeProperties,
  applyPowerPointTextFrameProperties,
  resolvePowerPointSlide,
  summarizePowerPointShape,
} from "../powerpoint-helpers";
import { createPowerPointComponentAssetPreview } from "../powerpoint-asset-previews";
import { isRecord, toNumber, trimString } from "../shared";

const POWERPOINT_COMPONENT_ACTIONS = new Set(["addReusableComponent"]);

interface ComponentTextShape {
  shape: PowerPoint.Shape;
  textFrame: PowerPoint.TextFrame;
  text: string;
  fontSize: number;
  fontColor: string;
  bold?: boolean | undefined;
  margin?: number | undefined;
}

interface ComponentLayout {
  left: number;
  top: number;
  width: number;
  height: number;
  accentColor: string;
  fillColor: string;
  textColor: string;
  mutedTextColor: string;
  subtleFillColor: string;
}

export function isPowerPointComponentAction(type: string): boolean {
  return POWERPOINT_COMPONENT_ACTIONS.has(type);
}

function optionText(options: Record<string, unknown>, keys: string[], fallback = ""): string {
  const textSlots = isRecord(options.textSlots) ? options.textSlots : {};
  for (const key of keys) {
    const value = trimString(options[key]) ?? trimString(textSlots[key]);
    if (value) {
      return value;
    }
  }
  return fallback;
}

function componentAltText(component: PowerPointReusableSlideComponentDefinition, part: string): string {
  return `Pi-Office reusable slide component ${component.id}: ${part}.`;
}

function shapeOptions(left: number, top: number, width: number, height: number): PowerPoint.ShapeAddOptions {
  return { left, top, width, height };
}

function toLayout(
  component: PowerPointReusableSlideComponentDefinition,
  options: Record<string, unknown>,
): ComponentLayout {
  return {
    left: toNumber(options.left) ?? 72,
    top: toNumber(options.top) ?? 120,
    width: toNumber(options.width) ?? component.defaultSize.width,
    height: toNumber(options.height) ?? component.defaultSize.height,
    accentColor: trimString(options.accentColor) ?? trimString(options.lineColor) ?? "#2563EB",
    fillColor: trimString(options.fillColor) ?? "#F8FAFC",
    textColor: trimString(options.textColor) ?? trimString(options.fontColor) ?? "#111827",
    mutedTextColor: trimString(options.mutedTextColor) ?? "#64748B",
    subtleFillColor: trimString(options.subtleFillColor) ?? "#E0F2FE",
  };
}

function applyTextShape(textShape: ComponentTextShape): void {
  textShape.textFrame.load("isNullObject");
}

function finishTextShape(textShape: ComponentTextShape): void {
  const textFrame = textShape.textFrame;
  if (!textFrame || textFrame.isNullObject) {
    return;
  }
  textFrame.textRange.text = textShape.text;
  textFrame.textRange.font.size = textShape.fontSize;
  textFrame.textRange.font.color = textShape.fontColor;
  if (typeof textShape.bold === "boolean") {
    textFrame.textRange.font.bold = textShape.bold;
  }
  applyPowerPointTextFrameProperties(textFrame, {
    wordWrap: true,
    autoSizeSetting: "TextToFitShape",
    verticalAlignment: "Middle",
    topMargin: textShape.margin ?? 4,
    rightMargin: textShape.margin ?? 4,
    bottomMargin: textShape.margin ?? 4,
    leftMargin: textShape.margin ?? 4,
  });
}

function addComponentTextBox(
  slide: PowerPoint.Slide,
  component: PowerPointReusableSlideComponentDefinition,
  part: string,
  text: string,
  bounds: PowerPoint.ShapeAddOptions,
  style: {
    fontSize: number;
    fontColor: string;
    bold?: boolean | undefined;
    margin?: number | undefined;
  },
): ComponentTextShape {
  const shape = slide.shapes.addTextBox("", bounds);
  const textFrame = shape.getTextFrameOrNullObject();
  applyPowerPointShapeProperties(shape, {
    name: `Component ${component.label} - ${part}`,
    altTextTitle: `${component.label} ${part}`,
    altTextDescription: componentAltText(component, part),
    clearFill: true,
    lineVisible: false,
  });
  return { shape, textFrame, text, ...style };
}

function addMetricCard(
  slide: PowerPoint.Slide,
  component: PowerPointReusableSlideComponentDefinition,
  layout: ComponentLayout,
  options: Record<string, unknown>,
): { shapes: PowerPoint.Shape[]; textShapes: ComponentTextShape[]; preview: Record<string, unknown> } {
  const metricLabel = optionText(options, ["metricLabel", "label", "eyebrow"], "Metric");
  const metricValue = optionText(options, ["metricValue", "value", "headline"], "42%");
  const metricDelta = optionText(options, ["metricDelta", "delta", "supportingText"], "Change vs. plan");

  const container = slide.shapes.addGeometricShape("RoundRectangle" as PowerPoint.GeometricShapeType, shapeOptions(
    layout.left,
    layout.top,
    layout.width,
    layout.height,
  ));
  applyPowerPointShapeProperties(container, {
    name: `Component ${component.label} - Container`,
    altTextTitle: component.label,
    altTextDescription: componentAltText(component, "container"),
    fillColor: layout.fillColor,
    lineColor: "#CBD5E1",
    lineWeight: 1,
  });

  const rail = slide.shapes.addGeometricShape("Rectangle" as PowerPoint.GeometricShapeType, shapeOptions(
    layout.left,
    layout.top,
    7,
    layout.height,
  ));
  applyPowerPointShapeProperties(rail, {
    name: `Component ${component.label} - Accent rail`,
    altTextTitle: `${component.label} accent`,
    altTextDescription: componentAltText(component, "accent rail"),
    fillColor: layout.accentColor,
    lineVisible: false,
  });

  const textShapes = [
    addComponentTextBox(slide, component, "label", metricLabel, shapeOptions(layout.left + 22, layout.top + 14, layout.width - 36, 28), {
      fontSize: 12,
      fontColor: layout.mutedTextColor,
      bold: true,
    }),
    addComponentTextBox(slide, component, "value", metricValue, shapeOptions(layout.left + 22, layout.top + 43, layout.width - 36, 50), {
      fontSize: 30,
      fontColor: layout.textColor,
      bold: true,
    }),
    addComponentTextBox(slide, component, "delta", metricDelta, shapeOptions(layout.left + 22, layout.top + 94, layout.width - 36, 24), {
      fontSize: 11,
      fontColor: layout.mutedTextColor,
    }),
  ];

  return {
    shapes: [container, rail, ...textShapes.map((entry) => entry.shape)],
    textShapes,
    preview: { metricLabel, metricValue, metricDelta },
  };
}

function addQuoteCallout(
  slide: PowerPoint.Slide,
  component: PowerPointReusableSlideComponentDefinition,
  layout: ComponentLayout,
  options: Record<string, unknown>,
): { shapes: PowerPoint.Shape[]; textShapes: ComponentTextShape[]; preview: Record<string, unknown> } {
  const quote = optionText(options, ["quote", "body", "text"], "A clear quote or customer insight goes here.");
  const attribution = optionText(options, ["attribution", "source", "caption"], "");

  const container = slide.shapes.addGeometricShape("RoundRectangle" as PowerPoint.GeometricShapeType, shapeOptions(
    layout.left,
    layout.top,
    layout.width,
    layout.height,
  ));
  applyPowerPointShapeProperties(container, {
    name: `Component ${component.label} - Container`,
    altTextTitle: component.label,
    altTextDescription: componentAltText(component, "container"),
    fillColor: layout.fillColor,
    lineColor: "#CBD5E1",
    lineWeight: 1,
  });

  const rail = slide.shapes.addGeometricShape("Rectangle" as PowerPoint.GeometricShapeType, shapeOptions(
    layout.left + 18,
    layout.top + 22,
    6,
    layout.height - 44,
  ));
  applyPowerPointShapeProperties(rail, {
    name: `Component ${component.label} - Accent rail`,
    altTextTitle: `${component.label} accent`,
    altTextDescription: componentAltText(component, "accent rail"),
    fillColor: layout.accentColor,
    lineVisible: false,
  });

  const textShapes = [
    addComponentTextBox(slide, component, "quote", quote, shapeOptions(layout.left + 38, layout.top + 24, layout.width - 62, layout.height - 62), {
      fontSize: 20,
      fontColor: layout.textColor,
      bold: true,
      margin: 2,
    }),
  ];
  if (attribution) {
    textShapes.push(
      addComponentTextBox(slide, component, "attribution", attribution, shapeOptions(layout.left + 38, layout.top + layout.height - 38, layout.width - 62, 24), {
        fontSize: 11,
        fontColor: layout.mutedTextColor,
      }),
    );
  }

  return {
    shapes: [container, rail, ...textShapes.map((entry) => entry.shape)],
    textShapes,
    preview: { quote, attribution: attribution || undefined },
  };
}

function addSectionDivider(
  slide: PowerPoint.Slide,
  component: PowerPointReusableSlideComponentDefinition,
  layout: ComponentLayout,
  options: Record<string, unknown>,
): { shapes: PowerPoint.Shape[]; textShapes: ComponentTextShape[]; preview: Record<string, unknown> } {
  const eyebrow = optionText(options, ["eyebrow", "label"], "Section");
  const title = optionText(options, ["title", "headline"], "New section");
  const subtitle = optionText(options, ["subtitle", "supportingText"], "");

  const block = slide.shapes.addGeometricShape("Rectangle" as PowerPoint.GeometricShapeType, shapeOptions(
    layout.left,
    layout.top,
    24,
    layout.height,
  ));
  applyPowerPointShapeProperties(block, {
    name: `Component ${component.label} - Accent block`,
    altTextTitle: `${component.label} accent`,
    altTextDescription: componentAltText(component, "accent block"),
    fillColor: layout.accentColor,
    lineVisible: false,
  });

  const backdrop = slide.shapes.addGeometricShape("Rectangle" as PowerPoint.GeometricShapeType, shapeOptions(
    layout.left + 24,
    layout.top,
    layout.width - 24,
    layout.height,
  ));
  applyPowerPointShapeProperties(backdrop, {
    name: `Component ${component.label} - Backdrop`,
    altTextTitle: component.label,
    altTextDescription: componentAltText(component, "backdrop"),
    fillColor: layout.subtleFillColor,
    lineVisible: false,
  });

  const textShapes = [
    addComponentTextBox(slide, component, "eyebrow", eyebrow, shapeOptions(layout.left + 52, layout.top + 30, layout.width - 84, 24), {
      fontSize: 12,
      fontColor: layout.accentColor,
      bold: true,
    }),
    addComponentTextBox(slide, component, "title", title, shapeOptions(layout.left + 52, layout.top + 60, layout.width - 84, 62), {
      fontSize: 34,
      fontColor: layout.textColor,
      bold: true,
    }),
  ];
  if (subtitle) {
    textShapes.push(
      addComponentTextBox(slide, component, "subtitle", subtitle, shapeOptions(layout.left + 52, layout.top + 124, layout.width - 84, 34), {
        fontSize: 14,
        fontColor: layout.mutedTextColor,
      }),
    );
  }

  return {
    shapes: [backdrop, block, ...textShapes.map((entry) => entry.shape)],
    textShapes,
    preview: { eyebrow, title, subtitle: subtitle || undefined },
  };
}

export async function applyPowerPointComponentAction(
  context: PowerPoint.RequestContext,
  action: OfficeHostAction,
  type: string,
  options: Record<string, unknown>,
): Promise<unknown> {
  if (type !== "addReusableComponent") {
    throw new Error(`Unsupported PowerPoint component action: ${type}`);
  }

  const componentId =
    trimString(options.componentId) ??
    trimString(options.component) ??
    trimString(options.componentName) ??
    trimString(action.content);
  const component = resolvePowerPointReusableSlideComponent(componentId);
  if (!component) {
    throw new Error(
      `PowerPoint addReusableComponent requires a supported componentId: metric-card, quote-callout, or section-divider.`,
    );
  }

  const slide = await resolvePowerPointSlide(context, action.target, true);
  slide.load("id,index");
  const layout = toLayout(component, options);
  const result =
    component.id === "metric-card"
      ? addMetricCard(slide, component, layout, options)
      : component.id === "quote-callout"
        ? addQuoteCallout(slide, component, layout, options)
        : addSectionDivider(slide, component, layout, options);

  for (const textShape of result.textShapes) {
    applyTextShape(textShape);
  }
  await context.sync();
  for (const textShape of result.textShapes) {
    finishTextShape(textShape);
  }
  for (const shape of result.shapes) {
    shape.load("id,name,type");
  }
  await context.sync();

  context.presentation.setSelectedSlides([slide.id]);
  slide.setSelectedShapes(result.shapes.map((shape) => shape.id));
  await context.sync();

  return {
    ok: true,
    host: "powerpoint",
    action: type,
    slideId: slide.id,
    slideIndex: slide.index + 1,
    assetSourceId: "reusable-slide-component",
    componentId: component.id,
    componentLabel: component.label,
    componentDescription: component.description,
    insertionMode: "native-shape-component",
    createdShapeCount: result.shapes.length,
    createdShapes: result.shapes.map((shape) => summarizePowerPointShape(slide, shape)),
    componentPreview: result.preview,
    assetPreview: createPowerPointComponentAssetPreview({
      assetSourceId: "reusable-slide-component",
      componentId: component.id,
      componentLabel: component.label,
      componentPreview: result.preview,
      insertionMode: "native-shape-component",
      createdShapeCount: result.shapes.length,
      slideId: slide.id,
      slideIndex: slide.index + 1,
    }),
    verificationHint: "Select the created component shapes and run verify_slide_visual to classify reusable component metadata.",
  };
}
