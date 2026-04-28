export interface PowerPointReusableSlideComponentDefinition {
  id: string;
  label: string;
  aliases: string[];
  description: string;
  createdParts: string[];
  textSlots: string[];
  defaultSize: { width: number; height: number };
  insertionPath: string;
  verificationPath: string;
}

const POWERPOINT_REUSABLE_SLIDE_COMPONENTS: readonly PowerPointReusableSlideComponentDefinition[] = [
  {
    id: "metric-card",
    label: "Metric card",
    aliases: ["kpi-card", "stat-card", "number-card", "metric"],
    description: "A deck-ready KPI card with a quiet container, accent rail, label, value, and optional change line.",
    createdParts: ["container", "accent rail", "label", "value", "delta"],
    textSlots: ["metricLabel", "metricValue", "metricDelta"],
    defaultSize: { width: 220, height: 128 },
    insertionPath: "insert_slide_element with operation add_reusable_component and componentId metric-card.",
    verificationPath: "verify_slide_visual can classify selected component shapes from Pi-Office alt text metadata.",
  },
  {
    id: "quote-callout",
    label: "Quote callout",
    aliases: ["quote-card", "callout", "testimonial"],
    description: "A polished quote block with an accent rail, quote text, and optional attribution line.",
    createdParts: ["container", "accent rail", "quote", "attribution"],
    textSlots: ["quote", "attribution"],
    defaultSize: { width: 420, height: 150 },
    insertionPath: "insert_slide_element with operation add_reusable_component and componentId quote-callout.",
    verificationPath: "verify_slide_visual can classify selected component shapes from Pi-Office alt text metadata.",
  },
  {
    id: "section-divider",
    label: "Section divider",
    aliases: ["section-header", "divider", "chapter-divider"],
    description: "A slide section header with a vertical accent block, eyebrow, title, and optional subtitle.",
    createdParts: ["accent block", "eyebrow", "title", "subtitle"],
    textSlots: ["eyebrow", "title", "subtitle"],
    defaultSize: { width: 560, height: 190 },
    insertionPath: "insert_slide_element with operation add_reusable_component and componentId section-divider.",
    verificationPath: "verify_slide_visual can classify selected component shapes from Pi-Office alt text metadata.",
  },
];

export function getPowerPointReusableSlideComponentCatalog(): PowerPointReusableSlideComponentDefinition[] {
  return POWERPOINT_REUSABLE_SLIDE_COMPONENTS.map((component) => ({
    ...component,
    aliases: [...component.aliases],
    createdParts: [...component.createdParts],
    textSlots: [...component.textSlots],
    defaultSize: { ...component.defaultSize },
  }));
}

export function resolvePowerPointReusableSlideComponent(
  componentId: string | undefined,
): PowerPointReusableSlideComponentDefinition | undefined {
  const normalized = componentId?.trim().toLowerCase();
  if (!normalized) {
    return undefined;
  }

  return POWERPOINT_REUSABLE_SLIDE_COMPONENTS.find((component) =>
    component.id === normalized ||
    component.label.toLowerCase() === normalized ||
    component.aliases.some((alias) => alias === normalized),
  );
}
