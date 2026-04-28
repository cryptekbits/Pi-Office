export interface PowerPointIconCatalogEntry {
  id: string;
  name: string;
  keywords: string[];
  glyph: string;
  svgContent: string;
}

const POWERPOINT_ICON_CATALOG: PowerPointIconCatalogEntry[] = [
  {
    id: "trend-up",
    name: "Trend Up",
    keywords: ["growth", "up", "arrow", "chart", "revenue"],
    glyph: "📈",
    svgContent: '<polyline points="28 106 64 70 88 94 132 50"/><polyline points="96 50 132 50 132 86"/><line x1="28" y1="128" x2="132" y2="128"/>',
  },
  {
    id: "trend-down",
    name: "Trend Down",
    keywords: ["decline", "down", "arrow", "chart"],
    glyph: "📉",
    svgContent: '<polyline points="28 58 64 94 88 70 132 114"/><polyline points="96 114 132 114 132 78"/><line x1="28" y1="128" x2="132" y2="128"/>',
  },
  {
    id: "bar-chart",
    name: "Bar Chart",
    keywords: ["chart", "bar", "analytics", "data"],
    glyph: "📊",
    svgContent: '<line x1="28" y1="128" x2="132" y2="128"/><rect x="38" y="82" width="18" height="46" rx="3"/><rect x="70" y="58" width="18" height="70" rx="3"/><rect x="102" y="34" width="18" height="94" rx="3"/>',
  },
  {
    id: "target",
    name: "Target",
    keywords: ["goal", "focus", "objective", "bullseye"],
    glyph: "🎯",
    svgContent: '<circle cx="80" cy="80" r="50"/><circle cx="80" cy="80" r="28"/><circle cx="80" cy="80" r="8"/><line x1="80" y1="18" x2="80" y2="36"/><line x1="80" y1="124" x2="80" y2="142"/><line x1="18" y1="80" x2="36" y2="80"/><line x1="124" y1="80" x2="142" y2="80"/>',
  },
  {
    id: "rocket",
    name: "Rocket",
    keywords: ["launch", "growth", "speed", "startup"],
    glyph: "🚀",
    svgContent: '<path d="M62 104 48 112 56 86 92 50c16-16 34-20 48-18 2 16-4 32-18 48l-36 36-26 8 8-14"/><circle cx="104" cy="68" r="8"/><path d="M48 112 34 126"/><path d="M68 124 54 138"/>',
  },
  {
    id: "shield",
    name: "Shield",
    keywords: ["security", "compliance", "guard", "protection"],
    glyph: "🛡️",
    svgContent: '<path d="M80 24 126 42v34c0 30-17 52-46 68-29-16-46-38-46-68V42z"/><path d="M58 82 74 98 104 62"/>',
  },
  {
    id: "lock",
    name: "Lock",
    keywords: ["security", "privacy", "restricted", "safe"],
    glyph: "🔒",
    svgContent: '<rect x="42" y="72" width="76" height="58" rx="10"/><path d="M58 72V56c0-16 10-28 22-28s22 12 22 28v16"/><line x1="80" y1="94" x2="80" y2="110"/>',
  },
  {
    id: "globe",
    name: "Globe",
    keywords: ["global", "world", "internet", "web"],
    glyph: "🌐",
    svgContent: '<circle cx="80" cy="80" r="54"/><path d="M28 80h104"/><path d="M80 26c18 18 28 36 28 54s-10 36-28 54"/><path d="M80 26c-18 18-28 36-28 54s10 36 28 54"/><path d="M42 48c24 10 52 10 76 0"/><path d="M42 112c24-10 52-10 76 0"/>',
  },
  {
    id: "people",
    name: "People",
    keywords: ["team", "users", "audience", "customer"],
    glyph: "👥",
    svgContent: '<circle cx="64" cy="58" r="18"/><circle cx="106" cy="64" r="14"/><path d="M34 124c4-26 18-40 30-40s26 14 30 40"/><path d="M88 122c4-18 14-28 25-28 8 0 18 8 23 28"/>',
  },
  {
    id: "calendar",
    name: "Calendar",
    keywords: ["schedule", "timeline", "date", "plan"],
    glyph: "📅",
    svgContent: '<rect x="34" y="40" width="92" height="90" rx="10"/><line x1="34" y1="66" x2="126" y2="66"/><line x1="58" y1="28" x2="58" y2="50"/><line x1="102" y1="28" x2="102" y2="50"/><line x1="56" y1="88" x2="104" y2="88"/><line x1="56" y1="110" x2="86" y2="110"/>',
  },
  {
    id: "gear",
    name: "Gear",
    keywords: ["settings", "process", "automation", "system"],
    glyph: "⚙️",
    svgContent: '<circle cx="80" cy="80" r="18"/><path d="M80 24v18"/><path d="M80 118v18"/><path d="M24 80h18"/><path d="M118 80h18"/><path d="M40 40l13 13"/><path d="M107 107l13 13"/><path d="M120 40l-13 13"/><path d="M53 107l-13 13"/><circle cx="80" cy="80" r="42"/>',
  },
  {
    id: "lightbulb",
    name: "Lightbulb",
    keywords: ["idea", "insight", "innovation", "concept"],
    glyph: "💡",
    svgContent: '<path d="M56 78c0-18 10-34 24-34s24 16 24 34c0 13-8 20-14 30H70c-6-10-14-17-14-30z"/><line x1="68" y1="118" x2="92" y2="118"/><line x1="70" y1="132" x2="90" y2="132"/><path d="M80 20v12"/><path d="M42 38l9 9"/><path d="M118 38l-9 9"/>',
  },
];

function escapeSvgAttribute(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function normalizeSvgColor(value: string | undefined, fallback: string): string {
  const candidate = value?.trim();
  if (!candidate) {
    return fallback;
  }

  if (/^#[0-9a-f]{3,8}$/i.test(candidate) || /^rgb(a)?\([^)]+\)$/i.test(candidate) || /^[a-z]+$/i.test(candidate)) {
    return candidate;
  }

  return fallback;
}

export function buildPowerPointIconSvg(
  icon: PowerPointIconCatalogEntry,
  options?: { color?: string | undefined; strokeWidth?: number | undefined },
): string {
  const color = escapeSvgAttribute(normalizeSvgColor(options?.color, "#2563EB"));
  const strokeWidth = Number.isFinite(options?.strokeWidth) ? Math.max(4, Math.min(14, options?.strokeWidth ?? 8)) : 8;
  return [
    '<svg xmlns="http://www.w3.org/2000/svg" width="160" height="160" viewBox="0 0 160 160" role="img">',
    `<title>${escapeSvgAttribute(icon.name)}</title>`,
    `<g fill="none" stroke="${color}" stroke-width="${strokeWidth}" stroke-linecap="round" stroke-linejoin="round">`,
    icon.svgContent,
    "</g>",
    "</svg>",
  ].join("");
}

function tokenizeIconQuery(value: string): string[] {
  return value
    .toLowerCase()
    .split(/[\s,.;:|/_-]+/g)
    .map((token) => token.trim())
    .filter(Boolean);
}

export function searchPowerPointIcons(query: string, maxResults: number): PowerPointIconCatalogEntry[] {
  const tokens = tokenizeIconQuery(query);
  if (!tokens.length) {
    return [];
  }

  const scored = POWERPOINT_ICON_CATALOG.map((icon) => {
    const haystack = `${icon.id} ${icon.name} ${icon.keywords.join(" ")}`.toLowerCase();
    let score = 0;
    for (const token of tokens) {
      if (icon.id.includes(token)) score += 4;
      if (icon.name.toLowerCase().includes(token)) score += 3;
      if (icon.keywords.some((keyword) => keyword.includes(token))) score += 2;
      if (haystack.includes(token)) score += 1;
    }
    return { icon, score };
  })
    .filter((entry) => entry.score > 0)
    .sort((left, right) => right.score - left.score);

  return scored.slice(0, maxResults).map((entry) => entry.icon);
}

export function resolvePowerPointIcon(iconQuery: string): PowerPointIconCatalogEntry | undefined {
  const normalized = iconQuery.trim().toLowerCase();
  if (!normalized) {
    return undefined;
  }

  const exact =
    POWERPOINT_ICON_CATALOG.find((icon) => icon.id === normalized) ??
    POWERPOINT_ICON_CATALOG.find((icon) => icon.name.toLowerCase() === normalized) ??
    POWERPOINT_ICON_CATALOG.find((icon) => icon.keywords.some((keyword) => keyword === normalized));
  if (exact) {
    return exact;
  }

  return searchPowerPointIcons(iconQuery, 1)[0];
}
