export interface PowerPointIconCatalogEntry {
  id: string;
  name: string;
  keywords: string[];
  glyph: string;
}

const POWERPOINT_ICON_CATALOG: PowerPointIconCatalogEntry[] = [
  { id: "trend-up", name: "Trend Up", keywords: ["growth", "up", "arrow", "chart", "revenue"], glyph: "📈" },
  { id: "trend-down", name: "Trend Down", keywords: ["decline", "down", "arrow", "chart"], glyph: "📉" },
  { id: "bar-chart", name: "Bar Chart", keywords: ["chart", "bar", "analytics", "data"], glyph: "📊" },
  { id: "target", name: "Target", keywords: ["goal", "focus", "objective", "bullseye"], glyph: "🎯" },
  { id: "rocket", name: "Rocket", keywords: ["launch", "growth", "speed", "startup"], glyph: "🚀" },
  { id: "shield", name: "Shield", keywords: ["security", "compliance", "guard", "protection"], glyph: "🛡️" },
  { id: "lock", name: "Lock", keywords: ["security", "privacy", "restricted", "safe"], glyph: "🔒" },
  { id: "globe", name: "Globe", keywords: ["global", "world", "internet", "web"], glyph: "🌐" },
  { id: "people", name: "People", keywords: ["team", "users", "audience", "customer"], glyph: "👥" },
  { id: "calendar", name: "Calendar", keywords: ["schedule", "timeline", "date", "plan"], glyph: "📅" },
  { id: "gear", name: "Gear", keywords: ["settings", "process", "automation", "system"], glyph: "⚙️" },
  { id: "lightbulb", name: "Lightbulb", keywords: ["idea", "insight", "innovation", "concept"], glyph: "💡" },
];

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
