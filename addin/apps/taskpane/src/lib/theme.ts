import type { CSSProperties } from "react";
import type { OfficeStateUpdate } from "@pi-office/pi-office-pack/protocol";
import type { OfficeThemeSnapshot } from "./office";

export type ThemeStyle = CSSProperties & Record<`--${string}`, string>;

export function normalizeHex(input: string | undefined, fallback: string): string {
  const value = input?.trim();
  if (!value) return fallback;

  const hex = value.startsWith("#") ? value.slice(1) : value;
  if (/^[0-9a-fA-F]{6}$/.test(hex)) return `#${hex.toUpperCase()}`;
  if (/^[0-9a-fA-F]{3}$/.test(hex)) {
    return `#${hex.split("").map((p) => `${p}${p}`).join("").toUpperCase()}`;
  }
  return fallback;
}

export function hexToRgb(input: string): [number, number, number] {
  const n = normalizeHex(input, "#000000").slice(1);
  return [
    Number.parseInt(n.slice(0, 2), 16),
    Number.parseInt(n.slice(2, 4), 16),
    Number.parseInt(n.slice(4, 6), 16),
  ];
}

export function withAlpha(input: string, alpha: number): string {
  const [r, g, b] = hexToRgb(input);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

export function mixColors(left: string, right: string, ratio: number): string {
  const [lr, lg, lb] = hexToRgb(left);
  const [rr, rg, rb] = hexToRgb(right);
  const c = Math.max(0, Math.min(1, ratio));
  const blend = (s: number, t: number) => Math.round(s + (t - s) * c);
  return `#${[blend(lr, rr), blend(lg, rg), blend(lb, rb)].map((p) => p.toString(16).padStart(2, "0")).join("").toUpperCase()}`;
}

export function getHostAccent(host: OfficeStateUpdate["host"] | undefined): string {
  switch (host) {
    case "excel": return "#107C41";
    case "powerpoint": return "#C43E1C";
    case "word":
    default: return "#185ABD";
  }
}

export function buildThemeStyle(
  theme: OfficeThemeSnapshot | undefined,
  host: OfficeStateUpdate["host"] | undefined,
): ThemeStyle {
  const bodyBg = normalizeHex(theme?.bodyBackgroundColor, "#F3F2F1");
  const bodyFg = normalizeHex(theme?.bodyForegroundColor, "#201F1E");
  const ctrlBg = normalizeHex(theme?.controlBackgroundColor, "#FFFFFF");
  const ctrlFg = normalizeHex(theme?.controlForegroundColor, "#201F1E");
  const dark = theme?.isDarkTheme ?? false;
  const accent = getHostAccent(host);

  return {
    "--accent": accent,
    "--shell-bg": mixColors(bodyBg, dark ? "#0B0B0B" : "#FFFFFF", dark ? 0.16 : 0.22),
    "--surface": mixColors(ctrlBg, bodyBg, dark ? 0.24 : 0.46),
    "--surface-elevated": ctrlBg,
    "--surface-muted": withAlpha(bodyFg, dark ? 0.12 : 0.045),
    "--surface-queue": withAlpha(accent, dark ? 0.16 : 0.07),
    "--ink": bodyFg,
    "--muted": withAlpha(bodyFg, dark ? 0.72 : 0.66),
    "--line": withAlpha(bodyFg, dark ? 0.22 : 0.12),
    "--line-strong": withAlpha(bodyFg, dark ? 0.38 : 0.22),
    "--button-bg": ctrlBg,
    "--button-fg": ctrlFg,
    "--accent-ring": withAlpha(accent, dark ? 0.24 : 0.14),
    "--chat-user-bg": accent,
    "--chat-user-fg": "#FFFFFF",
    "--status-ready": withAlpha("#107C10", dark ? 0.28 : 0.14),
    "--status-offline": withAlpha("#A4262C", dark ? 0.32 : 0.14),
    "--status-connecting": withAlpha(accent, dark ? 0.28 : 0.14),
    "--settings-backdrop": withAlpha("#000000", dark ? 0.42 : 0.22),
    "--shadow": dark ? "0 10px 22px rgba(0, 0, 0, 0.24)" : "0 6px 16px rgba(0, 0, 0, 0.08)",
  };
}
