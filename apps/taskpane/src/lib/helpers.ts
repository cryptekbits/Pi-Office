import type { OfficeStateUpdate, PromptImagePayload } from "@pi-office/pi-office-pack/protocol";

export type ChatRole = "user" | "assistant" | "system" | "error";

export interface ToolCallEntry {
  toolCallId: string;
  toolName: string;
  status: "running" | "done" | "error";
  errorMessage?: string | undefined;
  input?: unknown;
  result?: unknown;
}

export interface ChatImagePayload {
  data: string;
  mimeType: string;
}

export interface ThinkingSegment {
  text: string;
  startedAt: number;
  durationMs?: number;
}

export interface ChatEntry {
  id: string;
  role: ChatRole;
  text: string;
  pending?: boolean;
  thinking?: string | undefined;
  thinkingSegments?: ThinkingSegment[] | undefined;
  toolCalls?: ToolCallEntry[] | undefined;
  images?: ChatImagePayload[] | undefined;
}

export interface LocalQueueEntry {
  id: string;
  text: string;
  images?: PromptImagePayload[] | undefined;
  status: "queued" | "sending";
}

export interface RemoteQueueEntry {
  id: string;
  text: string;
  mode: "steer" | "followUp";
}

export const quickPrompts = [
  "Summarize the current selection and identify gaps.",
  "Rewrite the current selection for a more executive tone.",
  "Turn the current content into a review checklist.",
];

const visualPromptPattern =
  /\b(image|figure|chart|diagram|format|formatting|style|spacing|margin|margins|tab|tabs|ruler|layout|visual|theme|align|alignment|slide|screenshot|snapshot)\b/i;

export function createEntry(
  role: ChatRole,
  text: string,
  pending = false,
  extra?: Pick<ChatEntry, "thinking" | "toolCalls"> & { thinkingSegments?: ThinkingSegment[] },
): ChatEntry {
  return { id: crypto.randomUUID(), role, text, pending, ...extra };
}

export function parseApiError(raw: unknown): string {
  if (!raw) return "";
  const str = typeof raw === "string" ? raw : JSON.stringify(raw);
  // Try to extract a nested JSON error body (e.g. "400 {\"type\":\"error\",\"error\":{...}}")
  const jsonMatch = str.match(/\{[\s\S]*\}/);
  if (jsonMatch) {
    try {
      const parsed = JSON.parse(jsonMatch[0]);
      const msg = parsed?.error?.message ?? parsed?.message ?? parsed?.error ?? "";
      if (typeof msg === "string" && msg) return msg;
    } catch { /* not JSON */ }
  }
  return str;
}

export function connectionLabel(state: string): string {
  if (state === "ready") return "Bridge ready";
  if (state === "offline") return "Bridge offline";
  if (state === "reconnecting") return "Reconnecting";
  return "Connecting";
}

export function extractDisplayUserText(rawText: string): string {
  const marker = "\n\nUser request:\n";
  const index = rawText.lastIndexOf(marker);
  return (index === -1 ? rawText : rawText.slice(index + marker.length)).trim();
}

export function extractMessageText(content: unknown): string {
  if (typeof content === "string") return extractDisplayUserText(content);
  if (!Array.isArray(content)) return "";
  const text = content
    .filter(
      (part): part is { type: string; text: string } =>
        Boolean(part && typeof part === "object" && (part as { type?: string }).type === "text"),
    )
    .map((part) => part.text)
    .join("\n");
  return extractDisplayUserText(text);
}

export function hasNonTextMessageContent(content: unknown): boolean {
  if (!Array.isArray(content)) return false;
  return content.some((part) => {
    if (!part || typeof part !== "object") return false;
    const type = (part as { type?: unknown }).type;
    return typeof type === "string" && type !== "text";
  });
}

export function extractMessageImages(content: unknown): ChatImagePayload[] | undefined {
  if (!Array.isArray(content)) return undefined;
  const images = content
    .filter((part): part is { type: "image"; data: string; mimeType: string } =>
      Boolean(
        part &&
        typeof part === "object" &&
        (part as { type?: string }).type === "image" &&
        typeof (part as { data?: unknown }).data === "string",
      ),
    )
    .map((part) => ({ data: part.data, mimeType: part.mimeType || "image/png" }));
  return images.length ? images : undefined;
}

export function shouldAutoAttachVisuals(text: string, officeState: OfficeStateUpdate | undefined): boolean {
  if (!officeState) return false;
  if ((officeState.selection.imageCount ?? 0) > 0) return true;
  if (officeState.host === "powerpoint" && (officeState.selection.objectCount ?? 0) > 0) return true;
  return visualPromptPattern.test(text);
}

export function clampNumber(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

export function formatTokenCount(value: number | null | undefined): string {
  if (value == null || Number.isNaN(value)) return "?";
  if (value < 1_000) return String(Math.round(value));
  if (value < 10_000) return `${(value / 1_000).toFixed(1)}k`;
  if (value < 1_000_000) return `${Math.round(value / 1_000)}k`;
  return `${(value / 1_000_000).toFixed(value < 10_000_000 ? 1 : 0)}M`;
}

export function getContextRingColor(percent: number | null | undefined): string {
  if (percent == null || Number.isNaN(percent)) return "rgba(127, 140, 141, 0.48)";
  const progress = clampNumber(percent, 0, 100) / 100;
  const hue = 140 - progress * 140;
  return `hsl(${hue} 72% 46%)`;
}
