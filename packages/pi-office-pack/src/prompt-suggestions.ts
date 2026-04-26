import type { OfficeDocumentState, PromptSuggestion } from "./protocol.js";

export const MAX_PROMPT_SUGGESTIONS = 3;
const MAX_PROMPT_SUGGESTION_CHARS = 140;

export interface ParsePromptSuggestionOptions {
  documentState?: OfficeDocumentState | undefined;
  maxCount?: number | undefined;
}

export interface PromptSuggestionUiState {
  nextPromptSuggestionsEnabled: boolean;
  isBusy: boolean;
  draft: string;
  sessionId?: string | undefined;
  selectionFingerprint?: string | undefined;
}

const GENERIC_SUGGESTIONS = new Set([
  "continue",
  "continue please",
  "go on",
  "next",
  "next step",
  "next steps",
  "what next",
  "what should i ask next",
  "tell me more",
  "can you elaborate",
  "elaborate",
  "explain more",
  "make it better",
  "improve it",
  "fix it",
  "do it",
  "ok",
  "okay",
  "yes",
  "no",
  "thanks",
  "thank you",
]);

const NOISY_PATTERNS = [
  /^https?:\/\//i,
  /^www\./i,
  /```/,
  /<[^>]+>/,
  /\{[\s\S]*\}/,
  /^\[?\s*no suggestions?\s*\]?$/i,
  /^n\/a$/i,
];

const UNSAVED_LOCAL_FILE_PATTERNS = [
  /\bAGENTS\.md\b/i,
  /\bSKILL\.md\b/i,
  /\b(local\s+)?(filesystem|workspace|repo|repository)\b/i,
  /\b(read|scan|search|list|open|inspect)\s+(the\s+)?(local\s+)?(files?|folders?|directories|workspace|repo|repository)\b/i,
  /\b(document\s+folder|saved\s+folder|file\s+system)\b/i,
];

function stripFence(text: string): string {
  const trimmed = text.trim();
  const match = trimmed.match(/^```(?:json)?\s*([\s\S]*?)\s*```$/i);
  return (match?.[1] ?? trimmed).trim();
}

function tryParseJson(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return undefined;
  }
}

function parseJsonCandidate(raw: string): unknown {
  const cleaned = stripFence(raw);
  const direct = tryParseJson(cleaned);
  if (direct !== undefined) return direct;

  const objectStart = cleaned.indexOf("{");
  const objectEnd = cleaned.lastIndexOf("}");
  if (objectStart >= 0 && objectEnd > objectStart) {
    const objectParsed = tryParseJson(cleaned.slice(objectStart, objectEnd + 1));
    if (objectParsed !== undefined) return objectParsed;
  }

  const arrayStart = cleaned.indexOf("[");
  const arrayEnd = cleaned.lastIndexOf("]");
  if (arrayStart >= 0 && arrayEnd > arrayStart) {
    const arrayParsed = tryParseJson(cleaned.slice(arrayStart, arrayEnd + 1));
    if (arrayParsed !== undefined) return arrayParsed;
  }

  return undefined;
}

function extractSuggestionItems(value: unknown): unknown[] {
  if (Array.isArray(value)) return value;
  if (!value || typeof value !== "object") return [];
  const record = value as Record<string, unknown>;
  if (Array.isArray(record.suggestions)) return record.suggestions;
  if (Array.isArray(record.prompts)) return record.prompts;
  if (Array.isArray(record.nextPrompts)) return record.nextPrompts;
  return [];
}

function extractSuggestionText(value: unknown): string | undefined {
  if (typeof value === "string") return value;
  if (!value || typeof value !== "object") return undefined;
  const record = value as Record<string, unknown>;
  for (const key of ["text", "prompt", "label"]) {
    const entry = record[key];
    if (typeof entry === "string") return entry;
  }
  return undefined;
}

function canonicalizeSuggestion(text: string): string {
  return text
    .toLowerCase()
    .replace(/["'`]+/g, "")
    .replace(/[^\w]+/g, " ")
    .trim();
}

function trimSuggestionText(text: string): string {
  return text
    .replace(/^[\s"'`*\-]+/g, "")
    .replace(/[\s"'`]+$/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function isUsefulSuggestion(text: string, documentState: OfficeDocumentState | undefined): boolean {
  if (!text) return false;
  if (text.length > MAX_PROMPT_SUGGESTION_CHARS) return false;
  if (NOISY_PATTERNS.some((pattern) => pattern.test(text))) return false;

  const canonical = canonicalizeSuggestion(text);
  if (!canonical) return false;
  if (GENERIC_SUGGESTIONS.has(canonical)) return false;
  if (canonical.split(/\s+/).length < 2) return false;

  if (documentState === "unsaved" && UNSAVED_LOCAL_FILE_PATTERNS.some((pattern) => pattern.test(text))) {
    return false;
  }

  return true;
}

function createPromptSuggestionId(text: string, index: number): string {
  let hash = 0;
  for (let i = 0; i < text.length; i += 1) {
    hash = ((hash << 5) - hash + text.charCodeAt(i)) | 0;
  }
  return `sug-${index + 1}-${Math.abs(hash).toString(36)}`;
}

export function parsePromptSuggestions(raw: string, options: ParsePromptSuggestionOptions = {}): PromptSuggestion[] {
  const parsed = parseJsonCandidate(raw);
  const items = extractSuggestionItems(parsed);
  const maxCount = Math.max(0, Math.min(options.maxCount ?? MAX_PROMPT_SUGGESTIONS, MAX_PROMPT_SUGGESTIONS));
  const seen = new Set<string>();
  const suggestions: PromptSuggestion[] = [];

  for (const item of items) {
    const text = trimSuggestionText(extractSuggestionText(item) ?? "");
    const canonical = canonicalizeSuggestion(text);
    if (!isUsefulSuggestion(text, options.documentState) || seen.has(canonical)) continue;
    seen.add(canonical);
    suggestions.push({ id: createPromptSuggestionId(text, suggestions.length), text });
    if (suggestions.length >= maxCount) break;
  }

  return suggestions;
}

export function shouldShowPromptSuggestions(
  suggestions: readonly PromptSuggestion[],
  state: Pick<PromptSuggestionUiState, "nextPromptSuggestionsEnabled" | "isBusy" | "draft">,
): boolean {
  return (
    state.nextPromptSuggestionsEnabled &&
    !state.isBusy &&
    !state.draft.trim() &&
    suggestions.length > 0
  );
}

export function shouldClearPromptSuggestionsForUiChange(
  previous: PromptSuggestionUiState | undefined,
  next: PromptSuggestionUiState,
): boolean {
  if (!next.nextPromptSuggestionsEnabled) return true;
  if (next.isBusy) return true;
  if (next.draft.trim()) return true;
  if (!previous) return false;
  if (previous.sessionId !== next.sessionId) return true;
  if (previous.selectionFingerprint !== next.selectionFingerprint) return true;
  return false;
}
