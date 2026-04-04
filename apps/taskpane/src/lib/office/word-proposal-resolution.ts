export interface WordParagraphLocator {
  index: number;
  text: string;
  paragraphId?: string | undefined;
}

export interface WordProposalLocatorInput {
  kind?: string | undefined;
  anchor?: string | undefined;
  paragraphId?: string | undefined;
  searchText?: string | undefined;
  oldText?: string | undefined;
}

export type WordProposalResolutionSource = "paragraphId" | "anchor" | "searchText" | "none";

export interface WordProposalResolution {
  index: number | undefined;
  via: WordProposalResolutionSource;
  query: string | undefined;
}

function trimString(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  const trimmed = value.trim();
  return trimmed ? trimmed : undefined;
}

function normalizeCaseFold(value: string): string {
  return value.toLocaleLowerCase();
}

function getEditQuery(edit: WordProposalLocatorInput): string | undefined {
  return trimString(edit.searchText) ?? trimString(edit.oldText);
}

export function formatParagraphAnchor(index: number): string {
  return `paragraph:${index + 1}`;
}

export function parseParagraphAnchorIndex(anchor: string | undefined): number | undefined {
  const value = trimString(anchor);
  if (!value) {
    return undefined;
  }

  const numericMatch = /^(?:paragraph|para|p)\s*[:#]\s*(\d+)$/i.exec(value);
  if (!numericMatch) {
    return undefined;
  }

  const parsed = Number.parseInt(numericMatch[1] ?? "", 10);
  if (!Number.isFinite(parsed) || parsed < 1) {
    return undefined;
  }

  return parsed - 1;
}

function paragraphContainsQuery(paragraph: WordParagraphLocator | undefined, query: string): boolean {
  if (!paragraph) {
    return false;
  }
  return paragraph.text.includes(query);
}

function findParagraphIndexByParagraphId(
  paragraphs: WordParagraphLocator[],
  paragraphId: string | undefined,
): number | undefined {
  if (!paragraphId) {
    return undefined;
  }
  return paragraphs.findIndex((paragraph) => paragraph.paragraphId === paragraphId);
}

function findParagraphIndexByAnchor(
  paragraphs: WordParagraphLocator[],
  anchor: string | undefined,
): number | undefined {
  const value = trimString(anchor);
  if (!value) {
    return undefined;
  }

  const numericAnchorIndex = parseParagraphAnchorIndex(value);
  if (typeof numericAnchorIndex === "number" && numericAnchorIndex >= 0 && numericAnchorIndex < paragraphs.length) {
    return numericAnchorIndex;
  }

  const normalizedAnchor = normalizeCaseFold(value);
  const paragraph = paragraphs.find((entry) => normalizeCaseFold(entry.text).includes(normalizedAnchor));
  return paragraph?.index;
}

function findParagraphIndexBySearchText(paragraphs: WordParagraphLocator[], query: string | undefined): number | undefined {
  if (!query) {
    return undefined;
  }

  const paragraph = paragraphs.find((entry) => paragraphContainsQuery(entry, query));
  return paragraph?.index;
}

function resolveLocatorCandidateIndex(
  paragraphs: WordParagraphLocator[],
  query: string | undefined,
  source: "paragraphId" | "anchor",
  candidateIndex: number | undefined,
): WordProposalResolution | undefined {
  if (typeof candidateIndex !== "number" || candidateIndex < 0) {
    return undefined;
  }

  const candidateParagraph = paragraphs[candidateIndex];
  if (!candidateParagraph) {
    return undefined;
  }

  if (!query || paragraphContainsQuery(candidateParagraph, query)) {
    return {
      index: candidateIndex,
      via: source,
      query,
    };
  }

  return undefined;
}

export function resolveAcceptedEditParagraphIndex(
  paragraphs: WordParagraphLocator[],
  edit: WordProposalLocatorInput,
): WordProposalResolution {
  const query = getEditQuery(edit);
  const paragraphId = trimString(edit.paragraphId);
  const anchor = trimString(edit.anchor);

  const byParagraphId = resolveLocatorCandidateIndex(
    paragraphs,
    query,
    "paragraphId",
    findParagraphIndexByParagraphId(paragraphs, paragraphId),
  );
  if (byParagraphId) {
    return byParagraphId;
  }

  const byAnchor = resolveLocatorCandidateIndex(
    paragraphs,
    query,
    "anchor",
    findParagraphIndexByAnchor(paragraphs, anchor),
  );
  if (byAnchor) {
    return byAnchor;
  }

  const searchTextIndex = findParagraphIndexBySearchText(paragraphs, query);
  if (typeof searchTextIndex === "number") {
    return {
      index: searchTextIndex,
      via: "searchText",
      query,
    };
  }

  return {
    index: undefined,
    via: "none",
    query,
  };
}

function getAllocationBucket(
  allocations: Map<string, Set<number>>,
  query: string,
): Set<number> {
  const existing = allocations.get(query);
  if (existing) {
    return existing;
  }

  const created = new Set<number>();
  allocations.set(query, created);
  return created;
}

export function resolveProposalParagraphIndex(
  paragraphs: WordParagraphLocator[],
  edit: WordProposalLocatorInput,
  allocations: Map<string, Set<number>>,
): WordProposalResolution {
  const query = getEditQuery(edit);
  const paragraphId = trimString(edit.paragraphId);
  const anchor = trimString(edit.anchor);

  const byParagraphId = resolveLocatorCandidateIndex(
    paragraphs,
    query,
    "paragraphId",
    findParagraphIndexByParagraphId(paragraphs, paragraphId),
  );
  if (byParagraphId) {
    return byParagraphId;
  }

  const byAnchor = resolveLocatorCandidateIndex(
    paragraphs,
    query,
    "anchor",
    findParagraphIndexByAnchor(paragraphs, anchor),
  );
  if (byAnchor) {
    return byAnchor;
  }

  if (query) {
    const candidates = paragraphs
      .filter((paragraph) => paragraphContainsQuery(paragraph, query))
      .map((paragraph) => paragraph.index);

    if (candidates.length > 0) {
      const used = getAllocationBucket(allocations, query);
      const nextCandidate = candidates.find((index) => !used.has(index)) ?? candidates[0];
      if (typeof nextCandidate !== "number") {
        return {
          index: undefined,
          via: "none",
          query,
        };
      }
      used.add(nextCandidate);
      return {
        index: nextCandidate,
        via: "searchText",
        query,
      };
    }
  }

  return {
    index: undefined,
    via: "none",
    query,
  };
}

export function buildParagraphContextPreview(
  paragraphText: string | undefined,
  query: string | undefined,
  radius = 40,
): string | undefined {
  const text = paragraphText ?? "";
  if (!text) {
    return undefined;
  }

  if (!query) {
    return text.slice(0, Math.min(text.length, radius * 2));
  }

  const directIndex = text.indexOf(query);
  const insensitiveIndex = directIndex >= 0 ? directIndex : normalizeCaseFold(text).indexOf(normalizeCaseFold(query));
  if (insensitiveIndex < 0) {
    return text.slice(0, Math.min(text.length, radius * 2));
  }

  const start = Math.max(0, insensitiveIndex - radius);
  const end = Math.min(text.length, insensitiveIndex + query.length + radius);
  return text.slice(start, end);
}
