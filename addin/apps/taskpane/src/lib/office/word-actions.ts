import type { OfficeAnchor, OfficeContextPayload, OfficeHostAction, OfficeStateUpdate, OfficeSelectionSummary } from "@pi-office/pi-office-pack/protocol";
import type { OfficeCaptureOptions } from "../office-host-adapter-types";
import {
  supportsRequirementSet,
  normalizeTextPreview,
  truncateText,
  pluralize,
  uniqueDetails,
  containsImageMarkup,
  formatPoints,
  formatPercent,
  serializeWordPages,
  buildSelectionSummary,
  buildStructuredSelectionPreview,
  buildSelectionMeta,
  trimString,
  truncateLabel,
  matchesTextQuery,
  parseAnchorOrdinal,
  uniqueAnchors,
  getActionOptions,
  toNumber,
  toBoolean,
  toStringMatrix,
  wordInsertLocationFromPlacement,
  isWordPictureShape,
  countSelectedWordImages,
  getSelectedTextAsync,
  getSelectedMarkupAsync,
  getSelectedImageAsync,
  optimizeVisual,
  getWordSnapshotWidthPx,
  renderHtmlSelectionSnapshot,
  createSummary,
  serializeOfficeRuntimeError,
} from "./shared";
import { matchesWordHeading, matchesWordParagraph } from "./word-navigate";
import { countWordOoxmlMathObjects, createWordEquationOoxml } from "./word-equations";

export type NormalizedWordBreakType =
  | "Page"
  | "Next"
  | "SectionNext"
  | "SectionContinuous"
  | "SectionEven"
  | "SectionOdd"
  | "Line";

export function normalizeWordBreakType(value: unknown): NormalizedWordBreakType {
  const raw = trimString(value) ?? "page";
  const compact = raw.replace(/[\s_-]/g, "").toLowerCase();
  switch (compact) {
    case "page":
      return "Page";
    case "next":
      return "Next";
    case "sectionnext":
      return "SectionNext";
    case "sectioncontinuous":
      return "SectionContinuous";
    case "sectioneven":
      return "SectionEven";
    case "sectionodd":
      return "SectionOdd";
    case "line":
      return "Line";
    default:
      throw new Error(
        `Unsupported Word breakType "${raw}". Use page, sectionNext, sectionContinuous, sectionEven, sectionOdd, or line.`,
      );
  }
}

export function countWordOoxmlBreaks(ooxml: string, breakType: NormalizedWordBreakType): number {
  if (breakType !== "Page") {
    return 0;
  }
  return (ooxml.match(/<w:br\b[^>]*\bw:type=["']page["'][^>]*(?:\/>|>)/gi) ?? []).length;
}

function decodeXmlText(value: string): string {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function normalizeOoxmlText(value: string | undefined): string {
  return (value ?? "").replace(/\s+/g, " ").trim();
}

function wordParagraphsFromOoxml(ooxml: string): string[] {
  return ooxml.match(/<w:p\b[\s\S]*?<\/w:p>/gi) ?? [];
}

function paragraphTextFromOoxml(paragraphOoxml: string): string {
  const textParts = Array.from(paragraphOoxml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/gi))
    .map((match) => decodeXmlText(match[1] ?? ""));
  const tabs = paragraphOoxml.match(/<w:tab\b[^>]*\/>/gi)?.length ?? 0;
  return normalizeOoxmlText(`${textParts.join("")}${tabs ? " ".repeat(tabs) : ""}`);
}

export function wordOoxmlHasAdjacentPageBreak(
  ooxml: string,
  targetText: string | undefined,
  placement: string | undefined,
): boolean {
  const normalizedTarget = normalizeOoxmlText(targetText);
  if (!normalizedTarget) {
    return false;
  }

  const paragraphs = wordParagraphsFromOoxml(ooxml);
  const targetIndex = paragraphs.findIndex((paragraph) => {
    const paragraphText = paragraphTextFromOoxml(paragraph);
    if (!paragraphText) {
      return false;
    }
    return paragraphText === normalizedTarget ||
      paragraphText.includes(normalizedTarget) ||
      normalizedTarget.includes(paragraphText);
  });
  if (targetIndex < 0) {
    return false;
  }
  if (countWordOoxmlBreaks(paragraphs[targetIndex] ?? "", "Page") > 0) {
    return true;
  }

  const direction = placement === "before" ? -1 : 1;
  for (let index = targetIndex + direction; index >= 0 && index < paragraphs.length; index += direction) {
    const paragraph = paragraphs[index] ?? "";
    if (countWordOoxmlBreaks(paragraph, "Page") > 0) {
      return true;
    }
    if (paragraphTextFromOoxml(paragraph)) {
      return false;
    }
  }
  return false;
}

export async function applyWordAction(action: OfficeHostAction): Promise<unknown> {
  return Word.run(async (context) => {
    const body = context.document.body;
    const type = trimString(action.type) ?? "insertText";
    const content = action.content ?? "";
    const placement = wordInsertLocationFromPlacement(action.placement);
    const options = getActionOptions(action);
    const supportsParagraphIds = supportsRequirementSet("WordApi", "1.6");
    const supportsFieldMetadata = supportsRequirementSet("WordApi", "1.5");
    const supportsContentControlSubtypes = supportsRequirementSet("WordApi", "1.3");
    let resolvedTarget:
      | {
          range: Word.Range;
          paragraph?: Word.Paragraph | undefined;
          comment?: Word.Comment | undefined;
          revision?: { change: Word.TrackedChange; revisionId: string; revisionIndex: number } | undefined;
          contentControl?: Word.ContentControl | undefined;
          field?: { field: Word.Field; fieldIndex: number } | undefined;
        }
      | undefined;

    const resolveWordNoteTarget = async (target: OfficeAnchor) => {
      if ((target.kind !== "footnote" && target.kind !== "endnote") || !supportsRequirementSet("WordApi", "1.5")) {
        return undefined;
      }

      const notes = target.kind === "footnote" ? body.footnotes : body.endnotes;
      notes.load("items/type,items/body/text,items/reference/text");
      await context.sync();

      const indexedNote = parseAnchorOrdinal(target.id, target.kind) ?? parseAnchorOrdinal(target.label, target.kind);
      const noteIndex =
        typeof indexedNote === "number"
          ? indexedNote - 1
          : notes.items.findIndex(
              (entry) =>
                matchesTextQuery(entry.body.text, target.text || target.label) ||
                matchesTextQuery(entry.reference.text, target.text || target.label),
            );
      const note = noteIndex >= 0 ? notes.items[noteIndex] : undefined;
      return note ? { note, noteIndex: noteIndex + 1 } : undefined;
    };

    const resolveWordCommentTarget = async (target: OfficeAnchor) => {
      if (!supportsRequirementSet("WordApi", "1.4")) {
        return undefined;
      }

      const comments = body.getComments();
      comments.load("items/id,items/content,items/authorName,items/resolved");
      await context.sync();

      return comments.items.find(
        (entry) =>
          (target.commentId && entry.id === target.commentId) ||
          matchesTextQuery(entry.content, target.text || target.label),
      );
    };

    const resolveWordRevisionTarget = async (target: OfficeAnchor) => {
      if (!supportsRequirementSet("WordApi", "1.6")) {
        return undefined;
      }

      const trackedChanges = body.getTrackedChanges();
      trackedChanges.load("items/author,items/date,items/text,items/type");
      await context.sync();

      const indexedRevision =
        parseAnchorOrdinal(target.revisionId, "revision") ??
        parseAnchorOrdinal(target.id, "revision") ??
        parseAnchorOrdinal(target.label, "revision");
      const changeIndex =
        typeof indexedRevision === "number"
          ? indexedRevision - 1
          : trackedChanges.items.findIndex((entry) => matchesTextQuery(entry.text, target.text || target.label));
      const change = changeIndex >= 0 ? trackedChanges.items[changeIndex] : undefined;
      return change ? { change, revisionId: `revision:${changeIndex + 1}`, revisionIndex: changeIndex + 1 } : undefined;
    };

    const resolveWordParagraphTarget = async (target: OfficeAnchor) => {
      const paragraphs = body.paragraphs;
      paragraphs.load(
        supportsParagraphIds
          ? "items/text,items/style,items/styleBuiltIn,items/uniqueLocalId"
          : "items/text,items/style,items/styleBuiltIn",
      );
      await context.sync();

      return target.kind === "heading"
        ? paragraphs.items.find((entry) => matchesWordHeading(entry, target))
        : paragraphs.items.find((entry) => matchesWordParagraph(entry, target));
    };

    const resolveWordContentControlTarget = async (target: OfficeAnchor) => {
      if (!supportsRequirementSet("WordApi", "1.1")) {
        return undefined;
      }

      const contentControls = body.contentControls;
      const directId = parseAnchorOrdinal(target.id, "contentControl");
      if (typeof directId === "number") {
        const directMatch = contentControls.getByIdOrNullObject(directId);
        directMatch.load("id,title,tag,text,type");
        await context.sync();
        if (!directMatch.isNullObject) {
          return directMatch;
        }
      }

      contentControls.load("items/id,items/title,items/tag,items/text,items/type");
      await context.sync();

      return contentControls.items.find(
        (entry) =>
          matchesTextQuery(entry.title, target.text || target.label) ||
          matchesTextQuery(entry.tag, target.text || target.label) ||
          matchesTextQuery(entry.text, target.text || target.label),
      );
    };

    const resolveWordFieldTarget = async (target: OfficeAnchor) => {
      if (!supportsRequirementSet("WordApi", "1.4")) {
        return undefined;
      }

      const fields = body.fields;
      fields.load(supportsFieldMetadata ? "items/code,items/type,items/result/text" : "items/code,items/result/text");
      await context.sync();

      const indexedField = parseAnchorOrdinal(target.id, "field") ?? parseAnchorOrdinal(target.label, "field");
      const fieldIndex =
        typeof indexedField === "number"
          ? indexedField - 1
          : fields.items.findIndex(
              (entry) =>
                matchesTextQuery(entry.code, target.text || target.label) ||
                matchesTextQuery(entry.result.text, target.text || target.label) ||
                (supportsFieldMetadata && matchesTextQuery(String(entry.type), target.text || target.label)),
            );
      const field = fieldIndex >= 0 ? fields.items[fieldIndex] : undefined;
      return field ? { field, fieldIndex: fieldIndex + 1 } : undefined;
    };

    const resolveWordSearchResultTarget = async (target: OfficeAnchor) => {
      const query = trimString(target.searchQuery) ?? trimString(target.text) ?? trimString(target.label);
      if (!query) {
        return undefined;
      }
      const matches = body.search(query, {
        matchCase: false,
        matchWholeWord: false,
      });
      matches.load("items/text");
      await context.sync();
      const index = typeof target.searchResultIndex === "number"
        ? Math.max(0, Math.trunc(target.searchResultIndex) - 1)
        : Math.max(0, Math.trunc(target.occurrenceIndex ?? 1) - 1);
      return matches.items[index];
    };

    const resolveWordBookmarkTarget = async (target: OfficeAnchor) => {
      if (!supportsRequirementSet("WordApiDesktop", "1.4")) {
        return undefined;
      }
      const bookmarks = context.document.bookmarks;
      bookmarks.load("items/name,items/start,items/end");
      await context.sync();
      return bookmarks.items.find((entry) =>
        matchesTextQuery(entry.name, target.bookmarkName || target.label || target.id || target.text),
      );
    };

    const resolveWordHyperlinkTarget = async (target: OfficeAnchor) => {
      if (!supportsRequirementSet("WordApiDesktop", "1.3")) {
        return undefined;
      }
      const hyperlinks = body.getRange("Content").hyperlinks;
      hyperlinks.load("items/address,items/subAddress,items/screenTip,items/textToDisplay,items/name");
      await context.sync();
      return hyperlinks.items.find((entry) =>
        matchesTextQuery(entry.address, target.hyperlinkAddress || target.text || target.label) ||
        matchesTextQuery(entry.textToDisplay, target.text || target.label) ||
        matchesTextQuery(entry.name, target.hyperlinkId || target.id),
      );
    };

    const resolveTargetRange = async () => {
      if (resolvedTarget) {
        return resolvedTarget;
      }

      const selection = context.document.getSelection();
      if (!action.target || action.target.kind === "selection") {
        resolvedTarget = { range: selection };
        return resolvedTarget;
      }

      if (action.target.kind === "comment") {
        const comment = await resolveWordCommentTarget(action.target);
        if (!comment) {
          throw new Error(`Could not find the requested Word comment: ${action.target.label || action.target.text || action.target.commentId || "comment"}.`);
        }
        resolvedTarget = { range: comment.getRange(), comment };
        return resolvedTarget;
      }

      if (action.target.kind === "revision") {
        const revision = await resolveWordRevisionTarget(action.target);
        if (!revision) {
          throw new Error(`Could not find the requested Word revision: ${action.target.label || action.target.text || action.target.revisionId || "revision"}.`);
        }
        resolvedTarget = { range: revision.change.getRange(), revision };
        return resolvedTarget;
      }

      if (action.target.kind === "heading" || action.target.kind === "paragraph") {
        const paragraph = await resolveWordParagraphTarget(action.target);
        if (!paragraph) {
          throw new Error(`Could not find the requested Word paragraph: ${action.target.label || action.target.text || action.target.paragraphId || action.target.kind}.`);
        }
        resolvedTarget = { range: paragraph.getRange(), paragraph };
        return resolvedTarget;
      }

      if (action.target.kind === "contentControl") {
        const contentControl = await resolveWordContentControlTarget(action.target);
        if (!contentControl) {
          throw new Error(`Could not find the requested Word content control: ${action.target.label || action.target.text || action.target.id || "content control"}.`);
        }
        resolvedTarget = { range: contentControl.getRange("Content"), contentControl };
        return resolvedTarget;
      }

      if (action.target.kind === "field") {
        const field = await resolveWordFieldTarget(action.target);
        if (!field) {
          throw new Error(`Could not find the requested Word field: ${action.target.label || action.target.text || action.target.id || "field"}.`);
        }
        resolvedTarget = { range: field.field.result, field };
        return resolvedTarget;
      }

      if (action.target.kind === "searchResult" || action.target.kind === "range") {
        const match = await resolveWordSearchResultTarget(action.target);
        if (!match) {
          throw new Error(`Could not find the requested Word search result: ${action.target.searchQuery || action.target.text || action.target.label || action.target.kind}.`);
        }
        resolvedTarget = { range: match };
        return resolvedTarget;
      }

      if (action.target.kind === "bookmark") {
        const bookmark = await resolveWordBookmarkTarget(action.target);
        if (!bookmark) {
          throw new Error(`Could not find the requested Word bookmark: ${action.target.bookmarkName || action.target.label || action.target.id || "bookmark"}.`);
        }
        bookmark.select();
        await context.sync();
        resolvedTarget = { range: context.document.getSelection() };
        return resolvedTarget;
      }

      if (action.target.kind === "hyperlink") {
        const hyperlink = await resolveWordHyperlinkTarget(action.target);
        if (!hyperlink) {
          throw new Error(`Could not find the requested Word hyperlink: ${action.target.hyperlinkAddress || action.target.text || action.target.label || "hyperlink"}.`);
        }
        resolvedTarget = { range: hyperlink.range };
        return resolvedTarget;
      }

      if (action.target.kind === "footnote" || action.target.kind === "endnote") {
        const note = await resolveWordNoteTarget(action.target);
        if (!note) {
          throw new Error(`Could not find the requested Word ${action.target.kind}: ${action.target.label || action.target.text || action.target.id || action.target.kind}.`);
        }
        const noteTarget = action.target.noteTarget === "reference" ? "reference" : "body";
        resolvedTarget = { range: noteTarget === "reference" ? note.note.reference : note.note.body.getRange("Content") };
        return resolvedTarget;
      }

      throw new Error(`Unsupported Word target anchor kind for this action: ${action.target.kind}.`);
    };

    const resolveCommentActionTarget = async () => {
      if (action.target?.kind === "comment") {
        const targetComment = await resolveWordCommentTarget(action.target);
        if (targetComment) {
          return targetComment;
        }
      }

      const selectionComments = context.document.getSelection().getComments();
      selectionComments.load("items/id,items/content,items/authorName,items/resolved");
      await context.sync();
      if (selectionComments.items[0]) {
        return selectionComments.items[0];
      }

      if (action.target) {
        const fallbackComment = await resolveWordCommentTarget(action.target);
        if (fallbackComment) {
          return fallbackComment;
        }
      }

      throw new Error("A Word comment target is required for this action.");
    };

    const resolveRevisionActionTarget = async () => {
      if (action.target?.kind === "revision") {
        const targetRevision = await resolveWordRevisionTarget(action.target);
        if (targetRevision) {
          return targetRevision;
        }
      }

      const selectionChanges = context.document.getSelection().getTrackedChanges();
      selectionChanges.load("items/author,items/date,items/text,items/type");
      await context.sync();
      if (selectionChanges.items[0]) {
        return {
          change: selectionChanges.items[0],
          revisionId: "revision:selection",
          revisionIndex: 1,
        };
      }

      if (action.target) {
        const fallbackRevision = await resolveWordRevisionTarget(action.target);
        if (fallbackRevision) {
          return fallbackRevision;
        }
      }

      throw new Error("A Word revision target is required for this action.");
    };

    if (type === "insertHtml") {
      if (action.target?.kind === "document") {
        const documentPlacement =
          action.placement === "start" || action.placement === "end" || action.placement === "replace"
            ? action.placement
            : "replace";
        const documentInsertLocation: Word.InsertLocation.start | Word.InsertLocation.end | Word.InsertLocation.replace =
          documentPlacement === "start"
            ? Word.InsertLocation.start
            : documentPlacement === "end"
              ? Word.InsertLocation.end
              : Word.InsertLocation.replace;
        body.insertHtml(content, documentInsertLocation);
        await context.sync();
        return { ok: true, host: "word", action: type, target: { kind: "document" }, placement: documentPlacement };
      }
      const { range } = await resolveTargetRange();
      range.insertHtml(content, placement);
      await context.sync();
      return { ok: true, host: "word", action: type, target: action.target, placement: action.placement ?? "replace" };
    }

    if (type === "insertText") {
      if (action.target?.kind === "document") {
        const documentPlacement =
          action.placement === "start" || action.placement === "end" || action.placement === "replace"
            ? action.placement
            : "replace";
        const documentInsertLocation: Word.InsertLocation.start | Word.InsertLocation.end | Word.InsertLocation.replace =
          documentPlacement === "start"
            ? Word.InsertLocation.start
            : documentPlacement === "end"
              ? Word.InsertLocation.end
              : Word.InsertLocation.replace;
        body.insertText(content, documentInsertLocation);
        await context.sync();
        return { ok: true, host: "word", action: type, target: { kind: "document" }, placement: documentPlacement };
      }
      const { range } = await resolveTargetRange();
      range.insertText(content, placement);
      await context.sync();
      return { ok: true, host: "word", action: type, target: action.target, placement: action.placement ?? "replace" };
    }

    if (type === "applyParagraphFormat" || type === "applyTextFormat" || type === "clearFormatting") {
      const { range, paragraph } = await resolveTargetRange();
      const changed: Record<string, unknown> = {};
      const warnings: string[] = [];

      if (type === "clearFormatting") {
        range.clear();
        await context.sync();
        return {
          ok: true,
          host: "word",
          action: type,
          target: action.target,
          summary: "Cleared formatting from the resolved Word range.",
        };
      }

      const style = trimString(options.style ?? action.style);
      const builtInStyle = trimString(options.styleBuiltIn ?? action.styleBuiltIn);
      if (style) {
        range.style = style;
        changed.style = style;
      }
      if (builtInStyle) {
        range.styleBuiltIn = builtInStyle as Word.BuiltInStyleName;
        changed.styleBuiltIn = builtInStyle;
      }

      const font = range.font;
      const fontName = trimString(options.fontName ?? action.fontName);
      const fontSize = toNumber(options.fontSize ?? action.fontSize);
      const color = trimString(options.color ?? action.color);
      const highlightColor = trimString(options.highlightColor ?? action.highlightColor);
      const bold = toBoolean(options.bold ?? action.bold);
      const italic = toBoolean(options.italic ?? action.italic);
      if (fontName) {
        font.name = fontName;
        changed.fontName = fontName;
      }
      if (typeof fontSize === "number") {
        font.size = fontSize;
        changed.fontSize = fontSize;
      }
      if (color) {
        font.color = color;
        changed.color = color;
      }
      if (highlightColor) {
        font.highlightColor = highlightColor;
        changed.highlightColor = highlightColor;
      }
      if (typeof bold === "boolean") {
        font.bold = bold;
        changed.bold = bold;
      }
      if (typeof italic === "boolean") {
        font.italic = italic;
        changed.italic = italic;
      }

      const paragraphCollection = paragraph ? undefined : range.paragraphs;
      if (paragraphCollection) {
        paragraphCollection.load("items/text");
        await context.sync();
      }

      const alignment = trimString(options.alignment ?? action.alignment);
      const leftIndent = toNumber(options.leftIndent ?? action.leftIndent);
      const rightIndent = toNumber(options.rightIndent ?? action.rightIndent);
      const firstLineIndent = toNumber(options.firstLineIndent ?? action.firstLineIndent);
      const lineSpacing = toNumber(options.lineSpacing ?? action.lineSpacing);
      const spaceBefore = toNumber(options.spaceBefore ?? action.spaceBefore);
      const spaceAfter = toNumber(options.spaceAfter ?? action.spaceAfter);
      const paragraphItems = paragraph ? [paragraph] : paragraphCollection?.items ?? [];
      for (const item of paragraphItems) {
        if (alignment) item.alignment = alignment as Word.Alignment;
        if (typeof leftIndent === "number") item.leftIndent = leftIndent;
        if (typeof rightIndent === "number") item.rightIndent = rightIndent;
        if (typeof firstLineIndent === "number") item.firstLineIndent = firstLineIndent;
        if (typeof lineSpacing === "number") item.lineSpacing = lineSpacing;
        if (typeof spaceBefore === "number") item.spaceBefore = spaceBefore;
        if (typeof spaceAfter === "number") item.spaceAfter = spaceAfter;
      }
      if (alignment) changed.alignment = alignment;
      if (typeof leftIndent === "number") changed.leftIndent = leftIndent;
      if (typeof rightIndent === "number") changed.rightIndent = rightIndent;
      if (typeof firstLineIndent === "number") changed.firstLineIndent = firstLineIndent;
      if (typeof lineSpacing === "number") changed.lineSpacing = lineSpacing;
      if (typeof spaceBefore === "number") changed.spaceBefore = spaceBefore;
      if (typeof spaceAfter === "number") changed.spaceAfter = spaceAfter;

      if (!Object.keys(changed).length) {
        warnings.push("No supported formatting properties were provided.");
      }

      await context.sync();
      return {
        ok: true,
        host: "word",
        action: type,
        target: action.target,
        changed,
        warnings,
      };
    }

    if (type === "applyListFormat") {
      if (!supportsRequirementSet("WordApiDesktop", "1.3")) {
        throw new Error("Word list formatting requires WordApiDesktop 1.3 and is currently desktop-only.");
      }
      const { range } = await resolveTargetRange();
      const listFormat = range.listFormat;
      const listKind = trimString(options.listKind ?? options.kind ?? action.listKind ?? action.kind) ?? "bullet";
      const level = toNumber(options.level ?? action.level);
      const changed: Record<string, unknown> = { listKind };

      if (listKind === "bullet") {
        listFormat.applyBulletDefault("Word2010" as Word.DefaultListBehavior);
      } else if (listKind === "numbered" || listKind === "number") {
        listFormat.applyNumberDefault("Word2010" as Word.DefaultListBehavior);
      } else if (listKind === "outline" || listKind === "multilevel") {
        listFormat.applyOutlineNumberDefault("Word2010" as Word.DefaultListBehavior);
      } else if (listKind === "remove" || listKind === "none") {
        listFormat.removeNumbers("AllNumbers");
      } else {
        throw new Error(`Unsupported Word list kind: ${listKind}.`);
      }

      if (typeof level === "number") {
        listFormat.listLevelNumber = Math.max(1, Math.min(9, Math.trunc(level)));
        changed.level = listFormat.listLevelNumber;
      }

      await context.sync();
      return {
        ok: true,
        host: "word",
        action: type,
        target: action.target,
        changed,
        capability: "WordApiDesktop 1.3",
        desktopOnly: true,
      };
    }

    if (type === "manageHyperlink") {
      if (!supportsRequirementSet("WordApiDesktop", "1.3")) {
        throw new Error("Word hyperlink actions require WordApiDesktop 1.3 or newer.");
      }
      const { range } = await resolveTargetRange();
      const operation = trimString(options.operation ?? action.operation) ?? "add";
      const hyperlinks = range.hyperlinks;
      hyperlinks.load("items/address,items/subAddress,items/screenTip,items/textToDisplay,items/name");
      await context.sync();

      const address = trimString(options.address ?? action.address);
      const textToDisplay = trimString(options.textToDisplay ?? action.textToDisplay);
      const screenTip = trimString(options.screenTip ?? action.screenTip);
      const subAddress = trimString(options.subAddress ?? action.subAddress);
      const targetName = trimString(options.name ?? action.name ?? action.text ?? action.label);
      const target = hyperlinks.items.find((entry) =>
        (targetName && (entry.name === targetName || entry.textToDisplay === targetName || entry.address === targetName)) ||
        (address && entry.address === address),
      ) ?? hyperlinks.items[0];

      if (operation === "remove" || operation === "delete") {
        if (!target) throw new Error("Could not find a Word hyperlink to remove.");
        const removed = target.textToDisplay || target.address || target.name;
        target.delete();
        await context.sync();
        return { ok: true, host: "word", action: type, operation: "remove", removed };
      }

      if (operation === "update") {
        if (!target) throw new Error("Could not find a Word hyperlink to update.");
        if (address) target.address = address;
        if (textToDisplay) target.textToDisplay = textToDisplay;
        if (screenTip) target.screenTip = screenTip;
        if (subAddress) target.subAddress = subAddress;
        await context.sync();
        return { ok: true, host: "word", action: type, operation: "update", address, textToDisplay, screenTip, subAddress };
      }

      if (!address && !subAddress) {
        throw new Error("Adding a Word hyperlink requires address or subAddress.");
      }
      const link = hyperlinks.add(range, {
        ...(address ? { address } : {}),
        ...(subAddress ? { subAddress } : {}),
        ...(screenTip ? { screenTip } : {}),
        ...(textToDisplay ? { textToDisplay } : {}),
      });
      link.load("address,subAddress,screenTip,textToDisplay,name");
      await context.sync();
      return {
        ok: true,
        host: "word",
        action: type,
        operation: "add",
        address: link.address,
        subAddress: link.subAddress,
        screenTip: link.screenTip,
        textToDisplay: link.textToDisplay,
        name: link.name,
      };
    }

    if (type === "bookmarkAction") {
      if (!supportsRequirementSet("WordApiDesktop", "1.4")) {
        throw new Error("Word bookmark actions require WordApiDesktop 1.4 or newer.");
      }
      const operation = trimString(options.operation ?? action.operation) ?? "inventory";
      const bookmarks = context.document.bookmarks;
      bookmarks.load("items/name,items/start,items/end");
      await context.sync();

      if (operation === "inventory") {
        return {
          ok: true,
          host: "word",
          action: type,
          operation,
          bookmarks: bookmarks.items.slice(0, 50).map((bookmark, index) => ({
            id: `bookmark:${index + 1}`,
            bookmarkName: bookmark.name,
            start: bookmark.start,
            end: bookmark.end,
          })),
          count: bookmarks.items.length,
          readOnly: true,
        };
      }

      const name = trimString(options.name ?? action.name ?? action.target?.bookmarkName ?? action.target?.label);
      if (!name) {
        throw new Error(`Word bookmark ${operation} requires a bookmark name.`);
      }
      const bookmark = bookmarks.items.find((entry) => matchesTextQuery(entry.name, name));
      if (operation === "exists") {
        return { ok: true, host: "word", action: type, operation, bookmarkName: name, exists: Boolean(bookmark) };
      }
      if (operation === "select") {
        if (!bookmark) throw new Error(`Could not find Word bookmark "${name}".`);
        bookmark.select();
        await context.sync();
        return { ok: true, host: "word", action: type, operation, bookmarkName: bookmark.name };
      }

      throw new Error(`Unsupported read-only Word bookmark operation: ${operation}.`);
    }

    if (type === "insertOoxml") {
      const { range } = await resolveTargetRange();
      range.insertOoxml(content, placement);
      await context.sync();
      return { ok: true, host: "word", action: type };
    }

    if (type === "insertEquation") {
      const latex = trimString(options.latex ?? action.latex ?? action.content);
      if (!latex) {
        throw new Error("word_equation requires a non-empty latex parameter.");
      }
      const display = trimString(options.display ?? action.display) === "inline" ? "inline" : "block";
      const equation = createWordEquationOoxml(latex, display);
      const beforeOoxml = body.getOoxml();
      await context.sync();
      const beforeMathCount = countWordOoxmlMathObjects(beforeOoxml.value);

      if (action.target?.kind === "document") {
        const documentPlacement =
          action.placement === "start" || action.placement === "end" || action.placement === "replace"
            ? action.placement
            : "end";
        const documentInsertLocation: Word.InsertLocation.start | Word.InsertLocation.end | Word.InsertLocation.replace =
          documentPlacement === "start"
            ? Word.InsertLocation.start
            : documentPlacement === "replace"
              ? Word.InsertLocation.replace
              : Word.InsertLocation.end;
        body.insertOoxml(equation.ooxml, documentInsertLocation);
      } else {
        const { range } = await resolveTargetRange();
        range.insertOoxml(equation.ooxml, placement);
      }

      const afterOoxml = body.getOoxml();
      await context.sync();
      const afterMathCount = countWordOoxmlMathObjects(afterOoxml.value);
      if (afterMathCount <= beforeMathCount) {
        throw new Error(
          "word_equation inserted OfficeMath OOXML but post-write verification did not find a new persisted m:oMath object. Do not claim rendered equation success without verify_doc or verify_doc_visual evidence.",
        );
      }
      return {
        ok: true,
        host: "word",
        action: type,
        target: action.target,
        placement: action.target?.kind === "document" ? (action.placement ?? "end") : (action.placement ?? "replace"),
        requestedFormat: "latex",
        resolvedFormat: "omml",
        display: equation.display,
        normalizedLatex: equation.normalizedLatex,
        warnings: equation.warnings.length ? equation.warnings : undefined,
        unsupportedCommands: equation.unsupportedCommands.length ? equation.unsupportedCommands : undefined,
        verification: {
          method: "body.getOoxml",
          persisted: true,
          beforeMathCount,
          afterMathCount,
          evidence: equation.evidence,
          note: "Equation success is based on persisted OfficeMath (m:oMath) evidence; use verify_doc_visual when visual equation layout matters.",
        },
      };
    }

    if (type === "addComment") {
      const { range } = await resolveTargetRange();
      const comment = range.insertComment(content);
      comment.load("id,content,resolved");
      await context.sync();
      return { ok: true, host: "word", action: type, commentId: comment.id, content: comment.content, resolved: comment.resolved };
    }

    if (type === "replyToComment") {
      const comment = await resolveCommentActionTarget();
      const reply = comment.reply(content);
      reply.load("id,content");
      await context.sync();
      return { ok: true, host: "word", action: type, commentId: comment.id, replyId: reply.id, content: reply.content };
    }

    if (type === "resolveComment" || type === "reopenComment" || type === "setCommentResolved") {
      const comment = await resolveCommentActionTarget();
      const resolved =
        type === "setCommentResolved"
          ? (toBoolean(options.resolved ?? action.resolved) ?? true)
          : type === "resolveComment";
      comment.resolved = resolved;
      await context.sync();
      return { ok: true, host: "word", action: type, commentId: comment.id, resolved };
    }

    if (type === "deleteComment") {
      const comment = await resolveCommentActionTarget();
      const commentId = comment.id;
      const existingContent = comment.content;
      comment.delete();
      await context.sync();
      return { ok: true, host: "word", action: type, commentId, content: existingContent };
    }

    if (type === "insertInlinePicture") {
      const { range } = await resolveTargetRange();
      const picture = range.insertInlinePictureFromBase64(content, placement);
      const altText = trimString(options.altText) ?? trimString(action.altText);
      if (altText) {
        picture.altTextTitle = altText;
      }
      picture.load("width,height,altTextTitle");
      await context.sync();
      return { ok: true, host: "word", action: type, width: picture.width, height: picture.height, altTextTitle: picture.altTextTitle };
    }

    if (type === "insertFileFromBase64") {
      const { range } = await resolveTargetRange();
      range.insertFileFromBase64(content, placement);
      await context.sync();
      return { ok: true, host: "word", action: type };
    }

    if (type === "insertTable") {
      const { range } = await resolveTargetRange();
      const values = toStringMatrix(options.values ?? action.values);
      const rowCount = toNumber(options.rowCount) ?? values?.length ?? 2;
      const columnCount = toNumber(options.columnCount) ?? values?.[0]?.length ?? 2;
      const tablePlacement = action.placement === "before" ? Word.InsertLocation.before : Word.InsertLocation.after;
      range.insertTable(rowCount, columnCount, tablePlacement, values);
      await context.sync();
      return { ok: true, host: "word", action: type, rowCount, columnCount };
    }

    const resolveWordTableTarget = async () => {
      const tableIndex = Math.max(0, Math.trunc(toNumber(options.tableIndex ?? action.tableIndex) ?? 0));
      const tables = body.tables;
      tables.load("items/rowCount,items/values,items/style,items/styleBuiltIn,title");
      await context.sync();
      const table = tables.items[tableIndex];
      if (!table) {
        throw new Error(`Could not find Word table at index ${tableIndex}.`);
      }
      return { table, tableIndex };
    };

    if (type === "inspectTable") {
      if (!supportsRequirementSet("WordApi", "1.3")) {
        throw new Error("Word table inspection requires WordApi 1.3 or newer.");
      }
      const { table, tableIndex } = await resolveWordTableTarget();
      return {
        ok: true,
        host: "word",
        action: type,
        tableIndex,
        rowCount: table.rowCount,
        columnCount: table.values?.[0]?.length ?? 0,
        title: trimString(table.title),
        style: table.style || table.styleBuiltIn,
        valuesPreview: table.values?.slice(0, 8).map((row) => row.slice(0, 8)),
      };
    }

    if (type === "editTableCell") {
      if (!supportsRequirementSet("WordApi", "1.3")) {
        throw new Error("Word table cell edits require WordApi 1.3 or newer.");
      }
      const { table, tableIndex } = await resolveWordTableTarget();
      const rowIndex = Math.max(0, Math.trunc(toNumber(options.rowIndex ?? action.rowIndex) ?? 0));
      const columnIndex = Math.max(0, Math.trunc(toNumber(options.columnIndex ?? action.columnIndex) ?? 0));
      const cell = table.getCell(rowIndex, columnIndex);
      cell.value = content;
      await context.sync();
      return { ok: true, host: "word", action: type, tableIndex, rowIndex, columnIndex, text: content };
    }

    if (type === "modifyTable") {
      if (!supportsRequirementSet("WordApi", "1.3")) {
        throw new Error("Word table row/column operations require WordApi 1.3 or newer.");
      }
      const { table, tableIndex } = await resolveWordTableTarget();
      const operation = trimString(options.operation ?? action.operation);
      const count = Math.max(1, Math.trunc(toNumber(options.count ?? action.count) ?? 1));
      const insertLocation = (options.insertLocation ?? action.insertLocation) === "start" ? Word.InsertLocation.start : Word.InsertLocation.end;
      const values = toStringMatrix(options.values ?? action.values);
      if ((operation === "deleteRows" || operation === "deleteColumns" || operation === "deleteTable" || operation === "clear") && toBoolean(options.confirmDestructive ?? action.confirmDestructive) !== true) {
        throw new Error(`Word table ${operation} requires confirmDestructive=true.`);
      }
      if (operation === "addRows") {
        table.addRows(insertLocation, count, values);
      } else if (operation === "addColumns") {
        table.addColumns(insertLocation, count, values);
      } else if (operation === "deleteRows") {
        table.deleteRows(Math.max(0, Math.trunc(toNumber(options.rowIndex ?? action.rowIndex) ?? 0)), count);
      } else if (operation === "deleteColumns") {
        table.deleteColumns(Math.max(0, Math.trunc(toNumber(options.columnIndex ?? action.columnIndex) ?? 0)), count);
      } else if (operation === "clear") {
        table.clear();
      } else if (operation === "deleteTable") {
        table.delete();
      } else if (operation === "setStyle") {
        const style = trimString(options.style ?? action.style);
        const builtInStyle = trimString(options.styleBuiltIn ?? action.styleBuiltIn);
        if (style) table.style = style;
        if (builtInStyle) table.styleBuiltIn = builtInStyle as Word.BuiltInStyleName;
      } else {
        throw new Error(`Unsupported Word table operation: ${operation || "(missing)"}.`);
      }
      await context.sync();
      return { ok: true, host: "word", action: type, tableIndex, operation, count };
    }

    if (type === "sectionLayout") {
      const operation = trimString(options.operation ?? action.operation) ?? "inventory";
      const sectionIndex = Math.max(0, Math.trunc(toNumber(options.sectionIndex ?? action.sectionIndex) ?? 0));
      const sections = context.document.sections;
      sections.load("items/body/text");
      const documentPageSetup = context.document.pageSetup;
      documentPageSetup.load("topMargin,bottomMargin,leftMargin,rightMargin,pageWidth,pageHeight");
      await context.sync();
      const section = sections.items[sectionIndex];
      if (!section) {
        throw new Error(`Could not find Word section at index ${sectionIndex + 1}.`);
      }

      const headerFooterType = (trimString(options.headerFooterType ?? action.headerFooterType) ?? "Primary") as
        | "Primary"
        | "FirstPage"
        | "EvenPages";
      const headerFooterKind = trimString(options.part ?? action.part) === "footer" ? "footer" : "header";
      const targetHeaderFooter = () => headerFooterKind === "footer"
        ? section.getFooter(headerFooterType)
        : section.getHeader(headerFooterType);

      if (operation === "inventory") {
        return {
          ok: true,
          host: "word",
          action: type,
          sections: sections.items.map((entry, index) => ({
            sectionIndex: index + 1,
            bodyPreview: truncateLabel(entry.body.text, 180),
          })),
          pageSetup: {
            topMargin: formatPoints(documentPageSetup.topMargin),
            bottomMargin: formatPoints(documentPageSetup.bottomMargin),
            leftMargin: formatPoints(documentPageSetup.leftMargin),
            rightMargin: formatPoints(documentPageSetup.rightMargin),
            pageWidth: formatPoints(documentPageSetup.pageWidth),
            pageHeight: formatPoints(documentPageSetup.pageHeight),
          },
        };
      }

      if (operation === "setHeader" || operation === "setFooter" || operation === "clearHeader" || operation === "clearFooter") {
        if (operation.startsWith("clear") && toBoolean(options.confirmDestructive ?? action.confirmDestructive) !== true) {
          throw new Error(`Word ${operation} requires confirmDestructive=true.`);
        }
        const bodyPart = operation.endsWith("Footer") ? section.getFooter(headerFooterType) : targetHeaderFooter();
        const nextText = operation.startsWith("clear") ? "" : content;
        bodyPart.insertText(nextText, Word.InsertLocation.replace);
        await context.sync();
        return {
          ok: true,
          host: "word",
          action: type,
          operation,
          sectionIndex: sectionIndex + 1,
          headerFooterType,
        };
      }

      if (operation === "pageSetup") {
        if (!supportsRequirementSet("WordApiDesktop", "1.3")) {
          throw new Error("Word page setup changes require WordApiDesktop 1.3 or newer.");
        }
        const pageSetup = section.pageSetup;
        const margins = options.margins && typeof options.margins === "object" ? options.margins as Record<string, unknown> : {};
        const top = toNumber(margins.top ?? options.topMargin ?? action.topMargin);
        const bottom = toNumber(margins.bottom ?? options.bottomMargin ?? action.bottomMargin);
        const left = toNumber(margins.left ?? options.leftMargin ?? action.leftMargin);
        const right = toNumber(margins.right ?? options.rightMargin ?? action.rightMargin);
        const pageWidth = toNumber(margins.pageWidth ?? options.pageWidth ?? action.pageWidth);
        const pageHeight = toNumber(margins.pageHeight ?? options.pageHeight ?? action.pageHeight);
        const orientation = trimString(options.orientation ?? action.orientation);
        if (typeof top === "number") pageSetup.topMargin = top;
        if (typeof bottom === "number") pageSetup.bottomMargin = bottom;
        if (typeof left === "number") pageSetup.leftMargin = left;
        if (typeof right === "number") pageSetup.rightMargin = right;
        if (typeof pageWidth === "number") pageSetup.pageWidth = pageWidth;
        if (typeof pageHeight === "number") pageSetup.pageHeight = pageHeight;
        if (orientation) pageSetup.orientation = orientation as Word.PageOrientation;
        await context.sync();
        return {
          ok: true,
          host: "word",
          action: type,
          operation,
          sectionIndex: sectionIndex + 1,
          changed: { top, bottom, left, right, pageWidth, pageHeight, orientation },
        };
      }

      if (operation === "insertBreak") {
        const { range } = await resolveTargetRange();
        range.load("text");
        const requestedBreakType = trimString(options.breakType ?? action.breakType) ?? "page";
        const breakType = normalizeWordBreakType(requestedBreakType);
        const insertLocation = (action.placement === "before" ? Word.InsertLocation.before : Word.InsertLocation.after);
        const beforeOoxml = breakType === "Page" ? body.getOoxml() : undefined;
        await context.sync();
        const beforeBreakCount = beforeOoxml ? countWordOoxmlBreaks(beforeOoxml.value, breakType) : undefined;
        const targetText = trimString(
          action.target?.text ??
          action.target?.label ??
          action.target?.searchQuery ??
          range.text,
        );
        const allowDuplicatePageBreak =
          toBoolean(options.allowDuplicatePageBreak ?? options.allowDuplicate ?? action.allowDuplicatePageBreak) === true;
        if (
          breakType === "Page" &&
          !allowDuplicatePageBreak &&
          beforeOoxml?.value &&
          wordOoxmlHasAdjacentPageBreak(beforeOoxml.value, targetText, action.placement)
        ) {
          return {
            ok: true,
            host: "word",
            action: type,
            operation,
            requestedBreakType,
            resolvedBreakType: breakType,
            target: {
              kind: action.target?.kind ?? "selection",
              text: truncateLabel(range.text, 140),
            },
            placement: action.placement ?? "after",
            skipped: true,
            warnings: [
              "Skipped duplicate page-break insertion because a persisted page break already exists adjacent to the resolved target. Set allowDuplicatePageBreak=true only when an additional break is intentional.",
            ],
            verification: {
              method: "body.getOoxml",
              persisted: true,
              alreadyPresent: true,
              beforeBreakCount,
              afterBreakCount: beforeBreakCount,
              note: "The requested page break was already present next to the target, so no additional break was inserted.",
            },
          };
        }
        range.insertBreak(breakType as Word.BreakType, insertLocation);
        const afterOoxml = breakType === "Page" ? body.getOoxml() : undefined;
        await context.sync();
        const afterBreakCount = afterOoxml ? countWordOoxmlBreaks(afterOoxml.value, breakType) : undefined;
        if (
          breakType === "Page" &&
          typeof beforeBreakCount === "number" &&
          typeof afterBreakCount === "number" &&
          afterBreakCount <= beforeBreakCount
        ) {
          throw new Error(
            "word_section_layout.insertBreak requested a page break but post-write OOXML verification did not find a new persisted w:br w:type=\"page\". Do not claim page-break or page-count success without verify_doc, verify_doc_visual, or native page metadata evidence.",
          );
        }
        return {
          ok: true,
          host: "word",
          action: type,
          operation,
          requestedBreakType,
          resolvedBreakType: breakType,
          target: {
            kind: action.target?.kind ?? "selection",
            text: truncateLabel(range.text, 140),
          },
          placement: action.placement ?? "after",
          verification: breakType === "Page"
            ? {
                method: "body.getOoxml",
                persisted: typeof afterBreakCount === "number" ? afterBreakCount > (beforeBreakCount ?? 0) : undefined,
                beforeBreakCount,
                afterBreakCount,
                note: "Page-break success is based on persisted Word OOXML evidence; still call verify_doc or verify_doc_visual before claiming page count.",
              }
            : {
                method: "office-js-insertBreak",
                persisted: undefined,
                note: "Non-page break insertion was sent through Word.insertBreak; use verify_doc or native page/section metadata before claiming final layout.",
              },
        };
      }

      throw new Error(`Unsupported Word section/layout operation: ${operation}.`);
    }

    if (type === "insertContentControl") {
      if (!supportsRequirementSet("WordApi", "1.1")) {
        throw new Error("Word content controls require WordApi 1.1 or newer.");
      }

      const supportsTypedContentControls = supportsRequirementSet("WordApi", "1.5");
      const { range } = await resolveTargetRange();
      const warnings: string[] = [];
      const requestedType = trimString(options.contentControlType) ?? trimString(action.contentControlType);
      let contentRange = range;

      if ((action.placement === "before" || action.placement === "after") && content) {
        contentRange = range.insertText(content, placement);
      } else if ((action.placement === "before" || action.placement === "after") && !content) {
        warnings.push("Wrapped the target range because empty content controls cannot be inserted before or after a target natively.");
      }

      if (requestedType && !supportsTypedContentControls) {
        warnings.push(`Requested content control type "${requestedType}" but this host only supports default rich text insertion.`);
      }

      const contentControl =
        requestedType && supportsTypedContentControls
          ? contentRange.insertContentControl(requestedType as never)
          : contentRange.insertContentControl();

      if (content && contentRange === range) {
        contentControl.insertText(content, Word.InsertLocation.replace);
      }

      const title = trimString(options.title) ?? trimString(action.title);
      const tag = trimString(options.tag) ?? trimString(action.tag);
      const placeholderText = trimString(options.placeholderText) ?? trimString(action.placeholderText);
      const appearance = trimString(options.appearance) ?? trimString(action.appearance);
      const color = trimString(options.color) ?? trimString(action.color);
      const cannotDelete = toBoolean(options.cannotDelete ?? action.cannotDelete);
      const cannotEdit = toBoolean(options.cannotEdit ?? action.cannotEdit);
      const removeWhenEdited = toBoolean(options.removeWhenEdited ?? action.removeWhenEdited);

      if (title) {
        contentControl.title = title;
      }
      if (tag) {
        contentControl.tag = tag;
      }
      if (placeholderText) {
        contentControl.placeholderText = placeholderText;
      }
      if (appearance) {
        contentControl.appearance = appearance as Word.ContentControlAppearance;
      }
      if (color) {
        contentControl.color = color;
      }
      if (typeof cannotDelete === "boolean") {
        contentControl.cannotDelete = cannotDelete;
      }
      if (typeof cannotEdit === "boolean") {
        contentControl.cannotEdit = cannotEdit;
      }
      if (typeof removeWhenEdited === "boolean") {
        contentControl.removeWhenEdited = removeWhenEdited;
      }

      contentControl.load(
        supportsContentControlSubtypes ? "id,title,tag,type,subtype,text" : "id,title,tag,type,text",
      );
      await context.sync();
      return {
        ok: true,
        host: "word",
        action: type,
        id: `contentControl:${contentControl.id}`,
        contentControlId: contentControl.id,
        title: trimString(contentControl.title),
        tag: trimString(contentControl.tag),
        text: truncateLabel(contentControl.text, 180),
        contentControlType: contentControl.type,
        contentControlSubtype: supportsContentControlSubtypes ? contentControl.subtype : undefined,
        warnings,
      };
    }

    if (type === "contentControlEdit") {
      if (!supportsRequirementSet("WordApi", "1.1")) {
        throw new Error("Word content controls require WordApi 1.1 or newer.");
      }
      const operation = trimString(options.operation ?? action.operation) ?? "inventory";
      const supportsTypedContentControls = supportsRequirementSet("WordApi", "1.5");
      if (operation === "inventory") {
        const controls = body.contentControls;
        controls.load(supportsContentControlSubtypes
          ? "items/id,items/title,items/tag,items/type,items/subtype,items/appearance,items/cannotDelete,items/cannotEdit,items/removeWhenEdited,items/placeholderText,items/text"
          : "items/id,items/title,items/tag,items/type,items/appearance,items/cannotDelete,items/cannotEdit,items/removeWhenEdited,items/placeholderText,items/text");
        await context.sync();
        return {
          ok: true,
          host: "word",
          action: type,
          controls: controls.items.slice(0, 60).map((control) => ({
            id: `contentControl:${control.id}`,
            contentControlId: control.id,
            title: trimString(control.title),
            tag: trimString(control.tag),
            type: control.type,
            subtype: supportsContentControlSubtypes ? control.subtype : undefined,
            placeholderText: trimString(control.placeholderText),
            text: truncateLabel(control.text, 180),
            cannotDelete: control.cannotDelete,
            cannotEdit: control.cannotEdit,
            removeWhenEdited: control.removeWhenEdited,
            appearance: control.appearance,
          })),
        };
      }

      const target = action.target?.kind === "contentControl"
        ? action.target
        : { kind: "contentControl" as const, id: trimString(options.id ?? action.id), label: trimString(options.title ?? action.title), text: trimString(options.tag ?? action.tag) };
      const contentControl = await resolveWordContentControlTarget(target);
      if (!contentControl) {
        throw new Error(`Could not find the requested Word content control: ${target.id || target.label || target.text || "content control"}.`);
      }

      if (operation === "fill" || operation === "setText") {
        contentControl.insertText(trimString(options.text ?? action.text ?? action.content) ?? content, Word.InsertLocation.replace);
      } else if (operation === "clear") {
        if (toBoolean(options.confirmDestructive ?? action.confirmDestructive) !== true) {
          throw new Error("Word content-control clear requires confirmDestructive=true.");
        }
        contentControl.clear();
      } else if (operation === "select") {
        contentControl.select();
      } else if (operation === "delete") {
        if (toBoolean(options.keepContent ?? action.keepContent) !== true && toBoolean(options.confirmDestructive ?? action.confirmDestructive) !== true) {
          throw new Error("Word content-control deleteContent requires confirmDestructive=true.");
        }
        contentControl.delete(toBoolean(options.keepContent ?? action.keepContent) === true);
      } else if (operation === "setMetadata") {
        const title = trimString(options.title ?? action.title);
        const tag = trimString(options.tag ?? action.tag);
        const placeholderText = trimString(options.placeholderText ?? action.placeholderText);
        const cannotDelete = toBoolean(options.cannotDelete ?? action.cannotDelete);
        const cannotEdit = toBoolean(options.cannotEdit ?? action.cannotEdit);
        const removeWhenEdited = toBoolean(options.removeWhenEdited ?? action.removeWhenEdited);
        if (title) contentControl.title = title;
        if (tag) contentControl.tag = tag;
        if (placeholderText) contentControl.placeholderText = placeholderText;
        if (typeof cannotDelete === "boolean") contentControl.cannotDelete = cannotDelete;
        if (typeof cannotEdit === "boolean") contentControl.cannotEdit = cannotEdit;
        if (typeof removeWhenEdited === "boolean") contentControl.removeWhenEdited = removeWhenEdited;
      } else if (operation === "setCheckbox") {
        if (!supportsRequirementSet("WordApi", "1.7")) throw new Error("Word checkbox content controls require WordApi 1.7 or newer.");
        contentControl.checkboxContentControl.isChecked = toBoolean(options.checked ?? action.checked) === true;
      } else if (operation === "setDropdownItems" || operation === "setComboBoxItems") {
        if (!supportsRequirementSet("WordApi", "1.9")) throw new Error("Word dropdown/combo box item operations require WordApi 1.9 or newer.");
        const listHost = operation === "setDropdownItems" ? contentControl.dropDownListContentControl : contentControl.comboBoxContentControl;
        listHost.deleteAllListItems();
        const items = Array.isArray(options.items ?? action.items) ? options.items ?? action.items : [];
        for (const item of items as Array<{ displayText?: string; value?: string }>) {
          const displayText = trimString(item.displayText ?? item.value);
          if (displayText) listHost.addListItem(displayText, trimString(item.value) ?? displayText);
        }
      } else if (operation === "insert" || operation === "insertTyped") {
        const { range } = await resolveTargetRange();
        const requestedType = trimString(options.contentControlType ?? action.contentControlType);
        const control = requestedType && supportsTypedContentControls
          ? range.insertContentControl(requestedType as never)
          : range.insertContentControl();
        const title = trimString(options.title ?? action.title);
        const tag = trimString(options.tag ?? action.tag);
        if (title) control.title = title;
        if (tag) control.tag = tag;
        if (content) control.insertText(content, Word.InsertLocation.replace);
        control.load(supportsContentControlSubtypes ? "id,title,tag,type,subtype,text" : "id,title,tag,type,text");
        await context.sync();
        return { ok: true, host: "word", action: type, operation, id: `contentControl:${control.id}`, title: control.title, tag: control.tag, type: control.type, subtype: supportsContentControlSubtypes ? control.subtype : undefined, text: truncateLabel(control.text, 180) };
      } else {
        throw new Error(`Unsupported Word content-control operation: ${operation}.`);
      }

      contentControl.load(supportsContentControlSubtypes ? "id,title,tag,type,subtype,text" : "id,title,tag,type,text");
      await context.sync();
      return { ok: true, host: "word", action: type, operation, id: `contentControl:${contentControl.id}`, title: contentControl.title, tag: contentControl.tag, type: contentControl.type, subtype: supportsContentControlSubtypes ? contentControl.subtype : undefined, text: truncateLabel(contentControl.text, 180) };
    }

    if (type === "buildingBlock") {
      const operation = trimString(options.operation ?? action.operation) ?? "inventory";
      if (operation === "insertApprovedText") {
        const approvedText = trimString(options.text ?? action.text ?? action.content) ?? content;
        const provenance = trimString(options.provenance ?? action.provenance ?? options.name ?? action.name);
        if (!approvedText) {
          throw new Error("Word approved reusable text insertion requires text.");
        }
        if (!provenance) {
          throw new Error("Word approved reusable text insertion requires provenance.");
        }
        const { range } = await resolveTargetRange();
        range.insertText(approvedText, placement);
        await context.sync();
        return {
          ok: true,
          host: "word",
          action: type,
          operation,
          inserted: true,
          provenance,
          source: "user-approved-text",
        };
      }

      if (!supportsRequirementSet("WordApiDesktop", "1.3")) {
        throw new Error("Word building block/template operations require WordApiDesktop 1.3 or newer.");
      }
      const template = context.document.attachedTemplate;
      template.load("name,fullName");
      const entries = template.buildingBlockEntries;
      entries.load("items/name,items/type/name,items/category/name,items/description");
      await context.sync();

      const entryItems = (entries as unknown as { items?: Array<{
        name?: string;
        description?: string;
        category?: { name?: string };
        type?: { name?: string };
        insert?: (range: Word.Range, richText: boolean) => Word.Range;
      }> }).items ?? [];

      if (operation === "inventory") {
        return {
          ok: true,
          host: "word",
          action: type,
          template: { name: template.name, fullName: template.fullName },
          entries: entryItems.slice(0, Math.max(1, Math.trunc(toNumber(options.maxResults ?? action.maxResults) ?? 20))).map((entry, index) => ({
            id: `buildingBlock:${index + 1}`,
            name: entry.name,
            category: entry.category?.name,
            type: entry.type?.name,
            description: entry.description,
          })),
          provenanceRequired: true,
        };
      }

      if (operation === "insert") {
        const name = trimString(options.name ?? action.name);
        if (!name) {
          throw new Error("Word building block insertion requires an approved entry name.");
        }
        const entry = entryItems.find((item) => item.name === name);
        if (!entry) {
          throw new Error(`Could not find approved Word building block "${name}" in the attached template.`);
        }
        const { range } = await resolveTargetRange();
        const insertBuildingBlock = entry.insert;
        if (typeof insertBuildingBlock !== "function") {
          throw new Error("This host cannot insert the selected Word building block entry.");
        }
        insertBuildingBlock.call(entry, range, true);
        await context.sync();
        return { ok: true, host: "word", action: type, operation, name: entry.name, inserted: true };
      }

      throw new Error(`Unsupported Word building block operation: ${operation}.`);
    }

    if (type === "critiqueAnnotation") {
      const operation = trimString(options.operation ?? action.operation) ?? "propose";
      if (!supportsRequirementSet("WordApi", "1.7")) {
        return {
          ok: true,
          host: "word",
          action: type,
          operation,
          completion: "fallback",
          fallbackTool: "office_propose_edits",
          warning: "Word-native critique annotations require WordApi 1.7 and Microsoft 365 annotation service support.",
          proposal: {
            summary: trimString(options.summary ?? action.summary) ?? "Review suggestion prepared for taskpane approval.",
            edits: Array.isArray(options.edits ?? action.edits) ? options.edits ?? action.edits : [],
          },
        };
      }

      if (operation === "propose") {
        const { paragraph } = await resolveTargetRange();
        if (!paragraph) {
          return {
            ok: true,
            host: "word",
            action: type,
            operation,
            completion: "fallback",
            fallbackTool: "office_propose_edits",
            warning: "Native critique annotations currently require a paragraph/search anchor; falling back to reviewable proposals.",
          };
        }
        const start = Math.max(0, Math.trunc(toNumber(options.start ?? action.start) ?? 0));
        const length = Math.max(1, Math.trunc(toNumber(options.length ?? action.length) ?? 1));
        const colorScheme = trimString(options.colorScheme ?? action.colorScheme) ?? "Blue";
        const ids = paragraph.insertAnnotations({
          critiques: [
            {
              colorScheme: colorScheme as Word.CritiqueColorScheme,
              start,
              length,
              popupOptions: {
                brandingTextResourceId: "Pi-Office",
                titleResourceId: "pi-office-review",
                subtitleResourceId: "Pi-Office suggestion",
                suggestions: [
                  trimString(options.suggestion ?? action.suggestion ?? action.content) ?? "Review suggested change",
                ],
              },
            },
          ],
        });
        await context.sync();
        return { ok: true, host: "word", action: type, operation, annotationIds: ids.value };
      }

      if (operation === "accept" || operation === "reject" || operation === "delete") {
        const annotationId = trimString(options.annotationId ?? action.annotationId);
        if (!annotationId) {
          throw new Error(`Word annotation ${operation} requires annotationId.`);
        }
        const annotation = context.document.getAnnotationById(annotationId);
        annotation.load("id,state,critiqueAnnotation/critique");
        await context.sync();
        if (operation === "accept") annotation.critiqueAnnotation.accept();
        if (operation === "reject") annotation.critiqueAnnotation.reject();
        if (operation === "delete") annotation.delete();
        await context.sync();
        return { ok: true, host: "word", action: type, operation, annotationId };
      }

      throw new Error(`Unsupported Word critique annotation operation: ${operation}.`);
    }

    if (type === "insertField") {
      if (!supportsRequirementSet("WordApi", "1.5")) {
        throw new Error("Word field insertion requires WordApi 1.5 or newer.");
      }

      const { range } = await resolveTargetRange();
      const fieldType = trimString(options.fieldType) ?? trimString(action.fieldType) ?? "Empty";
      const fieldText = trimString(options.text) ?? trimString(action.fieldText) ?? trimString(content);
      const removeFormatting = toBoolean(options.removeFormatting ?? action.removeFormatting) ?? false;
      const field = range.insertField(placement, fieldType as Word.FieldType, fieldText, removeFormatting);
      field.load("code,type,locked,result/text");
      await context.sync();

      const documentFields = body.fields;
      documentFields.load("items/code,items/type,items/result/text");
      await context.sync();

      let fieldIndex = 0;
      for (let index = documentFields.items.length - 1; index >= 0; index -= 1) {
        const entry = documentFields.items[index];
        if (!entry) {
          continue;
        }
        if (
          entry.code === field.code &&
          entry.result.text === field.result.text &&
          String(entry.type || "") === String(field.type || "")
        ) {
          fieldIndex = index + 1;
          break;
        }
      }

      return {
        ok: true,
        host: "word",
        action: type,
        fieldId: fieldIndex > 0 ? `field:${fieldIndex}` : undefined,
        fieldIndex: fieldIndex || undefined,
        fieldType: field.type,
        fieldCode: truncateLabel(field.code, 180),
        text: truncateLabel(field.result.text, 180),
        locked: field.locked,
      };
    }

    if (type === "fieldAction") {
      if (!supportsRequirementSet("WordApi", "1.5")) {
        throw new Error("Word field actions require WordApi 1.5 or newer.");
      }
      const operation = trimString(options.operation ?? action.operation);
      if (operation === "inventory") {
        const fields = body.fields;
        fields.load("items/code,items/type,items/locked,items/result/text");
        await context.sync();
        return {
          ok: true,
          host: "word",
          action: type,
          operation,
          fields: fields.items.slice(0, 50).map((field, index) => ({
            id: `field:${index + 1}`,
            type: field.type,
            locked: field.locked,
            code: truncateLabel(field.code, 240),
            resultText: truncateLabel(field.result.text, 240),
          })),
        };
      }

      if (operation === "insert") {
        const { range } = await resolveTargetRange();
        const fieldType = trimString(options.fieldType ?? action.fieldType) ?? "Empty";
        const fieldText = trimString(options.text ?? action.fieldText ?? action.content) ?? "";
        const field = range.insertField(placement, fieldType as Word.FieldType, fieldText, false);
        field.load("code,type,locked,result/text");
        await context.sync();
        return { ok: true, host: "word", action: type, operation, typeName: field.type, code: field.code, resultText: field.result.text };
      }

      const target = await resolveTargetRange();
      if (!target.field) {
        throw new Error("Word field action requires a field target unless operation=inventory or insert.");
      }
      const field = target.field.field;
      if (operation === "update") {
        field.updateResult();
      } else if (operation === "lock") {
        field.locked = true;
      } else if (operation === "unlock") {
        field.locked = false;
      } else if (operation === "unlink") {
        if (!supportsRequirementSet("WordApiDesktop", "1.4")) {
          throw new Error("Word field unlink requires WordApiDesktop 1.4 or newer.");
        }
        if (toBoolean(options.confirmDestructive ?? action.confirmDestructive) !== true) {
          throw new Error("Word field unlink requires confirmDestructive=true.");
        }
        field.unlink();
      } else if (operation === "delete") {
        if (toBoolean(options.confirmDestructive ?? action.confirmDestructive) !== true) {
          throw new Error("Word field delete requires confirmDestructive=true.");
        }
        field.delete();
      } else if (operation === "select") {
        field.select();
      } else {
        throw new Error(`Unsupported Word field operation: ${operation || "(missing)"}.`);
      }
      await context.sync();
      return { ok: true, host: "word", action: type, operation, fieldId: target.field.fieldIndex };
    }

    if (type === "tocAction") {
      if (!supportsRequirementSet("WordApiDesktop", "1.4")) {
        throw new Error("Word table-of-contents actions require WordApiDesktop 1.4 or newer.");
      }
      const operation = trimString(options.operation ?? action.operation);
      const tocs = context.document.tablesOfContents;
      tocs.load("items");
      await context.sync();
      if (operation === "inventory") {
        return { ok: true, host: "word", action: type, operation, count: tocs.items.length };
      }
      const index = Math.max(0, Math.trunc(toNumber(options.tocIndex ?? action.tocIndex) ?? 1) - 1);
      if (operation === "add") {
        const { range } = await resolveTargetRange();
        const toc = tocs.add(range, {
          useBuiltInHeadingStyles: true,
          upperHeadingLevel: Math.max(1, Math.trunc(toNumber(options.upperHeadingLevel ?? action.upperHeadingLevel) ?? 1)),
          lowerHeadingLevel: Math.max(1, Math.trunc(toNumber(options.lowerHeadingLevel ?? action.lowerHeadingLevel) ?? 3)),
        });
        toc.load("upperHeadingLevel,lowerHeadingLevel");
        await context.sync();
        return { ok: true, host: "word", action: type, operation, upperHeadingLevel: toc.upperHeadingLevel, lowerHeadingLevel: toc.lowerHeadingLevel };
      }
      if (operation === "markEntry") {
        const { range } = await resolveTargetRange();
        const entry = trimString(options.entry ?? action.entry ?? action.content);
        const field = entry ? tocs.markTocEntry(range, { entry }) : tocs.markTocEntry(range);
        field.load("code,type,result/text");
        await context.sync();
        return { ok: true, host: "word", action: type, operation, code: field.code, resultText: field.result.text };
      }
      const toc = tocs.items[index];
      if (!toc) {
        throw new Error(`Could not find Word table of contents at index ${index + 1}.`);
      }
      if (operation === "update") {
        toc.updatePageNumbers();
      } else if (operation === "updatePageNumbers") {
        toc.updatePageNumbers();
      } else if (operation === "delete") {
        if (toBoolean(options.confirmDestructive ?? action.confirmDestructive) !== true) {
          throw new Error("Word TOC delete requires confirmDestructive=true.");
        }
        toc.delete();
      } else {
        throw new Error(`Unsupported Word TOC operation: ${operation || "(missing)"}.`);
      }
      await context.sync();
      return { ok: true, host: "word", action: type, operation, tocIndex: index + 1 };
    }

    if (type === "acceptRevision" || type === "rejectRevision") {
      const revision = await resolveRevisionActionTarget();
      if (type === "acceptRevision") {
        revision.change.accept();
      } else {
        revision.change.reject();
      }
      await context.sync();
      return {
        ok: true,
        host: "word",
        action: type,
        revisionId: revision.revisionId,
        text: truncateLabel(revision.change.text, 180),
        type: revision.change.type,
      };
    }

    if (type === "reviewExchange") {
      const operation = trimString(options.operation ?? action.operation) ?? "inventory";
      if (operation === "inventory") {
        const trackedChanges = body.getTrackedChanges();
        trackedChanges.load("items/author,items/date,items/text,items/type");
        await context.sync();
        return {
          ok: true,
          host: "word",
          action: type,
          operation,
          revisions: trackedChanges.items.slice(0, 50).map((change, index) => ({
            id: `revision:${index + 1}`,
            author: change.author,
            date: change.date?.toISOString?.(),
            type: change.type,
            text: truncateLabel(change.text, 220),
          })),
          count: trackedChanges.items.length,
        };
      }

      if (operation === "compare") {
        if (!supportsRequirementSet("WordApiDesktop", "1.4")) {
          throw new Error("Word document compare/redline requires WordApiDesktop 1.4 or newer.");
        }
        if (toBoolean(options.confirmReviewStateChange ?? action.confirmReviewStateChange) !== true) {
          throw new Error("Word compare review exchange requires confirmReviewStateChange=true.");
        }
        const filePath = trimString(options.filePath ?? action.filePath ?? action.content);
        if (!filePath) {
          throw new Error("Word compare requires an explicit baseline filePath.");
        }
        context.document.compare(filePath, {
          compareTarget: (options.compareTarget ?? action.compareTarget ?? "Current") as Word.CompareTarget,
          detectFormatChanges: toBoolean(options.detectFormatChanges ?? action.detectFormatChanges) ?? true,
        });
        await context.sync();
        return {
          ok: true,
          host: "word",
          action: type,
          operation,
          filePath,
          summary: "Requested Word-native compare/redline against the supplied baseline path.",
        };
      }

      if (operation === "acceptAll" || operation === "rejectAll") {
        if (toBoolean(options.confirmReviewStateChange ?? action.confirmReviewStateChange) !== true) {
          throw new Error(`Word ${operation} review exchange requires confirmReviewStateChange=true.`);
        }
        const targetChanges = body.getTrackedChanges();
        targetChanges.load("items/text");
        await context.sync();
        if (operation === "acceptAll") {
          targetChanges.acceptAll();
        } else {
          targetChanges.rejectAll();
        }
        await context.sync();
        return { ok: true, host: "word", action: type, operation, count: targetChanges.items.length };
      }

      throw new Error(`Unsupported Word review exchange operation: ${operation}.`);
    }

    if (type === "proofingStats") {
      const scope = trimString(options.scope ?? action.scope) ?? "document";
      const targetRange = scope === "selection" ? context.document.getSelection() : body.getRange();
      targetRange.load("text");
      await context.sync();
      const text = targetRange.text ?? "";
      const words = text.trim() ? text.trim().split(/\s+/).filter(Boolean).length : 0;
      const sentences = text.split(/[.!?]+/).filter((entry) => entry.trim()).length;
      const paragraphs = text.split(/\r|\n/).filter((entry) => entry.trim()).length;
      const characters = text.replace(/\s/g, "").length;
      const readabilityAvailable = supportsRequirementSet("WordApiDesktop", "1.4");
      let readability: Array<{ name?: string | undefined; value?: number | undefined }> = [];
      if (readabilityAvailable) {
        const stats = context.document.readabilityStatistics;
        stats.load("items/name,items/value");
        await context.sync();
        readability = stats.items.map((entry) => ({ name: entry.name, value: entry.value }));
      }
      return {
        ok: true,
        host: "word",
        action: type,
        scope,
        mutating: false,
        nativeMetrics: {
          words,
          sentences,
          paragraphs,
          characters,
          readability,
          readabilityAvailable,
        },
        note: readabilityAvailable
          ? "Native Word readability statistics were included."
          : "Basic counts are computed from Office.js text; native readability statistics require WordApiDesktop 1.4.",
      };
    }

    if (type === "protectionAwareness") {
      const operation = trimString(options.operation ?? action.operation) ?? "diagnostics";
      const supportsProtection = supportsRequirementSet("WordApiDesktop", "1.4");
      const trackedChanges = supportsRequirementSet("WordApi", "1.6") ? body.getTrackedChanges() : undefined;
      trackedChanges?.load("items/author,items/date,items/text,items/type");
      const comments = supportsRequirementSet("WordApi", "1.4") ? body.getComments() : undefined;
      comments?.load("items/id,items/authorName,items/content,items/resolved");
      if (supportsProtection) {
        context.document.load("protectionType");
        context.document.activeWindow?.view?.load("type");
        context.document.activeWindow?.view?.revisionsFilter?.load("markup,view,reviewers/items/isVisible");
      }
      await context.sync();

      if (operation === "protect" || operation === "unprotect") {
        throw new Error("word_collab_guard is diagnostics-only. Protection changes require a separate write-doc protection tool.");
      }

      const protectionType = supportsProtection ? String((context.document as unknown as { protectionType?: unknown }).protectionType ?? "Unknown") : "unsupported";
      const revisionCount = trackedChanges?.items.length ?? 0;
      const unresolvedComments = comments?.items.filter((comment) => !comment.resolved).length ?? 0;
      const warnings = [
        protectionType !== "NoProtection" && protectionType !== "unsupported" ? `Document protection is ${protectionType}.` : undefined,
        revisionCount ? `${revisionCount} tracked change${revisionCount === 1 ? "" : "s"} visible in current scope.` : undefined,
        unresolvedComments ? `${unresolvedComments} unresolved comment${unresolvedComments === 1 ? "" : "s"}.` : undefined,
      ].filter((entry): entry is string => Boolean(entry));
      return {
        ok: true,
        host: "word",
        action: type,
        operation,
        mutating: false,
        diagnostics: {
          protectionType,
          revisionCount,
          unresolvedComments,
          reviewers: supportsProtection
            ? (context.document.activeWindow?.view?.revisionsFilter?.reviewers?.items ?? []).map((reviewer, index) => ({
                id: `reviewer:${index + 1}`,
                isVisible: reviewer.isVisible,
              }))
            : [],
          revisionsFilter: supportsProtection
            ? {
                markup: context.document.activeWindow?.view?.revisionsFilter?.markup,
                view: context.document.activeWindow?.view?.revisionsFilter?.view,
              }
            : undefined,
          supportsNativeProtection: supportsProtection,
        },
        warnings,
      };
    }

    if (type === "acceptAllRevisions" || type === "rejectAllRevisions") {
      if (!supportsRequirementSet("WordApi", "1.6")) {
        throw new Error("Word revision actions require WordApi 1.6 or newer.");
      }

      const selectionChanges = context.document.getSelection().getTrackedChanges();
      selectionChanges.load("items/text");
      await context.sync();
      const targetChanges = selectionChanges.items.length ? selectionChanges : body.getTrackedChanges();

      if (!selectionChanges.items.length) {
        targetChanges.load("items/text");
        await context.sync();
      }

      if (type === "acceptAllRevisions") {
        targetChanges.acceptAll();
      } else {
        targetChanges.rejectAll();
      }
      await context.sync();
      return {
        ok: true,
        host: "word",
        action: type,
        scope: selectionChanges.items.length ? "selection" : "document",
        count: targetChanges.items.length,
      };
    }

    throw new Error(`Unsupported Word action: ${type}`);
  });
}


