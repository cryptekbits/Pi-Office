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

export function matchesWordHeading(paragraph: Word.Paragraph, anchor: OfficeAnchor): boolean {
  const style = String(paragraph.styleBuiltIn || paragraph.style || "");
  const text = normalizeTextPreview(paragraph.text)?.toLowerCase() ?? "";
  const needle = normalizeTextPreview(anchor.text || anchor.label)?.toLowerCase();
  return /heading/i.test(style) && (!needle || text.includes(needle));
}

export function matchesWordParagraph(paragraph: Word.Paragraph, anchor: OfficeAnchor): boolean {
  if (anchor.paragraphId && paragraph.uniqueLocalId === anchor.paragraphId) {
    return true;
  }

  const needle = normalizeTextPreview(anchor.text || anchor.label)?.toLowerCase();
  if (!needle) {
    return false;
  }

  return (normalizeTextPreview(paragraph.text)?.toLowerCase() ?? "").includes(needle);
}

export async function navigateWordAnchor(anchor: OfficeAnchor): Promise<unknown> {
  return Word.run(async (context) => {
    const body = context.document.body;
    const supportsParagraphIds = supportsRequirementSet("WordApi", "1.6");
    const supportsFieldMetadata = supportsRequirementSet("WordApi", "1.5");
    const supportsContentControlSubtypes = supportsRequirementSet("WordApi", "1.3");

    if (anchor.kind === "searchResult") {
      const query = trimString(anchor.searchQuery) ?? trimString(anchor.text) ?? trimString(anchor.label);
      if (!query) {
        throw new Error("Word search result anchors require searchQuery or text.");
      }
      const matches = body.search(query, {
        matchCase: false,
        matchWholeWord: false,
      });
      matches.load("items/text");
      await context.sync();
      const index = typeof anchor.searchResultIndex === "number" ? Math.max(0, Math.trunc(anchor.searchResultIndex) - 1) : 0;
      const match = matches.items[index];
      if (match) {
        match.select();
        await context.sync();
        return {
          ok: true,
          host: "word",
          anchorKind: "searchResult",
          searchQuery: query,
          searchResultIndex: index + 1,
          text: truncateLabel(match.text, 180),
        };
      }
    }

    if ((anchor.kind === "footnote" || anchor.kind === "endnote") && supportsRequirementSet("WordApi", "1.5")) {
      const notes = anchor.kind === "footnote" ? body.footnotes : body.endnotes;
      notes.load("items/type,items/body/text,items/reference/text");
      await context.sync();

      const indexedNote =
        parseAnchorOrdinal(anchor.id, anchor.kind) ?? parseAnchorOrdinal(anchor.label, anchor.kind);
      const noteIndex =
        typeof indexedNote === "number"
          ? indexedNote - 1
          : notes.items.findIndex(
              (entry) =>
                matchesTextQuery(entry.body.text, anchor.text || anchor.label) ||
                matchesTextQuery(entry.reference.text, anchor.text || anchor.label),
            );
      const note = noteIndex >= 0 ? notes.items[noteIndex] : undefined;
      if (note) {
        note.reference.select();
        await context.sync();
        return {
          ok: true,
          host: "word",
          anchorKind: anchor.kind,
          noteId: `${anchor.kind}:${noteIndex + 1}`,
          referenceText: truncateLabel(note.reference.text, 80),
          text: truncateLabel(note.body.text, 180),
        };
      }
    }

    if (anchor.kind === "comment" && supportsRequirementSet("WordApi", "1.4")) {
      const comments = body.getComments();
      comments.load("items/id,items/content,items/authorName,items/resolved");
      await context.sync();
      const comment = comments.items.find(
        (entry) =>
          (anchor.commentId && entry.id === anchor.commentId) ||
          matchesTextQuery(entry.content, anchor.text || anchor.label),
      );
      if (comment) {
        comment.getRange().select();
        await context.sync();
        return {
          ok: true,
          host: "word",
          anchorKind: "comment",
          commentId: comment.id,
          text: truncateLabel(comment.content, 180),
          resolved: comment.resolved,
        };
      }
    }

    if (anchor.kind === "revision" && supportsRequirementSet("WordApi", "1.6")) {
      const trackedChanges = body.getTrackedChanges();
      trackedChanges.load("items/author,items/date,items/text,items/type");
      await context.sync();
      const indexedRevision =
        parseAnchorOrdinal(anchor.revisionId, "revision") ??
        parseAnchorOrdinal(anchor.id, "revision") ??
        parseAnchorOrdinal(anchor.label, "revision");
      const changeIndex =
        typeof indexedRevision === "number"
          ? indexedRevision - 1
          : trackedChanges.items.findIndex((entry) => matchesTextQuery(entry.text, anchor.text || anchor.label));
      const change = changeIndex >= 0 ? trackedChanges.items[changeIndex] : undefined;
      if (change) {
        change.getRange().select();
        await context.sync();
        return {
          ok: true,
          host: "word",
          anchorKind: "revision",
          revisionId: `revision:${changeIndex + 1}`,
          text: truncateLabel(change.text, 180),
          type: change.type,
        };
      }
    }

    if (anchor.kind === "contentControl" && supportsRequirementSet("WordApi", "1.1")) {
      const contentControls = body.contentControls;
      const directId = parseAnchorOrdinal(anchor.id, "contentControl");
      if (typeof directId === "number") {
        const directMatch = contentControls.getByIdOrNullObject(directId);
        directMatch.load(supportsContentControlSubtypes ? "id,title,tag,text,type,subtype" : "id,title,tag,text,type");
        await context.sync();
        if (!directMatch.isNullObject) {
          directMatch.getRange().select();
          await context.sync();
          return {
            ok: true,
            host: "word",
            anchorKind: "contentControl",
            id: `contentControl:${directMatch.id}`,
            contentControlId: directMatch.id,
            title: trimString(directMatch.title),
            tag: trimString(directMatch.tag),
            text: truncateLabel(directMatch.text, 180),
            contentControlType: directMatch.type,
            contentControlSubtype: supportsContentControlSubtypes ? directMatch.subtype : undefined,
          };
        }
      }

      contentControls.load(supportsContentControlSubtypes ? "items/id,items/title,items/tag,items/text,items/type,items/subtype" : "items/id,items/title,items/tag,items/text,items/type");
      await context.sync();
      const contentControl = contentControls.items.find(
        (entry) =>
          matchesTextQuery(entry.title, anchor.text || anchor.label) ||
          matchesTextQuery(entry.tag, anchor.text || anchor.label) ||
          matchesTextQuery(entry.text, anchor.text || anchor.label),
      );
      if (contentControl) {
        contentControl.getRange().select();
        await context.sync();
        return {
          ok: true,
          host: "word",
          anchorKind: "contentControl",
          id: `contentControl:${contentControl.id}`,
          contentControlId: contentControl.id,
          title: trimString(contentControl.title),
          tag: trimString(contentControl.tag),
          text: truncateLabel(contentControl.text, 180),
          contentControlType: contentControl.type,
          contentControlSubtype: supportsContentControlSubtypes ? contentControl.subtype : undefined,
        };
      }
    }

    if (anchor.kind === "field" && supportsRequirementSet("WordApi", "1.4")) {
      const fields = body.fields;
      fields.load(supportsFieldMetadata ? "items/code,items/type,items/result/text" : "items/code,items/result/text");
      await context.sync();

      const indexedField = parseAnchorOrdinal(anchor.id, "field") ?? parseAnchorOrdinal(anchor.label, "field");
      const fieldIndex =
        typeof indexedField === "number"
          ? indexedField - 1
          : fields.items.findIndex(
              (entry) =>
                matchesTextQuery(entry.code, anchor.text || anchor.label) ||
                matchesTextQuery(entry.result.text, anchor.text || anchor.label) ||
                (supportsFieldMetadata && matchesTextQuery(String(entry.type), anchor.text || anchor.label)),
            );
      const field = fieldIndex >= 0 ? fields.items[fieldIndex] : undefined;
      if (field) {
        if (supportsFieldMetadata) {
          field.select();
        } else {
          field.result.select();
        }
        await context.sync();
        return {
          ok: true,
          host: "word",
          anchorKind: "field",
          fieldId: `field:${fieldIndex + 1}`,
          fieldIndex: fieldIndex + 1,
          fieldCode: truncateLabel(field.code, 180),
          text: truncateLabel(field.result.text, 180),
          fieldType: supportsFieldMetadata ? field.type : undefined,
        };
      }
    }

    const paragraphs = body.paragraphs;
    paragraphs.load(
      supportsParagraphIds
        ? "items/text,items/style,items/styleBuiltIn,items/uniqueLocalId"
        : "items/text,items/style,items/styleBuiltIn",
    );
    await context.sync();

    const paragraph =
      anchor.kind === "heading"
        ? paragraphs.items.find((entry) => matchesWordHeading(entry, anchor))
        : paragraphs.items.find((entry) => matchesWordParagraph(entry, anchor));
    if (paragraph) {
      paragraph.select();
      await context.sync();
      return {
        ok: true,
        host: "word",
        anchorKind: anchor.kind,
        paragraphId: supportsParagraphIds ? paragraph.uniqueLocalId : undefined,
        text: truncateLabel(paragraph.text, 180),
      };
    }

    const query = trimString(anchor.text) ?? trimString(anchor.label);
    if (query) {
      const matches = body.search(query, {
        matchCase: false,
        matchWholeWord: false,
      });
      matches.load("items");
      await context.sync();
      if (matches.items[0]) {
        matches.items[0].select();
        await context.sync();
        return { ok: true, host: "word", anchorKind: anchor.kind, query };
      }
    }

    throw new Error(`Could not find the requested Word anchor: ${anchor.label || anchor.text || anchor.kind}.`);
  });
}

