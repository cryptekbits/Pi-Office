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

      if (action.target.kind === "footnote" || action.target.kind === "endnote") {
        const note = await resolveWordNoteTarget(action.target);
        if (!note) {
          throw new Error(`Could not find the requested Word ${action.target.kind}: ${action.target.label || action.target.text || action.target.id || action.target.kind}.`);
        }
        resolvedTarget = { range: note.note.reference };
        return resolvedTarget;
      }

      resolvedTarget = { range: selection };
      return resolvedTarget;
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
      const { range } = await resolveTargetRange();
      range.insertHtml(content, placement);
      await context.sync();
      return { ok: true, host: "word", action: type };
    }

    if (type === "insertText") {
      const { range } = await resolveTargetRange();
      range.insertText(content, placement);
      await context.sync();
      return { ok: true, host: "word", action: type };
    }

    if (type === "insertOoxml") {
      const { range } = await resolveTargetRange();
      range.insertOoxml(content, placement);
      await context.sync();
      return { ok: true, host: "word", action: type };
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


