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

export async function collectWordState(base: OfficeStateUpdate): Promise<OfficeStateUpdate> {
  try {
    return await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const paragraphs = selection.paragraphs;
      const inlinePictures = selection.inlinePictures;
      const supportsShapes = supportsRequirementSet("WordApiDesktop", "1.2");
      const supportsViewportPages = supportsRequirementSet("WordApiDesktop", "1.2");
      const supportsWindowMetadata = supportsRequirementSet("WordApiDesktop", "1.4");
      const shapes = supportsShapes ? selection.shapes : undefined;

      selection.load("text");
      paragraphs.load("items/text,items/style,items/styleBuiltIn");
      inlinePictures.load("items/altTextTitle,items/altTextDescription");
      shapes?.load("items/type,items/name,items/altTextDescription");
      await context.sync();

      const preview = normalizeTextPreview(selection.text);
      const details = paragraphs.items.map((paragraph) => paragraph.text.trim()).filter(Boolean).slice(0, 2);
      const imageShapes = shapes?.items.filter((shape) => isWordPictureShape(shape.type)) ?? [];
      const imageCount = countSelectedWordImages(inlinePictures.items.length, imageShapes.length);
      const objectCount = shapes?.items.length ?? 0;
      const altPreview =
        inlinePictures.items
          .map((picture) => picture.altTextTitle || picture.altTextDescription)
          .find((value) => Boolean(value?.trim())) ??
        imageShapes
          .map((shape) => shape.altTextDescription || shape.name)
          .find((value) => Boolean(value?.trim()));
      const structuredPreview = buildStructuredSelectionPreview(paragraphs.items);
      const selectionMeta = paragraphs.items.length > 0 ? buildSelectionMeta(paragraphs.items) : undefined;
      const selectionSummary: OfficeSelectionSummary = {
        ...buildSelectionSummary({
          textPreview: preview,
          imageCount,
          objectCount: imageCount > 0 ? Math.max(objectCount - imageShapes.length, 0) : objectCount,
          altPreview: normalizeTextPreview(altPreview),
          details: uniqueDetails([
            ...details,
            imageCount > 0 ? pluralize(imageCount, "image") : undefined,
            objectCount > imageShapes.length ? pluralize(objectCount - imageShapes.length, "shape") : undefined,
          ]),
          emptyLabel: "Insertion point",
        }),
        structuredPreview,
        selectionMeta,
      };

      return {
        ...base,
        selection: selectionSummary,
        capabilities: [
          "word.selection",
          "word.inlinePictures",
          ...(supportsShapes ? ["word.shapes"] : []),
          ...(supportsViewportPages ? ["word.viewportPages", "word.selectionPages"] : []),
          ...(supportsViewportPages ? ["word.viewportCapture"] : []),
          ...(supportsWindowMetadata ? ["word.activeWindow", "word.view"] : []),
          "word.insertText",
          "word.insertHtml",
          "office.getSelectedData",
        ],
      };
    });
  } catch {
    const [preview, html, ooxml] = await Promise.all([
      getSelectedTextAsync().catch(() => ""),
      getSelectedMarkupAsync(Office.CoercionType.Html).catch(() => ""),
      getSelectedMarkupAsync(Office.CoercionType.Ooxml).catch(() => ""),
    ]);
    const imageCount = containsImageMarkup(html) || containsImageMarkup(ooxml) ? 1 : 0;
    return {
      ...base,
      selection: buildSelectionSummary({
        textPreview: preview,
        imageCount,
        details: uniqueDetails([imageCount > 0 ? "Image content detected" : undefined]),
        emptyLabel: "Insertion point",
      }),
      capabilities: ["office.getSelectedData", "office.setSelectedData"],
    };
  }
}


export async function collectWordContext(base: OfficeStateUpdate, options: OfficeCaptureOptions): Promise<OfficeContextPayload> {
  const maxImages = Math.max(0, Math.min(options.maxImages ?? 0, 4));
  const payload = await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const body = context.document.body;
    const font = selection.font;
    const paragraphs = selection.paragraphs;
    const bodyParagraphs = body.paragraphs;
    const inlinePictures = selection.inlinePictures;
    const supportsShapes = supportsRequirementSet("WordApiDesktop", "1.2");
    const supportsViewportPages = supportsRequirementSet("WordApiDesktop", "1.2");
    const supportsPageSetup = supportsRequirementSet("WordApiDesktop", "1.3");
    const supportsWindowMetadata = supportsRequirementSet("WordApiDesktop", "1.4");
    const supportsComments = supportsRequirementSet("WordApi", "1.4");
    const supportsNotes = supportsRequirementSet("WordApi", "1.5");
    const supportsTrackedChanges = supportsRequirementSet("WordApi", "1.6");
    const supportsParagraphIds = supportsRequirementSet("WordApi", "1.6");
    const supportsReviewedText = supportsRequirementSet("WordApi", "1.4");
    const supportsFields = supportsRequirementSet("WordApi", "1.4");
    const supportsFieldMetadata = supportsRequirementSet("WordApi", "1.5");
    const supportsContentControls = supportsRequirementSet("WordApi", "1.1");
    const supportsContentControlSubtypes = supportsRequirementSet("WordApi", "1.3");
    const supportsDesktopLists = supportsRequirementSet("WordApiDesktop", "1.3");
    const supportsTables = supportsRequirementSet("WordApi", "1.3");
    const shapes = supportsShapes ? selection.shapes : undefined;
    const comments = supportsComments ? selection.getComments() : undefined;
    const footnotes = supportsNotes ? body.footnotes : undefined;
    const endnotes = supportsNotes ? body.endnotes : undefined;
    const trackedChanges = supportsTrackedChanges ? selection.getTrackedChanges() : undefined;
    const documentComments = supportsComments ? body.getComments() : undefined;
    const documentTrackedChanges = supportsTrackedChanges ? body.getTrackedChanges() : undefined;
    const selectionFields = supportsFields ? selection.fields : undefined;
    const documentFields = supportsFields ? body.fields : undefined;
    const selectionContentControls = supportsContentControls ? selection.contentControls : undefined;
    const documentContentControls = supportsContentControls ? body.contentControls : undefined;
    const documentTables = supportsTables ? body.tables : undefined;
    const documentSections = context.document.sections;
    const pageSetup = supportsPageSetup ? context.document.pageSetup : undefined;
    const selectionListFormat = supportsDesktopLists ? selection.listFormat : undefined;
    const activeWindow = supportsViewportPages ? context.document.activeWindow : undefined;
    const activePane = activeWindow?.activePane;
    const viewportPages = activePane?.pagesEnclosingViewport;
    const selectionPages = supportsViewportPages ? selection.pages : undefined;
    const view = supportsWindowMetadata ? activeWindow?.view : undefined;
    const reviewedSelectionCurrent = supportsReviewedText ? selection.getReviewedText("Current") : undefined;
    const reviewedSelectionOriginal = supportsReviewedText ? selection.getReviewedText("Original") : undefined;
    const reviewedDocumentCurrent = supportsReviewedText ? body.getReviewedText("Current") : undefined;
    const reviewedDocumentOriginal = supportsReviewedText ? body.getReviewedText("Original") : undefined;

    selection.load("text");
    font.load(["name", "size", "color", "bold", "italic", "underline", "underlineColor", "highlightColor"]);
    paragraphs.load(
      supportsParagraphIds
        ? "items/text,items/uniqueLocalId,items/style,items/styleBuiltIn,items/alignment,items/leftIndent,items/rightIndent,items/firstLineIndent,items/lineSpacing,items/spaceBefore,items/spaceAfter"
        : "items/text,items/style,items/styleBuiltIn,items/alignment,items/leftIndent,items/rightIndent,items/firstLineIndent,items/lineSpacing,items/spaceBefore,items/spaceAfter",
    );
    bodyParagraphs.load(
      supportsParagraphIds
        ? "items/text,items/uniqueLocalId,items/style,items/styleBuiltIn"
        : "items/text,items/style,items/styleBuiltIn",
    );
    inlinePictures.load("items/altTextTitle,items/altTextDescription,items/width,items/height,items/imageFormat");
    shapes?.load("items/type,items/name,items/left,items/top,items/width,items/height,items/rotation,items/altTextDescription");
    comments?.load("items/id,items/authorName,items/content,items/resolved,items/creationDate");
    footnotes?.load("items/type,items/body/text,items/reference/text");
    endnotes?.load("items/type,items/body/text,items/reference/text");
    trackedChanges?.load("items/author,items/date,items/text,items/type");
    documentComments?.load("items/id,items/authorName,items/content,items/resolved,items/creationDate");
    documentTrackedChanges?.load("items/author,items/date,items/text,items/type");
    selectionFields?.load(
      supportsFieldMetadata ? "items/code,items/type,items/locked,items/result/text" : "items/code,items/result/text",
    );
    documentFields?.load(
      supportsFieldMetadata ? "items/code,items/type,items/locked,items/result/text" : "items/code,items/result/text",
    );
    selectionContentControls?.load(
      supportsContentControlSubtypes
        ? "items/id,items/title,items/tag,items/type,items/subtype,items/appearance,items/cannotDelete,items/cannotEdit,items/removeWhenEdited,items/placeholderText,items/text"
        : "items/id,items/title,items/tag,items/type,items/appearance,items/cannotDelete,items/cannotEdit,items/removeWhenEdited,items/placeholderText,items/text",
    );
    documentContentControls?.load(
      supportsContentControlSubtypes
        ? "items/id,items/title,items/tag,items/type,items/subtype,items/appearance,items/cannotDelete,items/cannotEdit,items/removeWhenEdited,items/placeholderText,items/text"
        : "items/id,items/title,items/tag,items/type,items/appearance,items/cannotDelete,items/cannotEdit,items/removeWhenEdited,items/placeholderText,items/text",
    );
    documentTables?.load("items/rowCount,items/values,items/style,items/styleBuiltIn,items/title,items/description");
    documentSections.load("items/body/text");
    pageSetup?.load("topMargin,bottomMargin,leftMargin,rightMargin,pageWidth,pageHeight");
    selectionListFormat?.load("listType,listLevelNumber,listString,listValue");
    viewportPages?.load("items/index,items/width,items/height");
    selectionPages?.load("items/index,items/width,items/height");
    if (supportsWindowMetadata) {
      activeWindow?.load([
        "caption",
        "height",
        "width",
        "usableHeight",
        "usableWidth",
        "horizontalPercentScrolled",
        "verticalPercentScrolled",
        "areRulersDisplayed",
        "areThumbnailsDisplayed",
        "isSplit",
        "isHorizontalScrollBarDisplayed",
        "isVerticalScrollBarDisplayed",
        "isVisible",
        "left",
        "top",
        "windowNumber",
      ]);
    }
    view?.load("type,seekView");
    await context.sync();

    const inlineImages = inlinePictures.items.slice(0, maxImages);
    const base64Results = inlineImages.map((picture) => picture.getBase64ImageSrc());
    if (base64Results.length) {
      await context.sync();
    }

    const readSelectionSnippet = async (
      kind: "html" | "ooxml",
      getter: () => OfficeExtension.ClientResult<string>,
      maxLength: number,
    ): Promise<string | undefined> => {
      try {
        const result = getter();
        await context.sync();
        return truncateText(result.value, maxLength);
      } catch (error) {
        console.warn(`[office-word] Failed to read selection ${kind}.`, serializeOfficeRuntimeError(error));
        return undefined;
      }
    };

    const hasSelectedContent = Boolean(selection.text?.trim()) || inlinePictures.items.length > 0;
    const html = hasSelectedContent ? await readSelectionSnippet("html", () => selection.getHtml(), 2000) : undefined;
    const ooxml = hasSelectedContent ? await readSelectionSnippet("ooxml", () => selection.getOoxml(), 3000) : undefined;

    const preview = normalizeTextPreview(selection.text);
    const details = paragraphs.items.map((paragraph) => paragraph.text.trim()).filter(Boolean).slice(0, 2);
    const imageShapes = shapes?.items.filter((shape) => isWordPictureShape(shape.type)) ?? [];
    const imageCount = countSelectedWordImages(inlinePictures.items.length, imageShapes.length);
    const objectCount = shapes?.items.length ?? 0;
    const altPreview =
      inlinePictures.items
        .map((picture) => picture.altTextTitle || picture.altTextDescription)
        .find((value) => Boolean(value?.trim())) ??
      imageShapes
        .map((shape) => shape.altTextDescription || shape.name)
        .find((value) => Boolean(value?.trim()));
    const contextStructuredPreview = buildStructuredSelectionPreview(paragraphs.items);
    const contextSelectionMeta = paragraphs.items.length > 0
      ? buildSelectionMeta(paragraphs.items, { name: font.name, size: font.size, color: font.color, bold: font.bold, italic: font.italic })
      : undefined;
    const state: OfficeStateUpdate = {
      ...base,
      selection: {
        ...buildSelectionSummary({
          textPreview: preview,
          imageCount,
          objectCount: imageCount > 0 ? Math.max(objectCount - imageShapes.length, 0) : objectCount,
          altPreview: normalizeTextPreview(altPreview),
          details: uniqueDetails([
            ...details,
            imageCount > 0 ? pluralize(imageCount, "image") : undefined,
            objectCount > imageShapes.length ? pluralize(objectCount - imageShapes.length, "shape") : undefined,
          ]),
          emptyLabel: "Insertion point",
        }),
        structuredPreview: contextStructuredPreview,
        selectionMeta: contextSelectionMeta,
      },
      capabilities: [
        "word.selection",
        "word.document",
        "word.inlinePictures",
        ...(supportsShapes ? ["word.shapes"] : []),
        ...(supportsPageSetup ? ["word.pageSetup"] : []),
        ...(supportsViewportPages ? ["word.viewportPages", "word.selectionPages"] : []),
        ...(supportsViewportPages ? ["word.viewportCapture"] : []),
        ...(supportsWindowMetadata ? ["word.activeWindow", "word.view"] : []),
        ...(supportsComments ? ["word.comments"] : []),
        ...(supportsNotes ? ["word.footnotes", "word.endnotes"] : []),
        ...(supportsTrackedChanges ? ["word.trackedChanges"] : []),
        ...(supportsReviewedText ? ["word.reviewedText"] : []),
        ...(supportsFields ? ["word.fields"] : []),
        ...(supportsContentControls ? ["word.contentControls"] : []),
        ...(supportsDesktopLists ? ["word.listFormat"] : []),
        "word.insertText",
        "word.insertHtml",
        "word.insertOoxml",
        "word.insertFileFromBase64",
        ...(supportsContentControls ? ["word.insertContentControl"] : []),
        ...(supportsFields && supportsFieldMetadata ? ["word.insertField"] : []),
        ...(supportsComments ? ["word.commentThreads"] : []),
        ...(supportsTrackedChanges ? ["word.revisionActions"] : []),
        "office.getSelectedData",
      ],
    };

    const headingParagraphs = bodyParagraphs.items.filter((paragraph) => /heading/i.test(String(paragraph.styleBuiltIn || paragraph.style || "")));
    const reviewComments = documentComments?.items ?? comments?.items ?? [];
    const reviewChanges = documentTrackedChanges?.items ?? trackedChanges?.items ?? [];
    const documentFootnotes = footnotes?.items ?? [];
    const documentEndnotes = endnotes?.items ?? [];
    const documentFieldsList = documentFields?.items ?? selectionFields?.items ?? [];
    const selectedFields = selectionFields?.items ?? [];
    const documentContentControlList = documentContentControls?.items ?? selectionContentControls?.items ?? [];
    const selectedContentControls = selectionContentControls?.items ?? [];
    const documentTableList = documentTables?.items ?? [];
    const documentSectionList = documentSections.items ?? [];
    const paragraphMap = bodyParagraphs.items.slice(0, 40).map((paragraph, index) => ({
      kind: /heading/i.test(String(paragraph.styleBuiltIn || paragraph.style || "")) ? "heading" : "paragraph",
      index,
      text: truncateLabel(paragraph.text, 180),
      paragraphId: supportsParagraphIds ? paragraph.uniqueLocalId : undefined,
      style: paragraph.style || paragraph.styleBuiltIn,
      isHeading: /heading/i.test(String(paragraph.styleBuiltIn || paragraph.style || "")),
    }));
    const reviewedText = supportsReviewedText
      ? {
          selectionCurrentPreview: truncateText(reviewedSelectionCurrent?.value, 320),
          selectionOriginalPreview: truncateText(reviewedSelectionOriginal?.value, 320),
          documentCurrentPreview: truncateText(reviewedDocumentCurrent?.value, 480),
          documentOriginalPreview: truncateText(reviewedDocumentOriginal?.value, 480),
          selectionHasTrackedDifference:
            normalizeTextPreview(reviewedSelectionCurrent?.value) !== normalizeTextPreview(reviewedSelectionOriginal?.value),
          documentHasTrackedDifference:
            normalizeTextPreview(reviewedDocumentCurrent?.value) !== normalizeTextPreview(reviewedDocumentOriginal?.value),
        }
      : undefined;

    let visuals = await Promise.all(
      inlineImages.map(async (picture, index) =>
        optimizeVisual(
          {
            kind: "inline-picture",
            label: picture.altTextTitle || picture.altTextDescription || `Selected inline picture ${index + 1}`,
            data: base64Results[index]?.value ?? "",
            mimeType: String(picture.imageFormat ?? "png").toLowerCase() === "jpeg" ? "image/jpeg" : "image/png",
            width: picture.width,
            height: picture.height,
          },
          { maxDimension: 1400 },
        ),
      ),
    );

    if (!visuals.length && maxImages > 0) {
      const renderedSelection = await renderHtmlSelectionSnapshot({
        html: html ?? "",
        label: "Rendered selection snapshot",
        widthPx: getWordSnapshotWidthPx(pageSetup?.pageWidth, pageSetup?.leftMargin, pageSetup?.rightMargin),
      });
      if (renderedSelection) {
        visuals = [await optimizeVisual(renderedSelection, { maxDimension: 1600 })];
      }
    }

    const formatting = options.includeFormatting
      ? {
          selectionFont: {
            name: font.name,
            size: font.size,
            color: font.color,
            bold: font.bold,
            italic: font.italic,
            underline: font.underline,
            underlineColor: font.underlineColor,
            highlightColor: font.highlightColor,
          },
          pageSetup: pageSetup
            ? {
                pageWidth: formatPoints(pageSetup.pageWidth),
                pageHeight: formatPoints(pageSetup.pageHeight),
                topMargin: formatPoints(pageSetup.topMargin),
                bottomMargin: formatPoints(pageSetup.bottomMargin),
                leftMargin: formatPoints(pageSetup.leftMargin),
                rightMargin: formatPoints(pageSetup.rightMargin),
              }
            : undefined,
          viewport:
            supportsViewportPages && activeWindow
              ? {
                  window: supportsWindowMetadata
                    ? {
                        caption: activeWindow.caption,
                        width: formatPoints(activeWindow.width),
                        height: formatPoints(activeWindow.height),
                        usableWidth: formatPoints(activeWindow.usableWidth),
                        usableHeight: formatPoints(activeWindow.usableHeight),
                        left: formatPoints(activeWindow.left),
                        top: formatPoints(activeWindow.top),
                        horizontalPercentScrolled: formatPercent(activeWindow.horizontalPercentScrolled),
                        verticalPercentScrolled: formatPercent(activeWindow.verticalPercentScrolled),
                        areRulersDisplayed: activeWindow.areRulersDisplayed,
                        areThumbnailsDisplayed: activeWindow.areThumbnailsDisplayed,
                        isSplit: activeWindow.isSplit,
                        isHorizontalScrollBarDisplayed: activeWindow.isHorizontalScrollBarDisplayed,
                        isVerticalScrollBarDisplayed: activeWindow.isVerticalScrollBarDisplayed,
                        isVisible: activeWindow.isVisible,
                        windowNumber: activeWindow.windowNumber,
                      }
                    : undefined,
                  view: supportsWindowMetadata && view ? { type: view.type, seekView: view.seekView } : undefined,
                  pagesEnclosingViewport: viewportPages ? serializeWordPages(viewportPages.items) : [],
                  selectionPages: selectionPages ? serializeWordPages(selectionPages.items) : [],
                }
              : undefined,
          paragraphs: paragraphs.items.slice(0, 3).map((paragraph) => ({
            text: truncateText(paragraph.text, 280),
            style: paragraph.style || paragraph.styleBuiltIn,
            alignment: paragraph.alignment,
            firstLineIndent: formatPoints(paragraph.firstLineIndent),
            leftIndent: formatPoints(paragraph.leftIndent),
            rightIndent: formatPoints(paragraph.rightIndent),
            lineSpacing: formatPoints(paragraph.lineSpacing),
            spaceBefore: formatPoints(paragraph.spaceBefore),
            spaceAfter: formatPoints(paragraph.spaceAfter),
          })),
          selectedShapes: shapes?.items.slice(0, 4).map((shape) => ({
            name: shape.name,
            type: shape.type,
            left: formatPoints(shape.left),
            top: formatPoints(shape.top),
            width: formatPoints(shape.width),
            height: formatPoints(shape.height),
            rotation: shape.rotation,
            altTextDescription: shape.altTextDescription,
          })),
          selectionList: supportsDesktopLists && selectionListFormat
            ? {
                listType: selectionListFormat.listType,
                level: selectionListFormat.listLevelNumber,
                listString: selectionListFormat.listString,
                listValue: selectionListFormat.listValue,
              }
            : undefined,
        }
      : undefined;

    const result: OfficeContextPayload = {
      summary: "",
      state,
      anchors: uniqueAnchors([
        ...headingParagraphs.slice(0, 20).map((paragraph) => ({
          kind: "heading",
          label: truncateLabel(paragraph.text),
          text: truncateLabel(paragraph.text, 240),
          paragraphId: supportsParagraphIds ? paragraph.uniqueLocalId : undefined,
        }) as OfficeAnchor),
        ...bodyParagraphs.items.slice(0, 24).map((paragraph) => {
          const style = String(paragraph.styleBuiltIn || paragraph.style || "");
          const kind = /heading/i.test(style) ? "heading" : "paragraph";
          return {
            kind: kind as OfficeAnchor["kind"],
            label: truncateLabel(paragraph.text),
            text: truncateLabel(paragraph.text, 240),
            paragraphId: supportsParagraphIds ? paragraph.uniqueLocalId : undefined,
          } as OfficeAnchor;
        }),
        ...(reviewComments.slice(0, 12).map((comment) => ({
          kind: "comment",
          label: truncateLabel(comment.content) || `Comment by ${comment.authorName || "unknown author"}`,
          text: truncateLabel(comment.content, 240),
          commentId: comment.id,
        }) as OfficeAnchor) ?? []),
        ...(reviewChanges.slice(0, 12).map((change, index) => ({
          kind: "revision",
          id: `revision:${index + 1}`,
          label: truncateLabel(change.text) || `${change.type} revision`,
          text: truncateLabel(change.text, 240),
          revisionId: `revision:${index + 1}`,
        }) as OfficeAnchor) ?? []),
        ...documentFootnotes.slice(0, 12).map((note, index) => ({
          kind: "footnote",
          id: `footnote:${index + 1}`,
          label: `Footnote ${index + 1}`,
          text: truncateLabel(note.body.text || note.reference.text, 240),
          noteTarget: "body",
        }) as OfficeAnchor),
        ...documentEndnotes.slice(0, 12).map((note, index) => ({
          kind: "endnote",
          id: `endnote:${index + 1}`,
          label: `Endnote ${index + 1}`,
          text: truncateLabel(note.body.text || note.reference.text, 240),
          noteTarget: "body",
        }) as OfficeAnchor),
        ...documentContentControlList.slice(0, 12).map((control) => ({
          kind: "contentControl",
          id: `contentControl:${control.id}`,
          label: truncateLabel(control.title || control.tag || control.text || `Content control ${control.id}`),
          text: truncateLabel(control.text, 240),
        }) as OfficeAnchor),
        ...documentFieldsList.slice(0, 12).map((field, index) => ({
          kind: "field",
          id: `field:${index + 1}`,
          label: truncateLabel(
            `${supportsFieldMetadata ? `${field.type || "Field"} · ` : ""}${field.code || field.result.text || `Field ${index + 1}`}`,
            120,
          ),
          text: truncateLabel(field.code || field.result.text, 240),
        }) as OfficeAnchor),
      ]),
      formatting,
      snippets: {
        html,
        ooxml,
        documentStructure: {
          paragraphs: bodyParagraphs.items.length,
          headings: headingParagraphs.length,
          comments: reviewComments.length,
          footnotes: documentFootnotes.length,
          endnotes: documentEndnotes.length,
          revisions: reviewChanges.length,
          fields: documentFieldsList.length,
          contentControls: documentContentControlList.length,
          tables: documentTableList.length,
          sections: documentSectionList.length,
        },
        comments: reviewComments.slice(0, 8).map((comment) => ({
          id: comment.id,
          author: comment.authorName,
          content: truncateLabel(comment.content, 160),
          resolved: comment.resolved,
          creationDate: comment.creationDate?.toISOString?.(),
        })),
        trackedChanges: reviewChanges.slice(0, 8).map((change, index) => ({
          id: `revision:${index + 1}`,
          author: change.author,
          date: change.date?.toISOString?.(),
          type: change.type,
          text: truncateLabel(change.text, 160),
        })),
        headings: headingParagraphs.slice(0, 20).map((paragraph) => ({
          text: truncateLabel(paragraph.text, 160),
          paragraphId: supportsParagraphIds ? paragraph.uniqueLocalId : undefined,
        })),
        paragraphs: paragraphMap,
        footnotes: documentFootnotes.slice(0, 8).map((note, index) => ({
          id: `footnote:${index + 1}`,
          referenceText: truncateLabel(note.reference.text, 80),
          text: truncateLabel(note.body.text, 160),
        })),
        endnotes: documentEndnotes.slice(0, 8).map((note, index) => ({
          id: `endnote:${index + 1}`,
          referenceText: truncateLabel(note.reference.text, 80),
          text: truncateLabel(note.body.text, 160),
        })),
        fields: documentFieldsList.slice(0, 12).map((field, index) => ({
          id: `field:${index + 1}`,
          type: supportsFieldMetadata ? field.type : undefined,
          code: truncateLabel(field.code, 180),
          resultText: truncateLabel(field.result.text, 160),
          locked: supportsFieldMetadata ? field.locked : undefined,
        })),
        selectedFields: selectedFields.slice(0, 8).map((field, index) => ({
          id: `selectionField:${index + 1}`,
          type: supportsFieldMetadata ? field.type : undefined,
          code: truncateLabel(field.code, 160),
          resultText: truncateLabel(field.result.text, 120),
        })),
        contentControls: documentContentControlList.slice(0, 12).map((control) => ({
          id: `contentControl:${control.id}`,
          title: trimString(control.title),
          tag: trimString(control.tag),
          type: control.type,
          subtype: supportsContentControlSubtypes ? control.subtype : undefined,
          text: truncateLabel(control.text, 160),
          placeholderText: trimString(control.placeholderText),
          appearance: control.appearance,
          cannotDelete: control.cannotDelete,
          cannotEdit: control.cannotEdit,
          removeWhenEdited: control.removeWhenEdited,
        })),
        tables: documentTableList.slice(0, 8).map((table, index) => ({
          id: `table:${index + 1}`,
          title: trimString(table.title),
          description: trimString(table.description),
          rowCount: table.rowCount,
          columnCount: table.values?.[0]?.length ?? 0,
          style: table.style || table.styleBuiltIn,
          preview: table.values?.slice(0, 3).map((row) => row.slice(0, 5)) ?? [],
        })),
        sections: documentSectionList.slice(0, 8).map((section, index) => ({
          id: `section:${index + 1}`,
          index: index + 1,
          bodyPreview: truncateLabel(section.body.text, 200),
        })),
        selectedContentControls: selectedContentControls.slice(0, 8).map((control) => ({
          id: `contentControl:${control.id}`,
          title: trimString(control.title),
          tag: trimString(control.tag),
          type: control.type,
          text: truncateLabel(control.text, 120),
        })),
        reviewedText,
      },
      visuals,
    };
    result.summary = createSummary(result);
    return result;
  });

  if (!payload.visuals?.length && maxImages > 0 && (payload.state.selection.imageCount ?? 0) > 0) {
    const fallback = await getSelectedImageAsync().catch(() => undefined);
    if (fallback) {
      payload.visuals = [await optimizeVisual(fallback, { maxDimension: 1400 })];
      payload.summary = createSummary(payload);
    }
  }

  return payload;
}

function wordSearchMatchesText(text: string | undefined, query: string): boolean {
  return Boolean(text && text.toLowerCase().includes(query.toLowerCase()));
}

export async function searchWordDocument(params: Record<string, unknown>): Promise<unknown> {
  const query = trimString(params.query);
  if (!query) {
    return { ok: false, error: "word_search requires query." };
  }

  const maxResults = typeof params.maxResults === "number"
    ? Math.max(1, Math.min(50, Math.trunc(params.maxResults)))
    : 20;

  return Word.run(async (context) => {
    const body = context.document.body;
    const supportsParagraphIds = supportsRequirementSet("WordApi", "1.6");
    const supportsComments = supportsRequirementSet("WordApi", "1.4");
    const supportsNotes = supportsRequirementSet("WordApi", "1.5");
    const supportsFields = supportsRequirementSet("WordApi", "1.4");
    const supportsContentControls = supportsRequirementSet("WordApi", "1.1");
    const supportsTrackedChanges = supportsRequirementSet("WordApi", "1.6");

    const paragraphs = body.paragraphs;
    const comments = supportsComments ? body.getComments() : undefined;
    const footnotes = supportsNotes ? body.footnotes : undefined;
    const endnotes = supportsNotes ? body.endnotes : undefined;
    const fields = supportsFields ? body.fields : undefined;
    const contentControls = supportsContentControls ? body.contentControls : undefined;
    const trackedChanges = supportsTrackedChanges ? body.getTrackedChanges() : undefined;
    const nativeMatches = body.search(query, { matchCase: false, matchWholeWord: false });

    paragraphs.load(
      supportsParagraphIds
        ? "items/text,items/style,items/styleBuiltIn,items/uniqueLocalId"
        : "items/text,items/style,items/styleBuiltIn",
    );
    nativeMatches.load("items/text");
    comments?.load("items/id,items/content,items/authorName,items/resolved");
    footnotes?.load("items/body/text,items/reference/text");
    endnotes?.load("items/body/text,items/reference/text");
    fields?.load("items/code,items/result/text");
    contentControls?.load("items/id,items/title,items/tag,items/text,type");
    trackedChanges?.load("items/author,items/text,items/type");
    await context.sync();

    const results: Array<Record<string, unknown>> = [];
    const push = (entry: Record<string, unknown>) => {
      if (results.length < maxResults) results.push(entry);
    };

    paragraphs.items.forEach((paragraph, index) => {
      if (!wordSearchMatchesText(paragraph.text, query)) return;
      const style = String(paragraph.styleBuiltIn || paragraph.style || "");
      push({
        anchor: {
          kind: /heading/i.test(style) ? "heading" : "paragraph",
          label: truncateLabel(paragraph.text),
          text: truncateLabel(paragraph.text, 240),
          paragraphId: supportsParagraphIds ? paragraph.uniqueLocalId : undefined,
        },
        objectType: /heading/i.test(style) ? "heading" : "paragraph",
        rank: index + 1,
        contextPreview: truncateText(paragraph.text, 320),
      });
    });

    nativeMatches.items.slice(0, maxResults).forEach((match, index) => {
      if (results.length >= maxResults) return;
      push({
        anchor: { kind: "range", label: `Search match ${index + 1}`, text: truncateLabel(match.text, 180) },
        objectType: "range",
        rank: index + 1,
        contextPreview: truncateText(match.text, 320),
      });
    });

    (comments?.items ?? []).forEach((comment, index) => {
      if (!wordSearchMatchesText(comment.content, query)) return;
      push({
        anchor: { kind: "comment", commentId: comment.id, label: `Comment ${index + 1}`, text: truncateLabel(comment.content, 180) },
        objectType: "comment",
        rank: index + 1,
        contextPreview: truncateText(comment.content, 320),
        metadata: { authorName: comment.authorName, resolved: comment.resolved },
      });
    });

    (trackedChanges?.items ?? []).forEach((change, index) => {
      if (!wordSearchMatchesText(change.text, query)) return;
      push({
        anchor: { kind: "revision", revisionId: `revision:${index + 1}`, label: `Revision ${index + 1}`, text: truncateLabel(change.text, 180) },
        objectType: "revision",
        rank: index + 1,
        contextPreview: truncateText(change.text, 320),
        metadata: { author: change.author, type: change.type },
      });
    });

    (fields?.items ?? []).forEach((field, index) => {
      const fieldText = `${field.code ?? ""} ${field.result?.text ?? ""}`;
      if (!wordSearchMatchesText(fieldText, query)) return;
      push({
        anchor: { kind: "field", id: `field:${index + 1}`, label: `Field ${index + 1}`, text: truncateLabel(field.result?.text || field.code, 180) },
        objectType: "field",
        rank: index + 1,
        contextPreview: truncateText(fieldText, 320),
      });
    });

    (contentControls?.items ?? []).forEach((control) => {
      const controlText = `${control.title ?? ""} ${control.tag ?? ""} ${control.text ?? ""}`;
      if (!wordSearchMatchesText(controlText, query)) return;
      push({
        anchor: { kind: "contentControl", id: `contentControl:${control.id}`, label: control.title || control.tag || `Content control ${control.id}`, text: truncateLabel(control.text, 180) },
        objectType: "contentControl",
        rank: control.id,
        contextPreview: truncateText(controlText, 320),
      });
    });

    (footnotes?.items ?? []).forEach((note, index) => {
      if (!wordSearchMatchesText(note.body.text, query) && !wordSearchMatchesText(note.reference.text, query)) return;
      push({
        anchor: { kind: "footnote", id: `footnote:${index + 1}`, label: `Footnote ${index + 1}`, text: truncateLabel(note.body.text, 180), noteTarget: "body" },
        objectType: "footnote",
        rank: index + 1,
        contextPreview: truncateText(note.body.text, 320),
      });
    });

    (endnotes?.items ?? []).forEach((note, index) => {
      if (!wordSearchMatchesText(note.body.text, query) && !wordSearchMatchesText(note.reference.text, query)) return;
      push({
        anchor: { kind: "endnote", id: `endnote:${index + 1}`, label: `Endnote ${index + 1}`, text: truncateLabel(note.body.text, 180), noteTarget: "body" },
        objectType: "endnote",
        rank: index + 1,
        contextPreview: truncateText(note.body.text, 320),
      });
    });

    return {
      ok: true,
      host: "word",
      query,
      totalReturned: results.length,
      maxResults,
      results,
      capabilities: {
        paragraphIds: supportsParagraphIds,
        comments: supportsComments,
        notes: supportsNotes,
        fields: supportsFields,
        contentControls: supportsContentControls,
        trackedChanges: supportsTrackedChanges,
      },
    };
  });
}


