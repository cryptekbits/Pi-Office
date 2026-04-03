import type { OfficeAnchor, OfficeContextPayload, OfficeHostAction, OfficeStateUpdate, OfficeSelectionMeta, OfficeVisualSnapshot } from "@pi-office/pi-office-pack/protocol";
import type { OfficeCaptureOptions } from "../office-host-adapter-types";
import {
  createPowerPointChartInPresentationBase64,
  inspectPowerPointPresentationBase64,
  replaceSlideNotesInPowerPointPresentationBase64,
  updatePowerPointChartInPresentationBase64,
} from "../powerpoint-transform";
import {
  supportsRequirementSet,
  normalizeTextPreview,
  truncateText,
  pluralize,
  uniqueDetails,
  formatPoints,
  buildSelectionSummary,
  isRecord,
  trimString,
  truncateLabel,
  matchesTextQuery,
  parsePositiveInteger,
  clampPercentage,
  uniqueAnchors,
  getActionOptions,
  toNumber,
  toBoolean,
  toStringMatrix,
  isPowerPointImageShape,
  isPowerPointTableShape,
  isPowerPointGroupShape,
  getPowerPointShapeContentKind,
  truncateStringMatrix,
  resolveZeroBasedIndex,
  resolvePositiveCount,
  getSelectedTextAsync,
  setSelectedTextAsync,
  setSelectedImageAsync,
  getActionImagePayload,
  getSelectedImageAsync,
  optimizeVisual,
  createSummary,
  serializeOfficeRuntimeError,
  getStringArray,
  getNumberArray,
  getRecordArray,
  getChartValueArray,
  getChartSeriesInput,
} from "./shared";
import { inspectCurrentPowerPointPresentationPackage } from "./powerpoint-helpers";

export async function collectPowerPointState(base: OfficeStateUpdate): Promise<OfficeStateUpdate> {
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.getSelectedSlides();
      const supportsRichSelection = supportsRequirementSet("PowerPointApi", "1.5");
      const shapes = supportsRichSelection ? context.presentation.getSelectedShapes() : undefined;
      const textRange = supportsRichSelection ? context.presentation.getSelectedTextRangeOrNullObject() : undefined;

      slides.load("items/id,items/index");
      shapes?.load("items/type,items/name,items/id");
      textRange?.load("text,isNullObject");
      await context.sync();

      const preview =
        textRange && !textRange.isNullObject
          ? normalizeTextPreview(textRange.text)
          : normalizeTextPreview(await getSelectedTextAsync().catch(() => ""));
      const imageShapeCount = shapes?.items.filter((shape) => isPowerPointImageShape(shape.type)).length ?? 0;
      const objectCount = shapes?.items.length ?? 0;
      const details = uniqueDetails([
        ...slides.items.map((slide) => `Slide ${slide.index + 1}`).slice(0, 2),
        objectCount > 0 ? pluralize(objectCount, "shape") : undefined,
        imageShapeCount > 0 ? pluralize(imageShapeCount, "image") : undefined,
      ]);

      const pptShapeTypes = shapes?.items.map((s) => String(s.type ?? "shape")) ?? [];
      const pptStyleHistogram: Record<string, number> = {};
      for (const t of pptShapeTypes) {
        const kind = getPowerPointShapeContentKind(t);
        pptStyleHistogram[kind] = (pptStyleHistogram[kind] ?? 0) + 1;
      }
      const pptMeta: OfficeSelectionMeta | undefined = objectCount > 0 || preview
        ? {
            paragraphCount: preview ? 1 : undefined,
            firstParagraphStyle: pptShapeTypes[0] ? getPowerPointShapeContentKind(pptShapeTypes[0]) : undefined,
            styleHistogram: Object.keys(pptStyleHistogram).length > 0 ? pptStyleHistogram : undefined,
          }
        : undefined;

      return {
        ...base,
        selection: {
          ...buildSelectionSummary({
            textPreview: preview,
            imageCount: imageShapeCount,
            objectCount: imageShapeCount > 0 ? Math.max(objectCount - imageShapeCount, 0) : objectCount,
            altPreview: shapes?.items[0]?.name,
            details,
            emptyLabel: slides.items.length > 0 ? `${pluralize(slides.items.length, "slide")} selected` : "Slide selection",
          }),
          selectionMeta: pptMeta,
        },
        capabilities: [
          "powerpoint.selection",
          ...(supportsRichSelection ? ["powerpoint.shapes", "powerpoint.textRange"] : []),
          "powerpoint.slideSnapshot",
          "office.setSelectedData",
        ],
      };
    });
  } catch {
    const preview = await getSelectedTextAsync().catch(() => "");
    return {
      ...base,
      selection: buildSelectionSummary({
        textPreview: preview,
        emptyLabel: "Slide selection",
      }),
      capabilities: ["office.getSelectedData", "office.setSelectedData"],
    };
  }
}


export async function collectPowerPointContext(base: OfficeStateUpdate, options: OfficeCaptureOptions): Promise<OfficeContextPayload> {
  const maxImages = Math.max(0, Math.min(options.maxImages ?? 0, 4));
  const payload = await PowerPoint.run(async (context) => {
    const presentation = context.presentation;
    const slides = presentation.getSelectedSlides();
    const presentationSlides = presentation.slides;
    const supportsRichSelection = supportsRequirementSet("PowerPointApi", "1.5");
    const supportsLayoutMetadata = supportsRequirementSet("PowerPointApi", "1.3");
    const supportsSlideSnapshots = supportsRequirementSet("PowerPointApi", "1.8");
    const supportsShapeMetadata = supportsRequirementSet("PowerPointApi", "1.8");
    const supportsTableShapes = supportsRequirementSet("PowerPointApi", "1.8");
    const supportsTableEditing = supportsRequirementSet("PowerPointApi", "1.9");
    const supportsPageSetup = supportsRequirementSet("PowerPointApi", "1.10");
    const supportsShapeSnapshots = supportsRequirementSet("PowerPointApi", "1.10");
    const slideMasters = supportsLayoutMetadata ? presentation.slideMasters : undefined;
    const shapes = supportsRichSelection ? presentation.getSelectedShapes() : undefined;
    const textRange = supportsRichSelection ? presentation.getSelectedTextRangeOrNullObject() : undefined;
    const pageSetup = supportsPageSetup ? presentation.pageSetup : undefined;

    slides.load("items/id,items/index");
    presentationSlides.load("items/id,items/index");
    slideMasters?.load("items/id,items/name");
    const shapeLoadProperties = [
      "items/id",
      "items/name",
      "items/type",
      "items/left",
      "items/top",
      "items/width",
      "items/height",
      "items/rotation",
      "items/altTextTitle",
      "items/altTextDescription",
      ...(supportsShapeMetadata ? ["items/zOrderPosition"] : []),
      ...(supportsShapeSnapshots ? ["items/visible", "items/creationId"] : []),
    ];
    shapes?.load(shapeLoadProperties.join(","));
    textRange?.load(
      "isNullObject,text,start,length,font/name,font/size,font/color,font/bold,font/italic,font/underline,paragraphFormat/horizontalAlignment,paragraphFormat/indentLevel",
    );
    pageSetup?.load("slideHeight,slideWidth");
    await context.sync();

    if (supportsLayoutMetadata) {
      for (const slide of presentationSlides.items) {
        slide.layout.load("id,name,type");
        slide.slideMaster.load("id,name");
      }
      for (const slideMaster of slideMasters?.items ?? []) {
        slideMaster.layouts.load("items/id,items/name,items/type");
      }
    }

    const selectedShapes = shapes?.items.slice(0, Math.max(maxImages, 4)) ?? [];
    const textFrames =
      supportsShapeSnapshots && selectedShapes.length
        ? selectedShapes.map((shape) => shape.getTextFrameOrNullObject())
        : [];
    const tables =
      supportsTableShapes && selectedShapes.length
        ? selectedShapes.map((shape) => (isPowerPointTableShape(shape.type) ? shape.getTable() : undefined))
        : [];
    for (const textFrame of textFrames) {
      textFrame.load(
        "isNullObject,hasText,autoSizeSetting,topMargin,leftMargin,rightMargin,bottomMargin,verticalAlignment,wordWrap,textRange/text,textRange/font/name,textRange/font/size,textRange/font/color,textRange/font/bold,textRange/font/italic,textRange/font/underline,textRange/paragraphFormat/horizontalAlignment,textRange/paragraphFormat/indentLevel",
      );
    }
    for (const table of tables) {
      table?.load("rowCount,columnCount,values");
      if (supportsTableEditing) {
        table?.styleSettings.load(
          "style,areColumnsBanded,areRowsBanded,isFirstColumnHighlighted,isFirstRowHighlighted,isLastColumnHighlighted,isLastRowHighlighted",
        );
      }
    }

    const slideImageResults =
      maxImages > 0 && supportsSlideSnapshots && slides.items.length > 0 && selectedShapes.length === 0
        ? slides.items.slice(0, maxImages).map((slide) => slide.getImageAsBase64({ width: 1400 }))
        : [];
    const shapeImageResults =
      maxImages > 0 && supportsShapeSnapshots && selectedShapes.length > 0
        ? selectedShapes.slice(0, maxImages).map((shape) => shape.getImageAsBase64({ format: "Png", width: 1400 }))
        : [];

    if (supportsLayoutMetadata || textFrames.length || tables.some(Boolean) || slideImageResults.length || shapeImageResults.length) {
      await context.sync();
    }

    const slideMetadataById = new Map(
      presentationSlides.items.map((slide) => [
        slide.id,
        {
          id: slide.id,
          index: slide.index + 1,
          layoutId: supportsLayoutMetadata ? slide.layout.id : undefined,
          layoutName: supportsLayoutMetadata ? slide.layout.name : undefined,
          layoutType: supportsLayoutMetadata ? slide.layout.type : undefined,
          slideMasterId: supportsLayoutMetadata ? slide.slideMaster.id : undefined,
          slideMasterName: supportsLayoutMetadata ? slide.slideMaster.name : undefined,
        },
      ]),
    );
    const slideMetadataByIndex = new Map(Array.from(slideMetadataById.values()).map((slide) => [slide.index, slide]));
    const selectedSlideSummaries = slides.items.map(
      (slide) =>
        slideMetadataById.get(slide.id) ?? {
          id: slide.id,
          index: slide.index + 1,
          layoutId: undefined,
          layoutName: undefined,
          layoutType: undefined,
          slideMasterId: undefined,
          slideMasterName: undefined,
        },
    );
    const selectedSlideNotesAnchors = selectedSlideSummaries.slice(0, 12).map(
      (slide) =>
        ({
          kind: "notesRegion",
          id: `notes:${slide.index}`,
          label: `Notes for Slide ${slide.index}`,
          slideId: slide.id,
          slideIndex: slide.index,
          text: `Speaker notes for slide ${slide.index}`,
        }) as OfficeAnchor,
    );
    const masterSummaries = (slideMasters?.items ?? []).map((slideMaster) => ({
      id: slideMaster.id,
      name: slideMaster.name,
      layoutCount: slideMaster.layouts.items.length,
      layouts: slideMaster.layouts.items.map((layout) => ({ id: layout.id, name: layout.name, type: layout.type })),
    }));
    const layoutUsageMap = new Map<
      string,
      {
        id: string | undefined;
        name: string | undefined;
        type: string | undefined;
        slideMasterId: string | undefined;
        slideMasterName: string | undefined;
        count: number;
      }
    >();
    for (const slide of slideMetadataById.values()) {
      const key = slide.layoutId ?? slide.layoutName ?? `${slide.index}`;
      const existing = layoutUsageMap.get(key);
      if (existing) {
        existing.count += 1;
        continue;
      }

      layoutUsageMap.set(key, {
        id: slide.layoutId,
        name: slide.layoutName,
        type: slide.layoutType,
        slideMasterId: slide.slideMasterId,
        slideMasterName: slide.slideMasterName,
        count: 1,
      });
    }
    const layoutUsage = Array.from(layoutUsageMap.values()).sort((left, right) => right.count - left.count);
    const selectedShapeDescriptors = selectedShapes.slice(0, 8).map((shape, index) => {
      const textFrame = textFrames[index];
      const table = tables[index];
      const slideSummary = slides.items[0] ? slideMetadataById.get(slides.items[0].id) : undefined;
      const textPreview =
        textFrame && !textFrame.isNullObject && textFrame.hasText ? truncateText(textFrame.textRange.text, 400) : undefined;

      return {
        id: shape.id,
        name: shape.name,
        label: shape.name || truncateLabel(shape.altTextTitle || shape.altTextDescription) || `Shape ${shape.id}`,
        type: shape.type,
        contentKind: getPowerPointShapeContentKind(shape.type),
        slideId: slideSummary?.id,
        slideIndex: slideSummary?.index,
        layoutId: slideSummary?.layoutId,
        layoutName: slideSummary?.layoutName,
        slideMasterId: slideSummary?.slideMasterId,
        slideMasterName: slideSummary?.slideMasterName,
        left: formatPoints(shape.left),
        top: formatPoints(shape.top),
        width: formatPoints(shape.width),
        height: formatPoints(shape.height),
        rotation: shape.rotation,
        visible: supportsShapeSnapshots ? shape.visible : undefined,
        zOrderPosition: supportsShapeMetadata ? shape.zOrderPosition : undefined,
        creationId: supportsShapeSnapshots ? shape.creationId : undefined,
        altTextTitle: shape.altTextTitle,
        altTextDescription: shape.altTextDescription,
        textPreview,
        textFrame:
          textFrame && !textFrame.isNullObject
            ? {
                hasText: textFrame.hasText,
                autoSizeSetting: textFrame.autoSizeSetting,
                topMargin: formatPoints(textFrame.topMargin),
                rightMargin: formatPoints(textFrame.rightMargin),
                bottomMargin: formatPoints(textFrame.bottomMargin),
                leftMargin: formatPoints(textFrame.leftMargin),
                verticalAlignment: textFrame.verticalAlignment,
                wordWrap: textFrame.wordWrap,
                fontName: textFrame.hasText ? textFrame.textRange.font.name : undefined,
                fontSize: textFrame.hasText ? textFrame.textRange.font.size : undefined,
                fontColor: textFrame.hasText ? textFrame.textRange.font.color : undefined,
                bold: textFrame.hasText ? textFrame.textRange.font.bold : undefined,
                italic: textFrame.hasText ? textFrame.textRange.font.italic : undefined,
                underline: textFrame.hasText ? textFrame.textRange.font.underline : undefined,
                horizontalAlignment: textFrame.hasText ? textFrame.textRange.paragraphFormat.horizontalAlignment : undefined,
                indentLevel: textFrame.hasText ? textFrame.textRange.paragraphFormat.indentLevel : undefined,
              }
            : undefined,
        table:
          table && isPowerPointTableShape(shape.type)
            ? {
                rowCount: table.rowCount,
                columnCount: table.columnCount,
                previewValues: truncateStringMatrix(table.values),
                styleSettings: supportsTableEditing
                  ? {
                      style: table.styleSettings.style,
                      areRowsBanded: table.styleSettings.areRowsBanded,
                      areColumnsBanded: table.styleSettings.areColumnsBanded,
                      isFirstRowHighlighted: table.styleSettings.isFirstRowHighlighted,
                      isFirstColumnHighlighted: table.styleSettings.isFirstColumnHighlighted,
                      isLastRowHighlighted: table.styleSettings.isLastRowHighlighted,
                      isLastColumnHighlighted: table.styleSettings.isLastColumnHighlighted,
                    }
                  : undefined,
              }
            : undefined,
      };
    });

    const preview =
      textRange && !textRange.isNullObject
        ? normalizeTextPreview(textRange.text)
        : normalizeTextPreview(await getSelectedTextAsync().catch(() => ""));
    const imageShapeCount = selectedShapes.filter((shape) => isPowerPointImageShape(shape.type)).length;
    const objectCount = selectedShapes.length;
    const details = uniqueDetails([
      ...selectedSlideSummaries
        .slice(0, 2)
        .map((slide) => (slide.layoutName ? `Slide ${slide.index} · ${slide.layoutName}` : `Slide ${slide.index}`)),
      objectCount > 0 ? pluralize(objectCount, "shape") : undefined,
      imageShapeCount > 0 ? pluralize(imageShapeCount, "image") : undefined,
    ]);
    const ctxPptShapeTypes = selectedShapes.map((s) => String(s.type ?? "shape"));
    const ctxPptStyleHistogram: Record<string, number> = {};
    for (const t of ctxPptShapeTypes) {
      const kind = getPowerPointShapeContentKind(t);
      ctxPptStyleHistogram[kind] = (ctxPptStyleHistogram[kind] ?? 0) + 1;
    }
    const ctxPptMeta: OfficeSelectionMeta | undefined = objectCount > 0 || preview
      ? {
          paragraphCount: preview ? 1 : undefined,
          firstParagraphStyle: ctxPptShapeTypes[0] ? getPowerPointShapeContentKind(ctxPptShapeTypes[0]) : undefined,
          styleHistogram: Object.keys(ctxPptStyleHistogram).length > 0 ? ctxPptStyleHistogram : undefined,
        }
      : undefined;
    const state: OfficeStateUpdate = {
      ...base,
      selection: {
        ...buildSelectionSummary({
          textPreview: preview,
          imageCount: imageShapeCount,
          objectCount: imageShapeCount > 0 ? Math.max(objectCount - imageShapeCount, 0) : objectCount,
          altPreview: selectedShapeDescriptors[0]?.textPreview ?? selectedShapes[0]?.name,
          details,
          emptyLabel: slides.items.length > 0 ? `${pluralize(slides.items.length, "slide")} selected` : "Slide selection",
        }),
        selectionMeta: ctxPptMeta,
      },
      capabilities: [
        "powerpoint.selection",
        ...(supportsRichSelection ? ["powerpoint.shapes", "powerpoint.textRange"] : []),
        ...(supportsLayoutMetadata ? ["powerpoint.slideMasters", "powerpoint.layouts"] : []),
        ...(supportsTableShapes ? ["powerpoint.tables"] : []),
        "powerpoint.notesRegionAnchors",
        ...(supportsSlideSnapshots ? ["powerpoint.slideSnapshot"] : []),
        ...(supportsShapeSnapshots ? ["powerpoint.shapeSnapshot"] : []),
        "office.setSelectedData",
      ],
    };

    const visuals: OfficeVisualSnapshot[] = [];
    for (let index = 0; index < shapeImageResults.length; index += 1) {
      const shape = selectedShapes[index];
      if (!shape) {
        continue;
      }

      visuals.push(
        await optimizeVisual(
          {
            kind: "shape",
            label: shape.name || `Selected shape ${index + 1}`,
            data: shapeImageResults[index]?.value ?? "",
            mimeType: "image/png",
            width: shape.width,
            height: shape.height,
          },
          { maxDimension: 1400 },
        ),
      );
    }

    for (let index = 0; index < slideImageResults.length; index += 1) {
      const slide = slides.items[index];
      if (!slide) {
        continue;
      }

      visuals.push(
        await optimizeVisual(
          {
            kind: "slide",
            label: `Slide ${slide.index + 1}`,
            data: slideImageResults[index]?.value ?? "",
            mimeType: "image/png",
          },
          { maxDimension: 1400 },
        ),
      );
    }

    const result: OfficeContextPayload = {
      summary: "",
      state,
      anchors: uniqueAnchors([
        ...Array.from(slideMetadataById.values()).slice(0, 30).map((slide) => ({
          kind: "slide",
          label: `Slide ${slide.index}`,
          slideId: slide.id,
          slideIndex: slide.index,
        }) as OfficeAnchor),
        ...selectedSlideNotesAnchors,
        ...masterSummaries.slice(0, 12).map((slideMaster) => ({
          kind: "slideMaster",
          id: slideMaster.id,
          label: slideMaster.name || `Master ${slideMaster.id}`,
        }) as OfficeAnchor),
        ...masterSummaries.slice(0, 12).flatMap((slideMaster) =>
          slideMaster.layouts.slice(0, 12).map((layout) => ({
            kind: "layout",
            id: layout.id,
            label: layout.name || layout.type || `Layout ${layout.id}`,
            text: slideMaster.name,
          }) as OfficeAnchor),
        ),
        ...selectedShapes.slice(0, 12).map((shape) => ({
          kind: "shape",
          label: shape.name || truncateLabel(shape.altTextTitle || shape.altTextDescription) || `Shape ${shape.id}`,
          slideId: slides.items[0]?.id,
          slideIndex: slides.items[0] ? slides.items[0].index + 1 : undefined,
          shapeId: shape.id,
          id: shape.id,
        }) as OfficeAnchor),
      ]),
      formatting: options.includeFormatting
        ? {
            pageSetup: pageSetup ? { slideWidth: formatPoints(pageSetup.slideWidth), slideHeight: formatPoints(pageSetup.slideHeight) } : undefined,
            textSelection:
              textRange && !textRange.isNullObject
                ? {
                    text: truncateText(textRange.text, 600),
                    start: textRange.start,
                    length: textRange.length,
                    fontName: textRange.font.name,
                    fontSize: textRange.font.size,
                    fontColor: textRange.font.color,
                    bold: textRange.font.bold,
                    italic: textRange.font.italic,
                    underline: textRange.font.underline,
                    horizontalAlignment: textRange.paragraphFormat.horizontalAlignment,
                    indentLevel: textRange.paragraphFormat.indentLevel,
                  }
                : undefined,
            selectedShapes: selectedShapeDescriptors.slice(0, 4),
            selectedSlides: selectedSlideSummaries.slice(0, 6).map((slide) => ({
              index: slide.index,
              layoutName: slide.layoutName,
              layoutType: slide.layoutType,
              slideMasterName: slide.slideMasterName,
            })),
          }
        : undefined,
      snippets: {
        documentStructure: {
          slides: presentationSlides.items.length,
          selectedSlides: slides.items.length,
          selectedShapes: selectedShapes.length,
          slideMasters: masterSummaries.length,
          layouts: masterSummaries.reduce((total, slideMaster) => total + slideMaster.layoutCount, 0),
        },
        slideDeck: Array.from(slideMetadataById.values()).slice(0, 30),
        selectedSlides: selectedSlideSummaries.slice(0, 6),
        selectedShapeDescriptors: selectedShapeDescriptors.slice(0, 8),
        slideMasters: masterSummaries.slice(0, 12),
        layoutUsage: layoutUsage.slice(0, 24),
      },
      visuals,
    };
    result.summary = createSummary(result);
    return result;
  });

  if (!payload.visuals?.length && maxImages > 0) {
    const fallback = await getSelectedImageAsync().catch(() => undefined);
    if (fallback) {
      payload.visuals = [await optimizeVisual({ ...fallback, kind: "slide", label: "PowerPoint selection snapshot" })];
      payload.summary = createSummary(payload);
    }
  }

  if (options.includeFormatting && supportsRequirementSet("PowerPointApi", "1.10")) {
    try {
      const packageSummary = await inspectCurrentPowerPointPresentationPackage();
      const slideIdByNumber = new Map<number, string>();
      for (const slide of getRecordArray(payload.snippets?.slideDeck)) {
        const slideNumber = toNumber(slide.index);
        const slideId = trimString(slide.id);
        if (typeof slideNumber === "number" && slideId) {
          slideIdByNumber.set(slideNumber, slideId);
        }
      }
      const selectedSlidesForNotes = getRecordArray(payload.snippets?.selectedSlides);
      payload.state.capabilities = Array.from(
        new Set([
          ...payload.state.capabilities,
          "powerpoint.presentationPackage",
          "powerpoint.slideNotes",
          "powerpoint.notesRegionAnchors",
          "powerpoint.charts",
        ]),
      );
      const slideNotesAnchors = packageSummary.notes.slice(0, 20).map((note) => {
        return {
          kind: "notesRegion",
          id: note.partName ?? `notes:${note.slideNumber}`,
          label: `Notes for Slide ${note.slideNumber}`,
          slideId: slideIdByNumber.get(note.slideNumber),
          slideIndex: note.slideNumber,
          text: truncateLabel(note.preview ?? note.text ?? (note.hasNotes ? `Speaker notes for slide ${note.slideNumber}` : undefined), 180),
        } as OfficeAnchor;
      });
      payload.anchors = uniqueAnchors([...(payload.anchors ?? []), ...slideNotesAnchors]);
      payload.formatting = {
        ...(payload.formatting ?? {}),
        presentationTheme: packageSummary.theme,
        presentationPackage: {
          partCounts: packageSummary.partCounts,
          notesSlides: packageSummary.notes.length,
          charts: packageSummary.charts.length,
        },
      };
      payload.snippets = {
        ...(payload.snippets ?? {}),
        presentationTheme: packageSummary.theme,
        selectedSlideNotes: selectedSlidesForNotes.slice(0, 12).map((slide) => {
          const slideNumber = toNumber(slide.index) ?? toNumber(slide.slideIndex) ?? 0;
          const note = packageSummary.notes.find((entry) => entry.slideNumber === slideNumber);
          return {
            slideNumber,
            slideId: trimString(slide.id) ?? slideIdByNumber.get(slideNumber),
            hasNotes: note?.hasNotes ?? false,
            preview: note?.preview,
            partName: note?.partName,
          };
        }),
        slideNotes: packageSummary.notes.slice(0, 20).map((note) => ({
          slideNumber: note.slideNumber,
          hasNotes: note.hasNotes,
          preview: note.preview,
          partName: note.partName,
        })),
        presentationCharts: packageSummary.charts.slice(0, 20).map((chart) => ({
          slideNumber: chart.slideNumber,
          chartIndex: chart.chartIndex,
          chartType: chart.chartType,
          title: chart.title,
          shapeName: chart.shapeName,
          embeddedWorkbookPartName: chart.embeddedWorkbookPartName,
          seriesCount: chart.seriesCount,
          categoryCount: chart.categoryCount,
          hasEmbeddedWorkbook: chart.hasEmbeddedWorkbook,
        })),
        presentationPackage: packageSummary.partCounts,
      };
      payload.summary = createSummary(payload);
    } catch (error) {
      payload.formatting = {
        ...(payload.formatting ?? {}),
        presentationPackageWarning:
          error instanceof Error ? error.message : "PowerPoint presentation package inspection was not available.",
      };
    }
  }

  return payload;
}

