import type { OfficeAnchor, OfficeHost, OfficeHostAction, OfficeToolRequest, OfficeToolResult } from "@pi-office/pi-office-pack/protocol";

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === "object" && !Array.isArray(value));
}

function getErrorRecord(error: unknown): Record<string, unknown> | undefined {
  return error && typeof error === "object" ? (error as Record<string, unknown>) : undefined;
}

function trimString(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function isSingleCellAddress(value: string | undefined): boolean {
  const normalized = trimString(value)?.replace(/\$/g, "");
  if (!normalized) {
    return false;
  }

  const localAddress = normalized.includes("!") ? normalized.split("!").pop() : normalized;
  if (!localAddress) {
    return false;
  }

  const [start, end] = localAddress.split(":");
  return !end || start === end;
}

function stringifyDetail(value: unknown): string | undefined {
  if (value == null) return undefined;
  if (typeof value === "string" && value.trim()) return value.trim();
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  if (Array.isArray(value)) {
    const parts = value.map((entry) => stringifyDetail(entry)).filter((entry): entry is string => Boolean(entry));
    return parts.length ? parts.join(" | ") : undefined;
  }

  try {
    return JSON.stringify(value);
  } catch {
    return undefined;
  }
}

function extractPayloadFailure(payload: unknown): string | undefined {
  if (!isRecord(payload)) {
    return undefined;
  }

  const hasErrorField = Object.prototype.hasOwnProperty.call(payload, "error");
  const payloadError = stringifyDetail(payload.error);

  if (payload.ok === false) {
    return payloadError ?? "Office tool returned ok=false.";
  }

  if (hasErrorField && payloadError) {
    return payloadError;
  }

  return undefined;
}

function toPayloadAwareResult(requestId: string, payload: unknown): OfficeToolResult {
  const payloadFailure = extractPayloadFailure(payload);
  if (payloadFailure) {
    return {
      requestId,
      success: false,
      error: payloadFailure,
    };
  }

  return {
    requestId,
    success: true,
    content: payload,
  };
}

function appendViewportCaptureSummary(payload: unknown, includeWindowFrameRequested: boolean): unknown {
  if (!isRecord(payload)) {
    return payload;
  }

  const formatting = isRecord(payload.formatting) ? payload.formatting : {};
  const viewportCapture = {
    mode: "officejs-context",
    includeWindowFrameRequested,
    includeWindowFrameCaptured: false,
    note: includeWindowFrameRequested
      ? "Full window-frame capture is unavailable in browser-only runtime."
      : "Viewport metadata captured from Office.js context.",
  };

  const summaryParts = [
    trimString(payload.summary),
    "Visible Word viewport metadata captured.",
    includeWindowFrameRequested ? "Full window-frame capture is unavailable in browser-only runtime." : undefined,
  ].filter((entry): entry is string => Boolean(entry));

  return {
    ...payload,
    summary: summaryParts.join("\n"),
    formatting: {
      ...formatting,
      viewportCapture,
    },
  };
}

function toWordDocumentVerificationPayload(payload: unknown, scope: string | undefined): unknown {
  if (!isRecord(payload)) {
    return payload;
  }

  return {
    summary: trimString(payload.summary) ?? "Word verification context captured.",
    details: {
      kind: "word-document-verification",
      mutating: false,
      host: "word",
      scope: scope ?? "document",
      context: {
        state: payload.state,
        anchors: payload.anchors,
        snippets: payload.snippets,
        formatting: payload.formatting,
      },
    },
  };
}

function toWordVisualVerificationPayload(payload: unknown, includeWindowFrameRequested: boolean): unknown {
  if (!isRecord(payload)) {
    return payload;
  }

  const visuals = Array.isArray(payload.visuals) ? payload.visuals : [];
  const formatting = isRecord(payload.formatting) ? payload.formatting : {};
  const viewportCapture = isRecord(formatting.viewportCapture) ? formatting.viewportCapture : {};
  const state = isRecord(payload.state) ? payload.state : {};
  const captureMode = trimString(viewportCapture.mode) ?? "officejs-context";
  const includeWindowFrameCaptured = viewportCapture.includeWindowFrameCaptured === true;
  const note =
    trimString(viewportCapture.note) ??
    (includeWindowFrameRequested
      ? "Full window-frame capture is unavailable in browser-only runtime."
      : "Viewport metadata captured from Office.js context.");

  return {
    summary: trimString(payload.summary) ?? "Visible Word viewport metadata captured.",
    visual: {
      kind: "word-viewport",
      captureMode,
      imageCount: visuals.length,
      includeWindowFrameRequested,
      includeWindowFrameCaptured,
      note,
    },
    details: {
      kind: "word-visual-verification",
      mutating: false,
      host: "word",
      viewport: formatting.viewport,
      viewportCapture: {
        mode: captureMode,
        includeWindowFrameRequested,
        includeWindowFrameCaptured,
        note,
      },
      selection: state.selection,
    },
    visuals,
  };
}

function toPowerPointSlidesVerificationPayload(payload: unknown, scope: string | undefined): unknown {
  if (!isRecord(payload)) {
    return payload;
  }

  return {
    summary: trimString(payload.summary) ?? "PowerPoint slide verification context captured.",
    details: {
      kind: "powerpoint-slide-verification",
      mutating: false,
      host: "powerpoint",
      scope: scope ?? "presentation",
      structure: {
        slideCount: payload.slideCount,
        selectedSlideCount: payload.selectedSlideCount,
        selectedSlideIds: payload.selectedSlideIds,
        slides: payload.slides,
        slidePreviews: payload.slidePreviews,
        slideMasters: payload.slideMasters,
      },
    },
  };
}

function toPowerPointVisualVerificationPayload(payload: unknown, maxImages: number): unknown {
  if (!isRecord(payload)) {
    return payload;
  }

  const visuals = Array.isArray(payload.visuals) ? payload.visuals : [];
  const snippets = isRecord(payload.snippets) ? payload.snippets : {};
  const formatting = isRecord(payload.formatting) ? payload.formatting : {};

  return {
    summary: trimString(payload.summary) ?? "PowerPoint visual verification context captured.",
    visual: {
      kind: "powerpoint-slide-snapshot",
      captureMode: "officejs-slide-snapshot",
      imageCount: visuals.length,
      maxImagesRequested: maxImages,
      note: "Visual verification uses supported Office.js slide/shape snapshot paths.",
    },
    details: {
      kind: "powerpoint-slide-visual-verification",
      mutating: false,
      host: "powerpoint",
      selectedSlides: snippets.selectedSlides,
      selectedShapeDescriptors: snippets.selectedShapeDescriptors,
      formatting: {
        selectedSlides: formatting.selectedSlides,
        selectedShapes: formatting.selectedShapes,
      },
    },
    visuals,
  };
}

export function summarizeOfficeToolError(error: unknown, toolName: OfficeToolRequest["toolName"]): string {
  const record = getErrorRecord(error);
  const message =
    trimString(record?.message) ||
    (error instanceof Error ? trimString(error.message) : undefined) ||
    trimString(error) ||
    "Unknown Office host error.";
  const code = trimString(record?.code);
  const debugInfo = getErrorRecord(record?.debugInfo);
  const errorLocation = trimString(debugInfo?.errorLocation);
  const statement = trimString(debugInfo?.statement);
  const traceMessages = stringifyDetail(record?.traceMessages);

  const details = [
    code ? `code=${code}` : undefined,
    errorLocation ? `location=${errorLocation}` : undefined,
    statement ? `statement=${statement}` : undefined,
    traceMessages ? `trace=${traceMessages}` : undefined,
  ].filter((entry): entry is string => Boolean(entry));

  return details.length ? `${toolName} failed: ${message} (${details.join(", ")})` : `${toolName} failed: ${message}`;
}

export function serializeOfficeToolError(error: unknown): Record<string, unknown> {
  const record = getErrorRecord(error);
  const debugInfo = getErrorRecord(record?.debugInfo);

  return {
    name: trimString(record?.name) || (error instanceof Error ? error.name : undefined),
    message:
      trimString(record?.message) ||
      (error instanceof Error ? trimString(error.message) : undefined) ||
      trimString(error) ||
      "Unknown Office host error.",
    code: trimString(record?.code),
    stack: trimString(record?.stack) || (error instanceof Error ? trimString(error.stack) : undefined),
    traceMessages: record?.traceMessages,
    debugInfo,
    details: getErrorRecord(record?.details) ?? getErrorRecord((error as { details?: unknown } | undefined)?.details),
    error,
  };
}

export function toAnchor(params: Record<string, unknown>): OfficeAnchor {
  const anchorRecord = isRecord(params.anchor) ? params.anchor : params;
  const explicitKind = trimString(anchorRecord.kind);
  const layoutId = trimString(anchorRecord.layoutId);
  const slideMasterId = trimString(anchorRecord.slideMasterId);
  const anchorId = trimString(anchorRecord.id);
  const isNotesAnchorId =
    typeof anchorId === "string" &&
    (anchorId.startsWith("notes:") || anchorId.startsWith("notesRegion:") || anchorId.startsWith("slideNotes:"));
  const inferredKind =
    layoutId
      ? "layout"
      : slideMasterId
        ? "slideMaster"
        : isNotesAnchorId
          ? "notesRegion"
        : typeof anchorRecord.slideIndex === "number"
          ? "slide"
            : trimString(anchorRecord.slideId)
              ? "slide"
            : trimString(anchorRecord.shapeId)
              ? "shape"
              : trimString(anchorRecord.address)
                ? isSingleCellAddress(trimString(anchorRecord.address)) ? "cell" : "range"
                : trimString(anchorRecord.chartName)
                  ? "chart"
                  : trimString(anchorRecord.pivotTableName)
                    ? "pivotTable"
                    : trimString(anchorRecord.tableName)
                      ? "table"
                      : trimString(anchorRecord.sheetName)
                        ? "sheet"
                        : trimString(anchorRecord.paragraphId)
                          ? "paragraph"
                          : trimString(anchorRecord.commentId)
                            ? "comment"
                            : trimString(anchorRecord.revisionId)
                              ? "revision"
                              : anchorId?.startsWith("contentControl:")
                                ? "contentControl"
                                : anchorId?.startsWith("field:")
                                  ? "field"
                                  : "selection";

  return {
    kind: (explicitKind ?? inferredKind) as OfficeAnchor["kind"],
    label: trimString(anchorRecord.label) ?? trimString(anchorRecord.layoutName) ?? trimString(anchorRecord.slideMasterName),
    id: anchorId ?? layoutId ?? slideMasterId,
    text: trimString(anchorRecord.text) ?? trimString(params.target),
    sheetName: trimString(anchorRecord.sheetName),
    address: trimString(anchorRecord.address),
    paragraphId: trimString(anchorRecord.paragraphId),
    commentId: trimString(anchorRecord.commentId),
    revisionId: trimString(anchorRecord.revisionId),
    slideId: trimString(anchorRecord.slideId),
    slideIndex: typeof anchorRecord.slideIndex === "number" ? anchorRecord.slideIndex : undefined,
    shapeId: trimString(anchorRecord.shapeId),
    tableName: trimString(anchorRecord.tableName),
    chartName: trimString(anchorRecord.chartName),
    pivotTableName: trimString(anchorRecord.pivotTableName),
    namedItemName: trimString(anchorRecord.namedItemName),
  };
}

function parseMatrixContent(content: string): unknown {
  try {
    return JSON.parse(content);
  } catch {
    return undefined;
  }
}

function firstString(...values: unknown[]): string | undefined {
  for (const value of values) {
    if (typeof value === "string") {
      return value;
    }
  }
  return undefined;
}

function normalizeActionType(value: string | undefined): string | undefined {
  if (!value) {
    return undefined;
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  if (compact === "replaceselection" || compact === "insertafterselection" || compact === "insertbeforeselection") {
    return undefined;
  }
  if (compact === "inserttext") {
    return "insertText";
  }
  if (compact === "inserthtml") {
    return "insertHtml";
  }
  if (compact === "setrangevalues") {
    return "setRangeValues";
  }

  return value;
}

function normalizeMode(value: string | undefined): string {
  if (!value) {
    return "replaceSelection";
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  if (compact === "insertafterselection" || compact === "insertafter") {
    return "insertAfterSelection";
  }
  if (compact === "insertbeforeselection" || compact === "insertbefore") {
    return "insertBeforeSelection";
  }
  if (compact === "setrangevalues") {
    return "setRangeValues";
  }
  if (compact === "replaceselection" || compact === "replace") {
    return "replaceSelection";
  }

  return value;
}

function normalizeFormat(value: string | undefined): string | undefined {
  if (!value) {
    return undefined;
  }

  const normalized = value.trim().toLowerCase();
  if (!normalized) {
    return undefined;
  }
  if (normalized === "htm") {
    return "html";
  }
  if (normalized === "values" || normalized === "table") {
    return "matrix";
  }

  return normalized;
}

function toStringArray(value: unknown): string[] {
  if (Array.isArray(value)) {
    return value
      .map((entry) => trimString(entry))
      .filter((entry): entry is string => Boolean(entry));
  }

  const single = trimString(value);
  return single ? [single] : [];
}

function toPowerPointStructureActionType(value: string | undefined): string | undefined {
  if (!value) {
    return undefined;
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  switch (compact) {
    case "addslide":
      return "addSlide";
    case "moveslide":
      return "moveSlide";
    case "reorderslides":
    case "reorderstoryline":
      return "reorderSlides";
    case "deleteslide":
      return "deleteSlide";
    case "deleteslides":
      return "deleteSlides";
    case "applylayout":
      return "applyLayout";
    case "selectslides":
      return "selectSlides";
    case "addagendaslide":
      return "addAgendaSlide";
    case "addtransitionslide":
      return "addTransitionSlide";
    case "combineslides":
      return "combineSlides";
    case "importslidesfrombase64":
    case "mergepresentationfrombase64":
      return "importSlidesFromBase64";
    default:
      return undefined;
  }
}

function toPowerPointElementInsertActionType(value: string | undefined): string | undefined {
  if (!value) {
    return undefined;
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  switch (compact) {
    case "addtextbox":
    case "inserttextbox":
      return "addTextBox";
    case "addgeometricshape":
    case "insertgeometricshape":
    case "addshape":
      return "addGeometricShape";
    case "addtable":
    case "inserttable":
      return "addTable";
    case "addline":
    case "insertline":
      return "addLine";
    case "addprocessflow":
      return "addProcessFlow";
    case "addsimplediagram":
      return "addSimpleDiagram";
    case "insertinlinepicture":
    case "insertimage":
    case "addimage":
      return "insertInlinePicture";
    default:
      return undefined;
  }
}

function toPowerPointElementRemoveActionType(value: string | undefined): string | undefined {
  if (!value) {
    return "deleteShape";
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  switch (compact) {
    case "removeshape":
    case "deleteshape":
      return "deleteShape";
    case "removeshapes":
    case "deleteshapes":
      return "deleteShapes";
    case "clearshapetext":
      return "clearShapeText";
    default:
      return undefined;
  }
}

function toPowerPointTextAction(
  value: string | undefined,
  params: Record<string, unknown>,
): { type: string; placement?: string | undefined } | undefined {
  if (!value) {
    return trimString(params.shapeId) ? { type: "setShapeText" } : { type: "insertText" };
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  switch (compact) {
    case "setshapetext":
    case "editshapetext":
    case "replaceshapetext":
      return { type: "setShapeText" };
    case "appendshapetext":
    case "appendtext":
      return { type: "setShapeText", placement: "after" };
    case "clearshapetext":
      return { type: "clearShapeText" };
    case "inserttext":
    case "editselectedtext":
    case "setselectedtext":
      return { type: "insertText" };
    default:
      return undefined;
  }
}

function toPowerPointXmlActionType(value: string | undefined): string | undefined {
  if (!value) {
    return undefined;
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  switch (compact) {
    case "inspectpresentationpackage":
    case "readpresentationpackage":
      return "inspectPresentationPackage";
    case "getpresentationtheme":
    case "inspecttheme":
      return "getPresentationTheme";
    case "getslidenotes":
    case "readslidenotes":
    case "inspectslidenotes":
      return "getSlideNotes";
    case "setslidenotes":
    case "editslidenotes":
      return "setSlideNotes";
    case "replaceslidenotes":
      return "replaceSlideNotes";
    case "importslidesfrombase64":
      return "importSlidesFromBase64";
    case "mergepresentationfrombase64":
      return "mergePresentationFromBase64";
    case "exportslidesasbase64":
      return "exportSlidesAsBase64";
    default:
      return undefined;
  }
}

function toPowerPointMasterActionType(value: string | undefined): string | undefined {
  if (!value) {
    return "applyLayout";
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  switch (compact) {
    case "applylayout":
    case "setlayout":
    case "editslidelayout":
    case "setslidemaster":
    case "applymaster":
      return "applyLayout";
    default:
      return undefined;
  }
}

function toPowerPointChartActionType(value: string | undefined): string | undefined {
  if (!value) {
    return "updateSlideChart";
  }

  const compact = value.replace(/[\s_-]/g, "").toLowerCase();
  switch (compact) {
    case "getslidecharts":
    case "inspectslidecharts":
    case "readslidecharts":
    case "inspectcharts":
    case "readcharts":
    case "listcharts":
      return "getSlideCharts";
    case "addslidechart":
    case "createslidechart":
    case "addcharttoslide":
    case "insertslidechart":
    case "createchart":
    case "addchart":
      return "addSlideChart";
    case "updateslidechart":
    case "setchartdata":
    case "updatechartdata":
    case "replacechartdata":
    case "editslidechart":
    case "editchart":
      return "updateSlideChart";
    default:
      return undefined;
  }
}

function toPowerPointStructureOptions(params: Record<string, unknown>): Record<string, unknown> {
  const next: Record<string, unknown> = {
    ...(isRecord(params.options) ? params.options : {}),
  };

  for (const [key, value] of Object.entries(params)) {
    if (
      key === "options" ||
      key === "operation" ||
      key === "mode" ||
      key === "type" ||
      key === "anchor" ||
      key === "target" ||
      key === "content"
    ) {
      continue;
    }
    next[key] = value;
  }

  return next;
}

function toPowerPointStructureTarget(params: Record<string, unknown>): OfficeAnchor | undefined {
  const hasSlideTarget = Boolean(trimString(params.slideId)) || typeof params.slideIndex === "number";
  const hasShapeTarget = Boolean(trimString(params.shapeId));
  const hasLayoutTarget = Boolean(trimString(params.layoutId) || trimString(params.layoutName));
  const hasMasterTarget = Boolean(trimString(params.slideMasterId) || trimString(params.slideMasterName));
  const hasExplicitAnchor = isRecord(params.anchor) && Object.keys(params.anchor).length > 0;

  if (!hasSlideTarget && !hasShapeTarget && !hasLayoutTarget && !hasMasterTarget && !hasExplicitAnchor) {
    return undefined;
  }

  return toAnchor(params);
}

export function toHostAction(params: Record<string, unknown>): OfficeHostAction {
  if (isRecord(params.action) && typeof params.action.type === "string") {
    const action = { ...(params.action as Record<string, unknown>) } as OfficeHostAction;
    if (isRecord(action.target)) {
      action.target = toAnchor({ anchor: action.target });
    }
    return action;
  }

  const requestedType = normalizeActionType(trimString(params.operation) ?? trimString(params.type));
  const mode = normalizeMode(trimString(params.mode));
  const inferredFormatFromType =
    requestedType === "insertHtml" ? "html" : requestedType === "setRangeValues" ? "matrix" : undefined;
  const format = normalizeFormat(trimString(params.format)) ?? inferredFormatFromType;
  const content = firstString(
    params.content,
    format === "html" ? params.html : undefined,
    params.text,
    params.html,
  ) ?? "";
  const target = toAnchor(params);
  const directValues = Array.isArray(params.values) ? (params.values as OfficeHostAction["values"]) : undefined;

  if (requestedType === "setRangeValues" || mode === "setRangeValues" || format === "matrix") {
    return {
      type: "setRangeValues",
      target,
      values: directValues ?? (parseMatrixContent(content) as OfficeHostAction["values"]),
      content,
      format,
    };
  }

  const placement = trimString(params.placement) ?? (mode === "insertAfterSelection" ? "after" : mode === "insertBeforeSelection" ? "before" : "replace");

  return {
    type: requestedType ?? (format === "html" ? "insertHtml" : "insertText"),
    target,
    content,
    format,
    placement,
  };
}

export interface OfficeToolExecutorDependencies {
  collectOfficeContext: (
    host: OfficeHost,
    options: { includeFormatting?: boolean; maxImages?: number; scope?: string },
  ) => Promise<unknown>;
  applyHostAction: (host: OfficeHost, action: OfficeHostAction) => Promise<unknown>;
  navigateOfficeAnchor: (host: OfficeHost, anchor: OfficeAnchor) => Promise<unknown>;
  readDocumentSection: (host: OfficeHost, startIndex: number, endIndex: number, includeStyles: boolean) => Promise<unknown>;
  executeOfficeJs: (host: OfficeHost, code: string) => Promise<unknown>;
  proposeEdits: (host: OfficeHost, params: Record<string, unknown>) => Promise<unknown>;
  logger?: Pick<Console, "error"> | undefined;
}

export function createOfficeToolExecutor(dependencies: OfficeToolExecutorDependencies) {
  return async function executeOfficeTool(request: OfficeToolRequest): Promise<OfficeToolResult> {
    try {
      if (request.toolName === "office_get_context") {
        const scope = trimString(request.params.scope);
        const state = await dependencies.collectOfficeContext(request.host, {
          includeFormatting: true,
          maxImages: 0,
          ...(scope ? { scope } : {}),
        });
        return {
          requestId: request.requestId,
          success: true,
          content: state,
        };
      }

      if (request.toolName === "office_apply_edit") {
        const result = await dependencies.applyHostAction(request.host, toHostAction(request.params));
        return {
          requestId: request.requestId,
          success: true,
          content: result,
        };
      }

      if (request.toolName === "edit_doc_text") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "edit_doc_text is only available for Word.",
          };
        }

        const result = await dependencies.applyHostAction(request.host, toHostAction(request.params));
        return {
          requestId: request.requestId,
          success: true,
          content: result,
        };
      }

      if (request.toolName === "edit_doc_list") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "edit_doc_list is only available for Word.",
          };
        }

        const result = await dependencies.proposeEdits(request.host, request.params);
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "office_navigate") {
        const result = await dependencies.navigateOfficeAnchor(request.host, toAnchor(request.params));
        return {
          requestId: request.requestId,
          success: true,
          content: result,
        };
      }

      if (request.toolName === "office_capture_snapshot") {
        const includeFormatting = request.params.includeFormatting !== false;
        const scope = trimString(request.params.scope);
        const maxImages =
          typeof request.params.maxImages === "number" && Number.isFinite(request.params.maxImages)
            ? Math.max(0, Math.min(4, Math.trunc(request.params.maxImages)))
            : 2;
        const result = await dependencies.collectOfficeContext(request.host, {
          includeFormatting,
          maxImages,
          ...(scope ? { scope } : {}),
        });
        return {
          requestId: request.requestId,
          success: true,
          content: result,
        };
      }

      if (request.toolName === "office_capture_viewport") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "office_capture_viewport is only available for Word.",
          };
        }
        const includeFormatting = request.params.includeFormatting !== false;
        const includeWindowFrameRequested = request.params.includeWindowFrame === true;
        // FUTURE SCOPE ONLY (owner-gated):
        // Full Word window-frame capture depended on the removed companion app.
        // Browser runtime can only access Office.js viewport metadata and selection visuals.
        // Do not re-enable companion-based window capture without explicit approval from Manan.
        // const maxImages = includeWindowFrameRequested ? 2 : 1;
        const result = await dependencies.collectOfficeContext(request.host, {
          includeFormatting,
          maxImages: 1,
          scope: "viewport",
        });
        return {
          requestId: request.requestId,
          success: true,
          content: appendViewportCaptureSummary(result, includeWindowFrameRequested),
        };
      }

      if (request.toolName === "office_read_section") {
        const startIndexValue = request.params.startIndex ?? request.params.start;
        const endIndexValue = request.params.endIndex ?? request.params.end;
        const startIndex = typeof startIndexValue === "number" ? Math.max(0, Math.trunc(startIndexValue)) : 0;
        const endIndex = typeof endIndexValue === "number" ? Math.trunc(endIndexValue) : startIndex + 20;
        const includeStyles = request.params.includeStyles !== false;
        const result = await dependencies.readDocumentSection(request.host, startIndex, endIndex, includeStyles);
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "verify_doc") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "verify_doc is only available for Word.",
          };
        }

        const scope = trimString(request.params.scope);
        const includeFormatting = request.params.includeFormatting !== false;
        const rawResult = await dependencies.collectOfficeContext(request.host, {
          includeFormatting,
          maxImages: 0,
          ...(scope ? { scope } : {}),
        });
        const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
        if (!payloadAware.success) {
          return payloadAware;
        }

        return {
          requestId: request.requestId,
          success: true,
          content: toWordDocumentVerificationPayload(payloadAware.content, scope),
        };
      }

      if (request.toolName === "verify_doc_visual") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "verify_doc_visual is only available for Word.",
          };
        }

        const includeFormatting = request.params.includeFormatting !== false;
        const includeWindowFrameRequested = request.params.includeWindowFrame === true;
        const rawResult = await dependencies.collectOfficeContext(request.host, {
          includeFormatting,
          maxImages: 1,
          scope: "viewport",
        });
        const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
        if (!payloadAware.success) {
          return payloadAware;
        }

        const viewportPayload = appendViewportCaptureSummary(payloadAware.content, includeWindowFrameRequested);
        return {
          requestId: request.requestId,
          success: true,
          content: toWordVisualVerificationPayload(viewportPayload, includeWindowFrameRequested),
        };
      }

      if (request.toolName === "get_presentation_structure") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "get_presentation_structure is only available for PowerPoint.",
          };
        }

        const result = await dependencies.applyHostAction(request.host, {
          type: "getPresentationStructure",
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "get_slide") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "get_slide is only available for PowerPoint.",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: "getSlide",
          ...(target ? { target } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "list_slide_shapes") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "list_slide_shapes is only available for PowerPoint.",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: "listSlideShapes",
          ...(target ? { target } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "modify_presentation_structure") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "modify_presentation_structure is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toPowerPointStructureActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "modify_presentation_structure requires a supported operation (add_slide, move_slide, reorder_slides, delete_slide, apply_layout, select_slides, add_agenda_slide, add_transition_slide, combine_slides, import_slides_from_base64).",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          ...(typeof request.params.content === "string" ? { content: request.params.content } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "duplicate_slide") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "duplicate_slide is only available for PowerPoint.",
          };
        }

        const requestedSlideIds = toStringArray(request.params.slideIds);
        const fallbackSlideId = trimString(request.params.slideId);
        const slideIds = requestedSlideIds.length ? requestedSlideIds : fallbackSlideId ? [fallbackSlideId] : [];
        const targetParams =
          slideIds.length && !fallbackSlideId && typeof request.params.slideIndex !== "number"
            ? { ...request.params, slideId: slideIds[0] }
            : request.params;
        const target = toPowerPointStructureTarget(targetParams);
        const result = await dependencies.applyHostAction(request.host, {
          type: "duplicateSlide",
          ...(target ? { target } : {}),
          options: {
            ...toPowerPointStructureOptions(request.params),
            ...(slideIds.length ? { slideIds } : {}),
          },
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "insert_slide_element") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "insert_slide_element is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toPowerPointElementInsertActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "insert_slide_element requires a supported operation (add_text_box, add_geometric_shape, add_table, add_line, add_process_flow, add_simple_diagram, insert_inline_picture).",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const content = firstString(request.params.content, request.params.text, request.params.base64);
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          ...(typeof content === "string" ? { content } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "remove_slide_element") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "remove_slide_element is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toPowerPointElementRemoveActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "remove_slide_element requires a supported operation (remove_shape, remove_shapes, clear_shape_text).",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "edit_slide_text") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "edit_slide_text is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const textAction = toPowerPointTextAction(operation, request.params);
        if (!textAction) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "edit_slide_text requires a supported operation (set_shape_text, append_shape_text, clear_shape_text, insert_text).",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const content = firstString(request.params.content, request.params.text, request.params.newText) ?? "";
        const actionOptions = toPowerPointStructureOptions(request.params);
        const placement = trimString(request.params.placement) ?? textAction.placement;
        if (placement && !("placement" in actionOptions)) {
          actionOptions.placement = placement;
        }
        const result = await dependencies.applyHostAction(request.host, {
          type: textAction.type,
          ...(target ? { target } : {}),
          ...(textAction.type !== "clearShapeText" ? { content } : {}),
          ...(placement ? { placement } : {}),
          options: actionOptions,
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "edit_slide_xml") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "edit_slide_xml is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toPowerPointXmlActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "edit_slide_xml requires a supported operation (inspect_presentation_package, get_presentation_theme, get_slide_notes, set_slide_notes, replace_slide_notes, import_slides_from_base64, merge_presentation_from_base64, export_slides_as_base64).",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const content = firstString(request.params.content, request.params.text, request.params.base64);
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          ...(typeof content === "string" ? { content } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "edit_slide_master") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "edit_slide_master is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toPowerPointMasterActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error: "edit_slide_master currently supports apply_layout operations.",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "edit_slide_chart") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "edit_slide_chart is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toPowerPointChartActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "edit_slide_chart requires a supported operation (get_slide_charts, add_slide_chart, update_slide_chart).",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const content = firstString(request.params.content, request.params.title);
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          ...(typeof content === "string" ? { content } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "copy_image_between_slides") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "copy_image_between_slides is only available for PowerPoint.",
          };
        }

        const targetParams: Record<string, unknown> = {
          ...request.params,
          slideId: trimString(request.params.targetSlideId) ?? trimString(request.params.slideId),
          slideIndex:
            typeof request.params.targetSlideIndex === "number"
              ? request.params.targetSlideIndex
              : request.params.slideIndex,
          shapeId: trimString(request.params.targetShapeId) ?? trimString(request.params.shapeId),
        };
        const target = toPowerPointStructureTarget(targetParams);
        const content = firstString(request.params.sourceImageBase64, request.params.base64, request.params.content);
        const result = await dependencies.applyHostAction(request.host, {
          type: "copyImageBetweenSlides",
          ...(target ? { target } : {}),
          ...(typeof content === "string" ? { content } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "search_icons") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "search_icons is only available for PowerPoint.",
          };
        }

        const query = firstString(request.params.query, request.params.search, request.params.content);
        if (!query?.trim()) {
          return {
            requestId: request.requestId,
            success: false,
            error: "search_icons requires a non-empty query.",
          };
        }

        const result = await dependencies.applyHostAction(request.host, {
          type: "searchIcons",
          content: query,
          options: {
            ...toPowerPointStructureOptions(request.params),
            query,
          },
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "insert_icon") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "insert_icon is only available for PowerPoint.",
          };
        }

        const iconId = firstString(request.params.iconId, request.params.iconName, request.params.query, request.params.content);
        if (!iconId?.trim()) {
          return {
            requestId: request.requestId,
            success: false,
            error: "insert_icon requires iconId, iconName, query, or content.",
          };
        }

        const targetParams: Record<string, unknown> = {
          ...request.params,
          slideId: trimString(request.params.targetSlideId) ?? trimString(request.params.slideId),
          slideIndex:
            typeof request.params.targetSlideIndex === "number"
              ? request.params.targetSlideIndex
              : request.params.slideIndex,
          shapeId: trimString(request.params.targetShapeId) ?? trimString(request.params.shapeId),
        };
        const target = toPowerPointStructureTarget(targetParams);
        const result = await dependencies.applyHostAction(request.host, {
          type: "insertIcon",
          ...(target ? { target } : {}),
          content: iconId,
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "verify_slides") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "verify_slides is only available for PowerPoint.",
          };
        }

        const scope = trimString(request.params.scope);
        const result = await dependencies.applyHostAction(request.host, {
          type: "getPresentationStructure",
          options: toPowerPointStructureOptions(request.params),
        });
        const payloadAware = toPayloadAwareResult(request.requestId, result);
        if (!payloadAware.success) {
          return payloadAware;
        }

        return {
          requestId: request.requestId,
          success: true,
          content: toPowerPointSlidesVerificationPayload(payloadAware.content, scope),
        };
      }

      if (request.toolName === "verify_slide_visual") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "verify_slide_visual is only available for PowerPoint.",
          };
        }

        const includeFormatting = request.params.includeFormatting !== false;
        const maxImages =
          typeof request.params.maxImages === "number" && Number.isFinite(request.params.maxImages)
            ? Math.max(1, Math.min(4, Math.trunc(request.params.maxImages)))
            : 1;
        const scope = trimString(request.params.scope) ?? "slide";
        const rawResult = await dependencies.collectOfficeContext(request.host, {
          includeFormatting,
          maxImages,
          scope,
        });
        const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
        if (!payloadAware.success) {
          return payloadAware;
        }

        return {
          requestId: request.requestId,
          success: true,
          content: toPowerPointVisualVerificationPayload(payloadAware.content, maxImages),
        };
      }

      if (request.toolName === "office_execute_js") {
        const code = firstString(request.params.code, request.params.script) ?? "";
        if (!code.trim()) {
          return { requestId: request.requestId, success: false, error: "No code provided." };
        }
        const result = await dependencies.executeOfficeJs(request.host, code);
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "office_propose_edits") {
        const result = await dependencies.proposeEdits(request.host, request.params);
        return toPayloadAwareResult(request.requestId, result);
      }

      return {
        requestId: request.requestId,
        success: false,
        error: `${request.toolName} is not supported by the taskpane bridge.`,
      };
    } catch (error) {
      dependencies.logger?.error?.(`[office-tool] ${request.toolName} failed`, {
        requestId: request.requestId,
        host: request.host,
        params: request.params,
        error: serializeOfficeToolError(error),
      });

      return {
        requestId: request.requestId,
        success: false,
        error: summarizeOfficeToolError(error, request.toolName),
      };
    }
  };
}
