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

export function toHostAction(params: Record<string, unknown>): OfficeHostAction {
  if (isRecord(params.action) && typeof params.action.type === "string") {
    const action = { ...(params.action as Record<string, unknown>) } as OfficeHostAction;
    if (isRecord(action.target)) {
      action.target = toAnchor({ anchor: action.target });
    }
    return action;
  }

  const mode = trimString(params.mode) ?? "replaceSelection";
  const content = typeof params.content === "string" ? params.content : "";
  const format = trimString(params.format);
  const target = toAnchor(params);

  if (mode === "setRangeValues" || format === "matrix") {
    return {
      type: "setRangeValues",
      target,
      values: parseMatrixContent(content) as OfficeHostAction["values"],
      content,
      format,
    };
  }

  return {
    type: format === "html" ? "insertHtml" : "insertText",
    target,
    content,
    format,
    placement: mode === "insertAfterSelection" ? "after" : "replace",
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

      if (request.toolName === "office_read_section") {
        const startIndex = typeof request.params.startIndex === "number" ? Math.max(0, Math.trunc(request.params.startIndex)) : 0;
        const endIndex = typeof request.params.endIndex === "number" ? Math.trunc(request.params.endIndex) : startIndex + 20;
        const includeStyles = request.params.includeStyles !== false;
        const result = await dependencies.readDocumentSection(request.host, startIndex, endIndex, includeStyles);
        return {
          requestId: request.requestId,
          success: true,
          content: result,
        };
      }

      if (request.toolName === "office_execute_js") {
        const code = typeof request.params.code === "string" ? request.params.code : "";
        if (!code.trim()) {
          return { requestId: request.requestId, success: false, error: "No code provided." };
        }
        const result = await dependencies.executeOfficeJs(request.host, code);
        return {
          requestId: request.requestId,
          success: true,
          content: result,
        };
      }

      if (request.toolName === "office_propose_edits") {
        const result = await dependencies.proposeEdits(request.host, request.params);
        return {
          requestId: request.requestId,
          success: true,
          content: result,
        };
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
