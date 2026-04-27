import type { OfficeToolRequest, OfficeToolResult } from "@pi-office/pi-office-pack/protocol";
import type { OfficeToolExecutorDependencies } from "./common";
import {
  toAnchor,
  toHostAction,
  toPayloadAwareResult,
  trimString,
} from "./common";

export async function executeCommonOfficeTool(
  request: OfficeToolRequest,
  dependencies: OfficeToolExecutorDependencies,
): Promise<OfficeToolResult | undefined> {
  if (request.toolName === "office_tool_search") {
    return undefined;
  }
  if (request.toolName === "office_tool_get") {
    return undefined;
  }
  if (request.toolName === "mcp_tool_search") {
    return undefined;
  }

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

  if (request.toolName === "office_capture_viewport") {
    return {
      requestId: request.requestId,
      success: false,
      error:
        "office_capture_viewport is a compatibility tool for true viewport/window screenshots and requires companion native capture. " +
        "Use office_capture_snapshot for Office.js context snapshots, verify_doc_visual for Word metadata/selection visuals, read_range_image for Excel active-selection snapshots, or verify_slide_visual for PowerPoint native slide/shape snapshots.",
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

  if (request.toolName === "office_execute_js") {
    const code = trimString(request.params.code) ?? trimString(request.params.script);
    if (!code) {
      throw new Error("office_execute_js requires code.");
    }
    const result = await dependencies.executeOfficeJs(request.host, code);
    return toPayloadAwareResult(request.requestId, result);
  }

  return undefined;
}
