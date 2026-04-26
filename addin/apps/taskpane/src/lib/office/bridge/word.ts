import type { OfficeToolRequest, OfficeToolResult } from "@pi-office/pi-office-pack/protocol";
import type { OfficeToolExecutorDependencies } from "./common";
import {
  appendViewportCaptureSummary,
  toHostAction,
  toPayloadAwareResult,
  toWordDocumentVerificationPayload,
  toWordVisualVerificationPayload,
  trimString,
} from "./common";

export async function executeWordOfficeTool(
  request: OfficeToolRequest,
  dependencies: OfficeToolExecutorDependencies,
): Promise<OfficeToolResult | undefined> {
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


      if (request.toolName === "office_propose_edits") {
        const result = await dependencies.proposeEdits(request.host, request.params);
        return toPayloadAwareResult(request.requestId, result);
      }


  return undefined;
}
