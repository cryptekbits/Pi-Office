import type { OfficeToolRequest, OfficeToolResult } from "@pi-office/pi-office-pack/protocol";
import {
  serializeOfficeToolError,
  summarizeOfficeToolError,
  toAnchor,
  toHostAction,
  type OfficeToolExecutorDependencies,
} from "./office/bridge/common";
import { executeCommonOfficeTool } from "./office/bridge/common-executor";
import { executeExcelOfficeTool } from "./office/bridge/excel";
import { executePowerPointOfficeTool } from "./office/bridge/powerpoint";
import {
  getOfficeToolCapabilityDetail,
  searchOfficeToolDefinitions,
  toToolCapabilitySearchResult,
} from "./office/tools/index";
import { executeWordOfficeTool } from "./office/bridge/word";

export {
  serializeOfficeToolError,
  summarizeOfficeToolError,
  toAnchor,
  toHostAction,
};
export type { OfficeToolExecutorDependencies };

function searchOfficeToolsForBridge(request: OfficeToolRequest): OfficeToolResult | undefined {
  if (request.toolName === "office_tool_search") {
    const query = typeof request.params.query === "string" ? request.params.query : undefined;
    const host = request.params.host === "word" || request.params.host === "excel" || request.params.host === "powerpoint"
      ? request.params.host
      : request.host;
    const category = typeof request.params.category === "string" ? request.params.category : undefined;
    const limit =
      typeof request.params.maxResults === "number"
        ? request.params.maxResults
        : typeof request.params.limit === "number"
          ? request.params.limit
          : undefined;
    return {
      requestId: request.requestId,
      success: true,
      content: {
        host,
        query,
        results: searchOfficeToolDefinitions({
          host,
          query,
          category: category as never,
          limit,
          includeSchemas: request.params.includeSchemas === true,
        }),
      },
    };
  }

  if (request.toolName === "office_tool_get") {
    const toolName = typeof request.params.toolName === "string" ? request.params.toolName : undefined;
    const id = typeof request.params.id === "string" ? request.params.id : undefined;
    const toolNames = Array.isArray(request.params.toolNames)
      ? request.params.toolNames.filter((entry): entry is string => typeof entry === "string" && entry.trim().length > 0)
      : [];
    const ids = Array.isArray(request.params.ids)
      ? request.params.ids.filter((entry): entry is string => typeof entry === "string" && entry.trim().length > 0)
      : [];
    const lookups = [toolName, id, ...toolNames, ...ids]
      .filter((entry): entry is string => typeof entry === "string" && entry.trim().length > 0);

    if (lookups.length === 0) {
      return {
        requestId: request.requestId,
        success: false,
        error: "office_tool_get requires toolName, id, toolNames, or ids.",
      };
    }
    const details = lookups.map((lookup) => getOfficeToolCapabilityDetail(lookup, request.host));
    return {
      requestId: request.requestId,
      success: true,
      content: details.length === 1 ? { detail: details[0] } : { details },
    };
  }

  if (request.toolName === "mcp_tool_search") {
    return {
      requestId: request.requestId,
      success: true,
      content: {
        query: typeof request.params.query === "string" ? request.params.query : undefined,
        results: [],
        note: "Live MCP tool search is provided by the taskpane runtime when browser-direct or companion connectors are active.",
      },
    };
  }

  if (request.toolName === "office_batch_execute") {
    return {
      requestId: request.requestId,
      success: false,
      error: "office_batch_execute is handled by the taskpane runtime so per-step policy can use the live session.",
    };
  }

  if (request.toolName === "mcp_batch_execute") {
    return {
      requestId: request.requestId,
      success: false,
      error: "mcp_batch_execute is handled by the taskpane runtime so connector policy can use the live session.",
    };
  }

  return undefined;
}

export function createOfficeToolExecutor(dependencies: OfficeToolExecutorDependencies) {
  return async function executeOfficeTool(request: OfficeToolRequest): Promise<OfficeToolResult> {
    try {
      const result =
        searchOfficeToolsForBridge(request) ??
        await executeCommonOfficeTool(request, dependencies) ??
        await executeWordOfficeTool(request, dependencies) ??
        await executeExcelOfficeTool(request, dependencies) ??
        await executePowerPointOfficeTool(request, dependencies);

      if (result) {
        return result;
      }

      return {
        requestId: request.requestId,
        success: false,
        error: request.toolName + " is not supported by the taskpane bridge.",
      };
    } catch (error) {
      dependencies.logger?.error?.("[office-tool] " + request.toolName + " failed", {
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
