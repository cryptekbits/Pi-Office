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
import { executeWordOfficeTool } from "./office/bridge/word";

export {
  serializeOfficeToolError,
  summarizeOfficeToolError,
  toAnchor,
  toHostAction,
};
export type { OfficeToolExecutorDependencies };

export function createOfficeToolExecutor(dependencies: OfficeToolExecutorDependencies) {
  return async function executeOfficeTool(request: OfficeToolRequest): Promise<OfficeToolResult> {
    try {
      const result =
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
