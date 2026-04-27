import type {
  McpBatchExecuteRequest,
  McpBatchExecuteResponse,
  OfficeBatchExecuteRequest,
  OfficeBatchExecuteResponse,
  OfficeHost,
  OfficeToolName,
  StructuredBatchStep,
  StructuredBatchStepResult,
  ToolCategory,
} from "@pi-office/pi-office-pack/protocol";
import { TOOL_CATEGORY_MAP } from "@pi-office/pi-office-pack/protocol";

type OfficeToolInvoker = (toolName: OfficeToolName, params: Record<string, unknown>) => Promise<unknown>;
type McpToolInvoker = (toolName: string, args: Record<string, unknown>) => Promise<unknown>;
type McpToolSearcher = (request: { query?: string; connectorId?: string; limit?: number }) => Promise<unknown> | unknown;

const MAX_BATCH_STEPS = 25;
const MAX_SUMMARY_CHARS = 700;
const OFFICE_BATCH_ALLOWED_TOOLS = new Set<OfficeToolName>([
  "office_get_context",
  "office_tool_search",
  "office_tool_get",
  "office_read_section",
  "verify_doc",
  "verify_doc_visual",
  "get_cell_ranges",
  "get_all_objects",
  "search_data",
  "get_range_as_csv",
  "extract_chart_xml",
  "get_presentation_structure",
  "get_slide",
  "list_slide_shapes",
  "verify_slides",
  "office_navigate",
  "office_apply_edit",
  "edit_doc_text",
  "edit_doc_list",
  "office_propose_edits",
]);

function normalizeSteps(steps: StructuredBatchStep[] | undefined): StructuredBatchStep[] {
  return Array.isArray(steps) ? steps.slice(0, MAX_BATCH_STEPS) : [];
}

function toolCategory(toolName: string | undefined): ToolCategory {
  return toolName ? TOOL_CATEGORY_MAP[toolName] ?? "connector" : "connector";
}

function isWriteCategory(category: ToolCategory): boolean {
  return category === "write-doc" || category === "write-external" || category === "connector";
}

function summarizeValue(value: unknown): string {
  const text = typeof value === "string" ? value : JSON.stringify(value);
  if (!text) return "";
  return text.length > MAX_SUMMARY_CHARS ? `${text.slice(0, MAX_SUMMARY_CHARS)}…` : text;
}

function succeeded(result: StructuredBatchStepResult): boolean {
  return result.status === "completed";
}

function batchOk(results: readonly StructuredBatchStepResult[]): boolean {
  return results.every(succeeded);
}

function stepParams(step: StructuredBatchStep): Record<string, unknown> {
  const fromParams = (step as { params?: unknown }).params;
  if (fromParams && typeof fromParams === "object" && !Array.isArray(fromParams)) {
    return fromParams as Record<string, unknown>;
  }
  if (step.arguments && typeof step.arguments === "object" && !Array.isArray(step.arguments)) {
    return step.arguments;
  }
  return {};
}

function stepOperation(step: StructuredBatchStep): string {
  return String((step as { operation?: unknown }).operation ?? step.type ?? "").trim();
}

export async function executeOfficeBatchPlan(input: {
  host: OfficeHost;
  request: OfficeBatchExecuteRequest;
  invokeOfficeTool: OfficeToolInvoker;
}): Promise<OfficeBatchExecuteResponse> {
  const steps = normalizeSteps(input.request.steps);
  const results: StructuredBatchStepResult[] = [];

  for (let index = 0; index < steps.length; index += 1) {
    const step = steps[index]!;
    const result = await executeOfficeBatchStep(step, index, input.invokeOfficeTool);
    results.push(result);
    if (!succeeded(result) && (input.request as { stopOnError?: unknown }).stopOnError !== false) {
      break;
    }
  }

  const completed = results.filter(succeeded).length;
  const failed = results.length - completed;
  return {
    ok: batchOk(results),
    summary: `Office batch executed ${completed}/${steps.length} step${steps.length === 1 ? "" : "s"}${failed ? ` with ${failed} failure${failed === 1 ? "" : "s"}` : ""}.`,
    steps: results,
    blocked: results.filter((entry) => entry.status === "blocked").length,
    failed,
  };
}

async function executeOfficeBatchStep(
  step: StructuredBatchStep,
  index: number,
  invokeOfficeTool: OfficeToolInvoker,
): Promise<StructuredBatchStepResult> {
  const toolName = String(step.toolName ?? "").trim() as OfficeToolName;
  const category = toolCategory(toolName);
  const base = {
    id: step.id,
    index,
    type: step.type || "tool",
    toolName,
  };

  if (!OFFICE_BATCH_ALLOWED_TOOLS.has(toolName)) {
    return {
      ...base,
      status: "blocked",
      summary: toolName === "office_execute_js"
        ? "Blocked raw Office.js escape-hatch execution inside a structured batch."
        : `Blocked unsupported Office batch tool ${toolName || "(missing)"}.`,
      error: toolName === "office_execute_js"
        ? "office_execute_js is not allowed inside structured batches; use the manual one-time escape hatch directly."
        : `${toolName || "toolName"} is not allowed inside structured Office batches.`,
    };
  }

  try {
    const output = await invokeOfficeTool(toolName, stepParams(step));
    return {
      ...base,
      status: "completed",
      summary: summarizeValue(output),
      content: output,
    };
  } catch (error) {
    return {
      ...base,
      status: "failed",
      summary: `${toolName} failed.`,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

export async function executeMcpBatchPlan(input: {
  request: McpBatchExecuteRequest;
  searchMcpTools: McpToolSearcher;
  invokeMcpTool: McpToolInvoker;
}): Promise<McpBatchExecuteResponse> {
  const steps = normalizeSteps(input.request.steps);
  const results: StructuredBatchStepResult[] = [];

  for (let index = 0; index < steps.length; index += 1) {
    const step = steps[index]!;
    const result = await executeMcpBatchStep(step, index, input.searchMcpTools, input.invokeMcpTool);
    results.push(result);
    if (!succeeded(result) && (input.request as { stopOnError?: unknown }).stopOnError !== false) {
      break;
    }
  }

  const completed = results.filter(succeeded).length;
  const failed = results.length - completed;
  return {
    ok: batchOk(results),
    summary: `MCP batch executed ${completed}/${steps.length} step${steps.length === 1 ? "" : "s"}${failed ? ` with ${failed} failure${failed === 1 ? "" : "s"}` : ""}.`,
    steps: results,
    blocked: results.filter((entry) => entry.status === "blocked").length,
    failed,
  };
}

async function executeMcpBatchStep(
  step: StructuredBatchStep,
  index: number,
  searchMcpTools: McpToolSearcher,
  invokeMcpTool: McpToolInvoker,
): Promise<StructuredBatchStepResult> {
  const operation = stepOperation(step);
  const base = {
    id: step.id,
    index,
    type: operation || step.type || "mcp",
    toolName: step.toolName,
  };

  if (operation === "search" || step.type === "mcp_tool_search") {
    const searchRequest: { query?: string; connectorId?: string; limit?: number } = {};
    if (step.query) searchRequest.query = step.query;
    const connectorId = (step as { connectorId?: unknown }).connectorId;
    if (typeof connectorId === "string" && connectorId.trim()) searchRequest.connectorId = connectorId.trim();
    if (typeof step.limit === "number") searchRequest.limit = step.limit;
    const output = await searchMcpTools(searchRequest);
    return {
      ...base,
      status: "completed",
      summary: summarizeValue(output),
      content: output,
    };
  }

  if (operation === "callTool" || step.type === "mcp_call") {
    const toolName = String(step.toolName ?? "").trim();
    if (!toolName) {
      return {
        ...base,
        status: "blocked",
        summary: "Blocked MCP call without toolName.",
        error: "MCP callTool batch steps require toolName.",
      };
    }
    if ((step as unknown as Record<string, unknown>).confirmed !== true) {
      return {
        ...base,
        status: "blocked",
        summary: "Blocked MCP call without explicit confirmation.",
        error: "MCP callTool batch steps require confirmed=true so connector policy cannot be bypassed by the batch envelope.",
      };
    }
    try {
      const output = await invokeMcpTool(toolName, stepParams(step));
      return {
        ...base,
        status: "completed",
        summary: summarizeValue(output),
        content: output,
      };
    } catch (error) {
      return {
        ...base,
        status: "failed",
        summary: `${toolName} failed.`,
        error: error instanceof Error ? error.message : String(error),
      };
    }
  }

  return {
    ...base,
    status: "blocked",
    summary: `Blocked unsupported MCP batch operation ${operation || "(missing)"}.`,
    error: `Unsupported MCP batch operation: ${operation || "(missing)"}.`,
  };
}
