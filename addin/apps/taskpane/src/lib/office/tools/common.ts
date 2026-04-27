import { Type } from "@sinclair/typebox";
import type { OfficeToolDefinition } from "./types";

const officeActionTargetSchema = Type.Object(
  {
    kind: Type.Optional(Type.String({ description: "Target kind such as document, selection, heading, paragraph, range, slide, shape, table, or chart." })),
    position: Type.Optional(Type.String({ description: "For document targets, use replace, start, or end to make document-scope placement explicit." })),
  },
  { additionalProperties: true },
);

const officeActionSchema = Type.Object(
  {
    type: Type.String({
      description:
        "Structured action type. Use insertHtml for Word HTML, replaceDocumentHtml/appendDocumentHtml for document-scope generation, insertText for literal text, or setRangeValues for Excel matrices.",
    }),
    content: Type.Optional(Type.String({ description: "Action payload. For insertHtml this must be valid HTML, not Markdown." })),
    html: Type.Optional(Type.String({ description: "HTML alias for action.content when type is insertHtml." })),
    text: Type.Optional(Type.String({ description: "Text alias for action.content when type is insertText." })),
    target: Type.Optional(officeActionTargetSchema),
    placement: Type.Optional(Type.String({ description: "replace, before, after, start, or end. For document append, use target.kind=document and placement=end." })),
    options: Type.Optional(Type.Object({}, { additionalProperties: true })),
    values: Type.Optional(Type.Any({ description: "2D array for setRangeValues only." })),
  },
  { additionalProperties: true },
);

export const COMMON_OFFICE_TOOL_DEFINITIONS: readonly OfficeToolDefinition[] = [
  {
    name: "office_tool_search",
    hosts: "all",
    category: "read",
    label: "Search Office Tools",
    description:
      "Search Pi-Office's structured Office tool registry and return concise, host-gated capability matches. Use this before reaching for long-tail Office tools or raw Office.js.",
    parameters: Type.Object({
      query: Type.String({ description: "Natural-language task or capability to search for, such as Word tables, paragraph styles, or PowerPoint charts." }),
      host: Type.Optional(Type.String({ description: "Optional host filter: word, excel, or powerpoint. Defaults to the active Office host." })),
      category: Type.Optional(Type.String({ description: "Optional permission category filter such as read, write-doc, or escape-hatch." })),
      maxResults: Type.Optional(Type.Number({ minimum: 1, maximum: 10, description: "Maximum matches to return. Defaults to 5." })),
      includeSchemas: Type.Optional(Type.Boolean({ description: "Include parameter schemas for returned tools. Defaults to false to keep context compact." })),
    }),
    executor: "office-bridge",
    deferred: false,
    capabilityTags: ["registry", "discovery", "office-tools", "deferred-tools"],
    riskLevel: "low",
    compactSummary: "Find supported Office capabilities without loading every tool schema.",
  },
  {
    name: "office_tool_get",
    hosts: "all",
    category: "read",
    label: "Get Office Tool Definition",
    description:
      "Fetch full schema and capability metadata for one or more structured Office tools discovered through office_tool_search.",
    parameters: Type.Object({
      toolNames: Type.Optional(Type.Array(Type.String(), { description: "Office tool names to retrieve." })),
      ids: Type.Optional(Type.Array(Type.String(), { description: "Capability or tool IDs returned by office_tool_search." })),
    }),
    executor: "office-bridge",
    deferred: false,
    capabilityTags: ["registry", "schema", "office-tools", "deferred-tools"],
    riskLevel: "low",
    compactSummary: "Load full schemas for selected Office capabilities only when needed.",
  },
  {
    name: "office_tool_call",
    hosts: "all",
    category: "write-doc",
    label: "Call Discovered Office Tool",
    description:
      "Execute a structured Office tool after discovering it with office_tool_search and loading its schema with office_tool_get. This conservative dispatcher is classified as document-write capable and cannot run office_execute_js.",
    parameters: Type.Object({
      toolName: Type.String({ description: "Exact Office tool name returned by office_tool_search / office_tool_get." }),
      arguments: Type.Optional(Type.Any({ description: "JSON arguments matching the discovered tool schema." })),
    }),
    executor: "runtime-registry",
    deferred: false,
    capabilityTags: ["registry", "execution", "office-tools", "deferred-tools"],
    riskLevel: "high",
    compactSummary: "Execute a discovered Office capability without publishing every specialized schema upfront.",
  },
  {
    name: "mcp_tool_search",
    hosts: "all",
    category: "read",
    label: "Search MCP Tools",
    description:
      "Search enabled browser-direct and companion MCP connector tools by connector, capability, source, and exact executable tool name.",
    parameters: Type.Object({
      query: Type.String({ description: "Natural-language connector task or tool name to search for." }),
      connectorId: Type.Optional(Type.String({ description: "Optional stored or catalog connector ID filter." })),
      maxResults: Type.Optional(Type.Number({ minimum: 1, maximum: 20, description: "Maximum connector tool matches to return. Defaults to 8." })),
      includeSchemas: Type.Optional(Type.Boolean({ description: "Include MCP input schemas when available. Defaults to false." })),
    }),
    executor: "office-bridge",
    deferred: false,
    capabilityTags: ["registry", "discovery", "mcp", "connectors", "deferred-tools"],
    riskLevel: "low",
    compactSummary: "Find enabled connector tools without injecting all connector schemas into the prompt.",
  },
  {
    name: "office_batch_execute",
    hosts: "all",
    category: "write-doc",
    label: "Execute Office Batch Plan",
    description:
      "Execute a constrained, typed batch plan over approved Office tools. This is not arbitrary code execution: escape-hatch Office.js is excluded and the whole batch is permissioned as document-write capable.",
    parameters: Type.Object({
      steps: Type.Array(Type.Any({ description: "Typed Office batch steps." })),
      stopOnError: Type.Optional(Type.Boolean({ description: "Stop after the first failed step. Defaults to true." })),
    }),
    executor: "runtime-registry",
    deferred: false,
    capabilityTags: ["batch", "office", "typed-plan", "structured-tools"],
    riskLevel: "high",
    compactSummary: "Run repeated Office reads/edits through a typed allowlisted plan with per-step policy.",
  },
  {
    name: "mcp_batch_execute",
    hosts: "all",
    category: "connector",
    label: "Execute MCP Batch Plan",
    description:
      "Execute a constrained, typed batch plan over enabled MCP connector tools. Connector policy is checked per call; disabled, destructive, or unknown tools remain unavailable.",
    parameters: Type.Object({
      steps: Type.Array(Type.Any({ description: "Typed MCP batch steps." })),
      stopOnError: Type.Optional(Type.Boolean({ description: "Stop after the first failed step. Defaults to true." })),
    }),
    executor: "runtime-registry",
    deferred: false,
    capabilityTags: ["batch", "mcp", "connector", "typed-plan"],
    riskLevel: "high",
    compactSummary: "Run repeated connector searches/calls through a typed allowlisted plan with connector policy.",
  },
  {
    name: "mcp_result_get",
    hosts: "all",
    category: "read",
    label: "Get MCP Result Page",
    description: "Retrieve a page from a cached MCP result handle returned by a large connector response.",
    parameters: Type.Object({
      handleId: Type.String({ description: "Opaque MCP result handle ID." }),
      page: Type.Optional(Type.Number({ minimum: 0, description: "Zero-based page number. Defaults to 0." })),
    }),
    executor: "runtime-registry",
    deferred: false,
    capabilityTags: ["mcp", "result-handle", "pagination", "connector-cache"],
    riskLevel: "low",
    compactSummary: "Read a paginated cached connector result by handle.",
  },
  {
    name: "mcp_result_summarize",
    hosts: "all",
    category: "read",
    label: "Summarize MCP Result",
    description: "Return a compact extract/summary from a cached MCP result handle without flooding the main chat context.",
    parameters: Type.Object({
      handleId: Type.String({ description: "Opaque MCP result handle ID." }),
      query: Type.Optional(Type.String({ description: "Optional focus query for the summary." })),
      maxChars: Type.Optional(Type.Number({ minimum: 200, maximum: 4000, description: "Maximum summary characters." })),
    }),
    executor: "runtime-registry",
    deferred: false,
    capabilityTags: ["mcp", "result-handle", "summarization", "connector-cache"],
    riskLevel: "low",
    compactSummary: "Summarize cached connector result evidence by handle.",
  },
  {
    name: "mcp_result_clear",
    hosts: "all",
    category: "connector",
    label: "Clear MCP Result Cache",
    description: "Clear one cached MCP result handle or all cached connector results for the current taskpane session.",
    parameters: Type.Object({
      handleId: Type.Optional(Type.String({ description: "Optional handle to clear; omit to clear all cached MCP results." })),
    }),
    executor: "runtime-registry",
    deferred: false,
    capabilityTags: ["mcp", "result-handle", "clear-data", "connector-cache"],
    riskLevel: "medium",
    compactSummary: "Clear cached connector result handles.",
  },
  {
    name: "office_get_context",
    hosts: "all",
    category: "read",
    label: "Office Context",
    description: "Read the current Office document or selection context from the active host.",
    parameters: Type.Object({
      scope: Type.Optional(
        Type.String({
          description:
            "Optional context hint (selection, document, worksheet, workbook, slide, presentation). Scope filtering is currently strongest for Excel and may be treated as a hint for Word/PowerPoint.",
        }),
      ),
    }),
    executor: "office-bridge",
    deferred: false,
    capabilityTags: ["context", "read", "selection", "document"],
    riskLevel: "low",
    compactSummary: "Read active Office document or selection context.",
  },
  {
    name: "office_apply_edit",
    hosts: "all",
    category: "write-doc",
    label: "Office Edit",
    description:
      "Apply native edits to the active Office document, worksheet, or slide. Prefer an object action; action must be an object, not a JSON string. For Word full-document generation use action: { type: \"replaceDocumentHtml\", content: \"<h1>...</h1>\" } or action: { type: \"insertHtml\", target: { kind: \"document\" }, placement: \"end\", content: \"<h2>...</h2>\" }. Legacy operation/text/html params are also supported.",
    parameters: Type.Object({
      mode: Type.Optional(Type.String({ description: "Legacy edit mode such as replaceSelection, insertAfterSelection, or setRangeValues." })),
      content: Type.Optional(Type.String({ description: "Legacy text, HTML, or JSON matrix payload to insert into Office." })),
      text: Type.Optional(Type.String({ description: "Alias for legacy text content. Use when operation/type is insertText." })),
      html: Type.Optional(Type.String({ description: "Alias for legacy HTML content. Bare html infers insertHtml when operation/type/format are omitted." })),
      format: Type.Optional(Type.String({ description: "Legacy content format such as text, html, or matrix." })),
      operation: Type.Optional(Type.String({ description: "Top-level action type alias, e.g., insertText, insertHtml, setRangeValues." })),
      type: Type.Optional(Type.String({ description: "Top-level action type alias when not wrapping with action.type." })),
      values: Type.Optional(Type.Any({ description: "2D array of values for setRangeValues actions only; do not use values[] for Word HTML insertion." })),
      action: Type.Optional(
        officeActionSchema,
      ),
    }),
    executor: "office-bridge",
    deferred: false,
    capabilityTags: ["edit", "structured-action", "office"],
    riskLevel: "medium",
    compactSummary: "Apply a structured Office action to the active host.",
  },
  {
    name: "office_navigate",
    hosts: "all",
    category: "write-doc",
    label: "Office Navigate",
    description: "Move to an Office anchor such as heading, range, or slide.",
    parameters: Type.Object({
      target: Type.Optional(Type.String({ description: "Legacy anchor label or text to navigate to." })),
      kind: Type.Optional(
        Type.String({
          description: "Anchor kind such as heading, paragraph, range, worksheet, slide, or shape.",
        }),
      ),
      anchor: Type.Optional(
        Type.Any({
          description: "Structured Office anchor with host-specific target fields.",
        }),
      ),
    }),
    executor: "office-bridge",
    deferred: false,
    capabilityTags: ["navigate", "anchor", "selection"],
    riskLevel: "medium",
    compactSummary: "Navigate to a structured Office anchor.",
  },
  {
    name: "office_capture_snapshot",
    hosts: "all",
    category: "read",
    label: "Office Snapshot",
    description: "Capture visual snapshots and formatting metadata for the active Office surface.",
    parameters: Type.Object({
      scope: Type.Optional(
        Type.String({
          description:
            "Optional context hint (selection, document, worksheet, workbook, slide, shape). Scope is currently most effective for Excel and may be treated as a hint in Word/PowerPoint.",
        }),
      ),
      includeFormatting: Type.Optional(Type.Boolean({ description: "Include formatting and layout metadata alongside the visuals." })),
      maxImages: Type.Optional(Type.Number({ minimum: 0, maximum: 4, description: "Maximum number of visual snapshots to include." })),
    }),
    executor: "office-bridge",
    deferred: false,
    capabilityTags: ["snapshot", "visual", "metadata", "formatting"],
    riskLevel: "low",
    compactSummary: "Capture Office.js visual/context snapshots and formatting metadata.",
  },
  {
    name: "office_capture_viewport",
    hosts: "all",
    category: "read",
    label: "Office Viewport Screenshot",
    description:
      "Capture a true viewport/window screenshot through the optional companion native capture backend. Office.js document reads/writes still run through the taskpane.",
    parameters: Type.Object({
      includeFormatting: Type.Optional(
        Type.Boolean({ description: "Include companion native capture metadata alongside the screenshot when available." }),
      ),
      includeWindowFrame: Type.Optional(
        Type.Boolean({
          description:
            "Reserved for future native capture support. Browser-only runtime acknowledges this flag but cannot capture the full OS window frame.",
        }),
      ),
    }),
    executor: "companion-native-capture",
    deferred: true,
    requiresCompanion: true,
    capabilityTags: ["viewport", "screenshot", "native-capture", "companion"],
    riskLevel: "low",
    compactSummary: "True viewport screenshot when companion native capture is available.",
  },
  {
    name: "office_execute_js",
    hosts: "all",
    category: "escape-hatch",
    label: "Execute Office.js",
    description:
      "Execute Office.js code as an escape hatch when structured tools are insufficient. This tool is a best-effort restricted subset enforced by regex checks (not an isolated sandbox) and blocks network, storage, eval, and system-access patterns.",
    parameters: Type.Object({
      code: Type.Optional(
        Type.String({
          description:
            "Office.js code to execute. Must use the active host run function (Word.run, Excel.run, or PowerPoint.run). Return a JSON-serializable value. The runtime enforces a best-effort restricted subset (regex checks only, not an isolated sandbox) and blocks network, storage, eval, and system-access patterns.",
        }),
      ),
      script: Type.Optional(Type.String({ description: "Alias for code." })),
    }),
    executor: "office-bridge",
    deferred: false,
    capabilityTags: ["escape-hatch", "office-js", "manual-approval"],
    riskLevel: "critical",
    compactSummary: "Manual one-time approved raw Office.js escape hatch.",
  },
];
