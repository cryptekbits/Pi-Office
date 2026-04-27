import type { ExtensionFactory } from "@mariozechner/pi-coding-agent";
import { Type } from "@sinclair/typebox";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { HOST_LABELS } from "./defaults.js";
import { OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH } from "./protocol.js";
import type {
  AskUserQuestion,
  AskUserQuestionAnswer,
  AskUserRequest,
  AskUserResponse,
  ImageReasoningEffort,
  OfficeApplyEditParams,
  OfficeContextPayload,
  OfficeHost,
  OfficeNavigateParams,
  OfficeStateUpdate,
  OfficeToolName,
  OfficeVisualSnapshot,
} from "./protocol.js";

export interface ImageGenToolResult {
  base64: string;
  mimeType: string;
  width: number;
  height: number;
  modelKey: string;
  modelName: string;
}

export interface ImageGenToolParams {
  prompt: string;
  aspectRatio?: string | undefined;
  size?: string | undefined;
  quality?: string | undefined;
  insert?: boolean | undefined;
}

export interface OfficeExtensionOptions {
  getHost(): OfficeHost;
  getState(): OfficeStateUpdate | undefined;
  isToolDisabled?(toolName: string): boolean;
  invokeTool(toolName: OfficeToolName, params: Record<string, unknown>): Promise<unknown>;
  invokeAskUser(request: AskUserRequest): Promise<AskUserResponse>;
  invokeEditProposal?(proposal: import("./protocol.js").OfficeEditProposal): Promise<import("./protocol.js").OfficeEditProposalDecision>;
  isImageGenerationEnabled?(): boolean;
  getImageReasoningEffort?(): ImageReasoningEffort;
  getDefaultImageModel?(): string;
  generateImage?(params: ImageGenToolParams): Promise<ImageGenToolResult>;
}

function stripBinaryData(result: unknown): unknown {
  if (!result || typeof result !== "object") {
    return result;
  }

  if (Array.isArray(result)) {
    return result.map((entry) => stripBinaryData(entry));
  }

  const next = { ...(result as Record<string, unknown>) };
  if (Array.isArray(next.visuals)) {
    next.visuals = next.visuals.map((visual) => {
      if (!visual || typeof visual !== "object") {
        return visual;
      }

      const copy = { ...(visual as Record<string, unknown>) };
      if (typeof copy.data === "string") {
        copy.data = `[base64 ${copy.data.length} chars]`;
      }
      return copy;
    });
  }

  return Object.fromEntries(Object.entries(next).map(([key, value]) => [key, stripBinaryData(value)]));
}

function isContextPayload(result: unknown): result is OfficeContextPayload {
  return Boolean(result && typeof result === "object" && "summary" in result && "state" in result);
}

function hasVisuals(result: unknown): result is { visuals: OfficeVisualSnapshot[] } {
  return Boolean(
    result &&
      typeof result === "object" &&
      Array.isArray((result as { visuals?: unknown }).visuals) &&
      (result as { visuals: unknown[] }).visuals.every(
        (visual) =>
          visual &&
          typeof visual === "object" &&
          typeof (visual as OfficeVisualSnapshot).data === "string" &&
          typeof (visual as OfficeVisualSnapshot).mimeType === "string",
      ),
  );
}

function buildReadableParagraphMap(paragraphs: unknown): string | undefined {
  if (!Array.isArray(paragraphs) || !paragraphs.length) return undefined;
  const lines: string[] = [];
  for (const p of paragraphs.slice(0, 40)) {
    if (!p || typeof p !== "object") continue;
    const record = p as Record<string, unknown>;
    const text = typeof record.text === "string" ? record.text.trim() : "";
    const style = typeof record.style === "string" ? record.style : "";
    const isHeading = record.isHeading === true;
    if (!text && !style) continue;
    const prefix = isHeading ? (style.match(/\d/) ? "#".repeat(Number(style.match(/\d/)?.[0]) || 1) + " " : "# ") : "";
    lines.push(`${prefix}${text || "(empty)"}  [${style || "Normal"}]`);
  }
  return lines.length ? lines.join("\n") : undefined;
}

function toToolText(result: unknown): string {
  if (typeof result === "string") {
    return result;
  }

  if (isContextPayload(result)) {
    const snippets = (result as { snippets?: Record<string, unknown> }).snippets;
    const paragraphPreview = snippets ? buildReadableParagraphMap(snippets.paragraphs) : undefined;
    const structured = JSON.stringify(stripBinaryData(result), null, 2);
    return [
      result.summary,
      paragraphPreview ? `\nDocument structure:\n${paragraphPreview}` : undefined,
      `\nStructured data:\n${structured}`,
    ].filter(Boolean).join("\n");
  }

  return JSON.stringify(stripBinaryData(result), null, 2);
}

function toToolContent(result: unknown) {
  const content: Array<{ type: "text"; text: string } | { type: "image"; data: string; mimeType: string }> = [
    { type: "text", text: toToolText(result) },
  ];

  if (hasVisuals(result)) {
    for (const visual of result.visuals) {
      content.push({
        type: "image",
        data: visual.data,
        mimeType: visual.mimeType,
      });
    }
  }

  return content;
}

const getContextParams = Type.Object({
  scope: Type.Optional(
    Type.String({
      description:
        "Optional context hint (selection, document, worksheet, workbook, slide, presentation). Scope filtering is currently strongest for Excel and may be treated as a hint for Word/PowerPoint.",
    }),
  )
});

const applyEditParams = Type.Object({
  mode: Type.Optional(Type.String({ description: "Legacy edit mode such as replaceSelection, insertAfterSelection, or setRangeValues." })),
  content: Type.Optional(Type.String({ description: "Legacy text, HTML, or JSON matrix payload to insert into Office." })),
  text: Type.Optional(Type.String({ description: "Alias for legacy text content. Use when operation/type is insertText." })),
  html: Type.Optional(Type.String({ description: "Alias for legacy HTML content. Bare html infers insertHtml when operation/type/format are omitted." })),
  format: Type.Optional(Type.String({ description: "Legacy content format such as text, html, or matrix." })),
  operation: Type.Optional(Type.String({ description: "Top-level action type alias, e.g., insertText, insertHtml, setRangeValues." })),
  type: Type.Optional(Type.String({ description: "Top-level action type alias when not wrapping with action.type." })),
  values: Type.Optional(Type.Any({ description: "2D array of values for setRangeValues actions only; do not use values[] for Word HTML insertion." })),
  action: Type.Optional(
    Type.Any({
      description:
        "Structured host action. Prefer { type: \"insertHtml\", content: \"...\" } for Word document generation with HTML, and prefer this over legacy mode/content for workbook, slide, comment, shape image, chart, table, and navigation-aware edits. Destructive actions should set action.options.confirmDestructive=true so the host can distinguish intentional deletes/clears from accidental ones.",
    }),
  ),
});

const navigateParams = Type.Object({
  target: Type.Optional(Type.String({ description: "Legacy anchor label or text to navigate to." })),
  kind: Type.Optional(
    Type.String({
      description:
        "heading, comment, revision, footnote, endnote, paragraph, field, contentControl, cell, range, sheet, workbook, slide, notesRegion, layout, slideMaster, shape, chart, or pivotTable.",
    }),
  ),
  anchor: Type.Optional(
    Type.Any({
      description: "Structured navigation anchor. Prefer this when sheet names, range addresses, slide IDs, or paragraph IDs are known.",
    }),
  ),
});

const captureSnapshotParams = Type.Object({
  scope: Type.Optional(
    Type.String({
      description:
        "Optional context hint (selection, document, worksheet, workbook, slide, shape). Scope is currently most effective for Excel and may be treated as a hint in Word/PowerPoint.",
    }),
  ),
  includeFormatting: Type.Optional(Type.Boolean({ description: "Include formatting and layout metadata alongside Office.js context snapshots." })),
  maxImages: Type.Optional(Type.Number({ minimum: 0, maximum: 4, description: "Maximum number of Office.js context snapshots to include." })),
});

const captureViewportParams = Type.Object({
  includeFormatting: Type.Optional(
    Type.Boolean({ description: "Include Word viewport metadata such as visible pages, scroll position, and view mode." }),
  ),
  includeWindowFrame: Type.Optional(
    Type.Boolean({
      description:
        "Reserved for future native capture support. In browser-only runtime this is acknowledged but cannot capture the full OS window frame.",
    }),
  ),
});

const readSectionParams = Type.Object({
  startIndex: Type.Optional(Type.Number({ description: "Zero-based paragraph start index (inclusive)." })),
  endIndex: Type.Optional(Type.Number({ description: "Zero-based paragraph end index (exclusive). Defaults to startIndex + 20." })),
  start: Type.Optional(Type.Number({ description: "Alias for startIndex." })),
  end: Type.Optional(Type.Number({ description: "Alias for endIndex." })),
  includeStyles: Type.Optional(Type.Boolean({ description: "Include paragraph styles and heading levels. Defaults to true." })),
});

const verifyDocParams = Type.Object({
  scope: Type.Optional(
    Type.String({
      description: "Verification target for Word context capture (selection or document). Defaults to document-level verification context.",
    }),
  ),
  includeFormatting: Type.Optional(
    Type.Boolean({
      description: "Include formatting and review metadata in the structured verification details. Defaults to true.",
    }),
  ),
});

const verifyDocVisualParams = Type.Object({
  includeFormatting: Type.Optional(
    Type.Boolean({
      description: "Include viewport formatting metadata in the structured visual verification details. Defaults to true.",
    }),
  ),
  includeWindowFrame: Type.Optional(
    Type.Boolean({
      description:
        "Reserved for future native capture support. Browser-only runtime acknowledges this flag but cannot capture the full OS window frame.",
    }),
  ),
});

const getCellRangesParams = Type.Object({
  sheetName: Type.Optional(Type.String({ description: "Worksheet name that contains the range. Defaults to the active worksheet." })),
  address: Type.Optional(Type.String({ description: "A1-style cell/range address. Defaults to the current selection." })),
  includeValues: Type.Optional(Type.Boolean({ description: "Include range values in the response. Defaults to true." })),
  includeText: Type.Optional(Type.Boolean({ description: "Include rendered text values in the response. Defaults to true." })),
  includeFormulas: Type.Optional(Type.Boolean({ description: "Include range formulas in the response. Defaults to true." })),
  includeNumberFormat: Type.Optional(Type.Boolean({ description: "Include number formats in the response. Defaults to true." })),
}, { additionalProperties: true });

const setCellRangeParams = Type.Object({
  sheetName: Type.Optional(Type.String({ description: "Worksheet name that contains the destination range. Defaults to active worksheet." })),
  address: Type.Optional(Type.String({ description: "A1-style destination address. Defaults to current selection." })),
  values: Type.Optional(Type.Any({ description: "2D matrix values to write into the target range." })),
  content: Type.Optional(Type.String({ description: "JSON matrix alias when values is omitted." })),
  options: Type.Optional(Type.Any({ description: "Additional write options forwarded to the host adapter." })),
}, { additionalProperties: true });

const clearCellRangeParams = Type.Object({
  sheetName: Type.Optional(Type.String({ description: "Worksheet name that contains the range to clear." })),
  address: Type.Optional(Type.String({ description: "A1-style address to clear. Defaults to current selection." })),
  applyTo: Type.Optional(
    Type.String({
      description: "Clear mode: all, contents, formats, hyperlinks, removeHyperlinks. Defaults to all.",
    }),
  ),
  confirmDestructive: Type.Optional(
    Type.Boolean({
      description: "Set true to acknowledge this destructive clear operation.",
    }),
  ),
  options: Type.Optional(Type.Any({ description: "Additional clear options forwarded to the host adapter." })),
}, { additionalProperties: true });

const resizeRangeParams = Type.Object({
  sheetName: Type.Optional(Type.String({ description: "Worksheet name that contains the source range." })),
  address: Type.Optional(Type.String({ description: "A1-style source range address. Defaults to current selection." })),
  rowCount: Type.Optional(Type.Number({ minimum: 1, description: "Final row count for the resized range." })),
  columnCount: Type.Optional(Type.Number({ minimum: 1, description: "Final column count for the resized range." })),
  rowDelta: Type.Optional(Type.Number({ description: "Relative row delta when rowCount is not provided." })),
  columnDelta: Type.Optional(Type.Number({ description: "Relative column delta when columnCount is not provided." })),
  activate: Type.Optional(Type.Boolean({ description: "Activate/select the resized range after resolving it." })),
  options: Type.Optional(Type.Any({ description: "Additional resize options forwarded to the host adapter." })),
}, { additionalProperties: true });

const copyToParams = Type.Object({
  sourceSheetName: Type.Optional(Type.String({ description: "Worksheet name for the source range. Defaults to active worksheet." })),
  sourceAddress: Type.Optional(Type.String({ description: "A1-style source range address. Defaults to current selection." })),
  destinationSheetName: Type.Optional(Type.String({ description: "Worksheet name for the destination range." })),
  destinationAddress: Type.String({ description: "A1-style destination range address." }),
  copyType: Type.Optional(
    Type.String({
      description: "Excel copy type: All, Formats, Formulas, Values, or Link.",
    }),
  ),
  skipBlanks: Type.Optional(Type.Boolean({ description: "Skip blank cells while copying. Defaults to false." })),
  transpose: Type.Optional(Type.Boolean({ description: "Transpose copied rows/columns. Defaults to false." })),
  options: Type.Optional(Type.Any({ description: "Additional copy options forwarded to the host adapter." })),
}, { additionalProperties: true });

const modifySheetStructureParams = Type.Object({
  operation: Type.String({
    description: "Worksheet structure operation: create_worksheet, rename_worksheet, duplicate_worksheet, or delete_worksheet.",
  }),
  sheetName: Type.Optional(Type.String({ description: "Worksheet name targeted by rename/duplicate/delete operations." })),
  name: Type.Optional(Type.String({ description: "Worksheet name for create/rename/duplicate operations." })),
  relativeTo: Type.Optional(Type.String({ description: "Worksheet name used as placement anchor for duplication." })),
  positionType: Type.Optional(Type.String({ description: "Worksheet copy placement type when duplicating (before/after)." })),
  confirmDestructive: Type.Optional(
    Type.Boolean({
      description: "Set true when performing destructive operations such as delete_worksheet.",
    }),
  ),
  options: Type.Optional(Type.Any({ description: "Additional structure options forwarded to the host adapter." })),
}, { additionalProperties: true });

const modifyObjectParams = Type.Object({
  operation: Type.String({
    description:
      "Excel object mutation operation (format_range, create_table, format_table, apply_table_filter, clear_table_filter, clear_table_filters, reapply_table_filters, create_chart, update_chart, create_pivot_table, update_pivot_table, sort_pivot_field, sort_pivot_by_labels, sort_pivot_by_values, refresh_pivot_table, set_worksheet_gridlines, set_worksheet_headings, set_print_area, set_data_validation, clear_data_validation, add_conditional_format, clear_conditional_formats, insert_inline_picture).",
  }),
  sheetName: Type.Optional(Type.String({ description: "Worksheet name for range/table/chart/pivot operations." })),
  address: Type.Optional(Type.String({ description: "A1-style range address when an operation targets a worksheet range." })),
  tableName: Type.Optional(Type.String({ description: "Target table name for table-oriented operations." })),
  chartName: Type.Optional(Type.String({ description: "Target chart name for chart-oriented operations." })),
  pivotTableName: Type.Optional(Type.String({ description: "Target PivotTable name for pivot-oriented operations." })),
  confirmDestructive: Type.Optional(Type.Boolean({ description: "Set true for destructive operations when required." })),
  options: Type.Optional(Type.Any({ description: "Additional object-operation options forwarded to the host adapter." })),
}, { additionalProperties: true });

const getAllObjectsParams = Type.Object({
  scope: Type.Optional(
    Type.String({
      description: "Inventory scope hint: selection, worksheet, or workbook. Defaults to workbook.",
    }),
  ),
  includeFormatting: Type.Optional(
    Type.Boolean({
      description: "Include worksheet formatting metadata in addition to object inventory. Defaults to true.",
    }),
  ),
  objectTypes: Type.Optional(
    Type.Array(Type.String(), {
      description: "Optional object kinds to include (table, chart, pivotTable, namedItem, worksheet, cell).",
    }),
  ),
}, { additionalProperties: true });

const searchDataParams = Type.Object({
  query: Type.String({ description: "Case-insensitive query used to search workbook/worksheet objects and cited cells." }),
  scope: Type.Optional(
    Type.String({
      description: "Search scope hint: selection, worksheet, or workbook. Defaults to workbook.",
    }),
  ),
  objectTypes: Type.Optional(
    Type.Array(Type.String(), {
      description: "Object kinds to search: table, chart, pivotTable, namedItem, worksheet, or cell.",
    }),
  ),
  limit: Type.Optional(Type.Number({ minimum: 1, maximum: 200, description: "Maximum number of matches to return." })),
}, { additionalProperties: true });

const getRangeAsCsvParams = Type.Object({
  sheetName: Type.Optional(Type.String({ description: "Worksheet name that contains the source range. Defaults to active worksheet." })),
  address: Type.Optional(Type.String({ description: "A1-style source range address. Defaults to current selection." })),
  delimiter: Type.Optional(Type.String({ description: "CSV delimiter. Defaults to comma." })),
  quoteValues: Type.Optional(Type.Boolean({ description: "Wrap and escape all CSV cells in quotes. Defaults to false." })),
  includeHeaders: Type.Optional(Type.Boolean({ description: "Include header row. Defaults to true." })),
  includeFormulas: Type.Optional(
    Type.Boolean({
      description: "Export formulas instead of displayed values for auditable formula-first reviews. Defaults to false.",
    }),
  ),
}, { additionalProperties: true });

const readRangeImageParams = Type.Object({
  sheetName: Type.Optional(Type.String({ description: "Worksheet name hint. The capture still uses the active Excel selection." })),
  address: Type.Optional(Type.String({ description: "A1-style range hint. Select or navigate to this range before capture; browser-only runtime does not render arbitrary offscreen ranges." })),
  scope: Type.Optional(
    Type.String({
      description: "Visual capture scope hint. Defaults to the active selection.",
    }),
  ),
  includeFormatting: Type.Optional(Type.Boolean({ description: "Include formatting metadata in the visual payload. Defaults to true." })),
  maxImages: Type.Optional(Type.Number({ minimum: 1, maximum: 4, description: "Maximum number of active-selection snapshots to include." })),
}, { additionalProperties: true });

const extractChartXmlParams = Type.Object({
  sheetName: Type.Optional(Type.String({ description: "Worksheet name that contains the chart. Defaults to active worksheet." })),
  chartName: Type.Optional(Type.String({ description: "Chart name to extract. Required when chartId/chartIndex are not provided." })),
  chartId: Type.Optional(Type.String({ description: "Optional chart id alias when chartName is unknown." })),
  chartIndex: Type.Optional(Type.Number({ minimum: 1, description: "Optional one-based chart index when chartName is unknown." })),
}, { additionalProperties: true });

const getPresentationStructureParams = Type.Object({
  maxSlides: Type.Optional(
    Type.Number({
      minimum: 1,
      description: "Optional maximum number of slide previews to include. Defaults to a bounded preview size.",
    }),
  ),
  includeSlideText: Type.Optional(
    Type.Boolean({
      description: "Include per-slide title/body text previews when available. Defaults to true.",
    }),
  ),
}, { additionalProperties: true });

const getSlideParams = Type.Object({
  slideId: Type.Optional(Type.String({ description: "PowerPoint slide ID to read." })),
  slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based slide index to read." })),
  includeShapes: Type.Optional(Type.Boolean({ description: "Include shape summaries for the resolved slide. Defaults to true." })),
  includeSlideText: Type.Optional(Type.Boolean({ description: "Include title/body text previews for the resolved slide. Defaults to true." })),
}, { additionalProperties: true });

const listSlideShapesParams = Type.Object({
  slideId: Type.Optional(Type.String({ description: "PowerPoint slide ID whose shapes should be listed." })),
  slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based slide index whose shapes should be listed." })),
  maxShapes: Type.Optional(Type.Number({ minimum: 1, description: "Optional maximum number of shapes to return." })),
}, { additionalProperties: true });

const modifyPresentationStructureParams = Type.Object({
  operation: Type.String({
    description:
      "Presentation structure operation (add_slide, move_slide, reorder_slides, delete_slide, apply_layout, select_slides, add_agenda_slide, add_transition_slide, combine_slides, import_slides_from_base64).",
  }),
  slideId: Type.Optional(Type.String({ description: "Target slide ID for the operation." })),
  slideIds: Type.Optional(Type.Array(Type.String(), { description: "Ordered list of slide IDs for multi-slide operations." })),
  slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based slide index target." })),
  targetSlideId: Type.Optional(Type.String({ description: "Insertion target slide ID where applicable." })),
  formatting: Type.Optional(Type.String({ description: "PowerPoint insert formatting mode when supported." })),
  confirmDestructive: Type.Optional(
    Type.Boolean({
      description: "Required for destructive operations such as delete_slide/delete_slides.",
    }),
  ),
  content: Type.Optional(Type.String({ description: "Optional text payload used by supported slide-creation helpers." })),
  options: Type.Optional(Type.Any({ description: "Additional operation-specific options forwarded to the host adapter." })),
}, { additionalProperties: true });

const duplicateSlideParams = Type.Object({
  slideId: Type.Optional(Type.String({ description: "Single source slide ID to duplicate." })),
  slideIds: Type.Optional(Type.Array(Type.String(), { description: "One or more source slide IDs to duplicate in order." })),
  slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based source slide index when slideId is not known." })),
  targetSlideId: Type.Optional(Type.String({ description: "Slide ID to insert duplicates after." })),
  formatting: Type.Optional(Type.String({ description: "PowerPoint insert formatting mode when supported." })),
  options: Type.Optional(Type.Any({ description: "Additional duplication options forwarded to the host adapter." })),
}, { additionalProperties: true });

const insertSlideElementParams = Type.Object({
  operation: Type.String({
    description:
      "PowerPoint element insertion operation (add_text_box, add_geometric_shape, add_table, add_line, add_process_flow, add_simple_diagram, insert_inline_picture).",
  }),
  slideId: Type.Optional(Type.String({ description: "Target slide ID for inserting the new element." })),
  slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based slide index target when slideId is not known." })),
  shapeId: Type.Optional(Type.String({ description: "Optional shape target for grouped operations." })),
  content: Type.Optional(Type.String({ description: "Primary text payload (or base64 image payload for insert_inline_picture)." })),
  text: Type.Optional(Type.String({ description: "Alias for content when inserting text." })),
  values: Type.Optional(Type.Any({ description: "Matrix payload for table insertion when applicable." })),
  options: Type.Optional(Type.Any({ description: "Additional insertion options forwarded to the host adapter." })),
}, { additionalProperties: true });

const removeSlideElementParams = Type.Object({
  operation: Type.String({
    description:
      "PowerPoint element removal operation (remove_shape, remove_shapes, clear_shape_text).",
  }),
  slideId: Type.Optional(Type.String({ description: "Target slide ID for the removal operation." })),
  slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based slide index target when slideId is not known." })),
  shapeId: Type.Optional(Type.String({ description: "Primary shape ID to remove or clear." })),
  shapeIds: Type.Optional(Type.Array(Type.String(), { description: "One or more shape IDs for multi-shape removal." })),
  confirmDestructive: Type.Optional(
    Type.Boolean({
      description: "Required for destructive operations such as remove_shape/remove_shapes.",
    }),
  ),
  options: Type.Optional(Type.Any({ description: "Additional removal options forwarded to the host adapter." })),
}, { additionalProperties: true });

const editSlideTextParams = Type.Object({
  operation: Type.Optional(
    Type.String({
      description:
        "PowerPoint text-edit operation (set_shape_text, append_shape_text, clear_shape_text, insert_text). Defaults to set_shape_text when shapeId is provided, otherwise insert_text.",
    }),
  ),
  slideId: Type.Optional(Type.String({ description: "Target slide ID for text updates." })),
  slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based slide index target when slideId is not known." })),
  shapeId: Type.Optional(Type.String({ description: "Shape ID whose text should be updated." })),
  content: Type.Optional(Type.String({ description: "Text payload to apply." })),
  text: Type.Optional(Type.String({ description: "Alias for content." })),
  placement: Type.Optional(Type.String({ description: "Optional placement hint (replace or after)." })),
  options: Type.Optional(Type.Any({ description: "Additional text-edit options forwarded to the host adapter." })),
}, { additionalProperties: true });

const editSlideXmlParams = Type.Object({
  operation: Type.String({
    description:
      "PowerPoint XML/serialized operation (inspect_presentation_package, get_presentation_theme, get_slide_notes, set_slide_notes, replace_slide_notes, import_slides_from_base64, merge_presentation_from_base64, export_slides_as_base64).",
  }),
  slideId: Type.Optional(Type.String({ description: "Target slide ID for slide-scoped XML operations." })),
  slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based slide index target when slideId is not known." })),
  content: Type.Optional(Type.String({ description: "Text or base64 payload used by mutating XML operations." })),
  base64: Type.Optional(Type.String({ description: "Alias for content when providing serialized PPTX payloads." })),
  formatting: Type.Optional(Type.String({ description: "Insert formatting mode for base64 import operations when supported." })),
  options: Type.Optional(Type.Any({ description: "Additional XML operation options forwarded to the host adapter." })),
}, { additionalProperties: true });

const editSlideMasterParams = Type.Object({
  operation: Type.Optional(
    Type.String({
      description:
        "PowerPoint layout operation. Currently supports apply_layout/set_layout only; this legacy-named tool does not edit slide masters.",
    }),
  ),
  slideId: Type.Optional(Type.String({ description: "Target slide ID whose layout should be updated." })),
  slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based slide index target when slideId is not known." })),
  layoutId: Type.Optional(Type.String({ description: "Layout ID to apply." })),
  layoutName: Type.Optional(Type.String({ description: "Layout name to apply." })),
  slideMasterId: Type.Optional(Type.String({ description: "Optional slide master ID used only to resolve the requested layout." })),
  slideMasterName: Type.Optional(Type.String({ description: "Optional slide master name used only to resolve the requested layout." })),
  options: Type.Optional(Type.Any({ description: "Additional layout-application options forwarded to the host adapter." })),
}, { additionalProperties: true });

const editSlideChartParams = Type.Object({
  operation: Type.Optional(
    Type.String({
      description:
        "PowerPoint chart operation (get_slide_charts, add_slide_chart, update_slide_chart). Defaults to update_slide_chart when omitted.",
    }),
  ),
  slideId: Type.Optional(Type.String({ description: "Target slide ID for chart inspection or mutation." })),
  slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based slide index target when slideId is not known." })),
  shapeId: Type.Optional(Type.String({ description: "Optional chart shape ID for precise chart targeting." })),
  chartIndex: Type.Optional(Type.Number({ minimum: 1, description: "Optional one-based chart index within the target slide package." })),
  shapeName: Type.Optional(Type.String({ description: "Optional chart shape name used to resolve chart edits." })),
  title: Type.Optional(Type.String({ description: "Chart title for create/update operations." })),
  categories: Type.Optional(Type.Any({ description: "Ordered category labels for chart create/update operations." })),
  series: Type.Optional(Type.Any({ description: "Series payload for chart create/update operations." })),
  replaceOriginal: Type.Optional(Type.Boolean({ description: "When true (default), replace the source slide after serialized chart updates." })),
  formatting: Type.Optional(Type.String({ description: "Insert formatting mode when serialized chart updates insert replacement slides." })),
  options: Type.Optional(Type.Any({ description: "Additional chart operation options forwarded to the host adapter." })),
}, { additionalProperties: true });

const copyImageBetweenSlidesParams = Type.Object({
  sourceSlideId: Type.Optional(Type.String({ description: "Source slide ID that contains the image shape to copy." })),
  sourceSlideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based source slide index when sourceSlideId is unknown." })),
  sourceShapeId: Type.Optional(Type.String({ description: "Source image shape ID to copy from." })),
  sourceImageBase64: Type.Optional(Type.String({ description: "Optional image base64 override when source shape export is unavailable." })),
  targetSlideId: Type.Optional(Type.String({ description: "Destination slide ID for image placement/replacement." })),
  targetSlideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based destination slide index when targetSlideId is unknown." })),
  targetShapeId: Type.Optional(Type.String({ description: "Destination shape ID to update. If omitted, inserts a new image shape." })),
  left: Type.Optional(Type.Number({ description: "Optional destination left position in points for inserted images." })),
  top: Type.Optional(Type.Number({ description: "Optional destination top position in points for inserted images." })),
  width: Type.Optional(Type.Number({ description: "Optional destination width in points for inserted images." })),
  height: Type.Optional(Type.Number({ description: "Optional destination height in points for inserted images." })),
  options: Type.Optional(Type.Any({ description: "Additional media-copy options forwarded to the host adapter." })),
}, { additionalProperties: true });

const searchIconsParams = Type.Object({
  query: Type.String({ description: "Icon search query text." }),
  maxResults: Type.Optional(Type.Number({ minimum: 1, maximum: 50, description: "Maximum number of icon matches to return." })),
  style: Type.Optional(Type.String({ description: "Optional style/category hint used by the icon catalog search." })),
  options: Type.Optional(Type.Any({ description: "Additional icon-search options forwarded to the host adapter." })),
}, { additionalProperties: true });

const insertIconParams = Type.Object({
  iconId: Type.Optional(Type.String({ description: "Icon ID returned by search_icons." })),
  iconName: Type.Optional(Type.String({ description: "Icon name alias when iconId is unknown." })),
  query: Type.Optional(Type.String({ description: "Fallback query used when selecting an icon by search text." })),
  slideId: Type.Optional(Type.String({ description: "Target slide ID for icon insertion." })),
  slideIndex: Type.Optional(Type.Number({ minimum: 1, description: "One-based target slide index when slideId is unknown." })),
  shapeId: Type.Optional(Type.String({ description: "Optional target shape ID for icon replacement workflows." })),
  left: Type.Optional(Type.Number({ description: "Optional icon left position in points." })),
  top: Type.Optional(Type.Number({ description: "Optional icon top position in points." })),
  width: Type.Optional(Type.Number({ description: "Optional icon width in points." })),
  height: Type.Optional(Type.Number({ description: "Optional icon height in points." })),
  fillColor: Type.Optional(Type.String({ description: "Optional icon fill/text color (hex/rgb)." })),
  lineColor: Type.Optional(Type.String({ description: "Optional icon outline color (hex/rgb)." })),
  options: Type.Optional(Type.Any({ description: "Additional icon insertion options forwarded to the host adapter." })),
}, { additionalProperties: true });

const verifySlidesParams = Type.Object({
  scope: Type.Optional(
    Type.String({
      description: "Structural verification scope for PowerPoint slides (presentation or selection). Defaults to presentation structure.",
    }),
  ),
  maxSlides: Type.Optional(Type.Number({ minimum: 1, description: "Maximum number of slide previews to include in verification details." })),
  includeSlideText: Type.Optional(Type.Boolean({ description: "Include bounded slide text previews in structural verification output." })),
  includeFormatting: Type.Optional(Type.Boolean({ description: "Include layout/master metadata where supported. Defaults to true." })),
}, { additionalProperties: true });

const verifySlideVisualParams = Type.Object({
  scope: Type.Optional(
    Type.String({
      description: "Visual verification scope. Defaults to slide selection snapshots.",
    }),
  ),
  includeFormatting: Type.Optional(Type.Boolean({ description: "Include slide/shape formatting metadata in visual verification details. Defaults to true." })),
  maxImages: Type.Optional(Type.Number({ minimum: 1, maximum: 4, description: "Maximum number of slide/shape snapshot images to include." })),
}, { additionalProperties: true });

const executeJsParams = Type.Object({
  code: Type.Optional(
    Type.String({
      description:
        "Office.js code to execute. Must use the active host run function (Word.run, Excel.run, or PowerPoint.run) and return a JSON-serializable value. The runtime enforces a best-effort restricted subset (regex checks only, not an isolated sandbox) and blocks network, storage, eval, and system-access patterns.",
    }),
  ),
  script: Type.Optional(Type.String({ description: "Alias for code." })),
});

const proposeEditsParams = Type.Object({
  edits: Type.Array(
    Type.Object({
      kind: Type.String({ description: "insert, replace, or delete." }),
      searchText: Type.Optional(Type.String({
        maxLength: OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH,
        description:
          `Text to locate in the document for replace/delete. Must be under ${OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH} characters.`,
      })),
      oldText: Type.Optional(Type.String({ description: "Expected existing text (for replace/delete verification)." })),
      newText: Type.Optional(Type.String({ description: "Replacement text (for insert/replace)." })),
      anchor: Type.Optional(Type.String({ description: "Heading or paragraph label to scope the search." })),
      paragraphId: Type.Optional(Type.String({ description: "Paragraph unique ID for precise targeting." })),
      explanation: Type.Optional(Type.String({ description: "Brief rationale for this edit." })),
    }),
    { description: "Ordered list of proposed edits." },
  ),
  summary: Type.String({ description: "One-sentence summary of all proposed changes." }),
});

const askUserOptionSchema = Type.Union([
  Type.String(),
  Type.Object({
    title: Type.String(),
    description: Type.Optional(Type.String()),
  }),
]);

const askUserQuestionSchema = Type.Object({
  question: Type.String({ description: "The question to ask." }),
  context: Type.Optional(Type.String({ description: "Relevant context summary shown before the question." })),
  options: Type.Optional(Type.Array(askUserOptionSchema, { description: "Multiple-choice options." })),
});

const askUserParams = Type.Object({
  question: Type.Optional(Type.String({ description: "A single question to ask the user." })),
  context: Type.Optional(Type.String({ description: "Context for a single question." })),
  options: Type.Optional(Type.Array(askUserOptionSchema, { description: "Options for a single question." })),
  questions: Type.Optional(Type.Array(askUserQuestionSchema, { description: "Multiple questions to ask in a carousel." })),
});

const generateImageParams = Type.Object({
  prompt: Type.String({ description: "Description of the image to generate." }),
  aspectRatio: Type.Optional(
    Type.String({ description: "Aspect ratio: 1:1, 16:9, 9:16, 4:3, 3:4, 3:2, 2:3. Defaults based on document type." }),
  ),
  size: Type.Optional(Type.String({ description: "Resolution tier: 1K, 2K, or 4K. Auto-resolved based on document context." })),
  quality: Type.Optional(Type.String({ description: "Quality: auto, low, medium, high. Defaults to auto." })),
  insert: Type.Optional(
    Type.Boolean({ description: "Whether to insert the generated image into the document. Defaults to true." }),
  ),
});

export function getOfficeSkillPaths(): string[] {
  const distDir = dirname(fileURLToPath(import.meta.url));
  return [
    join(distDir, "../skills/office-host.SKILL.md"),
    join(distDir, "../skills/workspace-handoff.SKILL.md"),
  ];
}

export function createOfficeExtension(options: OfficeExtensionOptions): ExtensionFactory {
  return (pi) => {
    const isDisabled = options.isToolDisabled ?? (() => false);
    const executeReviewableWordEdits = async (
      toolName: "office_propose_edits" | "edit_doc_list",
      params: Record<string, unknown>,
    ) => {
      const result = await options.invokeTool(toolName, params);
      const resultObj = result as Record<string, unknown> | undefined;
      if (!resultObj || resultObj.error || !Array.isArray(resultObj.edits) || !options.invokeEditProposal) {
        return { content: toToolContent(result), details: result };
      }

      const proposal: import("./protocol.js").OfficeEditProposal = {
        requestId: crypto.randomUUID(),
        edits: resultObj.edits as import("./protocol.js").OfficeProposedEdit[],
        summary: String(resultObj.summary ?? "Proposed edits"),
      };

      const decision = await options.invokeEditProposal(proposal);
      const accepted = decision.decisions.filter((d) => d.accepted);
      const rejected = decision.decisions.filter((d) => !d.accepted);

      const reasonLabels: Record<string, string> = {
        keep_original: "Keep original",
        rewrite_differently: "Rewrite differently",
        not_relevant: "Not relevant",
      };

      const appResult = decision.applicationResult;
      let applicationLine: string;
      if (appResult) {
        if (appResult.failed > 0) {
          applicationLine = `Application result: ${appResult.applied} applied, ${appResult.failed} failed. Errors: ${appResult.errors.join("; ")}`;
        } else if (appResult.applied > 0) {
          applicationLine = `All ${appResult.applied} accepted edit${appResult.applied !== 1 ? "s" : ""} applied successfully.`;
        } else {
          applicationLine = "No edits were applied.";
        }
      } else {
        applicationLine = accepted.length ? "Accepted edits were applied to the document." : "No edits were applied.";
      }

      const lines: string[] = [
        `Edit proposal reviewed: ${accepted.length} accepted, ${rejected.length} rejected out of ${proposal.edits.length} edits.`,
        applicationLine,
      ];

      if (rejected.length) {
        lines.push("", "Rejected edits:");
        for (const d of rejected) {
          const edit = proposal.edits.find((e) => e.id === d.editId);
          const preview = (edit?.searchText ?? edit?.oldText ?? "").slice(0, 60);
          const reason = d.rejectReason ? reasonLabels[d.rejectReason] ?? d.rejectReason : "No reason given";
          const note = d.rejectNote ? ` (note: "${d.rejectNote}")` : "";
          lines.push(`- "${preview}${preview.length >= 60 ? "..." : ""}": ${reason}${note}`);
        }
      }

      if (decision.globalFeedback) {
        lines.push("", `User feedback: "${decision.globalFeedback}"`);
      }

      return {
        content: [{ type: "text" as const, text: lines.join("\n") }],
        details: {
          proposalId: proposal.requestId,
          summary: proposal.summary,
          totalEdits: proposal.edits.length,
          accepted: accepted.length,
          rejected: rejected.length,
          decisions: decision.decisions,
          globalFeedback: decision.globalFeedback,
        },
      };
    };

    if (!isDisabled("office_get_context"))
    pi.registerTool({
      name: "office_get_context",
      label: "Office Context",
      description: "Read the current Office document or selection context from the active host.",
      parameters: getContextParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("office_get_context", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("office_apply_edit"))
    pi.registerTool({
      name: "office_apply_edit",
      label: "Office Edit",
      description: "Apply a native edit to the active Office document, selection, range, worksheet, or slide. For Word HTML insertion use action: { type: \"insertHtml\", content: \"<p>...</p>\" }.",
      parameters: applyEditParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("office_apply_edit", params as OfficeApplyEditParams as Record<string, unknown>);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("edit_doc_text"))
    pi.registerTool({
      name: "edit_doc_text",
      label: "Edit Word Text",
      description:
        "Word-only first-class text editing. Use this for direct clause/sentence updates through native Word actions when no per-edit review card is required.",
      parameters: applyEditParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("edit_doc_text", params as OfficeApplyEditParams as Record<string, unknown>);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("office_navigate"))
    pi.registerTool({
      name: "office_navigate",
      label: "Office Navigate",
      description: "Move the user back to a document anchor such as a heading, comment, slide, or range when supported.",
      parameters: navigateParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("office_navigate", params as OfficeNavigateParams as Record<string, unknown>);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("office_capture_snapshot"))
    pi.registerTool({
      name: "office_capture_snapshot",
      label: "Office Snapshot",
      description: "Capture Office.js context snapshots and formatting metadata for the current Office selection when layout fidelity matters. This is not an OS/window screenshot.",
      parameters: captureSnapshotParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("office_capture_snapshot", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("office_capture_viewport"))
    pi.registerTool({
      name: "office_capture_viewport",
      label: "Office Viewport Screenshot",
      description:
        "Compatibility tool for companion-native true viewport/window screenshots. Register it only when capability resolution says companion native capture is available for the active host; use office_capture_snapshot or host visual verification tools otherwise.",
      parameters: captureViewportParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("office_capture_viewport", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("office_read_section"))
    pi.registerTool({
      name: "office_read_section",
      label: "Read Document Section",
      description:
        "Read a range of paragraphs from the active Word document by paragraph index. Use when the document is too large to fit in context and you need to page through content. Returns paragraph text with styles. Start from index 0 and advance by 20 to page through the full document.",
      parameters: readSectionParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("office_read_section", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("verify_doc"))
    pi.registerTool({
      name: "verify_doc",
      label: "Verify Word Document",
      description:
        "Collect a non-mutating, structured Word verification context with summary text and detailed anchors/snippets for document checks.",
      parameters: verifyDocParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("verify_doc", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("verify_doc_visual"))
    pi.registerTool({
      name: "verify_doc_visual",
      label: "Verify Word Visual",
      description:
        "Capture non-mutating Word visual verification context through the supported viewport path. Word-only; returns structured visual/details payloads.",
      parameters: verifyDocVisualParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("verify_doc_visual", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("get_cell_ranges"))
    pi.registerTool({
      name: "get_cell_ranges",
      label: "Get Cell Ranges",
      description:
        "Excel-only first-class range read tool for cell/range values, text, formulas, and number formats.",
      parameters: getCellRangesParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("get_cell_ranges", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("set_cell_range"))
    pi.registerTool({
      name: "set_cell_range",
      label: "Set Cell Range",
      description:
        "Excel-only first-class range write tool for setting values in a target cell/range.",
      parameters: setCellRangeParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("set_cell_range", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("clear_cell_range"))
    pi.registerTool({
      name: "clear_cell_range",
      label: "Clear Cell Range",
      description:
        "Excel-only first-class range clear tool. Destructive clears should set confirmDestructive=true.",
      parameters: clearCellRangeParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("clear_cell_range", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("resize_range"))
    pi.registerTool({
      name: "resize_range",
      label: "Resize Range",
      description:
        "Excel-only first-class range layout tool for computing/activating resized ranges by count or delta.",
      parameters: resizeRangeParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("resize_range", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("copy_to"))
    pi.registerTool({
      name: "copy_to",
      label: "Copy To Range",
      description:
        "Excel-only first-class range copy tool that copies a source range into a destination range.",
      parameters: copyToParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("copy_to", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("modify_sheet_structure"))
    pi.registerTool({
      name: "modify_sheet_structure",
      label: "Modify Sheet Structure",
      description:
        "Excel-only first-class worksheet structure tool for create, rename, duplicate, and delete operations.",
      parameters: modifySheetStructureParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("modify_sheet_structure", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("modify_object"))
    pi.registerTool({
      name: "modify_object",
      label: "Modify Excel Object",
      description:
        "Excel-only first-class object mutation tool for table/chart/pivot/worksheet object operations through native workbook actions.",
      parameters: modifyObjectParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("modify_object", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("get_all_objects"))
    pi.registerTool({
      name: "get_all_objects",
      label: "Get Excel Objects",
      description:
        "Excel-only first-class object inventory read for workbook/worksheet tables, charts, PivotTables, and named items.",
      parameters: getAllObjectsParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("get_all_objects", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("search_data"))
    pi.registerTool({
      name: "search_data",
      label: "Search Excel Data",
      description:
        "Excel-only first-class workbook/worksheet data search across tables, charts, PivotTables, named items, and cited cells.",
      parameters: searchDataParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("search_data", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("get_range_as_csv"))
    pi.registerTool({
      name: "get_range_as_csv",
      label: "Export Range as CSV",
      description:
        "Excel-only first-class CSV export for auditable range snapshots. Use includeFormulas=true when formula-first verification is required.",
      parameters: getRangeAsCsvParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("get_range_as_csv", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("read_range_image"))
    pi.registerTool({
      name: "read_range_image",
      label: "Read Range Image",
      description:
        "Excel-only active-selection visual snapshot using Office.js image coercion. Select or navigate to the target range first; this does not render arbitrary offscreen ranges by address.",
      parameters: readRangeImageParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("read_range_image", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("extract_chart_xml"))
    pi.registerTool({
      name: "extract_chart_xml",
      label: "Extract Chart XML",
      description:
        "Excel-only first-class chart XML extraction that returns a runtime-generated chart metadata XML snapshot (not full package OOXML).",
      parameters: extractChartXmlParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("extract_chart_xml", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("get_presentation_structure"))
    pi.registerTool({
      name: "get_presentation_structure",
      label: "Read Presentation Structure",
      description:
        "PowerPoint-only first-class presentation structure read. Returns slide order plus layout/master structure metadata and bounded slide previews.",
      parameters: getPresentationStructureParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("get_presentation_structure", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("get_slide"))
    pi.registerTool({
      name: "get_slide",
      label: "Read Slide",
      description:
        "PowerPoint-only first-class per-slide read. Resolve a slide by slideId/slideIndex (or selection) and return structured slide details.",
      parameters: getSlideParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("get_slide", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("list_slide_shapes"))
    pi.registerTool({
      name: "list_slide_shapes",
      label: "List Slide Shapes",
      description:
        "PowerPoint-only first-class shape inventory read. Returns structured shape summaries for the resolved slide.",
      parameters: listSlideShapesParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("list_slide_shapes", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("modify_presentation_structure"))
    pi.registerTool({
      name: "modify_presentation_structure",
      label: "Modify Presentation Structure",
      description:
        "PowerPoint-only first-class structure mutation tool for slide create/move/reorder/delete/layout operations through native host actions.",
      parameters: modifyPresentationStructureParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("modify_presentation_structure", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("duplicate_slide"))
    pi.registerTool({
      name: "duplicate_slide",
      label: "Duplicate Slide",
      description:
        "PowerPoint-only first-class slide duplication tool supporting one or multiple source slides and optional insertion target/formatting controls.",
      parameters: duplicateSlideParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("duplicate_slide", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("insert_slide_element"))
    pi.registerTool({
      name: "insert_slide_element",
      label: "Insert Slide Element",
      description:
        "PowerPoint-only first-class element insertion tool. Use for explicit shape/table/diagram/picture insertion through native PowerPoint actions.",
      parameters: insertSlideElementParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("insert_slide_element", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("remove_slide_element"))
    pi.registerTool({
      name: "remove_slide_element",
      label: "Remove Slide Element",
      description:
        "PowerPoint-only first-class element removal tool. Supports explicit shape/text removal operations and preserves destructive-action confirmation checks.",
      parameters: removeSlideElementParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("remove_slide_element", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("edit_slide_text"))
    pi.registerTool({
      name: "edit_slide_text",
      label: "Edit Slide Text",
      description:
        "PowerPoint-only first-class text editing for slide shapes/selection. Use this instead of generic office_apply_edit when the intent is text-focused slide authoring.",
      parameters: editSlideTextParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("edit_slide_text", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("edit_slide_xml"))
    pi.registerTool({
      name: "edit_slide_xml",
      label: "Edit Slide XML",
      description:
        "PowerPoint-only first-class serialized/XML editing tool for slide notes and package-level OOXML workflows (including base64 import/export paths).",
      parameters: editSlideXmlParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("edit_slide_xml", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("edit_slide_master"))
    pi.registerTool({
      name: "edit_slide_master",
      label: "Apply Slide Layout",
      description:
        "PowerPoint-only legacy-named tool for applying an existing slide layout. It does not mutate slide masters or layout definitions.",
      parameters: editSlideMasterParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("edit_slide_master", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("edit_slide_chart"))
    pi.registerTool({
      name: "edit_slide_chart",
      label: "Edit Slide Chart",
      description:
        "PowerPoint-only first-class chart workflow tool for chart inspection and serialized chart create/update paths.",
      parameters: editSlideChartParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("edit_slide_chart", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("copy_image_between_slides"))
    pi.registerTool({
      name: "copy_image_between_slides",
      label: "Copy Image Between Slides",
      description:
        "PowerPoint-only first-class media workflow tool that copies an image from a source slide/shape to a destination slide or shape.",
      parameters: copyImageBetweenSlidesParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("copy_image_between_slides", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("search_icons"))
    pi.registerTool({
      name: "search_icons",
      label: "Search Slide Icons",
      description:
        "PowerPoint-only first-class icon search. Returns icon matches from the supported runtime icon catalog without mutating slides.",
      parameters: searchIconsParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("search_icons", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("insert_icon"))
    pi.registerTool({
      name: "insert_icon",
      label: "Insert Slide Icon",
      description:
        "PowerPoint-only first-class icon insertion tool that inserts or updates an icon-like visual on the target slide.",
      parameters: insertIconParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("insert_icon", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("verify_slides"))
    pi.registerTool({
      name: "verify_slides",
      label: "Verify Slides",
      description:
        "PowerPoint-only first-class structural verification. Returns a non-mutating summary/details payload for slide/layout/master checks.",
      parameters: verifySlidesParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("verify_slides", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("verify_slide_visual"))
    pi.registerTool({
      name: "verify_slide_visual",
      label: "Verify Slide Visual",
      description:
        "PowerPoint-only first-class visual verification using supported Office.js slide/shape snapshot paths (not slideshow-frame capture).",
      parameters: verifySlideVisualParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("verify_slide_visual", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("office_execute_js"))
    pi.registerTool({
      name: "office_execute_js",
      label: "Execute Office.js",
      description:
        "Execute Office.js code directly in the active host as an escape hatch when structured tools cannot achieve the desired result. The code MUST use Word.run, Excel.run, or PowerPoint.run and return a JSON-serializable result. This tool is a best-effort restricted subset enforced by regex checks (not an isolated sandbox) and blocks network, storage, eval, and system-access patterns. Do not use it for simple edits that office_apply_edit can handle.",
      parameters: executeJsParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("office_execute_js", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("edit_doc_list"))
    pi.registerTool({
      name: "edit_doc_list",
      label: "Edit Word List",
      description:
        `Word-only first-class reviewable list editing. Use for list-like rewrites or multi-item legal edits that should be reviewed before apply. Keep each searchText under ${OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH} characters and include paragraphId/anchor locators when available.`,
      parameters: proposeEditsParams,
      execute: async (_toolCallId, params) => executeReviewableWordEdits("edit_doc_list", params),
    });

    if (!isDisabled("office_propose_edits"))
    pi.registerTool({
      name: "office_propose_edits",
      label: "Propose Document Edits",
      description:
        `Propose a batch of text edits to the active Word document for the user to review before applying. Each edit specifies a kind (insert/replace/delete), the text to find (searchText), and the replacement. The user sees a reviewable card for each edit and can accept, modify, or reject individually. CRITICAL: each edit's searchText MUST be under ${OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH} characters. Break large paragraph rewrites into multiple small, targeted edits — one per sentence or distinct phrase. Never use a full paragraph as searchText. For example, instead of one edit replacing a 3-sentence paragraph, create 3 separate edits each targeting one sentence. Use this instead of office_apply_edit when making multi-paragraph changes or when the user should verify changes first.`,
      parameters: proposeEditsParams,
      execute: async (_toolCallId, params) => executeReviewableWordEdits("office_propose_edits", params),
    });

    if (!isDisabled("ask_user"))
    pi.registerTool({
      name: "ask_user",
      label: "Ask User",
      description:
        "Ask the user one or more questions with selectable options. Use this proactively when the task involves subjective choices, ambiguous intent, or assumptions that would materially change the result (e.g. tone, audience, format, scope, style). After receiving the user's answers, immediately proceed to execute the original task using those answers — do not stop at acknowledging the choices. Supports single or multiple questions in one call.",
      parameters: askUserParams,
      execute: async (_toolCallId, params) => {
        const typed = params as {
          question?: string;
          context?: string;
          options?: (string | { title: string; description?: string })[];
          questions?: { question: string; context?: string; options?: (string | { title: string; description?: string })[] }[];
        };

        const normalizeOptions = (raw?: (string | { title: string; description?: string })[]) =>
          (raw ?? []).map((o) => (typeof o === "string" ? { title: o } : o));

        const questions: AskUserQuestion[] = [];
        if (typed.questions?.length) {
          for (const q of typed.questions) {
            questions.push({
              id: crypto.randomUUID(),
              question: q.question,
              context: q.context,
              options: normalizeOptions(q.options),
            });
          }
        } else if (typed.question) {
          questions.push({
            id: crypto.randomUUID(),
            question: typed.question,
            context: typed.context,
            options: normalizeOptions(typed.options),
          });
        } else {
          return {
            content: [{ type: "text" as const, text: "No question provided." }],
            details: { error: "missing_question" } as Record<string, unknown>,
          };
        }

        const request: AskUserRequest = { requestId: crypto.randomUUID(), questions };
        const response = await options.invokeAskUser(request);

        const lines = response.answers.map((a) => {
          const q = questions.find((q) => q.id === a.questionId);
          const answer = a.selectedOption ?? "None of the above";
          const notes = a.notes ? ` (notes: ${a.notes})` : "";
          return `Q: ${q?.question ?? "?"}\nA: ${answer}${notes}`;
        });

        return {
          content: [{ type: "text" as const, text: lines.join("\n\n") }],
          details: { questions, answers: response.answers } as Record<string, unknown>,
        };
      },
    });

    if (options.generateImage && !isDisabled("generate_image")) {
      const generateImage = options.generateImage;
      const isEnabled = options.isImageGenerationEnabled ?? (() => true);
      pi.registerTool({
        name: "generate_image",
        label: "Generate Image",
        description:
          "Generate an AI image from a text prompt and optionally insert it into the active Office document. Use for photos, illustrations, graphics, and visual content. Do not use for diagrams or charts.",
        parameters: generateImageParams,
        execute: async (_toolCallId, params) => {
          if (!isEnabled()) {
            return {
              content: [{ type: "text" as const, text: "Image generation is disabled. The user can enable it in Settings > Preferences > Image Generation." }],
              details: { enabled: false } as Record<string, unknown>,
            };
          }

          const typedParams = params as {
            prompt: string;
            aspectRatio?: string;
            size?: string;
            quality?: string;
            insert?: boolean;
          };

          const result = await generateImage({
            prompt: typedParams.prompt,
            aspectRatio: typedParams.aspectRatio,
            size: typedParams.size,
            quality: typedParams.quality,
            insert: typedParams.insert,
          });

          const shouldInsert = typedParams.insert !== false;
          if (shouldInsert) {
            try {
              await options.invokeTool("office_apply_edit", {
                action: {
                  type: "insertInlinePicture",
                  content: result.base64,
                  placement: "after",
                  options: { altText: typedParams.prompt.slice(0, 120) },
                },
              } as Record<string, unknown>);
            } catch (insertError) {
              const msg = insertError instanceof Error ? insertError.message : String(insertError);
              return {
                content: [
                  { type: "text" as const, text: `Image generated with ${result.modelName} (${result.width}x${result.height}) but insertion failed: ${msg}. The image is shown below.` },
                  { type: "image" as const, data: result.base64, mimeType: result.mimeType },
                ],
                details: { ...result, inserted: false, insertError: msg } as Record<string, unknown>,
              };
            }
          }

          return {
            content: [
              {
                type: "text" as const,
                text: `Image generated with ${result.modelName} (${result.width}x${result.height}).${shouldInsert ? " Inserted into document." : ""}`,
              },
              { type: "image" as const, data: result.base64, mimeType: result.mimeType },
            ],
            details: { ...result, inserted: shouldInsert } as Record<string, unknown>,
          };
        },
      });
    }

    pi.on("session_start", async (_event, ctx) => {
      const state = options.getState();
      const label = HOST_LABELS[options.getHost()];
      const suffix = state?.document.saved ? "saved document" : "unsaved document";
      ctx.ui.setStatus("office-host", `${label} · ${suffix}`);
    });
  };
}
