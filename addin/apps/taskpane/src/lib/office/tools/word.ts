import { Type } from "@sinclair/typebox";
import { OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH } from "@pi-office/pi-office-pack/protocol";
import type { OfficeToolDefinition } from "./types";

const proposeEditsParams = Type.Object({
  edits: Type.Array(
    Type.Object({
      kind: Type.String({ description: "insert, replace, or delete." }),
      searchText: Type.Optional(Type.String({
        maxLength: OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH,
        description: `Existing text to find. Must be <= ${OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH} characters; split large rewrites into smaller edits.`,
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

export const WORD_OFFICE_TOOL_DEFINITIONS: readonly OfficeToolDefinition[] = [
  {
    name: "edit_doc_text",
    hosts: ["word"],
    category: "write-doc",
    label: "Edit Word Text",
    description: "Word-only first-class text editing. Use for direct clause/sentence updates through native Word actions when no per-edit review card is required.",
    discovery: {
      tier: "specialized",
      capabilityIds: ["word.text", "word.formatting", "word.anchors"],
      keywords: ["word", "text", "clause", "sentence", "native edit", "paragraph", "selection"],
      summary: "Apply direct Word-native text edits to selections or structured anchors.",
      riskLevel: "medium",
      requirementSets: [{ name: "WordApi", minVersion: "1.1" }],
      fallback: "Use office_propose_edits when the user should review each change before apply.",
    },
    parameters: Type.Object({
      mode: Type.Optional(Type.String({ description: "Legacy edit mode such as replaceSelection, insertAfterSelection, or setRangeValues." })),
      content: Type.Optional(Type.String({ description: "Legacy text, HTML, or JSON matrix payload to insert into Office." })),
      text: Type.Optional(Type.String({ description: "Alias for legacy text content. Use when operation/type is insertText." })),
      html: Type.Optional(Type.String({ description: "Alias for legacy HTML content. Use when operation/type is insertHtml." })),
      format: Type.Optional(Type.String({ description: "Legacy content format such as text, html, or matrix." })),
      operation: Type.Optional(Type.String({ description: "Top-level action type alias, e.g., insertText, insertHtml, setRangeValues." })),
      type: Type.Optional(Type.String({ description: "Top-level action type alias when not wrapping with action.type." })),
      values: Type.Optional(Type.Any({ description: "2D array of values for setRangeValues actions." })),
      action: Type.Optional(Type.Any({ description: "Structured OfficeHostAction payload. Use this for precise host-specific actions." })),
    }),
    executor: "office-bridge",
  },
  {
    name: "office_read_section",
    hosts: ["word"],
    category: "read",
    label: "Read Document Section",
    description: "Read a paginated range of Word paragraphs by paragraph index.",
    discovery: {
      tier: "core",
      capabilityIds: ["word.read", "word.sections"],
      keywords: ["word", "read", "paragraph", "section", "page", "document"],
      summary: "Read a bounded paragraph window from a Word document without mutating it.",
      riskLevel: "low",
      requirementSets: [{ name: "WordApi", minVersion: "1.1" }],
    },
    parameters: Type.Object({
      startIndex: Type.Optional(Type.Number({ description: "Zero-based paragraph start index (inclusive)." })),
      endIndex: Type.Optional(Type.Number({ description: "Zero-based paragraph end index (exclusive). Defaults to startIndex + 20." })),
      start: Type.Optional(Type.Number({ description: "Alias for startIndex." })),
      end: Type.Optional(Type.Number({ description: "Alias for endIndex." })),
      includeStyles: Type.Optional(Type.Boolean({ description: "Include paragraph styles and heading levels. Defaults to true." })),
    }),
    executor: "office-bridge",
  },
  {
    name: "verify_doc",
    hosts: ["word"],
    category: "read",
    label: "Verify Word Document",
    description: "Collect a non-mutating, structured Word verification context with summary text and detailed anchors/snippets for document checks.",
    discovery: {
      tier: "core",
      capabilityIds: ["word.verify", "word.read", "word.anchors"],
      keywords: ["word", "verify", "review", "anchors", "comments", "revisions", "fields", "content controls"],
      summary: "Collect non-mutating Word verification context with anchors and formatting metadata.",
      riskLevel: "low",
      requirementSets: [{ name: "WordApi", minVersion: "1.1" }],
    },
    parameters: Type.Object({
      scope: Type.Optional(Type.String({ description: "Verification target for Word context capture (selection or document). Defaults to document-level verification context." })),
      includeFormatting: Type.Optional(Type.Boolean({ description: "Include formatting metadata in the verification payload. Defaults to true." })),
    }),
    executor: "office-bridge",
  },
  {
    name: "verify_doc_visual",
    hosts: ["word"],
    category: "read",
    label: "Verify Word Visual",
    description: "Capture non-mutating Word visual verification context through the supported viewport path. Word-only; returns structured visual/details payloads.",
    discovery: {
      tier: "core",
      capabilityIds: ["word.verify", "word.visual"],
      keywords: ["word", "visual", "layout", "viewport", "formatting", "snapshot"],
      summary: "Capture Word visual/formatting verification through Office.js context and available companion metadata.",
      riskLevel: "low",
      requirementSets: [{ name: "WordApi", minVersion: "1.1" }],
      fallback: "Use office_get_context for text-only verification when visual capture is unavailable.",
    },
    parameters: Type.Object({
      includeFormatting: Type.Optional(Type.Boolean({ description: "Include formatting metadata in the visual payload. Defaults to true." })),
      includeWindowFrame: Type.Optional(Type.Boolean({ description: "Request window-frame metadata when a companion native capture is available." })),
    }),
    executor: "office-bridge",
  },
  {
    name: "edit_doc_list",
    hosts: ["word"],
    category: "write-doc",
    label: "Edit Word List",
    description: `Word-only first-class reviewable list editing. Keep each searchText under ${OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH} characters and include paragraphId/anchor locators when available.`,
    discovery: {
      tier: "specialized",
      capabilityIds: ["word.list", "word.review"],
      keywords: ["word", "list", "bullet", "numbering", "legal", "review", "proposal"],
      summary: "Prepare reviewable Word list edits with deterministic short search targets.",
      riskLevel: "medium",
      requirementSets: [{ name: "WordApi", minVersion: "1.1" }],
      fallback: "Use office_propose_edits for non-list reviewable text edits.",
    },
    parameters: proposeEditsParams,
    executor: "reviewable-word-edit",
    deferred: true,
  },
  {
    name: "office_propose_edits",
    hosts: ["word"],
    category: "write-doc",
    label: "Propose Document Edits",
    description:
      `Propose a batch of text edits for user review before applying changes. ` +
      `CRITICAL: each edit's searchText MUST be under ${OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH} characters. ` +
      "Split large paragraph rewrites into multiple small, targeted edits.",
    discovery: {
      tier: "core",
      capabilityIds: ["word.review", "word.text"],
      keywords: ["word", "proposal", "review", "tracked", "redline", "replace", "delete", "insert"],
      summary: "Show reviewable Word edit cards before applying targeted text changes.",
      riskLevel: "medium",
      requirementSets: [{ name: "WordApi", minVersion: "1.1" }],
    },
    parameters: proposeEditsParams,
    executor: "reviewable-word-edit",
    deferred: true,
  },
];
