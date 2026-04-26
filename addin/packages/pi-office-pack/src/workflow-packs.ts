import type { OfficeHost } from "./protocol.js";

export type WorkflowPackId =
  | "word-research-paper"
  | "word-resume-polish"
  | "word-spec-review"
  | "word-legal-professional-review"
  | "word-business-user-stories"
  | "excel-dcf-review"
  | "excel-formula-audit"
  | "excel-table-chart-improvement"
  | "excel-narrative-export"
  | "powerpoint-pitch-deck-outline"
  | "powerpoint-slide-polish"
  | "powerpoint-visual-consistency"
  | "powerpoint-speaker-notes"
  | "powerpoint-data-backed-slides";

export interface OfficeWorkflowPack {
  id: WorkflowPackId;
  host: OfficeHost;
  title: string;
  userPrompt: string;
  intent: string;
  requiredContext: string[];
  preferredTools: string[];
  reviewGates: string[];
  completionChecks: string[];
}

export const OFFICE_WORKFLOW_PACKS: readonly OfficeWorkflowPack[] = [
  {
    id: "word-research-paper",
    host: "word",
    title: "Research Paper Review",
    userPrompt: "Run the research paper workflow on this document.",
    intent: "Improve academic structure, argument flow, citation placeholders, and section-level clarity without inventing sources.",
    requiredContext: ["document outline", "current selection", "heading structure", "citation or reference markers"],
    preferredTools: ["office_get_context", "office_read_section", "verify_doc", "office_propose_edits"],
    reviewGates: ["Ask before changing citation style or adding claims that need sources.", "Use reviewable edits for paragraph rewrites."],
    completionChecks: ["Summarize unresolved evidence gaps.", "Confirm edited passages were targeted with anchors or paragraph IDs."],
  },
  {
    id: "word-resume-polish",
    host: "word",
    title: "Resume Polish",
    userPrompt: "Polish this resume for a sharper professional story.",
    intent: "Sharpen impact bullets, role summaries, tense consistency, and ATS-friendly wording while preserving factual claims.",
    requiredContext: ["selection or full resume sections", "role headings", "bullet structure", "target role when available"],
    preferredTools: ["office_get_context", "edit_doc_list", "office_propose_edits", "verify_doc_visual"],
    reviewGates: ["Ask before inventing metrics, employers, dates, or credentials.", "Use reviewable list edits for bullets."],
    completionChecks: ["Call out bullets that still need user-provided metrics.", "Verify layout-sensitive edits when spacing or page fit matters."],
  },
  {
    id: "word-spec-review",
    host: "word",
    title: "Spec Review",
    userPrompt: "Review this spec and turn gaps into action items.",
    intent: "Find ambiguous requirements, missing acceptance criteria, edge cases, dependencies, and implementation risks.",
    requiredContext: ["section headings", "requirements language", "tables or lists", "current selection"],
    preferredTools: ["office_get_context", "office_read_section", "verify_doc", "office_propose_edits"],
    reviewGates: ["Ask before rewriting scope or changing requirement priority.", "Keep recommendations traceable to source sections."],
    completionChecks: ["Return gap list with section anchors.", "Offer reviewable edits for wording changes instead of silent rewrites."],
  },
  {
    id: "word-legal-professional-review",
    host: "word",
    title: "Legal and Professional Review",
    userPrompt: "Review this section for legal or professional risk.",
    intent: "Flag risky wording, inconsistent obligations, undefined terms, tone issues, and review-sensitive edits.",
    requiredContext: ["current selection", "nearby clauses", "comments or tracked-change context", "defined terms when visible"],
    preferredTools: ["office_get_context", "verify_doc", "edit_doc_list", "office_propose_edits"],
    reviewGates: ["Do not present legal advice as final judgment.", "Use reviewable edits for contract-like or policy text."],
    completionChecks: ["Separate risk flags from proposed language.", "Include anchors or paragraph IDs for every proposed edit."],
  },
  {
    id: "word-business-user-stories",
    host: "word",
    title: "Business User Stories",
    userPrompt: "Convert this content into business user stories with acceptance criteria.",
    intent: "Turn business notes or requirements into clear stories, acceptance criteria, open questions, and traceable assumptions.",
    requiredContext: ["source requirements", "audience or product area", "tables and bullet lists", "selected scope"],
    preferredTools: ["office_get_context", "office_read_section", "office_apply_edit", "office_propose_edits"],
    reviewGates: ["Ask before changing persona, priority, or release scope.", "Preserve source assumptions visibly."],
    completionChecks: ["Produce stories with acceptance criteria and open questions.", "Confirm inserted content lands in the intended section."],
  },
  {
    id: "excel-dcf-review",
    host: "excel",
    title: "DCF Review",
    userPrompt: "Run a DCF review on the active workbook.",
    intent: "Review valuation assumptions, forecast logic, terminal value mechanics, discount-rate inputs, and sensitivity outputs.",
    requiredContext: ["workbook object inventory", "model sheets", "assumption ranges", "formula ranges", "charts or sensitivity tables"],
    preferredTools: ["office_get_context", "get_all_objects", "search_data", "get_cell_ranges", "get_range_as_csv"],
    reviewGates: ["Do not overwrite model formulas without user approval.", "Ask before changing assumptions or valuation methodology."],
    completionChecks: ["Cite sheets and cell ranges for each finding.", "Include formula-level evidence where auditability matters."],
  },
  {
    id: "excel-formula-audit",
    host: "excel",
    title: "Formula Audit",
    userPrompt: "Audit formulas and dependencies in the current sheet.",
    intent: "Find formula inconsistencies, hard-coded cells, broken references, suspicious dependencies, and range mismatch risks.",
    requiredContext: ["selected range or sheet", "formulas", "number formats", "named ranges", "tables or PivotTables"],
    preferredTools: ["office_get_context", "get_cell_ranges", "search_data", "get_range_as_csv"],
    reviewGates: ["Ask before modifying formulas or clearing values.", "Preserve explicit cell references in findings."],
    completionChecks: ["Return issues grouped by severity with cell addresses.", "Suggest targeted fixes separately from audit findings."],
  },
  {
    id: "excel-table-chart-improvement",
    host: "excel",
    title: "Table and Chart Improvement",
    userPrompt: "Improve the selected table or chart for presentation.",
    intent: "Make workbook tables, charts, filters, validations, and formatting clearer for presentation or analysis.",
    requiredContext: ["active selection", "table or chart inventory", "formatting metadata", "visual snapshot when selected"],
    preferredTools: ["office_get_context", "get_all_objects", "read_range_image", "extract_chart_xml", "modify_object"],
    reviewGates: ["Ask before destructive filters, clears, or chart type changes.", "Use active-selection visual checks for layout-sensitive changes."],
    completionChecks: ["Verify the selected range or object after mutation.", "Report any visual limitation if only metadata was available."],
  },
  {
    id: "excel-narrative-export",
    host: "excel",
    title: "Narrative Export",
    userPrompt: "Turn this workbook analysis into an executive narrative.",
    intent: "Convert selected workbook evidence into a concise narrative with cited ranges, assumptions, and recommended next steps.",
    requiredContext: ["source ranges", "tables or charts", "key assumptions", "workbook object names"],
    preferredTools: ["office_get_context", "get_all_objects", "get_range_as_csv", "extract_chart_xml"],
    reviewGates: ["Do not imply cross-host insertion unless the active host supports it.", "Ask before summarizing hidden or unavailable sheets."],
    completionChecks: ["Cite every quantitative claim to a range or object.", "Separate observations, assumptions, and recommendations."],
  },
  {
    id: "powerpoint-pitch-deck-outline",
    host: "powerpoint",
    title: "Pitch Deck Outline",
    userPrompt: "Run the pitch deck workflow for this presentation.",
    intent: "Assess deck narrative, slide order, missing investor-story beats, and concise slide-level next actions.",
    requiredContext: ["presentation structure", "slide titles", "speaker notes when relevant", "selected slide"],
    preferredTools: ["get_presentation_structure", "get_slide", "verify_slides", "modify_presentation_structure"],
    reviewGates: ["Ask before deleting, reordering, or adding many slides.", "Keep proposed narrative changes slide-specific."],
    completionChecks: ["Return slide-by-slide recommendations.", "Verify structure after any slide-order changes."],
  },
  {
    id: "powerpoint-slide-polish",
    host: "powerpoint",
    title: "Slide Polish",
    userPrompt: "Polish the current slide for clarity and executive tone.",
    intent: "Improve slide text, information hierarchy, element choice, and executive readability on selected slides.",
    requiredContext: ["selected slide", "shape inventory", "slide text", "visual snapshot"],
    preferredTools: ["get_slide", "list_slide_shapes", "verify_slide_visual", "edit_slide_text", "insert_slide_element"],
    reviewGates: ["Ask before replacing a major visual or changing the slide message.", "Use structural tools before XML escape paths."],
    completionChecks: ["Verify slide text and visual snapshot after edits.", "Summarize any manual design checks still needed."],
  },
  {
    id: "powerpoint-visual-consistency",
    host: "powerpoint",
    title: "Visual Consistency",
    userPrompt: "Check this deck for visual consistency issues.",
    intent: "Find inconsistent layouts, fonts, spacing, icons, chart styles, slide structure, and theme drift.",
    requiredContext: ["presentation structure", "layout metadata", "shape summaries", "visual snapshots"],
    preferredTools: ["get_presentation_structure", "verify_slides", "verify_slide_visual", "search_icons", "insert_icon"],
    reviewGates: ["Ask before applying layout changes across multiple slides.", "Do not claim pixel-perfect verification when only Office.js snapshots are available."],
    completionChecks: ["Group issues by slide.", "Verify changed slides with structural and visual tools."],
  },
  {
    id: "powerpoint-speaker-notes",
    host: "powerpoint",
    title: "Speaker Notes",
    userPrompt: "Add speaker notes for the selected slides.",
    intent: "Create concise presenter notes that match slide text, audience, timing, and narrative arc.",
    requiredContext: ["slide content", "existing notes", "target audience or duration", "selected slide range"],
    preferredTools: ["get_slide", "get_presentation_structure", "edit_slide_xml", "verify_slides"],
    reviewGates: ["Ask when audience, tone, or talk length materially changes the notes.", "Keep notes aligned to visible slide claims."],
    completionChecks: ["Verify notes through the serialized notes path.", "Summarize slides that still need audience-specific detail."],
  },
  {
    id: "powerpoint-data-backed-slides",
    host: "powerpoint",
    title: "Data-Backed Slides",
    userPrompt: "Create or improve data-backed slides from the available content.",
    intent: "Build chart/table-oriented slides with clear claims, cited data inputs, and reviewable visual structure.",
    requiredContext: ["source slide or provided data", "chart requirements", "theme/layout metadata", "speaker narrative"],
    preferredTools: ["get_presentation_structure", "edit_slide_chart", "insert_slide_element", "verify_slide_visual"],
    reviewGates: ["Ask before inventing data or implying unavailable Excel access.", "Use chart tools instead of image-only charts when possible."],
    completionChecks: ["Verify chart payload and slide snapshot after edits.", "Cite source data or state that user-supplied data is needed."],
  },
] as const;

export function getWorkflowPacksForHost(host: OfficeHost): readonly OfficeWorkflowPack[] {
  return OFFICE_WORKFLOW_PACKS.filter((pack) => pack.host === host);
}

export function getWorkflowQuickPrompts(host: OfficeHost | undefined, maxCount = 3): string[] {
  if (!host) return [];
  return getWorkflowPacksForHost(host).slice(0, maxCount).map((pack) => pack.userPrompt);
}

export function formatWorkflowPackGuidance(host?: OfficeHost): string {
  const packs = host ? getWorkflowPacksForHost(host) : OFFICE_WORKFLOW_PACKS;
  const grouped = new Map<OfficeHost, OfficeWorkflowPack[]>();
  for (const pack of packs) {
    const entries = grouped.get(pack.host) ?? [];
    entries.push(pack);
    grouped.set(pack.host, entries);
  }

  const lines = ["## Professional Workflow Packs"];
  for (const [packHost, hostPacks] of grouped) {
    lines.push("", `${packHost}:`);
    for (const pack of hostPacks) {
      lines.push(
        `- ${pack.title}: intent=${pack.intent}; required context=${pack.requiredContext.join(", ")}; preferred tools=${pack.preferredTools.join(", ")}; review gates=${pack.reviewGates.join(" | ")}; completion checks=${pack.completionChecks.join(" | ")}.`,
      );
    }
  }
  return lines.join("\n");
}
