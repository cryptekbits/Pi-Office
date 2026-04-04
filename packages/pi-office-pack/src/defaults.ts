import type { OfficeHost, OfficeMode, OfficeStateUpdate, UserPreferences, ToolCategory } from "./protocol.js";
import {
  AUTONOMY_LEVEL_AUTO_APPROVE,
  AUTONOMY_LEVEL_LABELS,
  TOOL_CATEGORY_MAP,
  OFFICE_TOOL_NAMES,
  OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH,
  type AutonomyLevel,
} from "./protocol.js";

export const DEFAULT_COMPANION_PORT = 3443;
export const DEFAULT_COMPANION_HOST = "localhost";
export const SHORTCUT_ACTION_ID = "ShowTaskpane";
export const SHORTCUT_DEFAULT_KEY = "Ctrl+Alt+P";

export const HOST_LABELS: Record<OfficeHost, string> = {
  word: "Word",
  excel: "Excel",
  powerpoint: "PowerPoint",
};

export const OFFICE_APPEND_SYSTEM_PROMPT = `
You are running inside a Microsoft Office add-in backed by Pi.

Use native Office tools whenever the task is about reading or changing the active document.
Do not invent document state. If exact wording, table values, or slide content matters, call office_get_context first.
If visual layout, images, charts, spacing, margins, tabs, ruler-level formatting, or slide styling matter, call office_capture_snapshot and office_get_context before answering.
If the user is asking about what is currently visible in Word, or about alignment, page breaks, wrapping, clipping, margins, header/footer placement, page position, or any other viewport-dependent issue, call office_capture_viewport proactively.
Use office_capture_viewport only for Word. It returns Office.js viewport metadata and context-derived visuals; it is not a pixel-perfect OS/window screenshot and does not capture off-screen pages.
Prefer targeted edits to the current selection instead of rewriting an entire document unless the user clearly wants that.
For direct Word clause/sentence updates, use edit_doc_text first so edits route through native Word actions.
For Word list rewrites, legal-review-sensitive edits, or tracked-changes-heavy passages, use edit_doc_list (or office_propose_edits) so each change is reviewable before apply.
When using edit_doc_list or office_propose_edits, keep every searchText under ${OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH} characters and include paragraphId or anchor locators whenever available for deterministic targeting.
office_execute_js is a best-effort restricted subset enforced with regex checks (not an isolated sandbox). It blocks network, storage, eval, and system-access patterns and should only be used as an escape hatch when structured tools are insufficient.
When a task involves subjective choices (tone, audience, format, scope, style) or the request is ambiguous enough that different interpretations would produce materially different results, use ask_user to clarify before proceeding. Do not guess — ask. After receiving the user's answers from ask_user, immediately carry out the full task using those answers in the same turn. Never stop after merely acknowledging the user's choices.
The taskpane chat renders Mermaid and Draw.io diagrams inline. When the user asks for a diagram, flowchart, sequence diagram, or visual aid, prefer returning a fenced code block tagged with mermaid or drawio so the taskpane can render it and offer insertion into the document.

DRAW.IO XML RULES (critical for correct rendering):
- Output must be well-formed mxGraphModel XML. NEVER include XML comments (<!-- -->).
- Basic structure: <mxGraphModel adaptiveColors="auto"><root><mxCell id="0"/><mxCell id="1" parent="0"/><!-- cells here --></root></mxGraphModel>
- Every edge mxCell MUST contain a child <mxGeometry relative="1" as="geometry"/> element. Self-closing edge cells are invalid and will not render.
- Use edgeStyle=orthogonalEdgeStyle for right-angle connectors. Space nodes at least 60px apart (prefer 200px horizontal, 120px vertical).
- Use exitX/exitY/entryX/entryY (0-1) to control edge connection sides. Ensure at least 20px straight segment before targets for arrowhead clearance.
- Add explicit waypoints via <Array as="points"><mxPoint x="..." y="..."/></Array> inside mxGeometry when edges would overlap.
- Do NOT wrap edge labels in HTML markup. Just set the value attribute directly. Do NOT use <br> or any HTML tags in value attributes -- use plain text only.
- Escape special characters in attribute values: &amp; &lt; &gt; &quot;. Never put raw < > & in attribute values.
- Containers: use parent-child containment (parent="containerId") with relative coordinates. Use group style for invisible containers, swimlane for titled containers. Add pointerEvents=0 to containers that should not capture child connections.
- Align nodes to a grid (multiples of 10). Use consistent spacing.
- For dark mode adaptation, include adaptiveColors="auto" on mxGraphModel. Colors default to "default" (black in light, white in dark). Explicit colors specify light-mode; dark-mode is auto-inverted.
- Use rounded=1 and whiteSpace=wrap on most shapes for cleaner appearance.
- Every mxCell must have a unique id. Use descriptive ids when possible.
When the user asks for an image, illustration, photo, graphic, or visual content (not a diagram or chart), use the generate_image tool. If image generation is disabled, the tool will inform you and you should tell the user to enable it in Settings > Preferences > Image Generation. Consider the document type when choosing aspect ratio: use 16:9 for PowerPoint slides, 4:3 or 1:1 for Word documents. For slide backgrounds or hero images, request higher resolution. Do not use generate_image for diagrams or charts -- use Mermaid or Draw.io for those.
When the document is unsaved, do not assume local file access is available. Once the document is saved, workspace tools may become available through Pi and the document folder becomes the working directory.
When you use filesystem tools, keep them focused on the saved document's workspace and treat AGENTS.md and SKILL.md files as live guidance.
`;

export function getOfficeMode(saved: boolean): OfficeMode {
  return saved ? "workspace" : "document-only";
}

function formatSelectionMetaLine(meta: import("./protocol.js").OfficeSelectionMeta | undefined): string | undefined {
  if (!meta) return undefined;
  const parts: string[] = [];
  if (meta.paragraphCount) parts.push(`${meta.paragraphCount} paragraph${meta.paragraphCount === 1 ? "" : "s"}`);
  if (meta.styleHistogram) {
    const styles = Object.entries(meta.styleHistogram)
      .map(([style, count]) => `${style}(${count})`)
      .join(", ");
    if (styles) parts.push(`styles: ${styles}`);
  }
  if (meta.isListItem) parts.push("list content");
  if (meta.isInTable) parts.push("in table");
  if (meta.fontSummary) {
    const f = meta.fontSummary;
    const fontParts = [f.name, f.size ? `${f.size}pt` : undefined, f.bold ? "bold" : undefined, f.italic ? "italic" : undefined].filter(Boolean);
    if (fontParts.length) parts.push(`font: ${fontParts.join(" ")}`);
  }
  return parts.length ? `Selection metadata: ${parts.join(" | ")}` : undefined;
}

export function summarizeOfficeState(state: OfficeStateUpdate | undefined): string {
  if (!state) {
    return "Office host state is not available yet. Ask for context with office_get_context before making assumptions.";
  }

  const pathLine = state.document.documentPath ? `Path: ${state.document.documentPath}` : "Path: unsaved document";
  const structured = state.selection.structuredPreview?.trim();
  const preview = state.selection.textPreview?.trim();
  const previewLine = structured
    ? `Selected content:\n${structured}`
    : preview
      ? `Selection preview: ${preview}`
      : "Selection preview: none";
  const metaLine = formatSelectionMetaLine(state.selection.selectionMeta);
  const detailLines = state.selection.details?.length
    ? `Selection details: ${state.selection.details.join(" | ")}`
    : undefined;

  return [
    `Host: ${HOST_LABELS[state.host]}`,
    `Document: ${state.document.title}`,
    `Mode: ${getOfficeMode(state.document.saved)}`,
    pathLine,
    `Selection: ${state.selection.label}`,
    previewLine,
    metaLine,
    detailLines,
    `Capabilities: ${state.capabilities.join(", ") || "none reported"}`,
  ]
    .filter(Boolean)
    .join("\n");
}

export function composeOfficeAwarePrompt(input: string, state: OfficeStateUpdate | undefined): string {
  return `Office session state:\n${summarizeOfficeState(state)}\n\nUser request:\n${input.trim()}`;
}

const WORKSPACE_TOOL_NAMES = ["read", "grep", "find", "ls", "edit", "write", "bash"] as const;
const CATEGORY_LABELS: Record<string, string> = {
  read: "read",
  "write-doc": "doc-write",
  "read-external": "workspace-read",
  "write-external": "workspace-write",
  connector: "connector",
  interaction: "interaction",
};

export function composeAutonomyPrompt(
  preferences: UserPreferences,
  isWorkspace: boolean,
  availableToolNames?: readonly string[],
): string {
  const level = preferences.autonomyLevel;
  const overrides = preferences.toolPermissionOverrides ?? [];

  const allToolNames: string[] =
    availableToolNames?.length
      ? [...new Set(availableToolNames.map((name) => String(name)))]
      : [
          ...OFFICE_TOOL_NAMES,
          ...(isWorkspace ? WORKSPACE_TOOL_NAMES : []),
        ];

  const autoApproved: string[] = [];
  const requireApproval: string[] = [];
  const disabled: string[] = [];

  for (const toolName of allToolNames) {
    const override = overrides.find((o) => o.toolName === toolName);
    if (override?.autoApproveAtLevel === "disabled") {
      disabled.push(toolName);
      continue;
    }
    const category = (TOOL_CATEGORY_MAP[toolName] ?? "connector") as ToolCategory;
    const catLabel = CATEGORY_LABELS[category] ?? category;
    const effectiveLevel: AutonomyLevel = override ? override.autoApproveAtLevel as AutonomyLevel : level;
    const effectiveApproved = AUTONOMY_LEVEL_AUTO_APPROVE[effectiveLevel];
    if (effectiveApproved.has(category)) {
      autoApproved.push(`${toolName} (${catLabel})`);
    } else {
      requireApproval.push(`${toolName} (${catLabel})`);
    }
  }

  const lines = [
    `## Tool Autonomy`,
    ``,
    `Current autonomy level: ${AUTONOMY_LEVEL_LABELS[level]}`,
    `Auto-approved tools: ${autoApproved.join(", ") || "none"}`,
    `Require user approval: ${requireApproval.join(", ") || "none"}`,
    `Disabled (unavailable): ${disabled.join(", ") || "none"}`,
    ``,
    `Rules:`,
    `- Use auto-approved tools freely to accomplish tasks.`,
    `- For tools requiring approval: proceed with the call — the user will be prompted to approve or deny. Do NOT ask permission in chat before calling them.`,
    `- If a tool call is denied by the user, adapt your approach using the remaining available tools. If no alternative exists, explain what you cannot do and why.`,
    `- NEVER ask the user to change the autonomy level. If a needed tool is unavailable or denied, simply explain the limitation.`,
    `- Disabled tools do not exist in this session. Do not reference or attempt to call them.`,
  ];

  return lines.join("\n");
}
