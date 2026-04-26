# Office Host

Use this skill when a task depends on the active Microsoft Office document.

## Rules

- Call `office_get_context` before making precise claims about the current selection, worksheet, or slide content.
- For Word layout questions about what is visible on screen, call `office_capture_viewport` proactively. Use it for alignment, page breaks, wrapping, clipping, visible page position, margins, and header/footer placement.
- `office_capture_viewport` is only for Word and returns Office.js viewport metadata plus context-derived visuals. It is not a pixel-perfect OS/window screenshot and should not be used as a substitute for reading off-screen content or reviewing the whole document.
- For large Word documents where context shows only a textPreview (not full text), use `office_read_section` to page through content. Start at index 0 and advance by 20 paragraphs per call. Use headings from `office_get_context` to navigate to relevant sections.
- In Word, use `edit_doc_text` for direct clause/sentence edits that should apply immediately through native Word actions.
- In Word, use `edit_doc_list` (or `office_propose_edits`) for review-sensitive list rewrites, legal-text revisions, or tracked-changes-heavy passages so each edit is reviewable before apply.
- When using `edit_doc_list` or `office_propose_edits`, each edit's `searchText` MUST be under 200 characters. Break large changes into many small, per-sentence or per-phrase edits. Never use a full paragraph as searchText. Include `paragraphId` or `anchor` locators whenever available to keep accepted edits deterministic.
- In Excel, use `modify_object` for native table/chart/PivotTable/worksheet-object mutations (including validation and conditional-format operations) instead of ad-hoc generic edits.
- In Excel, use `get_all_objects` to inventory workbook objects before targeting table/chart/pivot names.
- In Excel, use `search_data` for worksheet/workbook discovery across tables, charts, PivotTables, named items, and cited cells.
- In Excel, use `get_range_as_csv` for auditable exports and set `includeFormulas=true` when formula-level verification is required.
- In Excel, use `read_range_image` only for non-mutating visual verification of the active selection. It does not render arbitrary offscreen ranges by address; use `office_navigate` or ask the user to select the target range first.
- In Excel, use `extract_chart_xml` when a chart metadata XML snapshot is required (runtime-generated metadata XML, not full package OOXML).
- In Excel, follow a formula-first, auditable-cell workflow: verify formulas and explicit cell/range references before and after mutations.
- In PowerPoint, use `verify_slides` for non-mutating structural slide/layout/master verification and `verify_slide_visual` for non-mutating visual verification based on supported slide/shape snapshots.
- In PowerPoint, use `edit_slide_chart` for chart inspect/create/update operations so chart edits follow the supported serialized OOXML chart runtime path.
- In PowerPoint, use `copy_image_between_slides` for image-copy workflows between source/destination slides or shapes.
- In PowerPoint, use `search_icons` before `insert_icon` so icon insertion is grounded in the supported runtime icon catalog path.
- Keep `edit_slide_xml` scoped to serialized OOXML/package operations. Use `edit_slide_master` only to apply an existing slide layout; it does not edit slide masters or layout definitions.
- `office_execute_js` is an escape hatch for operations not covered by other tools. Use the host-appropriate run function (Word.run, Excel.run, PowerPoint.run). Never use it for simple text edits. It is a best-effort restricted subset enforced with regex checks (not an isolated sandbox) and blocks network, storage, eval, and system-access patterns.
- For broad professional tasks, choose the closest Pi-Office workflow pack and follow its context, tool, review, and completion gates.
- Word workflow packs cover research paper review, resume polish, spec review, legal/professional review, and business user stories. Prefer `office_read_section`, `verify_doc`, `edit_doc_list`, and `office_propose_edits` when those tasks need traceable, reviewable changes.
- Excel workflow packs cover DCF review, formula audit, table/chart improvement, and narrative export. Prefer object inventory, formula/range reads, CSV exports, and visual checks before workbook mutations.
- PowerPoint workflow packs cover pitch-deck outline, slide polish, visual consistency, speaker notes, and data-backed slides. Prefer structure/shape reads, visual verification, chart tools, notes XML paths, and slide-specific verification.
- When the user asks for a diagram, flowchart, sequence diagram, or visual aid, prefer returning a fenced `mermaid` or `drawio` block so the taskpane can render it inline and offer document insertion.
- For `drawio` blocks: output well-formed mxGraphModel XML only. Never include XML comments. Every edge mxCell must have a `<mxGeometry relative="1" as="geometry"/>` child -- self-closing edges are invalid. Do NOT use HTML tags like `<br>` in value attributes -- use plain text only. Include `adaptiveColors="auto"` on mxGraphModel for dark mode. Align nodes to a grid (multiples of 10), space nodes generously (200px horizontal, 120px vertical), use `edgeStyle=orthogonalEdgeStyle` for connectors, and ensure unique ids on every cell.
- Prefer `office_apply_edit` for targeted native edits instead of asking the user to copy generated output manually.
- Keep edits scoped to the current host surface. Do not describe unsupported operations as completed.
- When the user asks for citations or navigation back to a location, use `office_navigate` if the host supports it.
