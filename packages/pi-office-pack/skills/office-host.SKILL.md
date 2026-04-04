# Office Host

Use this skill when a task depends on the active Microsoft Office document.

## Rules

- Call `office_get_context` before making precise claims about the current selection, worksheet, or slide content.
- For Word layout questions about what is visible on screen, call `office_capture_viewport` proactively. Use it for alignment, page breaks, wrapping, clipping, visible page position, margins, and header/footer placement.
- `office_capture_viewport` is only for Word and returns Office.js viewport metadata plus context-derived visuals. It is not a pixel-perfect OS/window screenshot and should not be used as a substitute for reading off-screen content or reviewing the whole document.
- For large Word documents where context shows only a textPreview (not full text), use `office_read_section` to page through content. Start at index 0 and advance by 20 paragraphs per call. Use headings from `office_get_context` to navigate to relevant sections.
- When making substantive multi-paragraph changes to a Word document, or when tracked changes are enabled, prefer `office_propose_edits` over `office_apply_edit`. This lets the user review each change before it is applied.
- When using `office_propose_edits`, each edit's `searchText` MUST be under 200 characters. Break large changes into many small, per-sentence or per-phrase edits. Never use a full paragraph as searchText. For a paragraph rewrite, create separate edits for each sentence or clause within it.
- `office_execute_js` is an escape hatch for operations not covered by other tools. Use the host-appropriate run function (Word.run, Excel.run, PowerPoint.run). Never use it for simple text edits. It is a best-effort restricted subset enforced with regex checks (not an isolated sandbox) and blocks network, storage, eval, and system-access patterns.
- When the user asks for a diagram, flowchart, sequence diagram, or visual aid, prefer returning a fenced `mermaid` or `drawio` block so the taskpane can render it inline and offer document insertion.
- For `drawio` blocks: output well-formed mxGraphModel XML only. Never include XML comments. Every edge mxCell must have a `<mxGeometry relative="1" as="geometry"/>` child -- self-closing edges are invalid. Do NOT use HTML tags like `<br>` in value attributes -- use plain text only. Include `adaptiveColors="auto"` on mxGraphModel for dark mode. Align nodes to a grid (multiples of 10), space nodes generously (200px horizontal, 120px vertical), use `edgeStyle=orthogonalEdgeStyle` for connectors, and ensure unique ids on every cell.
- Prefer `office_apply_edit` for targeted native edits instead of asking the user to copy generated output manually.
- Keep edits scoped to the current host surface. Do not describe unsupported operations as completed.
- When the user asks for citations or navigation back to a location, use `office_navigate` if the host supports it.
