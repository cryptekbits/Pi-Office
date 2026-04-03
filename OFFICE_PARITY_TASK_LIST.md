# Office Parity Task List

## Goal

Reach feature parity with the current Claude Office add-ins for raw Office document capabilities while keeping the implementation native to this codebase:

- built on Pi core as a dependency
- shaped around our existing Office bridge
- behavior-compatible, not code-derivative
- limited to Office document capabilities, not backend, auth, model-routing, connector, or analytics concerns

## Design Principles

- Parity by behavior, not by structure. We should match user-visible capabilities without mirroring Claude's internal tool taxonomy or code layout.
- Keep Pi integration stable. The Pi extension layer should remain the main control surface; richer Office actions should sit behind the existing companion <-> taskpane bridge.
- Prefer typed host actions over prompt-shaped strings. Complex Excel and PowerPoint work will not stay reliable if everything is funneled through `content: string`.
- Split host logic by domain. The current monolithic `apps/taskpane/src/lib/office.ts` should be decomposed into shared helpers plus `word`, `excel`, and `powerpoint` host adapters.
- Preserve native Office objects. When we create or edit charts, tables, shapes, comments, or tracked changes, the result must remain editable in Office.
- Favor Office.js first, OOXML/ZIP/document transforms second, and desktop-native helpers only when Office.js cannot close the gap cleanly.

## Progress

Implemented on 2026-04-02:

- Typed anchor and action contracts in `packages/pi-office-pack/src/protocol.ts`
- Tool schema widening in `packages/pi-office-pack/src/extension.ts`
- Real `office_navigate` dispatch in `apps/taskpane/src/lib/office-tools.ts`
- Anchor-aware `office_get_context` payloads for Word, Excel, and PowerPoint
- Standardized action/navigation result envelopes with touched and created object references
- Structured host actions in `apps/taskpane/src/lib/office.ts` for:
  - Word navigation to headings, paragraphs, comments, and revisions
  - Word native edits for text, HTML, comments, tables, and base64 file insertion
  - Excel navigation to sheets, named items, ranges, tables, charts, and pivot tables
  - Excel native actions for value writes, formulas, number formats, formatting, sorting, filtering, tables, charts, pivot tables, data validation, conditional formatting, and worksheet lifecycle operations
  - PowerPoint navigation to slides and shapes
  - PowerPoint native actions for text updates, slide add/move/select, text boxes, tables, lines, and slide import/export
- Broader `office_get_context` document summaries:
  - Word document-wide headings, comments, and revisions
  - Excel workbook sheet summaries, used-range summaries, and workbook-wide table/chart/pivot inventories
- Word document-wide footnote and endnote anchors plus navigation
- PowerPoint presentation-wide slide master and layout metadata in `office_get_context`
- PowerPoint layout-aware slide actions:
  - resolve layouts and masters by ID or name
  - apply a layout to an existing slide
  - return layout/master metadata in native slide results
- Broader PowerPoint native object authoring:
  - geometric shapes
  - shape grouping
  - shared fill/line/alt-text/hyperlink/z-order property application
- PowerPoint native slide lifecycle actions:
  - duplicate single or multiple slides
  - delete selected or targeted slides with a keep-one-slide safeguard
  - merge/import presentation payloads with created-slide references and reselection
- Richer result-envelope warnings for host fallbacks such as `setSelectedDataAsync`
- Workbook-aware Excel anchors and navigation for the workbook shell plus workbook-level tables, charts, and pivot tables
- Broader Excel native actions for:
  - row and column insert/delete
  - duplicate removal
  - worksheet gridline and heading visibility
  - worksheet print area
- Richer Excel conditional formatting families:
  - color scales
  - icon sets

Still pending from the plan:

- deeper Excel parity such as chart axes/labels editing, pivot schema/filter refinements, table-aware sort/filter affordances, and remaining formatting helpers like borders
- broader PowerPoint structure coverage such as notes regions and deeper layout/master editing beyond slide-level targeting
- fuller Word review/document structure coverage such as wider non-selection navigation affordances beyond headings/comments/revisions/notes

## Current Anchors In This Repo

- Protocol and tool surface:
  - `packages/pi-office-pack/src/protocol.ts`
  - `packages/pi-office-pack/src/extension.ts`
- Taskpane bridge and host implementation:
  - `apps/taskpane/src/lib/office-tools.ts`
  - `apps/taskpane/src/lib/office.ts`
- Companion orchestration:
  - `apps/companion/src/office-session.ts`
  - `apps/companion/src/viewport-capture.ts`

## Out Of Scope

- Claude account parity
- LLM gateway/auth/OAuth work
- skills/connectors/instructions UX parity
- analytics, audit, retention, or observability parity
- model selector behavior beyond what is needed for Office tool execution

## Implementation Strategy

### Phase 0. Foundation Refactor

- [x] Split `apps/taskpane/src/lib/office.ts` into host-specific adapters:
  - `apps/taskpane/src/lib/office/shared/*`
  - `apps/taskpane/src/lib/office/word/*`
  - `apps/taskpane/src/lib/office/excel/*`
  - `apps/taskpane/src/lib/office/powerpoint/*`
- [x] Introduce a typed Office action model in `packages/pi-office-pack/src/protocol.ts`:
  - richer context payloads
  - anchor descriptors
  - structured edit actions
  - structured navigation targets
  - host capability descriptors
- [x] Keep the Pi tool layer recognizable, but widen the payloads:
  - `office_get_context`
  - `office_apply_edit`
  - `office_navigate`
  - `office_capture_snapshot`
  - `office_capture_viewport`
- [x] Rework `office_apply_edit` from string-only edits into typed operations by host.
- [x] Make `office_navigate` real instead of placeholder behavior.
- [x] Add shared result envelopes with:
  - touched objects
  - created objects
  - navigation anchors
  - warnings for partial fallbacks

Acceptance criteria:

- The companion can ask for structured context and dispatch typed document actions without host-specific string parsing.
- Navigation is functional in all supported hosts.
- Existing simple prompts continue to work.

### Phase 1. Shared Document Model And Navigation

- [ ] Add host-neutral anchor records:
  - [x] Word: heading, paragraph, comment, revision, footnote, range
  - [x] Excel: worksheet, named item, table, chart, pivot table, range, cell
  - [ ] PowerPoint: slide, layout, shape, chart, table, notes region
- [x] Extend `office_get_context` to return:
  - active selection
  - nearby anchors
  - document structure summary
  - stable IDs where Office exposes them
- [x] Implement `office_navigate` for:
  - Word headings/paragraphs/comments/revisions
  - Excel sheets/ranges/charts/pivots/tables
  - PowerPoint slides/shapes
- [x] Add citations/target references in results so Pi can point users back to exact Office objects.

Acceptance criteria:

- Pi can cite and revisit precise Word/Excel/PowerPoint locations.
- Navigation no longer returns scaffold placeholders.

### Phase 2. Excel Parity

#### Workbook Context

- [x] Move from selection-only Excel context to workbook-aware context.
- [x] Enumerate workbook structure:
  - worksheets
  - used ranges
  - tables
  - charts
  - pivot tables
  - named ranges
- [x] Add multi-sheet summaries and targeted sheet snapshots.
- [x] Add cell/range citation objects in responses.

#### Native Editing

- [x] Support value edits, formula edits, and mixed range writes.
- [x] Support row/column insert/delete where safe.
- [x] Support sheet creation, rename, duplicate, and delete with safeguards.
- [x] Support workbook and sheet navigation targets.

#### Formatting And Controls

- [x] Conditional formatting:
  - value rules
  - formula rules
  - data bars
- [x] Extend conditional formatting with:
  - color scales
  - icon sets
- [x] Sort and filter:
  - [x] worksheet ranges
  - [x] Excel tables
  - [x] pivot filters where supported
- [x] Data validation:
  - list dropdowns
  - numeric/date/text constraints
- [ ] Finance formatting helpers:
  - [x] gridlines on/off
  - [x] print area
  - [x] number formats
  - [x] alignment/fill/font
  - [x] borders

#### Charts And Pivot Tables

- [x] Create native Excel charts from selected or referenced data.
- [x] Edit existing charts:
  - [x] type
  - [x] titles
  - [x] axes
  - [x] legend
  - [x] labels
  - [x] source range
- [x] Create and edit pivot tables:
  - [x] source range or table
  - [x] row/column/value/filter fields
  - [x] sort adjustments
  - [x] filter adjustments
  - [x] schema changes

Acceptance criteria:

- Pi can work across multi-tab workbooks, cite cells/ranges, and perform native Excel operations comparable to Claude's public Excel feature set.

### Phase 3. PowerPoint Parity

#### Presentation Context

- [x] Add presentation-wide context:
  - slide inventory
  - selected slides
  - selected shapes
  - layout usage
  - theme/master metadata
  - notes presence
- [x] Add richer slide and shape descriptors:
  - text blocks
  - charts
  - tables
  - images
  - geometry
  - layout assignment

#### Template And Layout Awareness

- [x] Read slide master, custom layouts, theme fonts, and theme colors.
- [x] Let Pi target a specific layout when creating a slide.
- [x] Add slide creation that respects template/layout selection.

#### Native Slide And Object Editing

- [x] Create new slides natively, not as screenshots.
- [x] Duplicate, delete, reorder, and merge slides.
- [x] Support slide-level restructuring actions:
  - combine slides
  - insert agenda/transition slides
  - reorder storyline
- [x] Support pinpoint edits to selected objects:
  - [x] text replacements
  - [x] shape property updates
  - [x] image replacement
  - [x] object reposition/resize
- [x] Add table creation and editing.
- [x] Add native chart creation and editing.
- [x] Add simple diagram/process-flow generation from structured inputs.

#### Advanced Serialization Path

- [x] Introduce a presentation transformation path for operations that PowerPoint API alone cannot express cleanly.
- [x] Add serialized chart inspection and cache updates for existing charts.
- [x] Keep serialized chart edits synchronized with embedded workbook data.
- [ ] Prefer a repo-owned serialization layer for:
  - slide cloning
  - layout-preserving content insertion
  - native chart/table/shape assembly
- [x] Keep this implementation distinct from Claude's inferred OOXML path.

Acceptance criteria:

- Pi can build decks from templates, perform pinpoint edits, restructure decks, and create editable native charts/tables/diagrams.

### Phase 4. Word Parity

#### Document Structure And Review Context

- [x] Extend Word context beyond current selection snapshots.
- [x] Build a heading and paragraph map with stable references where available.
- [x] Enumerate comments and tracked changes.
- [x] Add footnote and field awareness where supported by the requirement set.
- [x] Include reviewed text metadata and revision summaries in context.

#### Navigation And Citations

- [x] Navigate to headings, paragraphs, comments, revisions, and footnotes.
- [x] Return paragraph/comment/revision citations that Pi can reference in replies.

#### Native Editing

- [x] Add document-fragment insertion beyond selection text/html replacement.
- [x] Add richer format-preserving replacements for paragraphs/ranges.
- [x] Add comment-aware and tracked-change-aware edit flows.
- [x] Add document/template insertion using a repo-owned base64/OOXML insertion path.

Acceptance criteria:

- Pi can reason over Word review artifacts and navigate back to exact comments/revisions/paragraph anchors.
- Word editing goes beyond text insertion and preserves more native document structure.

### Phase 5. Safety, Verification, And Tooling

- [x] Add host capability tests around protocol serialization and action dispatch.
- [x] Add adapter-level tests for:
  - Excel action planning
  - PowerPoint action planning
  - Word anchor resolution
- [x] Add smoke scripts for manual desktop verification in Word, Excel, and PowerPoint.
- [x] Add explicit fallback reporting when the host/API level cannot complete a requested action natively.
- [x] Add overwrite safeguards for destructive document actions.

Acceptance criteria:

- Every parity claim maps to a manual or automated verification case.
- Unsupported operations fail clearly instead of silently degrading.

## Recommended Execution Order

1. Foundation refactor and real navigation
2. Excel parity
3. PowerPoint parity
4. Word parity
5. Verification hardening

## Why This Order

- Excel parity has the clearest public feature target and the cleanest Office.js surface.
- PowerPoint parity is the largest gap, but it benefits from the structured action model and shared anchor system first.
- Word parity depends on the same anchor/citation infrastructure and will be cleaner after the shared navigation work lands.

## Definition Of Done

- Excel:
  - workbook-aware
  - multi-sheet
  - cell/range citations
  - native pivots/charts/formatting/data validation/sort/filter
- PowerPoint:
  - template-aware
  - native slide generation
  - pinpoint object editing
  - deck restructuring
  - native charts/tables/diagrams
- Word:
  - comments
  - tracked changes
  - paragraph/headings map
  - anchor navigation
  - richer insertion/document fragment support
- Shared:
  - typed protocol
  - true navigation
  - clear capability reporting
  - verification coverage

## Immediate Next Slice

Implementation parity backlog is complete:

- [x] add border-format helpers for Excel ranges/tables
- [x] add PowerPoint notes-region anchor coverage
- [x] add smoke scripts for later desktop validation
- [x] split host adapters out of `office.ts`

The remaining work is operational, not implementation scope: run the desktop smoke scripts in Word, Excel, and PowerPoint and close any host-specific issues they surface.
