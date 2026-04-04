# Claude Office Add-in Investigation and Tracking

## 1) Document Metadata

- Document ID: `INV-CLAUDE-PARITY-2026-04-04`
- Created: `2026-04-04`
- Repository: `C:\Users\manan\Code\Personal\office-word-addin`
- Purpose: Record investigation details, confirmed findings, and a trackable remediation backlog.
- Primary comparison target: parent-directory Claude Office add-in cache at `..\claude-powerpoint-addin-copy` plus live taskpane bundle at `https://pivot.claude.ai`.
- Companion planning artifact: `OFFICE_PARITY_TASK_LIST.md`.

## 2) What Was Investigated

- Whether the Claude Office add-in (in parent directory) exposes tools/capabilities that are missing in our implementation.
- Whether our current tool implementations are correct and consistent with their prompt/tool guidance.
- Whether our runtime tool registration and prompt instructions are aligned with actual execution behavior.
- Whether there are high-risk mismatches that should be prioritized before parity expansion.

## 3) Scope and Boundaries

- In scope:
- Tool surface comparison (Claude vs our add-in).
- Runtime/bridge correctness review for our Office tools.
- Prompt guidance and behavior-instruction comparison.
- Test-coverage gaps tied to critical failure modes.
- Out of scope:
- Connector/account OAuth parity and backend account systems.
- Claude proprietary server internals not available locally.
- Full UI/UX parity review beyond tooling and guidance behavior.

## 4) How the Investigation Was Done

### 4.1 Evidence Collection Steps

- [x] Enumerated parent-directory add-in artifacts from `..\claude-powerpoint-addin-copy`.
- [x] Inspected Claude cached Office manifest, extended manifest, app command cache, token/appstate metadata.
- [x] Confirmed parent folder is cache/manifest metadata, not full source.
- [x] Inspected our manifest and runtime/tool registration files.
- [x] Pulled live Claude taskpane HTML and shortcuts endpoint from `pivot.claude.ai`.
- [x] Pulled and scanned live bundled JS (`/m-addin/assets/index-BMkcZxzX.js`) for tool names and host guidance.
- [x] Cross-checked our implementation behavior in bridge and host tool code paths.
- [x] Ran multi-agent parallel review for independent validation.

### 4.2 Sources Used (Primary)

- Claude cache manifest:
- `..\claude-powerpoint-addin-copy\Wef\{5B7546C0-51DC-41A1-A796-F2EDDA9E67F0}\Omex\8Nhsvrs_h_VMw5ECZ+uMHQ==\Manifests\wa200010001_1.0.0.0`
- Claude cache extended manifest:
- `..\claude-powerpoint-addin-copy\Wef\{5B7546C0-51DC-41A1-A796-F2EDDA9E67F0}\Omex\8Nhsvrs_h_VMw5ECZ+uMHQ==\ExtendedManifest\wa200010001_1.0.0.0_en-US`
- Claude cache readme:
- `..\claude-powerpoint-addin-copy\README.md`
- Live Claude endpoints:
- `https://pivot.claude.ai`
- `https://pivot.claude.ai/shortcuts.json`
- `https://pivot.claude.ai/m-addin/assets/index-BMkcZxzX.js`
- Our core runtime/protocol files:
- `packages/pi-office-pack/src/protocol.ts`
- `packages/pi-office-pack/src/extension.ts`
- `packages/pi-office-pack/src/defaults.ts`
- `packages/pi-office-pack/skills/office-host.SKILL.md`
- `apps/taskpane/src/lib/runtime/inprocess-kernel.ts`
- `apps/taskpane/src/lib/office-bridge.ts`
- `apps/taskpane/src/lib/office/document-tools.ts`
- `scripts/office-tests/src/office-bridge.test.ts`

### 4.3 Confidence Model

- Confirmed: directly visible in local files/manifests or live bundle text.
- Inferred: derived from minified bundle semantics where names are clear but internals are not fully reconstructable.
- Unknown: server-side Claude behavior not visible from local cache/bundle text.

## 5) Current Architecture Components (Class/Function Names)

### 5.1 Core Runtime Classes (Our Code)

- `BrowserAuthStore` in `apps/taskpane/src/lib/runtime/inprocess-kernel.ts`.
- `BrowserCheckpointStore` in `apps/taskpane/src/lib/runtime/inprocess-kernel.ts`.
- `BrowserModelRegistry` in `apps/taskpane/src/lib/runtime/inprocess-kernel.ts`.
- `BrowserOfficeSession` in `apps/taskpane/src/lib/runtime/inprocess-kernel.ts`.
- `InProcessKernel` in `apps/taskpane/src/lib/runtime/inprocess-kernel.ts`.
- `LocalBridgeSocket` in `apps/taskpane/src/lib/runtime/inprocess-kernel.ts`.

### 5.2 Core Tool/Bridge Functions (Our Code)

- `createOfficeExtension(...)` in `packages/pi-office-pack/src/extension.ts`.
- `createOfficeToolExecutor(...)` in `apps/taskpane/src/lib/office-bridge.ts`.
- `toAnchor(...)` and `toHostAction(...)` in `apps/taskpane/src/lib/office-bridge.ts`.
- `executeOfficeJs(...)` in `apps/taskpane/src/lib/office/document-tools.ts`.
- `readDocumentSection(...)` in `apps/taskpane/src/lib/office/document-tools.ts`.
- `proposeDocumentEdits(...)` in `apps/taskpane/src/lib/office/document-tools.ts`.
- `applyAcceptedEdits(...)` in `apps/taskpane/src/lib/office/document-tools.ts`.

## 6) Findings: Tool Surface Comparison

## 6.1 Confirmed from Parent Cache Only (High Confidence)

- Claude parent cache confirms Office command surface:
- `ShowTaskpane` action.
- Ribbon button label `Open Claude`.
- Shortcut `Ctrl+Alt+C`.
- Host: PowerPoint (`Presentation`) only.
- Permissions: `ReadWriteDocument`.
- Runtime endpoint: `https://pivot.claude.ai`.

Result:
- No missing item against our manifest-level capability in this narrow command-surface layer.

## 6.2 Confirmed from Live Claude Bundle (Tool Registry and Guidance)

Additional tool names were observable in live bundle text, including Office document, spreadsheet, and slide-specialized tools (examples below):

- Word-focused examples:
- `edit_doc_text`, `edit_doc_list`, `propose_doc_edits`, `read_doc_section`, `verify_doc`, `verify_doc_visual`.
- PowerPoint-focused examples:
- `list_slide_shapes`, `read_slide_text`, `edit_slide_text`, `edit_slide_xml`, `edit_slide_chart`, `edit_slide_master`, `verify_slides`, `verify_slide_visual`, `insert_icon`, `search_icons`, `duplicate_slide`.
- Excel-focused examples:
- `get_cell_ranges`, `set_cell_range`, `clear_cell_range`, `resize_range`, `copy_to`, `modify_sheet_structure`, `modify_object`, `get_all_objects`, `get_range_as_csv`, `read_range_image`, `search_data`, `extract_chart_xml`.
- External/context utilities seen in guidance text:
- `refresh_mcp_connectors`, `read_skill`, `web_search`, `web_fetch`.

Result:
- Our implementation has robust general Office tools but does not expose many of these as first-class named tools.

## 6.3 Our Registered First-Class Tools (Current)

From `OFFICE_TOOL_NAMES` and runtime registration:

- `office_get_context`
- `office_apply_edit`
- `office_navigate`
- `office_capture_snapshot`
- `office_capture_viewport`
- `office_read_section`
- `office_execute_js`
- `office_propose_edits`
- `ask_user`
- `generate_image`

## 6.4 Gap List for Tracking (First-Class Tooling)

### Word Tooling Gaps

- [ ] `GAP-WORD-01` Add first-class equivalent of `edit_doc_text`.
- [ ] `GAP-WORD-02` Add first-class equivalent of `edit_doc_list`.
- [ ] `GAP-WORD-03` Add first-class equivalent of `verify_doc`.
- [ ] `GAP-WORD-04` Add first-class equivalent of `verify_doc_visual`.

### PowerPoint Tooling Gaps

- [ ] `GAP-PPT-01` Add first-class `list_slide_shapes` equivalent.
- [ ] `GAP-PPT-02` Add first-class `read_slide_text` equivalent.
- [ ] `GAP-PPT-03` Add first-class `edit_slide_text` equivalent.
- [ ] `GAP-PPT-04` Add first-class `edit_slide_xml` equivalent.
- [ ] `GAP-PPT-05` Add first-class `edit_slide_chart` equivalent.
- [ ] `GAP-PPT-06` Add first-class `edit_slide_master` equivalent.
- [ ] `GAP-PPT-07` Add first-class `duplicate_slide` equivalent.
- [ ] `GAP-PPT-08` Add first-class `copy_image_between_slides` equivalent.
- [ ] `GAP-PPT-09` Add first-class `insert_slide_element` equivalent.
- [ ] `GAP-PPT-10` Add first-class `remove_slide_element` equivalent.
- [ ] `GAP-PPT-11` Add first-class `verify_slides` equivalent.
- [ ] `GAP-PPT-12` Add first-class `verify_slide_visual` equivalent.
- [ ] `GAP-PPT-13` Add first-class `get_slide` equivalent.
- [ ] `GAP-PPT-14` Add first-class `get_presentation_structure` equivalent.
- [ ] `GAP-PPT-15` Add first-class `modify_presentation_structure` equivalent.
- [ ] `GAP-PPT-16` Add first-class icon workflow (`search_icons` + `insert_icon`) equivalent.

### Excel Tooling Gaps

- [ ] `GAP-XLS-01` Add first-class `get_cell_ranges` equivalent.
- [ ] `GAP-XLS-02` Add first-class `set_cell_range` equivalent.
- [ ] `GAP-XLS-03` Add first-class `clear_cell_range` equivalent.
- [ ] `GAP-XLS-04` Add first-class `resize_range` equivalent.
- [ ] `GAP-XLS-05` Add first-class `copy_to` equivalent.
- [ ] `GAP-XLS-06` Add first-class `modify_object` equivalent.
- [ ] `GAP-XLS-07` Add first-class `modify_sheet_structure` equivalent.
- [ ] `GAP-XLS-08` Add first-class `get_range_as_csv` equivalent.
- [ ] `GAP-XLS-09` Add first-class `read_range_image` equivalent.
- [ ] `GAP-XLS-10` Add first-class `search_data` equivalent.
- [ ] `GAP-XLS-11` Add first-class `get_all_objects` equivalent.
- [ ] `GAP-XLS-12` Add first-class `extract_chart_xml` equivalent.

### External Context Tooling Gaps (if intentionally in-scope)

- [ ] `GAP-EXT-01` Decide whether to expose `refresh_mcp_connectors` equivalent.
- [ ] `GAP-EXT-02` Decide whether to expose `read_skill` equivalent.
- [ ] `GAP-EXT-03` Decide whether to expose first-class `web_search`/`web_fetch` equivalents.

## 7) Findings: Correctness and Behavior Risks (Our Code)

### 7.1 High Severity

#### `COR-001` Success Flag Mismatch in Bridge Results

- Problem:
- `createOfficeToolExecutor` returns `success: true` for `office_read_section`, `office_execute_js`, and `office_propose_edits` as long as no exception is thrown.
- Underlying handlers can return payloads like `{ error: ... }` or `{ ok: false, error: ... }`, which then appear as successful tool calls.
- Impact:
- `BrowserOfficeSession.invokeOfficeTool(...)` checks `result.success`; false negatives can be treated as valid outputs.
- Affected functions/classes:
- `createOfficeToolExecutor(...)`
- `BrowserOfficeSession.invokeOfficeTool(...)`
- `readDocumentSection(...)`
- `executeOfficeJs(...)`
- `proposeDocumentEdits(...)`
- Tracking:
- [ ] `COR-001-A` Normalize payload-aware failures to `success: false` at bridge boundary.
- [ ] `COR-001-B` Add regression tests for error-object payload normalization.

#### `COR-002` Non-Deterministic Edit Targeting in Accepted Proposals

- Problem:
- `applyAcceptedEdits(...)` uses search-first-match behavior (`results.items[0]`) for target replacement.
- For repeated phrases, wrong occurrence can be edited.
- `anchor`/`paragraphId` are modeled but not enforced in apply path.
- Impact:
- Legal/document editing correctness risk.
- Affected functions/classes:
- `applyAcceptedEdits(...)`
- `proposeDocumentEdits(...)`
- `createOfficeExtension(...)` proposal schemas and review flow.
- Tracking:
- [ ] `COR-002-A` Apply edits using deterministic locator chain (`paragraphId` then anchored search then fallback).
- [ ] `COR-002-B` Add tests for repeated-text multi-occurrence edits.

### 7.2 Medium Severity

#### `COR-003` `office_execute_js` Enforcement Weaker Than Guidance

- Problem:
- Enforcement currently relies on regex denylist + `new Function(...)` execution.
- Guidance says “must not access network/storage/eval/system resources”, but enforcement is best-effort.
- Affected functions/classes:
- `executeOfficeJs(...)`
- `createOfficeExtension(...)` descriptions and policy text.
- Tracking:
- [ ] `COR-003-A` Define explicit threat model and allowed API subset.
- [ ] `COR-003-B` Harden execution boundary or reduce supported scope.
- [ ] `COR-003-C` Align prompt language to actual enforceable guarantees.

#### `COR-004` `searchText` Constraint Drift

- Problem:
- Prompt/skill states `searchText` MUST be under 200 chars.
- Implementation supports up to 255 and includes long-text fallback behavior.
- Affected functions/classes:
- `createOfficeExtension(...)`
- `executeOfficeJs(...)`
- `office-host.SKILL.md`
- Tracking:
- [ ] `COR-004-A` Choose canonical limit and enforce in schema + runtime.
- [ ] `COR-004-B` Update skill and prompt text to same threshold.

#### `COR-005` PowerPoint Snapshot Slice Integrity Risk

- Problem:
- In `captureDocumentSnapshot(...)`, per-slice failures can be ignored while overall operation resolves once slice count is reached.
- Impact:
- Potential partial/corrupt `presentationBase64` data accepted as success.
- Affected functions/classes:
- `captureDocumentSnapshot(...)`.
- Tracking:
- [ ] `COR-005-A` Fail fast if any slice retrieval fails.
- [ ] `COR-005-B` Add checksum/length validation before resolve.

### 7.3 Low Severity

#### `COR-006` Test Coverage Gaps for Critical Failure Modes

- Problem:
- Existing tests focus heavily on happy paths and exception flows.
- Missing direct coverage for payload-level error normalization and deterministic edit targeting.
- Tracking:
- [ ] `COR-006-A` Add bridge tests for `{ error }` payload -> failed tool result.
- [ ] `COR-006-B` Add proposal apply tests for repeated `searchText` collisions.
- [ ] `COR-006-C` Add PowerPoint snapshot partial-slice failure tests.

## 8) Findings: Prompt/Guidance Differences

### 8.1 Observed Difference

- Claude guidance (from live bundle text) is highly host-specific and prescriptive for Word legal workflows, Excel formula discipline, and PowerPoint OOXML/layout practices.
- Our guidance is strong but broader and less host-prescriptive in some areas.

### 8.2 Alignment Tasks

- [ ] `PROMPT-01` Add host-specific Word legal editing guardrails where desired.
- [ ] `PROMPT-02` Add explicit Excel “formula-first, auditable cells” rule set if desired.
- [ ] `PROMPT-03` Add deeper PowerPoint XML/layout rules where desired.
- [ ] `PROMPT-04` Ensure every prompt guarantee has corresponding enforceable runtime behavior.

## 9) Investigation Completion Checklist

- [x] Gathered local Claude cache evidence.
- [x] Gathered live Claude endpoint evidence.
- [x] Indexed our tool surface and runtime classes/functions.
- [x] Identified first-class tool gaps.
- [x] Identified correctness and guidance mismatches.
- [x] Produced actionable checkbox backlog for tracking.

## 10) Recommended Workstreams for Tracking

### Workstream A: Correctness First

- [ ] Complete `COR-001` through `COR-006` before broad parity expansion.

### Workstream B: First-Class Tool Surface Expansion

- [ ] Prioritize Word and PowerPoint first-class tools (`GAP-WORD-*`, `GAP-PPT-*`).
- [ ] Add Excel first-class tool wrappers (`GAP-XLS-*`) where runtime primitives already exist.

### Workstream C: Prompt + Runtime Contract Synchronization

- [ ] Complete `PROMPT-01` through `PROMPT-04`.
- [ ] Add CI checks to detect schema/prompt/runtime drift.

## 11) Notes and Unknowns

- Parent-directory Claude cache does not provide full source; live bundle inspection gives strong signal but not full server-side behavior.
- Any inferred behavior from minified client code should be validated against real runtime behavior before treating it as a hard parity target.

## 12) Status Block (For Future Updates)

- Last updated:
- Owner:
- Current phase:
- Blocking issues:
- Next milestone:

## 13) Manual Testing Scenarios and Steps

### Word

The following Word checklist is aligned to `npm run smoke:office -- --list` scenario IDs.

#### Scenario: `word-review-anchors` — Review-anchor navigation

- Steps:
  1. Open a Word document containing headings, comments, revisions, footnotes/endnotes, fields, and content controls.
  2. Run `office_get_context` with `includeFormatting=true` and capture returned anchors.
  3. Run `office_navigate` for at least one anchor in each review-anchor category.
- Expected:
  - Returned anchors include the review artifact families above.
  - Navigation lands on the expected anchor target (or nearest supported reference location).

#### Scenario: `word-structured-edits` — Structured Word edits

- Steps:
  1. Place caret in an editable paragraph and keep one unresolved comment plus one pending revision available.
  2. Run `edit_doc_text` for a direct sentence-level replacement or insertion.
  3. Run `office_apply_edit` actions for native Word objects (for example comment/revision/field/content-control actions) to confirm host-native behavior remains available.
- Expected:
  - `edit_doc_text` updates text through native Word edit paths.
  - Structured `office_apply_edit` actions operate on native Word objects (not flattened plain-text fallbacks).

#### Scenario: `word-visual-verification` — Word document + visual verification

- Steps:
  1. Run `verify_doc` with `scope=document` and confirm summary plus structured `details` payload.
  2. Run `verify_doc_visual` with `includeFormatting=true` and confirm returned `visual`, `details`, and `visuals` payload fields.
  3. Confirm visual verification notes include viewport/path limitations (not a full OS window-frame capture).
- Expected:
  - Verification outputs are non-mutating and structured for review workflows.
  - Visual verification stays bounded to the supported Word viewport capture path and reports explicit limits.

#### Scenario: `word-taskpane-stability` — Taskpane focus, scroll, prompt, and connector stability

- Steps:
  1. Send a long prompt from the taskpane composer and confirm keyboard focus stays in the taskpane.
  2. During streaming, scroll chat upward and verify manual scroll is preserved.
  3. Trigger one connector-backed request, then reconnect/reopen taskpane and send another prompt.
- Expected:
  - Focus, scroll, and streaming remain stable.
  - Connector outcomes are surfaced explicitly.
  - Session reconnect remains usable for subsequent prompts/tool calls.

### PowerPoint

The following PowerPoint checklist is aligned to `npm run smoke:office -- --list` scenario IDs.

#### Scenario: `powerpoint-anchor-navigation` — Slide, layout, shape, and notes anchors

- Steps:
  1. Open a deck with notes and at least one named shape on the selected slide.
  2. Run `office_get_context` with `includeFormatting=true` and confirm anchors include slide/shape/layout/master/notes targets.
  3. Run `office_navigate` with one anchor from each anchor family above.
- Expected:
  - Slide and shape anchors navigate to the exact target.
  - Layout/master/notes anchors report explicit fallback behavior when native direct navigation is partial.

#### Scenario: `powerpoint-slide-and-shape-authoring` — Native slide and shape operations

- Steps:
  1. In a disposable deck copy, run lifecycle operations such as `addAgendaSlide`, `addTransitionSlide`, `duplicateSlide`, `reorderSlides`, `combineSlides`, and `deleteSlide` (with destructive confirmation where required).
  2. Run text/media operations including `setShapeText`, `updateShapeProperties`, `replaceShapeImage`, and table updates (`setTableValues` or `updateTableCell`).
  3. Run process-flow style shape generation to confirm editable shapes/connectors are inserted.
- Expected:
  - Slide lifecycle operations mutate real slide objects (not flattened fallback output).
  - Resulting shapes/images/tables remain editable in native PowerPoint after completion.

#### Scenario: `powerpoint-notes-and-charts` — Serialized notes and chart workflows

- Steps:
  1. Run `getSlideNotes` then `setSlideNotes` for a target slide and confirm updated notes are reflected after reselection.
  2. Run `inspectPresentationPackage` or `getPresentationTheme` to validate serialized package metadata visibility.
  3. Run `edit_slide_chart` with inspect/create/update operations (for example `get_slide_charts`, `add_slide_chart`, `update_slide_chart`) on chart-bearing slides.
- Expected:
  - Notes operations remain aligned to serialized slide replacement flows with explicit result metadata.
  - Chart operations return chart metadata and preserve native chart editability with embedded workbook sync.

#### Scenario: `powerpoint-taskpane-stability` — Taskpane focus, scroll, prompt, and connector stability

- Steps:
  1. Send a long prompt while interacting with slides and confirm taskpane keyboard focus remains stable.
  2. Scroll chat during streaming output and confirm user-controlled scroll behavior remains intact.
  3. Trigger one connector-backed request, then reconnect/reopen taskpane and submit another prompt.
- Expected:
  - Prompting and scroll behavior stay stable during active slide interaction.
  - Connector outcomes remain visible and explicit.
  - Reconnected sessions continue to execute prompts and tools normally.

