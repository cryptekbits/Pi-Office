# Claude Office Add-in Investigation and Tracking

## 1) Document Metadata

- Document ID: `INV-CLAUDE-PARITY-2026-04-04`
- Created: `2026-04-04`
- Repository: `C:\Users\manan\Code\Personal\office-word-addin`
- Purpose: Record investigation details, confirmed findings, and a trackable remediation backlog.
- Primary comparison target: parent-directory Claude Office add-in cache at `..\claude-powerpoint-addin-copy` plus live taskpane bundle at `https://pivot.claude.ai`.
- Companion planning artifact: `docs/OFFICE_PARITY_TASK_LIST.md`.

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
- `addin/packages/pi-office-pack/src/protocol.ts`
- `addin/packages/pi-office-pack/src/extension.ts`
- `addin/packages/pi-office-pack/src/defaults.ts`
- `addin/packages/pi-office-pack/skills/office-host.SKILL.md`
- `addin/apps/taskpane/src/lib/runtime/inprocess-kernel.ts`
- `addin/apps/taskpane/src/lib/office-bridge.ts`
- `addin/apps/taskpane/src/lib/office/document-tools.ts`
- `addin/scripts/office-tests/src/office-bridge.test.ts`

### 4.3 Confidence Model

- Confirmed: directly visible in local files/manifests or live bundle text.
- Inferred: derived from minified bundle semantics where names are clear but internals are not fully reconstructable.
- Unknown: server-side Claude behavior not visible from local cache/bundle text.

## 5) Current Architecture Components (Class/Function Names)

### 5.1 Core Runtime Classes (Our Code)

- `BrowserAuthStore` in `addin/apps/taskpane/src/lib/runtime/inprocess-kernel.ts`.
- `BrowserCheckpointStore` in `addin/apps/taskpane/src/lib/runtime/inprocess-kernel.ts`.
- `BrowserModelRegistry` in `addin/apps/taskpane/src/lib/runtime/inprocess-kernel.ts`.
- `BrowserOfficeSession` in `addin/apps/taskpane/src/lib/runtime/inprocess-kernel.ts`.
- `InProcessKernel` in `addin/apps/taskpane/src/lib/runtime/inprocess-kernel.ts`.
- `LocalBridgeSocket` in `addin/apps/taskpane/src/lib/runtime/inprocess-kernel.ts`.

### 5.2 Core Tool/Bridge Functions (Our Code)

- `createOfficeExtension(...)` in `addin/packages/pi-office-pack/src/extension.ts`.
- `createOfficeToolExecutor(...)` in `addin/apps/taskpane/src/lib/office-bridge.ts`.
- `toAnchor(...)` and `toHostAction(...)` in `addin/apps/taskpane/src/lib/office-bridge.ts`.
- `executeOfficeJs(...)` in `addin/apps/taskpane/src/lib/office/document-tools.ts`.
- `readDocumentSection(...)` in `addin/apps/taskpane/src/lib/office/document-tools.ts`.
- `proposeDocumentEdits(...)` in `addin/apps/taskpane/src/lib/office/document-tools.ts`.
- `applyAcceptedEdits(...)` in `addin/apps/taskpane/src/lib/office/document-tools.ts`.

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

- [x] `GAP-EXT-01` Decide whether to expose `refresh_mcp_connectors` equivalent.
- [x] `GAP-EXT-02` Decide whether to expose `read_skill` equivalent.
- [x] `GAP-EXT-03` Decide whether to expose first-class `web_search`/`web_fetch` equivalents.

### 6.5 External Context Gap Closure Decisions (Bounded Independent-Taskpane Path)

- `GAP-EXT-01: resolved` - `refresh_mcp_connectors` remains a runtime/UI capability rather than a first-class agent tool. The supported refresh flow is Integrations -> Connected -> **Re-verify**, backed by `POST /v1/connectors/reverify`. Manual validation should execute re-verify during `word-taskpane-stability`, `excel-taskpane-stability`, and `powerpoint-taskpane-stability` scenarios and confirm surfaced status/log feedback.
- `GAP-EXT-02: resolved` - No first-class `read_skill` tool is exposed. Skill context is provided through packaged skill injection (`office-host.SKILL.md`, `workspace-handoff.SKILL.md`) and runtime prompt guidance. Saved-document gating remains explicit: unsaved documents stay in document-only mode, while saved documents can unlock workspace-scoped capabilities.
- `GAP-EXT-03: resolved` - No first-class `web_search`/`web_fetch` tools are exposed. Web-grounded context flows through research connectors in the connector catalog under enforced `hard-read-only` policy (`allowPrompts=false` plus read-safe allow/block patterns), without companion-coupled assumptions.

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

## 13) Resolution Ledger (Machine-Verifiable)

This ledger is the authoritative status record for every in-scope backlog item.  
Entry format is strict: `ID: status - note`.

### 13.1 Correctness Backlog (`COR-*`)

- COR-001: resolved - Bridge now normalizes payload-level failures into failed tool results instead of reporting false success.
- COR-001-A: resolved - `createOfficeToolExecutor` maps `{ error }` and `{ ok: false, error }` payloads to `success: false`.
- COR-001-B: resolved - `test:office` regressions cover payload-aware failure normalization for read-section, execute-js, and propose-edits.
- COR-002: resolved - Accepted Word proposals now use deterministic locator precedence before verified search fallback.
- COR-002-A: resolved - Apply path prioritizes `paragraphId`, then explicit anchors, then verified text-search fallback.
- COR-002-B: resolved - Repeated-text and locator-precedence regression tests were added to prevent first-match mis-edits.
- COR-003: resolved - `office_execute_js` contract is now explicit best-effort restricted execution, not sandbox language.
- COR-003-A: resolved - Threat model is encoded as blocked network/storage/eval/system-access categories.
- COR-003-B: resolved - Runtime rejects documented forbidden patterns before execution.
- COR-003-C: resolved - Prompt/tool descriptions were aligned to enforceable regex-based guarantees.
- COR-004: resolved - One canonical `searchText` limit is shared across schema, runtime enforcement, and guidance.
- COR-004-A: resolved - Over-limit `searchText` inputs are explicitly rejected at runtime.
- COR-004-B: resolved - Skill/prompt/runtime descriptions now publish the same `searchText` threshold.
- COR-005: resolved - PowerPoint snapshot capture now rejects incomplete or partial payloads before success.
- COR-005-A: resolved - Slice retrieval failure now fails fast instead of allowing partial assembly.
- COR-005-B: resolved - Missing/corrupt chunk validation runs before base64 assembly completion.
- COR-006: resolved - Critical failure-mode coverage was expanded across bridge, proposal, and snapshot paths.
- COR-006-A: resolved - Bridge regression tests assert payload-error normalization behavior.
- COR-006-B: resolved - Proposal-apply regressions assert repeated-text collision handling.
- COR-006-C: resolved - Snapshot regressions assert slice-failure and missing-chunk rejection behavior.

### 13.2 Prompt and Runtime Alignment (`PROMPT-*`)

- PROMPT-01: resolved - Word guidance now includes host-specific review/legal editing guardrails aligned with runtime behavior.
- PROMPT-02: resolved - Excel guidance now enforces formula-first, auditable-cell expectations aligned with supported tools.
- PROMPT-03: resolved - PowerPoint guidance now includes explicit XML/layout/master constraints aligned to supported paths.
- PROMPT-04: resolved - Prompt guarantees are synchronized with executable runtime/tool contracts and parity tests.

### 13.3 Word First-Class Tool Gaps (`GAP-WORD-*`)

- GAP-WORD-01: resolved - First-class Word text editing is exposed (`edit_doc_text` equivalent).
- GAP-WORD-02: resolved - First-class Word list editing is exposed (`edit_doc_list` equivalent).
- GAP-WORD-03: resolved - First-class Word document verification is exposed (`verify_doc` equivalent).
- GAP-WORD-04: resolved - First-class Word visual verification is exposed (`verify_doc_visual` equivalent).

### 13.4 PowerPoint First-Class Tool Gaps (`GAP-PPT-*`)

- GAP-PPT-01: resolved - First-class shape listing is exposed (`list_slide_shapes` equivalent).
- GAP-PPT-02: resolved - First-class per-slide text read is exposed (`read_slide_text` equivalent).
- GAP-PPT-03: resolved - First-class slide text editing is exposed (`edit_slide_text` equivalent).
- GAP-PPT-04: resolved - First-class slide XML editing is exposed (`edit_slide_xml` equivalent).
- GAP-PPT-05: resolved - First-class slide chart editing is exposed (`edit_slide_chart` equivalent).
- GAP-PPT-06: partially resolved - First-class layout application is exposed through the legacy-named `edit_slide_master` tool. It applies existing layouts via native resolution; it does not edit slide masters or layout definitions.
- GAP-PPT-07: resolved - First-class slide duplication is exposed (`duplicate_slide` equivalent).
- GAP-PPT-08: resolved - First-class image copy between slides is exposed.
- GAP-PPT-09: resolved - First-class slide-element insertion is exposed (`insert_slide_element` equivalent).
- GAP-PPT-10: resolved - First-class slide-element removal is exposed (`remove_slide_element` equivalent).
- GAP-PPT-11: resolved - First-class structural slide verification is exposed (`verify_slides` equivalent).
- GAP-PPT-12: resolved - First-class visual slide verification is exposed (`verify_slide_visual` equivalent).
- GAP-PPT-13: resolved - First-class slide read/navigation retrieval is exposed (`get_slide` equivalent).
- GAP-PPT-14: resolved - First-class presentation-structure read is exposed (`get_presentation_structure` equivalent).
- GAP-PPT-15: resolved - First-class presentation-structure mutation is exposed (`modify_presentation_structure` equivalent).
- GAP-PPT-16: resolved - First-class icon workflow is exposed (`search_icons` and `insert_icon` equivalents).

### 13.5 Excel First-Class Tool Gaps (`GAP-XLS-*`)

- GAP-XLS-01: resolved - First-class cell/range reads are exposed (`get_cell_ranges` equivalent).
- GAP-XLS-02: resolved - First-class cell/range writes are exposed (`set_cell_range` equivalent).
- GAP-XLS-03: resolved - First-class range clearing is exposed (`clear_cell_range` equivalent).
- GAP-XLS-04: resolved - First-class range resizing is exposed (`resize_range` equivalent).
- GAP-XLS-05: resolved - First-class range copy operations are exposed (`copy_to` equivalent).
- GAP-XLS-06: resolved - First-class Excel object mutation is exposed (`modify_object` equivalent).
- GAP-XLS-07: resolved - First-class sheet/workbook structure mutation is exposed (`modify_sheet_structure` equivalent).
- GAP-XLS-08: resolved - First-class CSV export is exposed (`get_range_as_csv` equivalent).
- GAP-XLS-09: partially resolved - `read_range_image` exposes an Office.js active-selection image snapshot. It is not an arbitrary offscreen range renderer; navigate/select the target range first.
- GAP-XLS-10: resolved - First-class workbook data/object search is exposed (`search_data` equivalent).
- GAP-XLS-11: resolved - First-class workbook object inventory is exposed (`get_all_objects` equivalent).
- GAP-XLS-12: resolved - First-class chart XML extraction is exposed (`extract_chart_xml` equivalent).

### 13.6 External Context Gaps (`GAP-EXT-*`)

- GAP-EXT-01: resolved - Connector refresh remains runtime/UI-only via Integrations re-verify (`POST /v1/connectors/reverify`) and is validated through taskpane-stability scenarios.
- GAP-EXT-02: resolved - On-demand `read_skill` is not first-class; packaged skill injection plus explicit saved-document gating is the supported closure path.
- GAP-EXT-03: resolved - `web_search`/`web_fetch` remain out of first-class inventory; bounded hard-read-only research connector flow is the supported closure path.

## 14) Manual Testing Scenarios and Steps

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

### Excel

The following Excel checklist is aligned to `npm run smoke:office -- --list` scenario IDs.

#### Scenario: `excel-context-citations` — Workbook context and citations

- Steps:
  1. Open a workbook with multiple worksheets plus at least one table, chart, and PivotTable.
  2. Run `get_all_objects` with `scope=workbook` to capture workbook object inventory.
  3. Run `search_data` for one table/chart/pivot name and one cited cell/formula term, then validate returned anchors.
  4. Run `office_navigate` to one returned `cell`, `range`, and `sheet` anchor.
- Expected:
  - Object inventory includes workbook-scoped tables, charts, PivotTables, worksheets, and named items.
  - Search results include structured anchors that navigate to the expected workbook location.

#### Scenario: `excel-formatting-and-tables` — Range formatting, borders, and table controls

- Steps:
  1. Select a populated range and run `modify_object` with `operation=format_range` using font/fill/alignment/border options.
  2. Run `modify_object` with `operation=format_table` (or table-filter operations) against an existing table.
  3. Run `get_range_as_csv` for the edited range with `includeFormulas=true` to capture a formula-first auditable export snapshot.
  4. Select the edited range, run `read_range_image`, and confirm active-selection visual payload plus `visuals` image data are returned.
- Expected:
  - Range and table updates remain native Excel object mutations.
  - CSV export and active-selection imagery provide auditable cell-level evidence for the edited region without implying offscreen range rendering.

#### Scenario: `excel-charts-and-pivots` — Existing chart and PivotTable editing

- Steps:
  1. Run `modify_object` with `operation=update_chart` or `operation=create_chart` on a target worksheet chart source.
  2. Run `extract_chart_xml` for the updated chart and inspect returned XML snapshot metadata.
  3. Run `modify_object` with `operation=update_pivot_table`, then `operation=sort_pivot_by_values` and `operation=refresh_pivot_table`.
- Expected:
  - Chart and PivotTable operations execute through native workbook objects.
  - `extract_chart_xml` returns a structured chart metadata XML snapshot tied to the targeted chart.

#### Scenario: `excel-taskpane-stability` — Taskpane focus, scroll, prompt, and connector stability

- Steps:
  1. Send a long prompt from the taskpane while interacting with worksheet selections.
  2. Scroll chat during streaming output and confirm manual scroll control is preserved.
  3. Trigger one connector-backed request, reconnect/reopen the taskpane, and submit a follow-up prompt.
- Expected:
  - Composer focus and prompt submission remain stable across worksheet interactions.
  - Chat scrolling remains responsive during streaming.
  - Reconnected sessions continue to execute prompts and tool calls reliably.

### PowerPoint

The following PowerPoint checklist is aligned to `npm run smoke:office -- --list` scenario IDs.

#### Scenario: `powerpoint-anchor-navigation` — Slide, layout, shape, and notes anchors

- Steps:
  1. Open a deck with notes and at least one named shape on the selected slide.
  2. Run `office_get_context` with `includeFormatting=true` and confirm anchors include slide/shape/layout/master/notes targets. For selected shapes on non-first slides, confirm the shape anchor's `slideId`/`slideIndex` match the owning slide.
  3. Run `office_navigate` with one anchor from each anchor family above.
- Expected:
  - Slide and shape anchors navigate to the exact target.
  - Layout/master/notes anchors report explicit fallback behavior when native direct navigation is partial, and layout editing is described as apply-existing-layout only.

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

### Integrations and External Context

Integrations/external-context validation is intentionally grouped under the published host stability scenarios from `npm run smoke:office -- --list`. Use the scenario IDs below verbatim.

#### Scenario: `word-taskpane-stability` — Connector re-verify and reconnection flow (Word)

- Steps:
  1. Open Integrations, run **Re-verify** for at least one connected connector, and wait for explicit status feedback.
  2. Submit a connector-backed prompt from Word and confirm the response includes connector-sourced context.
  3. Reopen the taskpane and rerun a connector-backed prompt.
- Expected:
  - Re-verify status is surfaced clearly (success/failure details).
  - Connector-backed requests remain read-only and return bounded research context.
  - Reconnected sessions preserve connector availability for follow-up prompts.

#### Scenario: `excel-taskpane-stability` — Connector re-verify and reconnection flow (Excel)

- Steps:
  1. From Excel, run Integrations **Re-verify** for the same connector used in Word.
  2. Submit a connector-backed prompt while interacting with workbook selections.
  3. Reopen the taskpane and repeat the prompt.
- Expected:
  - Re-verify remains available from Excel-hosted taskpane sessions.
  - Connector output remains explicit and does not mutate workbook state.
  - Connector-backed prompting remains stable after reconnect.

#### Scenario: `powerpoint-taskpane-stability` — Connector re-verify and reconnection flow (PowerPoint)

- Steps:
  1. From PowerPoint, run Integrations **Re-verify** and confirm status/log feedback is shown.
  2. Execute a connector-backed prompt while switching slides.
  3. Reopen taskpane and run one additional connector-backed prompt.
- Expected:
  - Re-verify and connector status remain visible in PowerPoint sessions.
  - Connector-backed responses remain bounded to read-only external-context flows.
  - Session reconnect keeps connector-backed prompting operational.

