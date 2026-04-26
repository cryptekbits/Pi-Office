# Independent Taskpane Transition Remediation Plan

## 1) Purpose

This document defines the implementation plan and execution checklist to close regressions observed during the transition from a **companion + taskpane** architecture to an **independent taskpane** runtime.

It is based on multi-worker audits across:
- architecture/runtime parity
- integrations/auth/model behavior
- Office UX/runtime behavior
- protocol flow correctness
- build/sideload/validation operations

### Progress Log

- **2026-04-03 (execution start)**
  - Implemented encrypted-at-rest browser auth storage with legacy plain-text migration path.
  - Implemented checkpoint persistence/load handling in in-process kernel and checkpoint restore hydration in taskpane.
  - Fixed saved-document autonomy prompt mode flag wiring.
  - Added dev preflight script for cert presence/trust + port 3443 availability, and wired into dev scripts.
  - Regenerated lockfile/workspace install state after companion removal (`npm install` pruned removed packages).
  - Replaced connector stubs with browser runtime routes for status/diagnostics/prepare/connect/test/reverify/remove/favorite/scope/logs.
  - Implemented connector export/import preview+apply persistence with conflict handling and encrypted localStorage-backed connector state.
  - Implemented connector OAuth callback completion + reauth recovery flow with state validation and expiring sign-in sessions.
  - Aligned model picker entries to executable/authenticated model states only.
  - Added disconnect-time cancellation for pending interactive requests (ask_user/tool_permission/edit_proposal).
  - Restored `office_capture_viewport` runtime path in taskpane bridge + in-process tool registration with Word-only capability mapping.
  - Added CI workflow gate for install/typecheck/build/bundle budget/manifest validation/office tests.
  - Added taskpane bundle budget checker script and package script (`check:bundle`).
  - Added protocol parity test suite for `ask_user` / `tool_permission` / `edit_proposal` / `rewind` / checkpoint persist-load with runtime polyfills.
  - Updated office test runner to normalize compiled relative ESM imports for Node execution compatibility.
  - Added runtime diagnostics panel in settings to surface recent route/auth/connector request failures.
  - Expanded Office smoke plan matrix with host-specific taskpane stability scenarios (focus/scroll/prompt/connector/reconnect).
  - Re-ran validators: `typecheck`, `build`, `check:bundle`, `validate:manifests`, `test:office`, `preflight:dev`.

---

## 2) Locked Product Decisions (from stakeholder responses)

1. **Connectors**: restore runtime now.
2. **Auth**: restore OAuth now.
3. **Workspace features**: parked/deferred for now (not required in current execution scope).
4. **Credential security**: implement encrypted browser storage (instead of plain localStorage).

---

## 3) Success Criteria

The transition is complete when all are true:

1. Connector flows (prepare/test/connect/oauth/reverify/log/export/import) are functional end-to-end.
2. OAuth provider onboarding works from taskpane UX.
3. Workspace filesystem/coding features are explicitly parked and excluded from the current release gate.
4. Secrets are no longer persisted as plain text in browser storage.
5. Protocol parity gaps are closed (checkpoint persistence, viewport capture, session lifecycle semantics).
6. Known UX regressions are fixed (input/scroll/focus stability in Word desktop).
7. Dev/sideload workflow is reliable and CI catches regressions.

---

## 4) Observation-to-Workstream Mapping

| Observation | Severity | Workstream |
|---|---:|---|
| Connectors runtime disabled/stubbed | High | WS2 Integrations Runtime Parity |
| OAuth unavailable | High | WS1 Auth & Secrets |
| Workspace filesystem/coding tools missing | High | WS3 Workspace Runtime |
| Saved-mode prompt incorrectly treated as unsaved | High | WS4 Protocol/Session Correctness |
| Viewport capture tool missing | Medium | WS4 Protocol/Session Correctness |
| Checkpoint persist/load bridge unhandled | High | WS4 Protocol/Session Correctness |
| Plain-text API keys in localStorage | Critical | WS1 Auth & Secrets |
| Model selection allows failing configurations | High | WS1 + WS2 |
| Focus/scroll instability in Office host | Medium/High | WS5 UX/Host Runtime Stability |
| Office state refresh race potential | Medium | WS5 UX/Host Runtime Stability |
| Port/cert/lockfile/CI operational gaps | Medium | WS6 Build, Sideload, CI Hardening |
| Release packaging/runtime assumptions unclear | Medium | WS6 Build, Sideload, CI Hardening |

---

## 5) Workstreams and Detailed Plan

## WS1) Auth & Secrets Hardening

### Goal
Restore OAuth and remove plain-text secret persistence risks while preserving independent taskpane UX.

### Tasks
1. Define encrypted credential storage abstraction for browser runtime.
   - Replace direct `localStorage` raw value persistence with encrypted payload storage.
   - Include key-rotation and migration path from current storage key.
2. Re-enable OAuth route behavior for taskpane runtime.
   - Implement provider OAuth flow orchestration and callback handling in runtime contract.
   - Ensure auth state propagation to provider/model catalog.
3. Tighten model/provider usability logic.
   - Ensure model picker cannot select models guaranteed to fail due to missing auth/config.
   - Add clear inline status for partially configured providers.
4. Add auth diagnostics surface.
   - Explicit error classes for auth unavailable/missing scopes/expired tokens.

### Deliverables
- Secure auth storage module.
- Restored `/v1/auth/start` behavior in runtime.
- Updated auth status responses and model picker gating.

### Acceptance
- OAuth login succeeds for supported providers.
- Stored secrets are not plain text.
- Selecting/configuring models does not lead to immediate avoidable runtime failures.

---

## WS2) Integrations Runtime Parity

### Goal
Restore real connector runtime behavior (not metadata-only stubs).

### Tasks
1. Replace connector stub endpoints with full runtime implementations:
   - status, diagnostics, prepare, connect, test, reverify, remove, favorite, scope update, logs.
2. Restore connector OAuth start + callback completion.
3. Restore connector export/import semantics:
   - export real configured connectors
   - import preview conflict detection
   - import apply with deterministic results
4. Restore audit logging behavior and diagnostics fidelity.
5. Align UI behavior with capability states:
   - no dead actions
   - explicit capability/state handling

### Deliverables
- Fully operational `/v1/connectors/*` routes.
- Connector state persistence + diagnostics/logging.
- UI parity with runtime capability.

### Acceptance
- Connector setup wizard succeeds end-to-end.
- Test/reverify/logs produce live data.
- Export/import roundtrip restores connector configs and scope metadata.

---

## WS3) Workspace Runtime Restoration (Saved Docs)

**Status:** Parked by stakeholder (deferred).

### Goal
Restore saved-document workspace-aware tooling (filesystem/coding-agent class features).

### Tasks
1. Reintroduce saved-mode workspace binding in session runtime.
   - Build workspace context from document path when saved.
2. Restore workspace tools in runtime tool registration (policy controlled).
3. Reinstate safe boundary controls:
   - workspace root restriction
   - explicit tool permission gating
4. Update system prompt composition and mode transitions:
   - saved vs unsaved behavior must reflect actual document state.
5. Ensure UX reflects active capability:
   - workspace-enabled messaging only when truly available.

### Deliverables
- Saved-doc workspace toolchain functional.
- Correct mode prompts and capability hints.

### Acceptance
- In saved documents, workspace file operations work within allowed boundary.
- In unsaved documents, workspace actions are blocked with clear guidance.

---

## WS4) Protocol & Session Correctness Parity

### Goal
Close missing/changed protocol semantics introduced during in-process conversion.

### Tasks
1. Fix saved-mode autonomy prompt flag handling.
2. Restore checkpoint persistence/load route handling.
3. Restore/replace viewport capture tool registration and execution path.
4. Align session disconnect behavior with pending request cancellation semantics.
5. Verify steer/follow-up handling consistency across streaming/non-streaming transitions.
6. Validate rewind behavior consistency (conversation + document restore paths).

### Deliverables
- Route/event parity checklist marked complete.
- Session behavior parity tests.

### Acceptance
- No silent drops of checkpoint or interactive-response events.
- Viewport-related prompts execute successfully.
- Disconnect/reconnect does not leave stale pending requests.

---

## WS5) UX and Office Host Runtime Stability

### Goal
Stabilize taskpane interaction model in desktop Office host.

### Tasks
1. Eliminate focus-steal regressions.
   - Remove/limit aggressive `window.focus()` usage to necessary interactions.
2. Improve modal accessibility/focus trap consistency for all dialogs/popups.
3. Refine chat autoscroll policy.
   - Avoid forcing smooth auto-scroll when user intentionally scrolled up.
4. Audit wheel/keyboard propagation boundaries across all scrollable regions.
5. De-race office state refresh loop.
   - dedupe/throttle overlapping selection-change + polling updates.

### Deliverables
- Stable input/scroll/focus behavior in Word taskpane.
- Documented host-interaction policy.

### Acceptance
- Typing in taskpane reliably stays in taskpane.
- Chat/messages remain vertically scrollable under heavy streaming.
- No recurring “Office state refresh failed” noise under normal interaction.

---

## WS6) Build, Sideload, and CI Hardening

### Goal
Make local dev + validation deterministic after architecture shift.

### Tasks
1. Resolve lockfile/workspace drift from companion removal.
2. Strengthen dev startup checks:
   - cert existence/trust check
   - hard fail with clear guidance on port 3443 conflicts
3. Ensure manifest/runtime source assumptions are validated pre-sideload.
4. Add CI pipeline gates for:
   - install consistency
   - typecheck
   - build
   - manifests validation
   - Office test/smoke script
5. Add bundle-size monitoring threshold and warnings as a gate.

### Deliverables
- Deterministic setup script and CI workflow.
- Stable sideload startup checklist.

### Acceptance
- Fresh checkout can run dev/sideload with predictable setup.
- CI catches drift/regressions before merge.

---

## 6) Phased Execution Plan

## Phase 0 (Immediate Stabilization, P0)

Focus: close critical correctness and security blockers first.

- WS1: encrypted secret storage + auth migration.
- WS4: saved-mode prompt bug, checkpoint route parity.
- WS5: focus/scroll blocking defects that impact core usability.
- WS6: lockfile consistency + startup preflight checks.

**Exit gate:** no critical/high unresolved issues that block daily use.

## Phase 1 (Core Parity Restoration, P1)

- WS2 full connectors runtime restoration.
- WS1 OAuth restoration in production path.
- WS3 workspace toolchain restoration for saved docs (**parked/deferred**).
- WS4 viewport capture + disconnect semantics alignment.

**Exit gate:** companion-era core capabilities functionally restored in independent runtime.

## Phase 2 (Hardening & Scale, P2)

- WS5 race/perf hardening.
- WS6 CI completeness + bundle budget policies.
- Regression automation for protocol + Office UX smoke suites.

**Exit gate:** release-quality stability and regression-proofing.

---

## 7) Detailed TODO List (Execution Checklist)

## P0 — Must Complete First

- [x] Implement encrypted browser credential storage and migrate existing auth records.
- [x] Remove plain-text secret writes from runtime auth store.
- [x] Fix saved-document autonomy/system prompt mode flag logic.
- [x] Implement checkpoint persist/load route handling in in-process runtime.
- [x] Resolve chat/taskpane focus and vertical scroll reliability regressions.
- [x] Add preflight checks for cert presence/trust and port `3443` conflicts.
- [x] Regenerate and clean lockfile/workspace references after companion removal.
- [ ] Add failing tests for the above regressions before fixes where feasible.

## P1 — Functional Parity

- [x] Restore connector status/prepare/connect/test/reverify/remove/scope/logs implementations.
- [x] Restore connector OAuth start flow and callback completion.
- [x] Restore connector import/export functional behavior.
- [ ] (Parked) Restore workspace coding/file tools in saved-document mode with policy guards.
- [x] Restore viewport capture tool path in runtime + UI capability mapping.
- [x] Align model/provider picker to executable/authenticated states only.
- [x] Ensure pending interactive requests are cancelled/recovered correctly on disconnect.

## P2 — Hardening and Scale

- [x] Add CI workflow for install/typecheck/build/manifest/test.
- [x] Add Office host stability smoke tests (focus, scroll, prompt, connector flows).
- [x] Add protocol parity tests for ask_user/tool_permission/edit_proposal/rewind/checkpoints.
- [x] Add bundle-size budget checks and chunking strategy follow-up.
- [x] Add operational diagnostics panel for runtime route/connector/auth failures.

---

## 8) Validation Matrix (Required Before Sign-off)

## Automated

- [x] `npm run typecheck`
- [x] `npm run build`
- [x] `npm run validate:manifests`
- [x] `npm run test:office`
- [x] targeted parity tests (auth/connectors/checkpoints)

## Manual (Word desktop first)

- [ ] Taskpane load + reconnect behavior after Word restart.
- [ ] Chat input focus, keyboard behavior, and vertical scroll behavior.
- [ ] Model switch + thinking-level updates.
- [ ] OAuth onboarding flow.
- [ ] Connector setup/test/use flow with at least two connector types.
- [ ] (Parked) Saved document workspace operations (read/list/search/edit constrained to workspace).
- [ ] Rewind with checkpoint persistence across refresh/reopen.
- [ ] Viewport capture scenario for formatting/layout prompts.

---

## 9) Risks and Mitigations During Execution

1. **Security migration risk (auth data corruption)**  
   Mitigation: migration guard + backup copy + one-time rollback switch.

2. **Connector runtime complexity in browser context**  
   Mitigation: staged adapter rollout and capability flags per connector.

3. **Workspace tool security scope drift**  
   Mitigation: explicit root guard tests and policy integration tests.

4. **Office host interaction fragility (Word webview quirks)**  
   Mitigation: dedicated smoke suite on targeted Office versions.

---

## 10) Ownership and Tracking Template

Use this for each task ticket:

- **Title**
- **Workstream**
- **Priority** (P0/P1/P2)
- **Owner**
- **Dependencies**
- **Acceptance criteria**
- **Validation evidence** (command output/screenshots)
- **Rollback plan**

---

## 11) Open Clarifications (if needed before implementation starts)

1. Which providers/connectors must be in the **first restored set** (MVP subset) vs later?
2. Do we require backward-compatible migration for any already stored auth/connectors from local companion installs?
3. What Office host/version matrix is mandatory for release sign-off (Word desktop build numbers, web, Mac)?
