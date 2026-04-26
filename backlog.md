# Pi-Office Backlog

## How To Use This Backlog

This is the canonical running backlog for this repository. It captures bugs, features, improvements, security concerns, and testing gaps that future agents must preserve and maintain even when the originating chat context is gone.

Every task must be specific enough for a future agent to understand the problem, why it matters, where to start, what it depends on, and how to know it is finished. Do not delete tasks just because they are old. Mark them `done` or `obsolete` with evidence.

## Task Format

Use this format for every new task:

```md
- [ ] CATEGORY-000: Short task subject
  - Category: Bug | Feature | Improvement | Security | Testing
  - Status: open | in_progress | blocked | done | obsolete
  - Priority: P0 | P1 | P2 | P3
  - Source: Where this came from, with date and context.
  - Details: Complete context for an agent with no chat history.
  - Dependencies: Explicit prerequisite tasks or "None".
  - Subtasks:
    - [ ] Concrete implementation or investigation step.
  - Acceptance Criteria:
    - [ ] Observable condition that proves the task is complete.
  - Notes/Evidence: Files, commands, review notes, or validation results.
```

Checkbox rule: checked boxes are only for tasks whose `Status` is `done` or `obsolete`. Keep active, blocked, or parked work unchecked.

Commit rule: when working on a backlog task, commit that task's code/doc/test changes right away after validation. Keep the commit scoped to that backlog item only, and use conventional commit style with the backlog ID in the subject, for example `fix(BUG-001): make taskpane build independent of local certs`.

## Active Backlog

### Security

- [ ] SECURITY-001: Connector OAuth can be marked connected without a real callback or token exchange
  - Category: Security
  - Status: open
  - Priority: P0
  - Source: 2026-04-26 worktree review.
  - Details: The connector OAuth UI currently allows a user to click a "Complete sign-in" action that sends `approved: true` for a pending OAuth state. The browser runtime then sets `credentialSource = "oauth"` and `oauthConnected = true` without proving that a provider callback occurred or that a usable access token/credential was stored. This can produce false "signed in" connector state and may cause later connector execution to fail silently or misleadingly.
  - Dependencies: None.
  - Subtasks:
    - [ ] Define the supported OAuth contract for connector setup: callback URL, state validation, token/credential source, and failure state.
    - [ ] Remove or gate any UI path that marks OAuth complete without a real callback or credential handoff.
    - [ ] Persist only verified OAuth state and keep expired or incomplete flows visible as `auth_required` or `auth_expired`.
    - [ ] Add tests for success, cancelled, expired, state-mismatch, and no-token OAuth completion paths.
  - Acceptance Criteria:
    - [ ] A connector cannot become `oauthConnected: true` unless a verified callback or credential exchange has completed.
    - [ ] The UI cannot manually approve OAuth completion in a way that bypasses the contract.
    - [ ] Connector status clearly reports incomplete, failed, and expired OAuth flows.
  - Notes/Evidence: Review pointed to `apps/taskpane/src/app/components/IntegrationsSection.tsx` sending `approved: true` and `apps/taskpane/src/lib/runtime/browser-connectors.ts` setting OAuth state from that flag.

### Bugs

- [ ] BUG-001: CI clean checkout can fail because Vite reads gitignored certificate files at config load
  - Category: Bug
  - Status: open
  - Priority: P0
  - Source: 2026-04-26 worktree review.
  - Details: The GitHub Actions workflow runs `npm run build`, which invokes the taskpane Vite build. The current Vite config reads `certs/localhost.pfx` and `certs/passphrase.txt` unconditionally when the config module loads. The `certs/` directory is intentionally gitignored, and CI does not run `npm run prepare:certs` before `npm run build`, so a fresh checkout can fail before the build starts.
  - Dependencies: None.
  - Subtasks:
    - [ ] Change taskpane Vite config so HTTPS cert files are only required for the dev server, not production builds.
    - [ ] Ensure CI either does not need local cert files for build or explicitly generates them before config load.
    - [ ] Add a regression check or CI comment documenting why clean checkout builds do not require `certs/`.
  - Acceptance Criteria:
    - [ ] `npm run build` works on a clean checkout with no `certs/` directory.
    - [ ] `npm run dev` still uses the trusted local HTTPS certs and fails with clear guidance when they are missing.
    - [ ] GitHub Actions can run install, typecheck, build, bundle budget, manifest validation, and Office tests without local-only files.
  - Notes/Evidence: Review pointed to `.github/workflows/ci.yml`, `.gitignore`, and `apps/taskpane/vite.config.ts`.

- [ ] BUG-002: Remote HTTP connectors are marked executable but do not appear to be exposed to the agent
  - Category: Bug
  - Status: open
  - Priority: P0
  - Source: 2026-04-26 worktree review plus transition plan WS2.
  - Details: Connector status metadata reports non-local connectors as browser-executable, but the runtime path that builds companion session connectors filters to `local_stdio`, and the agent-facing `mcp` tool is only registered when companion connector tool names exist. This creates a product gap where remote HTTP connectors may look available in Settings but are not actually callable by Pi during a session.
  - Dependencies: SECURITY-001 if OAuth-backed remote connectors are part of the first fixed path.
  - Subtasks:
    - [ ] Decide and document the execution path for remote HTTP MCP connectors in the independent taskpane architecture.
    - [ ] If browser execution is supported, expose read-only verified remote HTTP connector tools to the agent with the same allow/block policy as local connectors.
    - [ ] If browser execution is not supported yet, mark remote HTTP connectors as configured but not execution-available and update UI messaging.
    - [ ] Add route/runtime tests proving status metadata and actual agent tool availability cannot drift.
  - Acceptance Criteria:
    - [ ] A connector shown as execution-available is callable by Pi in the active session.
    - [ ] A connector not callable by Pi is visibly marked unavailable or setup-only.
    - [ ] Remote HTTP and local stdio connector behavior is covered by tests.
  - Notes/Evidence: Review pointed to `buildCompanionSessionConnectors` filtering `local_stdio`, non-local statuses receiving `executionAvailable: true`, and only companion-backed `mcp` tool registration.

- [ ] BUG-003: Local stdio connector credential and environment propagation is incomplete
  - Category: Bug
  - Status: open
  - Priority: P1
  - Source: 2026-04-26 worktree review.
  - Details: The companion connector bridge resolves credentials, but local stdio runtime creation does not inject resolved manual secrets or detected/env credentials into the child process. It also passes `runtime.env` as the process env for `StdioClientTransport`; if the SDK does not merge with `process.env`, custom connector env may drop `PATH` and other required system variables.
  - Dependencies: SECURITY-001 for secure credential lifecycle and storage expectations.
  - Subtasks:
    - [ ] Trace the MCP SDK `StdioClientTransport` env behavior and confirm whether it merges or replaces `process.env`.
    - [ ] Define how each credential source maps into local stdio process env without leaking secrets in UI/logs.
    - [ ] Merge inherited safe environment variables with connector-specific env when launching local stdio connectors.
    - [ ] Add tests or a local probe harness for manual secret, env-var secret, detected-env secret, and custom env cases.
  - Acceptance Criteria:
    - [ ] Local stdio connectors receive required credentials through the intended env variable.
    - [ ] Custom env does not break command discovery or PATH-dependent launches.
    - [ ] Secrets are not printed in diagnostics, status cards, logs, or exported connector bundles.
  - Notes/Evidence: Review pointed to `apps/companion/src/connector-bridge.ts` computing credentials but building local runtime without credential injection.

- [ ] BUG-004: Office state refresh race and deduping follow-up
  - Category: Bug
  - Status: open
  - Priority: P2
  - Source: `TASKPANE_INDEPENDENT_TRANSITION_REMEDIATION_PLAN.md` WS5 and observation mapping.
  - Details: The transition plan called out office state refresh race potential during the independent taskpane migration. Selection-change handlers, polling, and session state sync can overlap in Office hosts, especially Word desktop, leading to noisy refresh failures or stale document/selection state.
  - Dependencies: None.
  - Subtasks:
    - [ ] Trace all Office state refresh triggers in the taskpane app and host adapters.
    - [ ] Add throttling, deduping, or latest-only cancellation so stale refreshes cannot overwrite newer state.
    - [ ] Ensure errors from transient Office host state are surfaced only when actionable.
    - [ ] Add focused tests or smoke scenarios for rapid selection changes and taskpane reconnect.
  - Acceptance Criteria:
    - [ ] Rapid selection changes do not produce recurring "Office state refresh failed" noise.
    - [ ] Stale refresh responses cannot replace newer session state.
    - [ ] Word desktop smoke testing confirms selection/context updates remain stable.
  - Notes/Evidence: Transition plan lists "Office state refresh race potential" under WS5.

### Features

- [ ] FEATURE-001: Restore saved-document workspace and file tools with policy guards
  - Category: Feature
  - Status: open
  - Priority: P2
  - Source: `TASKPANE_INDEPENDENT_TRANSITION_REMEDIATION_PLAN.md` WS3; parked/deferred by stakeholder.
  - Details: Saved-document mode should eventually expose workspace-aware filesystem/coding-agent class capabilities, but only with strict saved-folder binding and policy controls. Unsaved documents must remain without local filesystem access. Current architecture treats the optional companion as the local file/MCP capability provider, so this task must align with that direction and avoid resurrecting companion-era runtime assumptions.
  - Dependencies: SECURITY-001, BUG-003, and any final architecture decision about optional companion versus browser-only file tooling.
  - Subtasks:
    - [ ] Confirm the intended saved-document workspace capability boundary and whether write/edit tools remain out of scope.
    - [ ] Bind workspace root to the saved document folder, not arbitrary repo or user profile folders.
    - [ ] Enforce root guards for read/list/search operations and reject traversal or absolute escape attempts.
    - [ ] Wire tool permission/autonomy policy so local file actions require the intended approvals.
    - [ ] Update context bar, composer hints, and system prompt text to reflect actual availability.
  - Acceptance Criteria:
    - [ ] Unsaved documents cannot access local files.
    - [ ] Saved documents can access only the allowed document-folder scope when the required local capability provider is connected.
    - [ ] Root-guard tests cover normal paths, traversal attempts, absolute paths, symlinks/junctions if supported, and missing folders.
  - Notes/Evidence: Transition plan marks workspace runtime restoration as parked/deferred; keep this task open until stakeholder explicitly removes or completes the feature.

### Improvements

- [ ] IMPROVEMENT-001: Clean up worktree hygiene for untracked archive and generated artifacts
  - Category: Improvement
  - Status: open
  - Priority: P2
  - Source: 2026-04-26 worktree review.
  - Details: The worktree contains untracked `.factory/`, `.github/`, `archive/`, and other generated or transitional artifacts. In particular, `archive/companion/node_modules` is present under an untracked archive tree. This can confuse future reviews, inflate diffs, and accidentally commit vendored dependencies or generated output.
  - Dependencies: None.
  - Subtasks:
    - [ ] Decide which untracked artifacts are intended source/documentation versus generated local state.
    - [ ] Add ignore rules for generated archive dependency folders and local factory/cache outputs that should never be committed.
    - [ ] If archive source is intentionally retained, keep source files only and exclude built output/dependencies unless there is a documented reason.
    - [ ] Re-run `git status --short` and document the intended remaining untracked files.
  - Acceptance Criteria:
    - [ ] No `node_modules`, built `dist`, caches, or generated local-only files remain staged or visible as intended source.
    - [ ] `git status --short` clearly distinguishes product changes from local/generated artifacts.
    - [ ] `.gitignore` covers repeatable generated artifacts without hiding important source files.
  - Notes/Evidence: Review observed untracked archive/vendor artifacts while inspecting the worktree.

- [ ] IMPROVEMENT-002: Clarify release packaging and runtime assumptions after the independent taskpane transition
  - Category: Improvement
  - Status: open
  - Priority: P2
  - Source: `TASKPANE_INDEPENDENT_TRANSITION_REMEDIATION_PLAN.md` WS6 and risk list.
  - Details: The architecture now separates the independent taskpane from an optional companion. Release packaging must make clear what ships as the Office taskpane, what is optional local companion functionality, how certs/dev-only assets are handled, and what users need for sideload versus packaged deployment.
  - Dependencies: BUG-001 and connector execution decisions in BUG-002.
  - Subtasks:
    - [ ] Document development, sideload, CI, and packaged release runtime assumptions in README or release docs.
    - [ ] Clarify whether the optional companion is distributed as npm, zip, binary, or source-only for now.
    - [ ] Ensure manifest URLs, icon cache busting, shortcut resources, and dev/prod host origins are described accurately.
    - [ ] Add validation steps that prove release docs match actual scripts and manifests.
  - Acceptance Criteria:
    - [ ] A fresh developer can identify which command starts taskpane-only development and which command starts the optional companion.
    - [ ] CI/release docs do not imply dev certs or local-only files are committed.
    - [ ] Optional companion limitations are explicit and do not conflict with taskpane-first architecture.
  - Notes/Evidence: Transition plan lists "Release packaging/runtime assumptions unclear" under WS6.

### Testing

- [ ] TESTING-001: Add regression tests for transition regressions where feasible
  - Category: Testing
  - Status: open
  - Priority: P1
  - Source: `TASKPANE_INDEPENDENT_TRANSITION_REMEDIATION_PLAN.md` P0 checklist.
  - Details: The transition plan marked many core regressions fixed, but left an unchecked task to add failing tests for those regressions where feasible. Future work should ensure that encrypted auth migration, saved-mode prompt behavior, checkpoint route handling, focus/scroll stability, preflight behavior, and lockfile/workspace drift are covered by automated or targeted smoke tests as appropriate.
  - Dependencies: BUG-001 for CI clean-checkout viability.
  - Subtasks:
    - [ ] Inventory which transition fixes already have automated tests.
    - [ ] Add missing tests for saved-document mode prompt/autonomy behavior.
    - [ ] Add route/event tests for checkpoint persist/load and session reconnect behavior if not already covered.
    - [ ] Add preflight script tests or dry-run assertions for missing certs and port conflicts.
    - [ ] Document any cases that remain manual-only because Office desktop host behavior cannot be automated here.
  - Acceptance Criteria:
    - [ ] Each critical transition regression has either an automated regression test or an explicit manual-only rationale.
    - [ ] `npm run test:office` includes the relevant protocol/runtime regression coverage.
    - [ ] CI can run the added tests without local Office desktop.
  - Notes/Evidence: Transition plan P0 line item "Add failing tests for the above regressions before fixes where feasible" remains unchecked.

- [ ] TESTING-002: Complete manual Office desktop validation matrix
  - Category: Testing
  - Status: open
  - Priority: P1
  - Source: `TASKPANE_INDEPENDENT_TRANSITION_REMEDIATION_PLAN.md` validation matrix.
  - Details: The automated validation matrix was marked complete in the transition plan, but the manual Word desktop first scenarios remain unchecked. These scenarios matter because Office taskpane focus, scroll, OAuth, connector execution, checkpoint rewind, and viewport capture behavior can differ in the real desktop host from browser or unit-test behavior.
  - Dependencies: BUG-001, BUG-002, SECURITY-001, and BUG-003 for meaningful connector/OAuth validation.
  - Subtasks:
    - [ ] Validate taskpane load and reconnect behavior after Word restart.
    - [ ] Validate chat input focus, keyboard handling, and vertical scrolling during long responses.
    - [ ] Validate model switch and thinking-level updates.
    - [ ] Validate OAuth onboarding flow after SECURITY-001 is fixed.
    - [ ] Validate connector setup/test/use flow with at least two connector types after BUG-002 and BUG-003 are fixed.
    - [ ] Validate rewind with checkpoint persistence across refresh/reopen.
    - [ ] Validate viewport capture scenario for formatting/layout prompts.
  - Acceptance Criteria:
    - [ ] Manual validation results are recorded in this backlog or a linked tracking doc with date, host, and outcome.
    - [ ] Any failed manual scenario creates or links a separate bug task.
    - [ ] Word desktop first validation is complete before claiming release readiness.
  - Notes/Evidence: Transition plan manual matrix lines remain unchecked.

## Done Or Obsolete Tasks

No tasks have been closed in this backlog yet.
