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

- [x] SECURITY-001: Connector OAuth can be marked connected without a real callback, token exchange, or imported secret
  - Category: Security
  - Status: done
  - Priority: P0
  - Source: 2026-04-26 worktree review; 2026-04-26 provider/auth/privacy subagent review.
  - Details: The connector OAuth UI currently allows a user to click a "Complete sign-in" action that sends `approved: true` for a pending OAuth state. The browser runtime then sets `credentialSource = "oauth"` and `oauthConnected = true` without proving that a provider callback occurred or that a usable access token/credential was stored. Connector bundle import can also recreate an OAuth-connected record from `credentialSource: "oauth"` even though exports intentionally omit secrets. These paths can produce false "signed in" connector state and may cause later connector execution to fail silently or misleadingly.
  - Dependencies: None.
  - Subtasks:
    - [x] Define the supported OAuth contract for connector setup: callback URL, state validation, token/credential source, and failure state.
    - [x] Remove or gate any UI path that marks OAuth complete without a real callback or credential handoff.
    - [x] Reset imported OAuth connectors to `auth_required` unless a verified secure credential migration path exists.
    - [x] Persist only verified OAuth state and keep expired or incomplete flows visible as `auth_required` or `auth_expired`.
    - [x] Add tests for success, cancelled, expired, state-mismatch, and no-token OAuth completion paths.
  - Acceptance Criteria:
    - [x] A connector cannot become `oauthConnected: true` unless a verified callback or credential exchange has completed.
    - [x] The UI cannot manually approve OAuth completion in a way that bypasses the contract.
    - [x] Importing a connector bundle never creates usable OAuth state without a verified credential or token.
    - [x] Connector status clearly reports incomplete, failed, and expired OAuth flows.
  - Notes/Evidence: Review pointed to `addin/apps/taskpane/src/app/components/IntegrationsSection.tsx` sending `approved: true`, `addin/apps/taskpane/src/lib/runtime/browser-connectors.ts` setting OAuth state from that flag, `getExportBundle` exporting `credentialSource` without secrets, and `applyImport` setting `oauthConnected` from the imported credential source. 2026-04-26 hardening removed the visible manual "Complete sign-in" UI path, changed callback completion to remain incomplete without a verified credential handoff, and reset imported OAuth connectors to `oauthConnected: false`. 2026-04-26 closure removed the remaining manual completion callback path from Settings/Integrations UI, replaced the callback contract with `ConnectorOAuthCredentialHandoff`, requires a verified access-token handoff before setting `oauthConnected: true`, preserves OAuth tokens only in secret storage fields, normalizes old/imported OAuth records without tokens back to `auth_required`, and surfaces re-auth reasons through `lastError`/diagnostics. Tests in `addin/scripts/office-tests/src/external-context-gaps.test.ts` cover no-token, cancelled, state-mismatch, expired, success, export-without-secret, and import-reset paths. Validation: `npm run test:office` passed with 92 tests; `npm run typecheck` passed.

- [x] SECURITY-002: Tool permission prompts time out as allow-by-default
  - Category: Security
  - Status: done
  - Priority: P0
  - Source: 2026-04-26 product-goal review and provider/auth/privacy subagent review.
  - Details: `requestToolPermission` resolves a pending approval as `{ allowed: true, scope: "once" }` after 120 seconds. Silence, hidden taskpane UI, host disconnect, or an unattended machine should not authorize document writes, connector calls, external reads, or future write-external tools. This conflicts with the privacy-conscious and reviewable-workbench product goal.
  - Dependencies: None.
  - Subtasks:
    - [x] Change permission timeout behavior to deny, expire, or abort the tool call instead of allowing it.
    - [x] Surface expired permission requests clearly in the UI and model/tool result.
    - [x] Ensure disconnect, session teardown, and pending-request cleanup cannot produce allow decisions.
    - [x] Add tests for timeout behavior across write-doc, connector, read-external, and write-external categories.
  - Acceptance Criteria:
    - [x] No permission request can become allowed without an explicit user action or a pre-existing approved policy.
    - [x] Timed-out requests are visible as expired/denied and do not execute.
    - [x] Automated tests prove timeout, disconnect, and cleanup paths fail closed.
  - Notes/Evidence: Review pointed to `addin/apps/taskpane/src/lib/runtime/inprocess-kernel.ts` resolving allowed on timeout and `addin/packages/pi-office-pack/src/protocol.ts` defining tool categories/autonomy levels. 2026-04-26 hardening changed timeout resolution to `allowed: false`, added a `tool_permission_expired` bridge event, clears the visible prompt when the matching request expires, and added a static regression test in `addin/scripts/office-tests/src/runtime-permission-policy.test.ts`. 2026-04-26 closure added protocol-level coverage in `addin/scripts/office-tests/src/protocol-parity.test.ts` proving write-doc, connector, read-external, and write-external permission timeouts resolve denied; existing disconnect coverage rejects pending permission prompts when the bridge closes; new cleanup coverage rejects pending permissions when a session is force-reopened/disposed. Validation: `npm run test:office` passed with 89 tests.

- [x] SECURITY-003: `office_execute_js` raw Office.js escape hatch can be auto-approved as a normal document write
  - Category: Security
  - Status: done
  - Priority: P0
  - Source: 2026-04-26 product-goal review; 2026-04-26 Office-host subagent review.
  - Details: `office_execute_js` executes arbitrary Office.js snippets through `new Function` after a best-effort regex blocklist. The tool is categorized as `write-doc`, and the default user preference is medium autonomy, which auto-approves `write-doc`. This makes the raw escape hatch much easier to invoke than its risk profile warrants. Structured native tools should remain the default path, and raw Office.js should require an explicit owner/user decision.
  - Dependencies: SECURITY-002 for fail-closed permission prompts.
  - Subtasks:
    - [x] Give `office_execute_js` a separate high-risk category or disable it by default.
    - [x] Require explicit per-call approval for `office_execute_js` regardless of medium/high document-write autonomy.
    - [x] Add UI and prompt wording that frames it as an escape hatch, not a normal edit tool.
    - [x] Add tests proving default autonomy cannot auto-approve `office_execute_js`.
    - [x] Inventory existing structured Office tools and route common use cases away from raw execution.
  - Acceptance Criteria:
    - [x] `office_execute_js` cannot execute under default medium autonomy without explicit approval.
    - [x] Regex gating is not presented as an isolated sandbox.
    - [x] Tests cover auto-approval, explicit approval, explicit denial, and blocked-code paths.
  - Notes/Evidence: Review pointed to `addin/apps/taskpane/src/lib/office/document-tools.ts` using `new Function`, `addin/packages/pi-office-pack/src/protocol.ts` mapping `office_execute_js` to `write-doc`, and the default `autonomyLevel: "medium"`. 2026-04-26 hardening introduced the `escape-hatch` tool category, moved `office_execute_js` into it, excluded it from all autonomy auto-approval sets, and added tests proving it stays manual-only. 2026-04-26 closure normalized escape-hatch permission responses to `scope: "once"` even if a client sends `scope: "session"`, limited the permission popup to one-time approval for escape-hatch requests, and added protocol/UI contract tests for medium/high/extreme autonomy, explicit denial, explicit approval, attempted session approval, and blocked-code paths. Existing first-class tool inventory/registration tests and tool descriptions route common Office work toward structured tools before raw execution. Validation: `npm run test:office` passed with 91 tests; `npm run typecheck:taskpane` passed.

- [x] SECURITY-004: Add taskpane CSP/security policy and public privacy/storage disclosure
  - Category: Security
  - Status: done
  - Priority: P1
  - Source: 2026-04-26 competitor review against ChatGPT, Claude, and Ghostwriter inspiration captures.
  - Details: The taskpane is intended to be open-source and privacy-conscious, but the current implementation stores provider API keys, connector state/secrets, and chat history in browser storage. API key and connector envelopes use AES-GCM, but the crypto keys are also stored in localStorage, so this is local obfuscation rather than strong protection against same-origin script access or XSS. The taskpane shell also lacks an obvious CSP comparable to the ChatGPT inspiration capture. Users need a plain privacy/auth panel and docs that state what leaves the machine, what stays in browser storage, provider/connector call boundaries, telemetry defaults, clear-data controls, and the limits of localStorage encryption.
  - Dependencies: SECURITY-001, BUG-002, BUG-005 for accurate connector/provider disclosures.
  - Subtasks:
    - [x] Add a taskpane CSP or equivalent deployment security-header policy compatible with Office add-in hosts.
    - [x] Add in-app privacy/auth disclosure covering provider calls, connector calls, local chat history, audit logs, and localStorage credential limits.
    - [x] Add clear-data controls for provider auth, connector config/logs, and chat history.
    - [x] Document telemetry defaults and ensure any telemetry or analytics are opt-in and visibly disclosed.
    - [x] Evaluate stronger storage options for packaged builds or the optional companion, such as OS keychain/token broker storage.
  - Acceptance Criteria:
    - [x] Users can tell where keys, prompts, document snippets, connector requests, and chat history are stored or sent.
    - [x] A user can clear locally stored sensitive state from the UI.
    - [x] The taskpane has an explicit CSP/security policy or a documented Office-host-compatible reason why a different mechanism is used.
  - Notes/Evidence: Review pointed to `BrowserAuthStore` and connector storage storing crypto keys in localStorage, `useChatHistory` persisting messages to localStorage, and inspiration add-ins with more explicit CSP/privacy surfaces. Closed 2026-04-26 by adding a taskpane `Content-Security-Policy` meta policy in `addin/apps/taskpane/index.html`, a Settings -> Privacy tab that discloses provider calls, connector calls, saved chat history, telemetry defaults, and the limits of localStorage-held AES-GCM keys, clear-data controls for all provider credentials, all connector config/secrets/scopes/OAuth state/logs, and saved local chat history, runtime routes `DELETE /v1/auth` and `DELETE /v1/connectors` that also remove their local crypto keys, and `docs/privacy-and-storage.md` linked from `README.md`. Stronger OS keychain/token-broker storage remains future provider/advanced-companion work and is documented as the intended hardening path. Regression coverage added in `addin/scripts/office-tests/src/privacy-storage.test.ts` for the CSP declaration, provider auth clear-all, connector clear-all, and chat-history clearing. Validation: `npm run typecheck:addin` passed; `npm run test:office` passed with 105 tests; `npm run build:addin` passed; `npm run typecheck:companion` passed; `npm run check:bundle` passed with `main.js=1436.3 KiB`; `npm run validate:manifests` validated Word, Excel, and PowerPoint.

- [x] SECURITY-005: Maintain originality/provenance audit for competitor-inspired capabilities
  - Category: Security
  - Status: done
  - Priority: P0
  - Source: 2026-04-26 plan implementation after user raised DMCA/IP risk from Claude, ChatGPT, Copilot, and other add-in inspiration.
  - Details: Pi-Office may study competitor add-ins to understand user expectations and capability gaps, but public release should not include copied source code, bundled assets, prompts, UI text, private API contracts, or distinctive expression from those products. Each competitor-inspired capability needs a short provenance trail showing the public API basis, original Pi-Office design, original implementation, and replacement status for any borrowed-looking artifact.
  - Dependencies: None.
  - Subtasks:
    - [x] Create or maintain a provenance matrix covering local inspiration captures and major public competitors.
    - [x] Audit prompts, UI strings, assets, icons, and code for copied or near-copied competitor expression.
    - [x] Replace any copied-looking artifact with original Pi-Office text, design, or licensed/public-domain material.
    - [x] Add contribution guidance so external contributors understand idea-level inspiration versus copying expression.
  - Acceptance Criteria:
    - [x] Public-release files have no copied competitor code, prompts, private API contracts, or bundled assets unless explicitly licensed.
    - [x] Every competitor-inspired feature can be traced to an original Pi-Office spec and implementation.
    - [x] `AGENTS.md` and contributor-facing docs explain the originality policy.
  - Notes/Evidence: 2026-04-26 added `AGENTS.md` originality/provenance policy and started `docs/provenance.md`. 2026-04-26 closure expanded `docs/provenance.md` into a release audit with scope, evidence commands, findings, capability provenance matrix, asset notes, contribution rules, and release gates; added `CONTRIBUTING.md` originality guidance; linked provenance docs from `README.md`; and replaced the copied-looking `Claude Excel inspired` CSS comment with neutral original wording. Audit commands searched active release source for competitor/private-API/prompt references and enumerated active prompt/skill/image/asset files. No active source file was found containing competitor code, copied prompt bundles, private endpoint contracts, or competitor-bundled assets. Separate third-party connector logo licensing/source review is tracked in `SECURITY-007`.

- [x] SECURITY-006: Design and enforce read-only companion shell sandbox before exposing bash
  - Category: Security
  - Status: done
  - Priority: P0
  - Source: 2026-04-26 plan implementation after user asked for Codex/OpenCode/pi sandbox analysis and a safe companion bridge design.
  - Details: SOTA coding models often use bash for analysis, but raw host bash from an Office chat can be prompt-injected into modifying files outside the document workspace. Pi-Office must not expose raw host bash. If shell capability is added, it must route through a companion-owned sandbox using Pi's pluggable tool operations, with user-selected roots mounted/readable only, writable scratch separated from user files, denied secret patterns, no network by default, timeouts/output caps, command audit logs, and fail-closed approvals. On Windows, prefer WSL2/Docker/Hyper-V-backed isolation until native sandbox behavior is proven.
  - Dependencies: FEATURE-006 and SECURITY-002.
  - Subtasks:
    - [x] Specify the companion shell policy, including readable roots, writable scratch, denied patterns, network policy, and command categories.
    - [x] Implement a custom Pi `BashOperations` backend that routes through the companion sandbox instead of local raw bash.
    - [x] Add Windows, WSL2/Linux, and macOS capability detection with safe fallback to "shell unavailable".
    - [x] Add destructive sandbox probes for write/delete outside scratch, `.env` reads, symlink escape, network calls, package installs, and git push.
  - Acceptance Criteria:
    - [x] No raw host bash is available from the add-in or companion by default.
    - [x] Shell commands cannot modify user workspace/document files outside the approved scratch path.
    - [x] Sandbox probes prove denied filesystem, secret, and network operations fail closed before shell is enabled.
  - Notes/Evidence: Codex uses platform sandbox modes and Linux bubblewrap/read-only defaults; OpenCode uses permission-driven plan/build modes. Pi-Office should combine both lessons: OS/process isolation first, permission prompts second. 2026-04-26 first slice added `docs/companion-shell-sandbox.md` with the required policy, platform backends, protocol contract, and destructive probes; README then stated shell/bash was unavailable until that policy was implemented; `addin/scripts/office-tests/src/external-context-gaps.test.ts` asserted the active session did not expose `bash`, `edit`, or `write` tools and the companion server had no shell/bash/exec route. 2026-04-26 closure added shared shell capability/request/result protocol types, `companion/src/shell-sandbox.ts` with fail-closed backend detection, policy validation, destructive probes, environment scrubbing, output caps, timeout plumbing, and `createCompanionBashOperations`; companion health/session routes now expose shell capability and a sandboxed execute endpoint that returns unavailable/denied unless the sandbox state is `available`; the taskpane only publishes `bash` for saved documents with a connected companion whose sandbox capability is available. Tests in `addin/scripts/office-tests/src/companion-shell-sandbox.test.ts` cover default-disabled detection, degraded Windows/Linux/macOS-style fallback, destructive probes, scratch-only writes, unavailable execution, env scrubbing, output caps, timeouts, and the BashOperations adapter. `addin/scripts/office-tests/src/external-context-gaps.test.ts` proves default sessions still expose no raw `bash`, `edit`, or `write` tools and that companion shell routing goes through `CompanionShellSandbox`. Validation: `npm run test:office` passed with 98 tests; `npm run typecheck:taskpane` and `npm run typecheck:companion` passed.

- [x] SECURITY-007: Audit third-party connector logo licensing and source provenance before public packaging
  - Category: Security
  - Status: obsolete
  - Priority: P1
  - Source: 2026-04-26 `SECURITY-005` provenance audit.
  - Details: The active taskpane contains connector logo/image assets under `addin/apps/taskpane/public/connectors`. They are not copied competitor-add-in assets, but they are third-party vendor-identification marks and should have source/license notes or neutral fallback badges before a public package/release is cut.
  - Dependencies: None.
  - Subtasks:
    - [x] Inventory every connector image/SVG under `addin/apps/taskpane/public/connectors`.
    - [x] Record source, license, trademark usage note, and replacement/fallback plan for each asset.
    - [x] Replace any asset that lacks acceptable source/license provenance with an original neutral badge or generated non-brand icon.
    - [x] Add a release check or doc section proving packaged connector assets match the approved inventory.
  - Acceptance Criteria:
    - [x] Public release artifacts do not bundle connector brand marks without documented source/license/trademark review.
    - [x] Any unapproved connector asset has a neutral fallback in the taskpane.
    - [x] `docs/provenance.md` or a linked asset inventory records the final approved state.
  - Notes/Evidence: `docs/provenance.md` now distinguishes competitor provenance from third-party connector logo licensing and explicitly tracks this as the remaining asset provenance gap. Obsoleted 2026-04-26 by stakeholder decision: connector logos are acceptable as-is and do not need a licensing/source audit for the current release path.

### Bugs

- [x] BUG-001: CI clean checkout can fail because Vite reads gitignored certificate files at config load
  - Category: Bug
  - Status: done
  - Priority: P0
  - Source: 2026-04-26 worktree review.
  - Details: The GitHub Actions workflow runs `npm run build`, which invokes the taskpane Vite build. The current Vite config reads `certs/localhost.pfx` and `certs/passphrase.txt` unconditionally when the config module loads. The `certs/` directory is intentionally gitignored, and CI does not run `npm run prepare:certs` before `npm run build`, so a fresh checkout can fail before the build starts.
  - Dependencies: None.
  - Subtasks:
    - [x] Change taskpane Vite config so HTTPS cert files are only required for the dev server, not production builds.
    - [x] Ensure CI either does not need local cert files for build or explicitly generates them before config load.
    - [x] Add a regression check or CI comment documenting why clean checkout builds do not require `certs/`.
  - Acceptance Criteria:
    - [x] `npm run build` works on a clean checkout with no `certs/` directory.
    - [x] `npm run dev` still uses the trusted local HTTPS certs and fails with clear guidance when they are missing.
    - [x] GitHub Actions can run install, typecheck, build, bundle budget, manifest validation, and Office tests without local-only files.
  - Notes/Evidence: Review pointed to `.github/workflows/ci.yml`, `.gitignore`, and `addin/apps/taskpane/vite.config.ts`. Fixed in `a8a5a0d` and `43b2a5b`: CI now runs install/typecheck/build/bundle-budget/manifest validation/Office tests, preflight scripts validate cert and sideload resources, and Vite reads cert files only for `serve`. Validation included `npm run typecheck`, `npm run build`, `npm run check:bundle`, `npm run validate:manifests`, `npm run test:office`, `npm run preflight:dev`, and `npm run build:taskpane` with `certs/` temporarily absent.

- [x] BUG-002: Remote HTTP connectors are marked executable but do not appear to be exposed to the agent
  - Category: Bug
  - Status: done
  - Priority: P0
  - Source: 2026-04-26 worktree review plus transition plan WS2.
  - Details: Connector status metadata reports non-local connectors as browser-executable, but the runtime path that builds companion session connectors filters to `local_stdio`, and the agent-facing `mcp` tool is only registered when companion connector tool names exist. This creates a product gap where remote HTTP connectors may look available in Settings but are not actually callable by Pi during a session.
  - Dependencies: SECURITY-001 if OAuth-backed remote connectors are part of the first fixed path.
  - Subtasks:
    - [x] Decide and document the execution path for remote HTTP MCP connectors in the independent taskpane architecture.
    - [x] Document that browser execution is not supported yet instead of exposing unimplemented remote HTTP tools.
    - [x] Mark remote HTTP connectors as configured but not execution-available and update UI messaging.
    - [x] Add route/runtime tests proving status metadata and actual agent tool availability cannot drift.
  - Acceptance Criteria:
    - [x] A connector shown as execution-available is callable by Pi in the active session.
    - [x] A connector not callable by Pi is visibly marked unavailable or setup-only.
    - [x] Remote HTTP and local stdio connector behavior is covered by tests.
  - Notes/Evidence: Review pointed to `buildCompanionSessionConnectors` filtering `local_stdio`, non-local statuses receiving `executionAvailable: true`, and only companion-backed `mcp` tool registration. 2026-04-26 fix chose setup-only browser behavior until remote MCP execution exists, updated runtime metadata/diagnostics/README, and added `remote HTTP connectors are marked setup-only until browser execution exists` in `addin/scripts/office-tests/src/external-context-gaps.test.ts`.

- [x] BUG-003: Local stdio connector credential and environment propagation is incomplete
  - Category: Bug
  - Status: done
  - Priority: P1
  - Source: 2026-04-26 worktree review.
  - Details: The companion connector bridge resolves credentials, but local stdio runtime creation does not inject resolved manual secrets or detected/env credentials into the child process. It also passes `runtime.env` as the process env for `StdioClientTransport`; if the SDK does not merge with `process.env`, custom connector env may drop `PATH` and other required system variables.
  - Dependencies: SECURITY-001 for secure credential lifecycle and storage expectations.
  - Subtasks:
    - [x] Trace the MCP SDK `StdioClientTransport` env behavior and confirm whether it merges or replaces `process.env`.
    - [x] Define how each credential source maps into local stdio process env without leaking secrets in UI/logs.
    - [x] Merge inherited safe environment variables with connector-specific env when launching local stdio connectors.
    - [x] Add tests or a local probe harness for manual secret, env-var secret, detected-env secret, and custom env cases.
  - Acceptance Criteria:
    - [x] Local stdio connectors receive required credentials through the intended env variable.
    - [x] Custom env does not break command discovery or PATH-dependent launches.
    - [x] Secrets are not printed in diagnostics, status cards, logs, or exported connector bundles.
  - Notes/Evidence: Review pointed to `companion/src/connector-bridge.ts` computing credentials but building local runtime without credential injection. 2026-04-26 trace confirmed the installed MCP SDK merges `getDefaultEnvironment()` with supplied stdio env; Pi-Office now also explicitly builds the stdio env from the SDK safe inherited key list, merges connector env, and injects manual/env/detected credentials only under the intended env variable. Manual local credentials with no env target now stay `auth_required` with `credential_env_key_required`, and taskpane setup preserves the local env target for manual secrets. Regression coverage added in `addin/scripts/office-tests/src/companion-connector-bridge.test.ts`; validation passed with `npm run typecheck:companion`, `npm run typecheck:taskpane`, `npm run test:office`, `npm run build`, `npm run validate:manifests`, and `npm run check:bundle`.

- [ ] BUG-004: Office state refresh race and deduping follow-up
  - Category: Bug
  - Status: open
  - Priority: P2
  - Source: `docs/TASKPANE_INDEPENDENT_TRANSITION_REMEDIATION_PLAN.md` WS5 and observation mapping.
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

- [x] BUG-005: Provider readiness reports stored credentials as ready without validating usability
  - Category: Bug
  - Status: done
  - Priority: P1
  - Source: 2026-04-26 product-goal review and provider/auth/privacy subagent review.
  - Details: Provider/model status treats the presence of a stored API key as configured/ready. The taskpane accepts and persists a string, `BrowserModelRegistry` reports models as configured via `hasAuth(providerId)`, and Settings displays ready counts from that flag. A mistyped, expired, revoked, or incompatible key can therefore appear ready until the first real model call fails. For a provider-flexible product, "stored" and "verified usable" need separate states.
  - Dependencies: FEATURE-002 for broader provider-auth capability modeling.
  - Subtasks:
    - [x] Split provider auth state into at least `credentialStored`, `verifiedUsable`, `verificationFailed`, and `notConfigured`.
    - [x] Add a lightweight validation path where provider APIs support it, or mark keys unverified until the first successful request.
    - [x] Demote provider/model readiness after 401/403/auth failures and surface actionable recovery text.
    - [x] Update Settings labels so unverified credentials do not read as fully ready.
    - [x] Add tests for saved-but-invalid, verified, expired/revoked, and recovered provider credentials.
  - Acceptance Criteria:
    - [x] A newly saved key is not labeled fully ready unless it has been verified or successfully used.
    - [x] Auth failures update provider status visibly.
    - [x] Model selection cannot imply a provider is usable when only an unverified credential string exists.
  - Notes/Evidence: Review pointed to `addin/apps/taskpane/src/lib/runtime/inprocess-kernel.ts` accepting API keys and using `hasAuth`, plus `addin/apps/taskpane/src/app/components/SettingsPage.tsx` deriving ready counts from model `configured`. Closed 2026-04-26 by adding shared provider auth states (`not_configured`, `credential_stored`, `verified_usable`, `verification_failed`), preserving legacy `configured` as credential-present while exposing explicit `credentialStored`/`verifiedUsable` metadata, migrating old stored keys to unverified, promoting providers after successful model/image requests, demoting 401/403/auth failures, and updating Settings/model-picker labels to show `Unverified` or `Auth failed` instead of `Ready`. Regression coverage added in `addin/scripts/office-tests/src/provider-auth-readiness.test.ts` for newly saved, legacy stored, auth-failed, and recovered credentials. Validation: `npm run typecheck:addin` passed; `npm run test:office` passed with 101 tests; `npm run build:addin` passed; `npm run typecheck:companion` passed; `npm run check:bundle` passed with `main.js=1436.3 KiB`; `npm run validate:manifests` validated Word, Excel, and PowerPoint.

- [x] BUG-006: Image-generation UI and catalog imply providers that browser runtime cannot execute
  - Category: Bug
  - Status: done
  - Priority: P2
  - Source: 2026-04-26 product-goal review and provider/auth/privacy subagent review.
  - Details: Settings copy tells users to configure OpenAI, Google, or OpenRouter for image models, and protocol constants include multiple image API styles. The browser runtime currently hard-errors unless the selected image model provider is OpenAI. This creates capability drift for visual reasoning and diagram/image workflows.
  - Dependencies: FEATURE-002 if non-OpenAI image providers are implemented through the broader provider matrix.
  - Subtasks:
    - [x] Decide whether v1 image generation is OpenAI-only or multi-provider.
    - [x] Restrict catalog/UI copy to OpenAI and explain other providers are not implemented yet.
    - [x] Defer Google/OpenRouter image execution paths behind future provider-specific runtime work.
    - [x] Add tests proving the configured image model catalog matches executable providers.
  - Acceptance Criteria:
    - [x] The UI never lists an image provider as usable unless runtime execution exists.
    - [x] Unsupported image providers fail at configuration/catalog time with clear messaging, not only during generation.
    - [x] Tests cover image-provider catalog/runtime consistency.
  - Notes/Evidence: Review pointed to `addin/apps/taskpane/src/app/components/SettingsPage.tsx` mentioning OpenAI/Google/OpenRouter and `addin/apps/taskpane/src/lib/runtime/inprocess-kernel.ts` throwing for non-OpenAI image generation. Closed 2026-04-26 by making Settings image-generation copy OpenAI-only, keeping `/v1/image-models` OpenAI-only, adding `imageGenerationSupported` provider capability metadata, and rejecting unsupported `defaultImageModel` preference writes before generation. Regression coverage in `addin/scripts/office-tests/src/provider-auth-readiness.test.ts` proves the image catalog is OpenAI-only and non-catalog image model preferences fail during configuration. Validation passed: `npm run typecheck:addin`, `npm run test:office`, `npm run build`, `npm run check:bundle`, and `npm run validate:manifests`.

- [x] BUG-007: Visual capture tools overstate screenshot and range-image fidelity
  - Category: Bug
  - Status: done
  - Priority: P1
  - Source: 2026-04-26 product-goal review; 2026-04-26 Office-host subagent review.
  - Details: The product goal depends on vision and layout reasoning for diagrams, images, pitch decks, resumes, and polished documents. Current Word viewport capture is explicitly metadata/context-derived and cannot capture a pixel-perfect window/page screenshot after the companion-based window capture was removed. Excel `read_range_image` is described as first-class range imagery but falls back to generic selected-image capture rather than rendering/copying the requested range. Settings also calls `office_capture_snapshot` a visual screenshot even though the actual contract is selection/context snapshots plus metadata.
  - Dependencies: TESTING-002 for manual Office desktop validation.
  - Subtasks:
    - [x] Reconcile tool names/descriptions so they match current fidelity exactly.
    - [x] Decide whether to restore a safe owner-approved viewport/window capture capability or keep metadata-only capture.
    - [x] Implement a real Excel range image path or downgrade `read_range_image` messaging until one exists.
    - [x] Add visual QA tests/manual scripts for Word viewport and Excel range-image scenarios.
  - Acceptance Criteria:
    - [x] Tool descriptions and Settings copy do not claim pixel screenshots or range imagery unless actually produced.
    - [x] Word layout prompts clearly distinguish metadata/context-derived views from true screenshots.
    - [x] Excel visual range workflows either return a real image of the requested range or report the limitation clearly.
  - Notes/Evidence: Review pointed to `office_capture_viewport` comments in `addin/apps/taskpane/src/lib/office-bridge.ts`, `read_range_image` fallback in `addin/apps/taskpane/src/lib/office/excel-context.ts`, and Settings tool descriptions. Closed 2026-04-26 by downgrading snapshot/range-image tool descriptions to Office.js context snapshots, documenting `read_range_image` as an active-selection snapshot rather than arbitrary offscreen range rendering, and making requested range mismatches fail with guidance to select/navigate first. Regression coverage added in `addin/scripts/office-tests/src/excel-object-export-tools.test.ts`; `docs/CLAUDE_ADDIN_INVESTIGATION_AND_TRACKING.md` now records the partial active-selection fidelity. Validation passed: `npm run typecheck:addin`, `npm run test:office` with 114 tests, `npm run check:bundle`, `npm run build`, and `npm run validate:manifests`.

- [x] BUG-008: Excel rewind restores values and number formats but drops formulas
  - Category: Bug
  - Status: done
  - Priority: P1
  - Source: 2026-04-26 Office-host subagent review.
  - Details: Excel checkpoint capture stores formulas, values, and number formats, but restore writes only `range.values` and `range.numberFormat`. For DCFs, financial models, and analytical workbooks, a rewind that flattens formulas into values can silently destroy the model while appearing successful.
  - Dependencies: None.
  - Subtasks:
    - [x] Define the intended Excel checkpoint fidelity contract for formulas, formats, tables, charts, validations, and workbook structure.
    - [x] Restore formulas when formula data exists, preserving values only where formulas are absent.
    - [x] Add safety messaging when a snapshot cannot fully restore workbook semantics.
    - [x] Add tests for formula preservation, mixed formula/value ranges, and number-format preservation.
  - Acceptance Criteria:
    - [x] Rewinding an Excel checkpoint preserves formulas for captured formula cells.
    - [x] The tool reports any unsupported workbook elements that were not restored.
    - [x] Tests protect DCF-style workbook formulas from value-only flattening.
  - Notes/Evidence: Review pointed to `addin/apps/taskpane/src/lib/office/document-tools.ts` loading `formulas` during capture but restoring only values and number formats. 2026-04-26 fix changed Excel restore to prefer the captured `range.formulas` matrix, which includes both formula cells and constant cells per Office.js, and falls back to values only when formula data is missing or shape-mismatched. Restore now returns warnings for values-only fallback and for the current checkpoint fidelity boundary: used-range formulas/constants/number formats only, not full replay of tables, charts, data validation, or workbook structure. The taskpane surfaces those restore warnings as system messages. Regression tests in `addin/scripts/office-tests/src/excel-rewind.test.ts` prove mixed formula/value matrices write `range.formulas`, values are not used when formulas are available, number formats are restored, and values-only fallback warns. Validation passed with `npm run typecheck:addin`, `npm run test:office`, `npm run build`, `npm run validate:manifests`, and `npm run check:bundle`.

- [x] BUG-009: PowerPoint shape anchoring and slide-master tooling can mislead agents
  - Category: Bug
  - Status: done
  - Priority: P1
  - Source: 2026-04-26 Office-host subagent review.
  - Details: PowerPoint selected-shape descriptors and anchors currently appear to attach selected shapes to `slides.items[0]`, which can produce bad anchors for cross-slide or multi-slide operations. Separately, `edit_slide_master` is named and described as layout/master editing, but the bridge currently supports only `apply_layout`. These gaps can make agents confidently target the wrong slide or overpromise master/layout mutation.
  - Dependencies: TESTING-002 for manual PowerPoint validation if desktop behavior differs from tests.
  - Subtasks:
    - [x] Verify selected-shape slide ownership for single-slide, multi-slide, and cross-slide selection scenarios.
    - [x] Fix selected-shape descriptors/anchors to include the actual owning slide where Office.js exposes it.
    - [x] Rename or narrow `edit_slide_master`, or implement real master/layout mutation beyond `apply_layout`.
    - [x] Add tests or fixtures for selected shape anchors and slide-layout operations.
  - Acceptance Criteria:
    - [x] Selected shape anchors resolve to the actual slide instead of defaulting to the first slide.
    - [x] `edit_slide_master` naming and behavior match exactly.
    - [x] Tests cover at least one multi-slide or non-first-slide shape operation.
  - Notes/Evidence: Review pointed to `addin/apps/taskpane/src/lib/office/powerpoint-context.ts` using `slides.items[0]` for selected shapes and `addin/apps/taskpane/src/lib/office-bridge.ts` limiting `edit_slide_master` to apply-layout operations. Closed 2026-04-26 by resolving selected shape descriptors and anchors through `PowerPoint.Shape.getParentSlideOrNullObject()` when Office.js exposes it, falling back to the selected slide only when parent-slide metadata is unavailable. `edit_slide_master` is now labeled and described as legacy-named apply-existing-layout only; `set_slide_master`/`apply_master` style operations fail with a clear no-master-editing error. Regression coverage was added for non-first-slide shape metadata, tool labels/descriptions in both extension and in-process runtime, and unsupported master-edit operations. Validation passed: `npm run typecheck:addin`, `npm run test:office` with 115 tests, `npm run check:bundle`, `npm run build`, and `npm run validate:manifests`.

- [x] BUG-010: Add-in install and build gates emit stale dependency and bundle warnings
  - Category: Bug
  - Status: done
  - Priority: P1
  - Source: 2026-04-26 follow-up after the repository split; user reported that `npm ci --prefix addin` still emitted npm audit findings and build still emitted the existing Vite large-chunk warning.
  - Details: After moving the add-in into `addin/`, the clean add-in install still reported vulnerable transitive packages and the taskpane build still printed Vite's large-chunk warning. The audit noise came from local Office CLI dev dependencies plus Mermaid's vulnerable `uuid` transitive. Removing the local Office CLI chain required keeping sideload/manifest validation available through pinned on-demand CLI execution. The Vite warning should be aligned with the existing bundle-budget gate and the initial bundle should stay below that gate.
  - Dependencies: None.
  - Subtasks:
    - [x] Remove vulnerable Office CLI packages from the local `addin` install path while preserving sideload and manifest validation commands.
    - [x] Override Mermaid's `uuid` transitive to the patched version without leaving an invalid npm dependency tree.
    - [x] Upgrade Vite to a patched 8.0.x release.
    - [x] Lazy-load Mermaid rendering so the initial taskpane bundle remains below the existing budget.
    - [x] Align Vite's chunk warning limit with the checked bundle-budget threshold.
  - Acceptance Criteria:
    - [x] `npm ci --prefix addin` completes with `found 0 vulnerabilities`.
    - [x] `npm run build:addin` completes without Vite's large-chunk warning.
    - [x] `npm run check:bundle` passes with the main JS bundle below budget.
    - [x] Existing typecheck, manifest validation, and Office regression tests still pass.
  - Notes/Evidence: Fixed by moving sideload scripts to pinned `npx --yes office-addin-debugging@6.0.7`, making manifest validation invoke pinned `office-addin-manifest@2.1.3` on demand, moving `mermaid` to the add-in root with a taskpane peer and `uuid@14.0.0` override, upgrading taskpane Vite to `8.0.10`, making Mermaid rendering a dynamic import, and resolving `vscode-jsonrpc` aliases through Node resolution after the clean lockfile changed hoisting. Validation on 2026-04-26: `npm ci --prefix addin` found 0 vulnerabilities; `npm audit --prefix addin --audit-level=low` found 0 vulnerabilities; `npm run typecheck:addin` passed; `npm run build:addin` passed with `index` 1,465.43 kB and no large-chunk warning; `npm run check:bundle` passed with `main.js=1431.1 KiB`; `npm run validate:manifests` validated Word, Excel, and PowerPoint; `npm run test:office` passed with 98 tests.

### Features

- [ ] FEATURE-001: Restore saved-document workspace and file tools with policy guards
  - Category: Feature
  - Status: open
  - Priority: P2
  - Source: `docs/TASKPANE_INDEPENDENT_TRANSITION_REMEDIATION_PLAN.md` WS3; parked/deferred by stakeholder.
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

- [ ] FEATURE-002: Define and implement provider/auth matrix for subscription-backed and API-key-backed AI access
  - Category: Feature
  - Status: in_progress
  - Priority: P1
  - Source: 2026-04-26 product-goal narrative and implementation review.
  - Details: The product thesis is that users should be able to bring an existing AI subscription or inference provider instead of buying another enterprise add-in subscription. Current browser provider catalog discovers Pi models and supports API-key storage, but reports `oauthSupported: false` for providers and `/v1/auth/start` throws that OAuth is unavailable in browser-only mode. Provider support needs an explicit capability matrix covering API key, OAuth, official SDK constraints, subscription-backed routes, browser compatibility, companion requirements, image support, and any legal/provider policy restrictions.
  - Dependencies: SECURITY-004 for user-facing disclosure; BUG-005 for honest readiness state.
  - Subtasks:
    - [x] Create a provider/auth matrix for Pi, OpenAI/ChatGPT, Anthropic-compatible official paths, GitHub Copilot, OpenCode, OpenRouter, Cloudflare, Vercel, and other target providers.
    - [x] Mark each provider as supported, planned, blocked, or research-only with the required auth method and runtime surface.
    - [ ] Implement provider auth flows one at a time behind honest capability flags.
    - [x] Ensure UI copy never advertises OAuth/subscription access until a real flow exists.
    - [x] Add provider-level tests for catalog flags, auth start behavior, readiness, and model execution.
  - Acceptance Criteria:
    - [x] Provider catalog flags match implemented auth/runtime capability.
    - [x] Users can distinguish API-key providers, OAuth providers, companion-required providers, and unsupported providers.
    - [x] At least one non-API-key provider path is implemented or explicitly deferred with documented constraints before any UI promise.
  - Notes/Evidence: Review pointed to `BrowserModelRegistry.getProviderCatalog()` hardcoding `oauthSupported: false`, `/v1/auth/start` throwing for browser-only mode, and the new `AGENTS.md` product goal requiring provider flexibility. 2026-04-26 first implementation slice added `docs/provider-auth-matrix.md`, shared provider capability metadata, runtime catalog fields for support status/runtime/auth methods/browser-callable/companion-required/subscription-backed/image support, Settings provider cards that show planned companion/OAuth providers as informational, `/v1/auth/api-key` rejection for non-browser API-key providers, and `/v1/auth/start` errors that distinguish companion-owned OAuth from unsupported browser OAuth. Regression coverage now proves OpenAI is browser/API-key/image capable, Codex/Copilot/Gemini CLI/Antigravity are planned companion OAuth paths, Bedrock is companion-only, OAuth start behavior is honest, and image providers remain OpenAI-only. Real companion-owned OAuth provider implementations remain open under this task. Validation passed: `npm run typecheck:addin`, `npm run test:office`, `npm run build`, `npm run check:bundle`, and `npm run validate:manifests`.

- [x] FEATURE-003: Add professional workflow packs and host playbooks for high-value Office artifacts
  - Category: Feature
  - Status: done
  - Priority: P1
  - Source: 2026-04-26 product-goal narrative and competitor/inspiration review.
  - Details: Pi-Office should be more than generic chat in a taskpane. Claude's extracted add-in has dense host-specific behavior and verification guidance, while Pi-Office currently has a strong tool surface but only a small skill/playbook library. The product needs curated workflows for research papers, pitch decks, resumes, specs, business user stories, DCFs, legal review, spreadsheets, and similar professional artifacts.
  - Dependencies: BUG-007 for visual truthfulness where workflows rely on layout/vision; FEATURE-004 for deeper native edit coverage.
  - Subtasks:
    - [x] Define workflow-pack structure for task intent, required context, preferred tools, review gates, and completion checks.
    - [x] Add Word workflows for research papers, resumes, specs, legal/professional review, and business user stories.
    - [x] Add Excel workflows for DCF/financial model review, formula auditing, table/chart improvement, and narrative export.
    - [x] Add PowerPoint workflows for pitch-deck outline, slide polish, visual consistency, speaker notes, and data-backed slides.
    - [x] Add tests or snapshot checks proving workflow prompts register and route to the expected tools.
  - Acceptance Criteria:
    - [x] The package ships multiple domain-specific workflow packs beyond generic Office tool prompts.
    - [x] Workflows include verification and review criteria, not only generation instructions.
    - [x] Users can invoke or discover workflows from the taskpane without reading code.
  - Notes/Evidence: Competitor review highlighted Claude's host playbooks and the local `addin/packages/pi-office-pack/skills` surface as a place to grow. Closed 2026-04-26 with typed workflow-pack registry in `addin/packages/pi-office-pack/src/workflow-packs.ts`, prompt/skill surfacing through `OFFICE_APPEND_SYSTEM_PROMPT` and `office-host.SKILL.md`, host-specific taskpane starter prompts, and provenance updates. The shipped packs cover Word research paper/resume/spec/legal/business-user-story work, Excel DCF/formula/table-chart/narrative work, and PowerPoint pitch-deck/slide-polish/visual-consistency/speaker-notes/data-backed-slide work. Regression coverage in `addin/scripts/office-tests/src/workflow-packs.test.ts` proves host coverage, expected Office tool references, prompt/skill injection, and taskpane discoverability. Validation passed: `npm run typecheck:addin`, `npm run test:office` with 121 tests, `npm run build`, `npm run check:bundle`, and `npm run validate:manifests`.

- [ ] FEATURE-004: Expand first-class native Office editing coverage for professional document work
  - Category: Feature
  - Status: open
  - Priority: P1
  - Source: 2026-04-26 Office-host subagent review.
  - Details: The current Office tool inventory is broad, but professional editing still needs more structured native operations so agents do not fall back to raw execution or broad rewrites. Word gaps include styles, paragraph/list formatting, table-cell edits, headers/footers, page setup, field updates, content-control updates, and footnote/endnote body targeting. PowerPoint and Excel also need deeper object-level actions for polished artifacts.
  - Dependencies: SECURITY-003 to reduce reliance on raw Office.js execution.
  - Subtasks:
    - [ ] Inventory native Office.js APIs for the highest-value Word/Excel/PowerPoint editing gaps.
    - [ ] Add structured Word tools for style/paragraph/list/table/header/footer/field/content-control operations.
    - [ ] Add deterministic targeting for footnote/endnote bodies rather than only the reference marker.
    - [ ] Add structured PowerPoint and Excel object edits where current tools require raw code or weak anchors.
    - [ ] Add tests for each new first-class operation and update prompt/tool guidance to prefer them.
  - Acceptance Criteria:
    - [ ] Common professional document edits can be expressed through structured tools instead of `office_execute_js`.
    - [ ] Word footnote/endnote edits target the note body correctly.
    - [ ] Tests cover representative Word, Excel, and PowerPoint native edits.
  - Notes/Evidence: Review pointed to Word context already exposing rich objects while action coverage remains thinner than the product bar.

- [ ] FEATURE-005: Add explicit cross-host artifact workflows
  - Category: Feature
  - Status: open
  - Priority: P2
  - Source: 2026-04-26 competitor/inspiration review.
  - Details: A best-in-class Office assistant should help users move work between Excel, Word, and PowerPoint with reviewable intent, such as turning Excel tables/charts into deck slides, converting specs into user stories, or creating Word summaries from workbook analysis. Current host tools are mostly per-host. Cross-host workflows need clear data movement, provenance, and user-visible confirmation.
  - Dependencies: FEATURE-003 and BUG-007 for workflow and visual verification foundations.
  - Subtasks:
    - [ ] Define initial cross-host flows, starting with Excel-to-PowerPoint and Excel-to-Word.
    - [ ] Track source ranges/slides/sections and generated target artifacts for provenance.
    - [ ] Add user-visible review steps before inserting or replacing content in another host.
    - [ ] Add tests or manual validation scripts for at least one cross-host flow.
  - Acceptance Criteria:
    - [ ] At least one cross-host workflow is usable from the taskpane with source provenance and review.
    - [ ] Generated target content clearly identifies source document/range/slide context where appropriate.
    - [ ] The workflow does not imply cross-host Office.js access that is not actually available in the active host.
  - Notes/Evidence: Competitor review recommended explicit cross-host handoff workflows as a way for Pi-Office to outperform generic add-ins.

- [ ] FEATURE-006: Add advanced mode where companion owns inference, providers, MCP, memory, and non-Office tools
  - Category: Feature
  - Status: open
  - Priority: P0
  - Source: 2026-04-26 plan implementation after user clarified "Companion owns all" for advanced mode.
  - Details: The current taskpane owns the Pi agent, model/provider catalog, provider auth, and tool loop, while the companion is a sidecar for read-only file and local MCP calls. Advanced mode should invert that ownership: the companion owns inference, providers, model auth, MCP, memory, non-Office tools, and tool calling. The taskpane remains the Office-hosted presentation layer and structured Office.js executor. Basic taskpane-only mode must remain usable when the companion is absent.
  - Dependencies: SECURITY-006, FEATURE-002, BUG-003, SECURITY-001.
  - Subtasks:
    - [ ] Define a versioned taskpane-companion protocol for capabilities, settings sync, chat streaming, tool requests, Office tool execution, and auth migration.
    - [ ] Add companion-side Pi agent/session ownership while keeping Office tools proxied back to the taskpane.
    - [ ] Add a Basic/Advanced mode switch with seamless migration of settings, preferences, enabled providers/models, and connector configuration.
    - [ ] Store or broker provider secrets through the companion using OS keychain-compatible storage where available.
    - [ ] Preserve taskpane-only fallback when the companion disconnects or is not installed.
  - Acceptance Criteria:
    - [ ] In Advanced mode, provider/model inference and non-Office tool calls are executed by the companion, not the browser taskpane.
    - [ ] Office.js calls still execute only inside the active Office taskpane.
    - [ ] Switching Basic -> Advanced preserves non-secret preferences automatically and handles secrets through an explicit safe migration flow.
    - [ ] Reconnect/fallback behavior is visible and tested.
  - Notes/Evidence: 2026-04-26 review found `BrowserOfficeSession` still constructs the Pi `Agent` in `addin/apps/taskpane/src/lib/runtime/inprocess-kernel.ts`, while `companion/src/server.ts` only exposes health, read-only file tools, and MCP execution.

### Improvements

- [x] IMPROVEMENT-001: Clean up worktree hygiene for untracked archive and generated artifacts
  - Category: Improvement
  - Status: done
  - Priority: P2
  - Source: 2026-04-26 worktree review.
  - Details: The worktree contains untracked `.factory/`, `.github/`, `archive/`, and other generated or transitional artifacts. In particular, `archive/companion/node_modules` is present under an untracked archive tree. This can confuse future reviews, inflate diffs, and accidentally commit vendored dependencies or generated output.
  - Dependencies: None.
  - Subtasks:
    - [x] Decide which untracked artifacts are intended source/documentation versus generated local state.
    - [x] Add ignore rules for generated archive dependency folders and local factory/cache outputs that should never be committed.
    - [x] If archive source is intentionally retained, keep source files only and exclude built output/dependencies unless there is a documented reason.
    - [x] Re-run `git status --short` and document the intended remaining untracked files.
  - Acceptance Criteria:
    - [x] No `node_modules`, built `dist`, caches, or generated local-only files remain staged or visible as intended source.
    - [x] `git status --short` clearly distinguishes product changes from local/generated artifacts.
    - [x] `.gitignore` covers repeatable generated artifacts without hiding important source files.
  - Notes/Evidence: Review observed untracked archive/vendor artifacts while inspecting the worktree. Fixed in `2b5f058` by ignoring `.factory/`; the grouped runtime commit retained archive source while generated `archive/companion/node_modules`, `archive/companion/dist`, and taskpane build output remained ignored. 2026-04-26 history cleanup removed active `.factory/services.yaml` test coupling, ignored future `archive/companion/` recreation, and purged tracked `.factory/` plus `archive/companion/` from branch history.

- [ ] IMPROVEMENT-002: Clarify release packaging and runtime assumptions after the independent taskpane transition
  - Category: Improvement
  - Status: open
  - Priority: P2
  - Source: `docs/TASKPANE_INDEPENDENT_TRANSITION_REMEDIATION_PLAN.md` WS6 and risk list.
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

- [ ] IMPROVEMENT-003: Consolidate Office tool contracts and user-facing descriptions
  - Category: Improvement
  - Status: open
  - Priority: P2
  - Source: 2026-04-26 Office-host and competitor/inspiration review.
  - Details: Tool contracts and descriptions are duplicated across `addin/packages/pi-office-pack`, the in-process kernel, the Office bridge, and Settings UI. This creates drift: for example, Settings calls `office_capture_snapshot` a visual screenshot while the actual implementation is selection/context snapshots plus metadata, and PowerPoint `edit_slide_master` wording is broader than its current apply-layout behavior. A single source or generated registry would make capability honesty easier to preserve.
  - Dependencies: BUG-007 and BUG-009 for known contract mismatches.
  - Subtasks:
    - [ ] Inventory all Office tool names, labels, descriptions, categories, parameters, and runtime support paths.
    - [ ] Choose a canonical registry or generation path for tool metadata used by prompts, Settings, and runtime registration.
    - [ ] Add drift tests that fail when Settings/prompt descriptions disagree with implemented support.
    - [ ] Update misleading descriptions found during the 2026-04-26 review.
  - Acceptance Criteria:
    - [ ] Each Office tool has one canonical capability description consumed by the UI and runtime where feasible.
    - [ ] Tests catch obvious drift between advertised and executable tool behavior.
    - [ ] User-facing copy clearly marks host-only, metadata-only, experimental, or escape-hatch tools.
  - Notes/Evidence: Review pointed to `addin/packages/pi-office-pack/src/extension.ts`, `addin/apps/taskpane/src/lib/runtime/inprocess-kernel.ts`, `addin/apps/taskpane/src/lib/office-bridge.ts`, and `addin/apps/taskpane/src/app/components/SettingsPage.tsx`.

- [ ] IMPROVEMENT-004: Build a professional PowerPoint visual asset and icon pipeline
  - Category: Improvement
  - Status: open
  - Priority: P2
  - Source: 2026-04-26 Office-host subagent review and ChatGPT inspiration comparison.
  - Details: PowerPoint `insert_icon` currently relies on a tiny taskpane runtime catalog and can fall back to glyph text boxes unless base64 content is supplied. For pitch decks and professional slides, Pi-Office needs a higher-quality asset pipeline for icons, generated images, slide screenshots/previews, and reusable visual components.
  - Dependencies: BUG-006 for image-provider execution honesty; BUG-007 for visual verification fidelity.
  - Subtasks:
    - [ ] Define supported asset sources for icons, generated images, user-provided images, and reusable slide components.
    - [ ] Replace glyph-textbox fallback with higher-quality Office-compatible icon/image insertion where possible.
    - [ ] Add preview/verification support for inserted assets.
    - [ ] Add tests or manual validation for icon insertion, generated-image insertion, and asset fallback behavior.
  - Acceptance Criteria:
    - [ ] Inserted icons/assets render professionally in PowerPoint rather than as plain glyph placeholders except when explicitly requested.
    - [ ] Asset insertion reports source, format, and fallback behavior.
    - [ ] At least one deck-quality asset workflow is validated in PowerPoint.
  - Notes/Evidence: Review pointed to the small runtime icon catalog and glyph-textbox insertion path in `addin/apps/taskpane/src/lib/office/powerpoint-actions.ts`, plus ChatGPT inspiration assets around generated slide stores and slide screenshots.

- [ ] IMPROVEMENT-005: Simplify provider, model, and settings UX into guided and advanced surfaces
  - Category: Improvement
  - Status: open
  - Priority: P1
  - Source: 2026-04-26 plan implementation after user noted the current provider/model/settings catalog is overwhelming for nontechnical users and includes regional/legacy models that can scare enterprises.
  - Details: Settings currently exposes a large provider/model list and many toggles without a strong basic/advanced information architecture. Default users should see a guided shortlist of recommended current models and plain-language provider choices. Advanced users should still be able to opt into the full catalog, including legacy, experimental, regional, and higher-risk providers. Provider/model metadata needs to include region/jurisdiction, enterprise-risk messaging, capability flags, auth methods, current-vs-legacy status, and replacement suggestions.
  - Dependencies: FEATURE-002 and BUG-005.
  - Subtasks:
    - [ ] Define model/provider metadata fields: visibility, status, replacedBy, region, enterpriseRisk, authMethods, capabilities, and companionRequired.
    - [ ] Curate a default guided shortlist for Word/Excel/PowerPoint professional work.
    - [ ] Move legacy, experimental, China-hosted/regional, and niche providers behind an explicit advanced catalog.
    - [ ] Group settings into basic, advanced, privacy/security, providers, companion, and diagnostics using plain-language labels.
    - [ ] Add tests that default enabled models/providers do not include advanced-only or region-risk entries unless the user opts in.
  - Acceptance Criteria:
    - [ ] A first-time nontechnical user can pick a recommended provider/model without reading a large model catalog.
    - [ ] Advanced users can still find and enable the complete catalog.
    - [ ] Regional/enterprise-risk providers are clearly labeled and not enabled by default.
    - [ ] Model lists avoid stale versions when newer replacements exist unless the user enables advanced/legacy mode.
  - Notes/Evidence: 2026-04-26 review found default enabled models/providers include broad Pi catalog entries, China-linked/regional providers, and old model revisions in `addin/apps/taskpane/src/hooks/usePreferences.ts`, with Settings rendering all enabled-provider models together. 2026-04-26 hardening narrowed the default shortlist to OpenAI, Anthropic, and Google current/recommended entries and added `addin/scripts/office-tests/src/model-curation.test.ts`; full metadata, regional labeling, and guided/advanced IA remain open.

### Testing

- [ ] TESTING-001: Add regression tests for transition regressions where feasible
  - Category: Testing
  - Status: open
  - Priority: P1
  - Source: `docs/TASKPANE_INDEPENDENT_TRANSITION_REMEDIATION_PLAN.md` P0 checklist.
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
  - Source: `docs/TASKPANE_INDEPENDENT_TRANSITION_REMEDIATION_PLAN.md` validation matrix.
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
