# AGENTS.md

## Project brief

This repository contains a Pi-powered Microsoft Office add-in stack centered on an independent shared taskpane for Word, Excel, and PowerPoint. Companion-era archive code is not part of the active runtime and should not be reintroduced into tracked source.

## Product goal

Pi-Office aims to be an open-source, state-of-the-art AI assistant for Microsoft Word, Excel, and PowerPoint that gives individuals and teams professional document-generation and editing help without forcing them into a single expensive enterprise subscription. The thesis is that the core capability should come from whichever AI subscription or inference provider a user already has access to, while the add-in supplies the Office-native UX, tool orchestration, privacy posture, and document-editing taste layer.

The product should be provider-flexible rather than model-vendor-bound. It should support major AI runtimes and auth patterns where feasible, including Pi via `pi-mono`, GitHub Copilot OAuth-style access, ChatGPT tiers with Codex-style agent capability, Anthropic-compatible paths where officially permitted, OpenCode offerings, OpenRouter, Cloudflare, Vercel, and future providers with comparable inference/tool APIs. Provider support must be honest about capabilities, restrictions, and token/auth state.

The competitive context includes Claude's Office add-in, ChatGPT/Codex-style add-ins, Microsoft Copilot, and other Office AI assistants. Those products can be powerful and well integrated, but they may be costly, subscription-tied, telemetry-opaque, access-restricted, uneven across document tasks, or weaker than desired for professional editing UX and user-specific writing taste. Pi-Office should compete by combining transparent local/Office behavior, broad provider choice, excellent Office.js leverage, LLM tool calling, optional visual reasoning for diagrams/images/layout, and a polished workflow for research papers, pitch decks, resumes, specs, business user stories, spreadsheets, DCF analysis, and similar high-value professional artifacts.

The core product bar is not just "chat in a taskpane." It is an interactive, trustworthy Office workbench: native Office edits where possible, strong context gathering, reviewable proposed changes, privacy-conscious connector behavior, clear auth/telemetry disclosure, saved-vs-unsaved document safety boundaries, and a UX that helps users create high-quality documents with less friction and more taste.

## Originality and provenance policy

- Office.js is the public platform foundation. Build on Microsoft-documented APIs, web standards, and original Pi-Office code.
- Competitor add-ins and local captures may be studied for capability mapping, user expectations, and gap analysis only. Do not copy their source code, assets, prompts, UI text, distinctive interaction patterns, private API contracts, or bundled resources.
- When a competitor-inspired capability is worth building, re-spec it in Pi-Office terms: user problem, public API basis, original UX, original implementation approach, and validation evidence.
- Keep a provenance trail for risky inspiration work in `docs/provenance.md` and `backlog.md` before public release.

## Companion architecture direction

- Basic mode is taskpane-only: the add-in owns the in-browser Pi runtime, provider configuration, chat loop, and Office.js tool execution.
- Advanced mode is add-in plus companion: the companion should own inference, provider auth/session state, MCP, memory, non-Office tools, tool calling, and local workspace context; the taskpane should become the presentation layer plus structured Office.js executor.
- Office tools stay in the add-in because only the Office host can safely and directly execute Office.js against the active document.
- Basic-to-advanced migration must preserve user settings, preferences, enabled providers/models, and connector configuration. Secret migration must be explicit and should prefer companion/OS keychain storage when available.
- Companion filesystem access defaults to read-only saved-document/workspace context. Shell access must not be exposed as raw host bash; if added, it must route through a tested read-only sandbox with writable scratch only, denied secret patterns, no network by default, and fail-closed approvals.

## Fresh session startup

- At the start of every fresh session, read `AGENTS.md` and `backlog.md` before substantial investigation, planning, or implementation.
- Correlate new user requests against existing `backlog.md` task IDs so related discussion, fixes, and commits stay tied to the running backlog.
- Use `backlog.md` to help guide the user toward sensible next work when they ask for direction, when a request is ambiguous, or when an existing backlog item is clearly relevant.

## Running project memory

- Maintain this file as durable project memory. Agents must keep the `Project summary` and `Last worked on` sections current without requiring explicit approval or discussion with the user.
- Update `Project summary` when architectural direction, committed state, major features, validation posture, repo workflow, or backlog ownership changes.
- Update `Last worked on` whenever an agent finishes meaningful investigation, planning, implementation, review, validation, commit, or handoff work.
- Keep these notes concise but specific enough that a future agent can resume without needing prior chat context.

## Project summary

- Pi-Office is moving toward an independent shared Office taskpane for Word, Excel, and PowerPoint, with Pi as the external AI runtime and Office host behavior kept separate from Pi auth/runtime concerns.
- The durable product goal is an open-source, provider-flexible, privacy-conscious Office AI workbench that can use a user's existing AI subscription/provider access instead of locking them into one expensive enterprise assistant.
- The originality policy is idea-level inspiration only: competitor captures can guide capability analysis, but Pi-Office must use original code, prompts, UI text, assets, and public API implementations.
- Advanced companion direction is now explicit: companion-owned inference/providers/MCP/non-Office tools, with taskpane-owned Office.js execution and seamless settings migration from taskpane-only mode.
- The active add-in surface now lives under `addin/`: `addin/apps/taskpane`, `addin/packages/pi-office-pack`, `addin/manifests`, and `addin/scripts`. The optional companion is split into the top-level `companion/` package; old `archive/companion` source and `.factory` scaffolding have been purged from tracked history and should not be restored.
- The repo root is a lightweight command router. Add-in dependencies and lockfile live in `addin/`; companion dependencies and lockfile live in `companion/`; there is no Turbo/Nx workspace layer.
- The repo now has a canonical `backlog.md` for review findings, feature gaps, bugs, improvements, security items, and testing work, with IDs intended to be referenced in commits and future implementation.
- The 2026-04-26 implementation review identified release-blocking "capability honesty" work around OAuth, remote connectors, provider readiness, raw Office.js execution, permission prompts, visual capture fidelity, and privacy/storage disclosure.
- Current hardening makes tool permissions fail closed on timeout, makes `office_execute_js` a manual-only escape hatch category, prevents connector imports/OAuth completion from creating false connected OAuth state, and marks remote HTTP connectors setup-only until execution exists.
- `SECURITY-001` is closed: connector OAuth completion now requires a verified credential handoff, manual UI completion is removed, tokenless/imported OAuth records report `auth_required`, and callback tests cover no-token/cancelled/mismatch/expired/success/import paths.
- `SECURITY-002` is closed: timeout, disconnect, and session cleanup paths now have automated fail-closed permission coverage across write-doc, connector, read-external, and write-external tool categories.
- `SECURITY-003` is closed: `office_execute_js` is an escape-hatch tool that cannot be auto-approved and any approval is normalized to a one-time decision, even if a client sends a broader scope.
- `SECURITY-005` is closed with an expanded `docs/provenance.md`, new `CONTRIBUTING.md` originality guidance, and a neutralized CSS comment; `SECURITY-007` is obsolete because the stakeholder accepted connector logos as-is without a licensing audit.
- `SECURITY-006` is closed: the companion now has shared shell capability protocol, fail-closed backend detection, policy validation, destructive probes, environment scrubbing, output caps/timeouts, a sandbox-routed Pi `BashOperations` adapter, and taskpane gating so `bash` only appears when a saved-document companion session reports shell `available`.
- `SECURITY-004` is closed: the taskpane now declares an Office-compatible CSP meta policy, Settings includes Privacy disclosure and clear-data controls, provider/connector clear-all routes remove stored encrypted envelopes plus local crypto keys, saved chat history can be cleared, and privacy/storage behavior is documented.
- `BUG-010` is closed: add-in clean install now reports zero npm audit vulnerabilities, Vite is patched to `8.0.10`, Mermaid renders through a lazy-loaded chunk, and the build no longer emits the large-chunk warning while the custom bundle-budget gate remains active.
- `BUG-005` is closed: provider/model readiness now separates stored credentials from verified usable credentials, labels unverified/auth-failed states in Settings and model selection, migrates old stored keys to unverified, promotes providers after successful use, and demotes 401/403/auth failures.
- `BUG-003` is closed: local stdio connector launches now merge safe inherited env, connector env, and the intended manual/env/detected credential variable; missing credential values or manual secrets without env targets stay auth-required without leaking secret values.
- `BUG-008` is closed: Excel rewind restores captured formula matrices instead of flattening formulas to values, falls back with warnings only when formula data is unavailable, and surfaces the current checkpoint fidelity boundary.
- `BUG-007` is closed: visual capture contracts now describe Office.js context/active-selection snapshots instead of OS screenshots or arbitrary offscreen Excel range renders, and `read_range_image` fails with selection guidance when the requested range does not match the active Excel selection.
- `BUG-009` is closed: PowerPoint selected-shape anchors now use each shape's owning parent slide when Office.js exposes it, and the legacy `edit_slide_master` tool is clearly apply-existing-layout only rather than slide-master mutation.
- `FEATURE-002` has its first implementation slice: provider catalog entries now carry support status, auth methods, runtime surface, browser-callable state, companion requirement, subscription-backed state, and image support. Browser taskpane API-key providers remain executable, while Codex/Copilot/Gemini CLI/Antigravity/Bedrock/Vertex/Azure-style paths are visible as planned companion/OAuth or companion-auth work instead of fake browser API-key setup.
- `BUG-006` is closed: image generation is explicitly OpenAI-only in the browser taskpane, Settings no longer advertises Google/OpenRouter image setup, `/v1/image-models` remains OpenAI-only, and unsupported image model preference writes fail during configuration.
- `FEATURE-003` is closed: the Office pack now ships typed professional workflow packs for Word, Excel, and PowerPoint, injects them into prompt/skill guidance, and exposes host-specific starter prompts in the taskpane.
- The broad independent taskpane transition is now committed: `a70cd9a` moves Pi session routes into the taskpane in-process kernel, reduces the companion to optional read-only file/MCP support, archives old companion source, and includes next-prompt suggestions, model curation, and runtime regression tests.
- Dev and sideload hardening is now committed as `a8a5a0d` plus `43b2a5b`, including the CI workflow, bundle budget, cert/port preflight, sideload resource preflight, manifest cache-bust version `1.0.0.2`, taskpane dev host on `https://localhost:3443`, and Vite cert loading scoped to the dev server only.

## Last worked on

- 2026-04-26: Created root `backlog.md`, added backlog stewardship and backlog-task commit discipline to `AGENTS.md`, and committed those documentation changes as `df51baf docs: add running backlog stewardship`.
- 2026-04-26: Added the standing rule that agents must automatically maintain `Project summary` and `Last worked on`, plus the fresh-session requirement to check `backlog.md` and correlate new work with existing backlog IDs.
- 2026-04-26: Recorded the expanded Pi-Office product goal in `AGENTS.md`, ran a multi-agent implementation review against local competitor/inspiration captures, and expanded `backlog.md` with the resulting security, provider, connector, visual-fidelity, and workflow tasks.
- 2026-04-26: Implemented the plan's first hardening pass: added originality/provenance and companion-mode direction, made permission timeouts fail closed, made `office_execute_js` manual-only, blocked false OAuth connected state, and marked remote HTTP connectors setup-only.
- 2026-04-26: Analyzed and grouped the dirty worktree into scoped commits: `2b5f058` for local factory ignore hygiene, `a8a5a0d` for BUG-001 dev/CI/preflight gates, `43b2a5b` for BUG-001 clean-checkout Vite cert loading, and `a70cd9a` for FEATURE-006 independent taskpane runtime work. Validation passed with `npm run typecheck`, `npm run build`, `npm run check:bundle`, `npm run validate:manifests`, `npm run test:office`, `npm run preflight:dev`, and a certs-absent `npm run build:taskpane` check.
- 2026-04-26: Closed `SECURITY-002` by adding protocol-level permission timeout coverage across write-doc, connector, read-external, and write-external categories plus session cleanup rejection coverage; `npm run test:office` passed with 89 tests.
- 2026-04-26: Closed `SECURITY-003` by enforcing one-time-only `office_execute_js` approvals in the runtime and permission popup, with protocol/UI tests for auto-approval, denial, attempted session approval, and blocked snippets; `npm run test:office` passed with 91 tests and `npm run typecheck:taskpane` passed.
- 2026-04-26: Closed `SECURITY-001` by requiring OAuth callback credential handoff before connected state, removing manual completion from connector UI, resetting tokenless imports/storage to `auth_required`, and adding callback contract tests; `npm run test:office` passed with 92 tests and `npm run typecheck` passed.
- 2026-04-26: Closed `SECURITY-005` by turning provenance notes into an audit matrix with evidence commands, adding contributor originality guidance, linking it from README, and replacing a copied-looking competitor reference in CSS; created `SECURITY-007` for third-party connector logo source/licensing review.
- 2026-04-26: Started `SECURITY-006` with a companion shell sandbox policy doc, README warning, and no-raw-shell regression coverage for active taskpane tools and companion routes; full sandbox backend/probes are still pending.
- 2026-04-26: Closed `SECURITY-006` by adding shared shell capability protocol, `CompanionShellSandbox`, capability/execute routes, taskpane `bash` gating behind available sandbox state, policy/destructive probe tests, environment scrubbing, output caps/timeouts, and a custom Pi `BashOperations` adapter; `npm run test:office` passed with 98 tests and companion/taskpane typechecks passed.
- 2026-04-26: Refactored the repository into independent `addin/`, `companion/`, `website/`, and `docs/` areas, keeping the root as a normal npm script router and preserving add-in/companion behavior with path-only config, CI, test, and documentation updates.
- 2026-04-26: Closed `BUG-010` by clearing the add-in npm audit findings, removing local vulnerable Office CLI packages from `npm ci`, pinning on-demand sideload/manifest CLIs, upgrading Vite to `8.0.10`, lazy-loading Mermaid, and verifying clean install/build/bundle/typecheck/manifest/test gates.
- 2026-04-26: Prepared a history cleanup for `.factory/` and `archive/companion/` by removing active test dependencies on `.factory/services.yaml`, ignoring future `archive/companion/` recreations, and removing both tracked trees before the rewrite.
- 2026-04-26: Closed `BUG-005` by adding provider auth states for stored/unverified/verified/failed credentials, updating Settings/model picker capability honesty, adding provider-readiness regression tests, refreshing stale local workspace package links, and validating with add-in typecheck, Office tests, add-in build, companion typecheck, bundle budget, and manifest validation.
- 2026-04-26: Closed `SECURITY-004` by adding the taskpane CSP meta policy, Privacy settings disclosure and clear-data controls, provider/connector clear-all runtime routes that remove local crypto keys, saved chat-history clearing, privacy/storage docs, and CSP/storage regression coverage; add-in typecheck, Office tests, add-in build, companion typecheck, bundle budget, and manifest validation passed.
- 2026-04-26: Closed `BUG-003` by tracing MCP stdio env behavior, adding explicit safe-env and credential injection for local stdio connectors, preserving manual local credential env targets from setup, and verifying companion/taskpane typechecks, Office tests, build, manifests, and bundle budget.
- 2026-04-26: Closed `BUG-008` by changing Excel checkpoint restore to write captured formula matrices, adding restore warnings for values-only fallback and unsupported workbook semantics, and validating add-in typecheck, Office tests, build, manifests, and bundle budget.
- 2026-04-26: Closed `BUG-007` by making snapshot and Excel range-image copy fidelity-honest, enforcing active-selection matching for `read_range_image`, updating the Claude parity tracker, and validating add-in typecheck, Office tests, build, manifests, and bundle budget.
- 2026-04-26: Closed `BUG-009` by resolving PowerPoint shape anchors through parent-slide metadata, narrowing the legacy slide-master tool to layout application only, updating PowerPoint guidance/tracker notes, and validating add-in typecheck, Office tests, build, manifests, and bundle budget.
- 2026-04-26: Obsoleted `SECURITY-007` after stakeholder clarified connector logos are acceptable as-is and do not need a licensing/source audit for the current release path.
- 2026-04-26: Advanced `FEATURE-002` and closed `BUG-006` by adding provider/auth capability metadata, `docs/provider-auth-matrix.md`, honest provider Settings cards, fail-closed browser auth routes for companion/OAuth providers, OpenAI-only image setup copy, image preference validation, and provider/image regression tests; add-in typecheck, Office tests, build, bundle budget, and manifest validation passed.
- 2026-04-26: Closed `FEATURE-003` by adding `OFFICE_WORKFLOW_PACKS`, host-specific workflow guidance and starter prompts for Word/Excel/PowerPoint professional artifacts, provenance notes, and regression coverage; add-in typecheck, Office tests, build, bundle budget, and manifest validation passed.

## Core rules

- Do not create Git commits unless the user explicitly asks, except when explicitly working a `backlog.md` task; backlog task work must be committed per the Backlog stewardship rules.
- Keep Git configuration repo-local only.
- Treat Pi as an external dependency, not vendored source.
- Keep host-side `Office.js` execution separate from Pi runtime and auth concerns.
- Do not restore old companion-era archive source into tracked history; use the top-level `companion/` package for active optional companion work.

## Backlog stewardship

- `backlog.md` is the canonical running backlog for bugs, features, improvements, security concerns, and testing gaps.
- Create backlog tasks automatically when you observe actionable gaps during reviews, investigations, plan-mode exits, implementation work, verification, or unrelated work that reveals a real issue.
- Update backlog tasks automatically when you explicitly work on them, prove them fixed, discover new evidence, split/merge scope, or realize a task is obsolete or no longer required.
- Do not delete backlog history. Mark tasks `done` or `obsolete` with evidence and keep enough context for a future agent with no chat history.
- Every backlog task must include category, status, priority, source, details, dependencies, subtasks, acceptance criteria, and notes/evidence.
- Keep task checkboxes unchecked unless the task status is `done` or `obsolete`.
- When working on a backlog task, commit that task's code/doc/test changes right away after validation, and keep the commit scoped to that backlog item only. Use conventional commit style and include the backlog ID in the subject, for example `fix(BUG-001): make taskpane build independent of local certs`.
- Backlog maintenance may be delegated to a subagent when the session explicitly allows delegation; prefer `gpt-5.3-codex` with high reasoning if model selection is available. The parent agent remains responsible for reviewing and integrating the backlog update.

## Architecture priorities

- Shared taskpane app for Word, Excel, and PowerPoint
- Independent local taskpane runtime without a live companion dependency
- Pi runtime as the primary AI core
- Internal Office Pi package for Office tools, prompts, skills, and tool gating
- Unsaved-document mode without local filesystem access
- Saved-document mode bound to the document folder for AGENTS/skills/file context

## Repo shape

- `addin/apps/taskpane`: Office-hosted React UI and Office bridge executor
- `addin/packages/pi-office-pack`: Office-specific Pi package
- `addin/manifests`: host-specific XML manifests
- `addin/scripts`: add-in validation and local setup helpers
- `companion`: optional local companion package for read-only file/MCP support and gated sandbox capability
- `website`: placeholder for the future public website
- `docs`: detailed project docs, trackers, provenance notes, and design references

## Quality bar

- Keep manifests valid and side-load friendly
- Prefer native Office edits over export/import hacks when possible
- Make bridge contracts explicit and versioned
- Validate build/typecheck/manifests before handoff
