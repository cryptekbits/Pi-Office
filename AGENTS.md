# AGENTS.md

## Project brief

This repository contains a Pi-powered Microsoft Office add-in stack centered on an independent shared taskpane for Word, Excel, and PowerPoint. Archived companion-era code may still exist in history or archive folders, but it is not the active runtime.

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
- The active surface is `apps/taskpane` plus `packages/pi-office-pack`; companion-era code is archival unless a task explicitly scopes it back in.
- The repo now has a canonical `backlog.md` for review findings, feature gaps, bugs, improvements, security items, and testing work, with IDs intended to be referenced in commits and future implementation.
- The 2026-04-26 implementation review identified release-blocking "capability honesty" work around OAuth, remote connectors, provider readiness, raw Office.js execution, permission prompts, visual capture fidelity, and privacy/storage disclosure.
- Current hardening makes tool permissions fail closed on timeout, makes `office_execute_js` a manual-only escape hatch category, prevents connector imports/OAuth completion from creating false connected OAuth state, and marks remote HTTP connectors setup-only until execution exists.
- `SECURITY-001` is closed: connector OAuth completion now requires a verified credential handoff, manual UI completion is removed, tokenless/imported OAuth records report `auth_required`, and callback tests cover no-token/cancelled/mismatch/expired/success/import paths.
- `SECURITY-002` is closed: timeout, disconnect, and session cleanup paths now have automated fail-closed permission coverage across write-doc, connector, read-external, and write-external tool categories.
- `SECURITY-003` is closed: `office_execute_js` is an escape-hatch tool that cannot be auto-approved and any approval is normalized to a one-time decision, even if a client sends a broader scope.
- `SECURITY-005` is closed with an expanded `docs/provenance.md`, new `CONTRIBUTING.md` originality guidance, and a neutralized CSS comment; third-party connector logo licensing remains open as `SECURITY-007`.
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

## Core rules

- Do not create Git commits unless the user explicitly asks, except when explicitly working a `backlog.md` task; backlog task work must be committed per the Backlog stewardship rules.
- Keep Git configuration repo-local only.
- Treat Pi as an external dependency, not vendored source.
- Keep host-side `Office.js` execution separate from Pi runtime and auth concerns.
- Treat companion-era references as archival unless a task explicitly says to touch them.

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

- `apps/taskpane`: Office-hosted React UI and Office bridge executor
- `apps/companion`: archived/removed active runtime; do not treat as the current architecture
- `packages/pi-office-pack`: Office-specific Pi package
- `manifests`: host-specific XML manifests
- `scripts`: validation and local setup helpers

## Quality bar

- Keep manifests valid and side-load friendly
- Prefer native Office edits over export/import hacks when possible
- Make bridge contracts explicit and versioned
- Validate build/typecheck/manifests before handoff
