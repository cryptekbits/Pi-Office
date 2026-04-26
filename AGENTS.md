# AGENTS.md

## Project brief

This repository contains a Pi-powered Microsoft Office add-in stack centered on an independent shared taskpane for Word, Excel, and PowerPoint. Archived companion-era code may still exist in history or archive folders, but it is not the active runtime.

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
