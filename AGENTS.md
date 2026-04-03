# AGENTS.md

## Project brief

This repository contains a Pi-powered Microsoft Office add-in stack with a shared taskpane for Word, Excel, and PowerPoint plus a local companion service.

## Core rules

- Do not create Git commits unless the user explicitly asks.
- Keep Git configuration repo-local only.
- Treat Pi as an external dependency, not vendored source.
- Keep host-side `Office.js` execution separate from Pi runtime and auth concerns.

## Architecture priorities

- Shared taskpane app for Word, Excel, and PowerPoint
- Local HTTPS companion serving the taskpane and Pi bridge from one origin
- Pi runtime as the primary AI core
- Internal Office Pi package for Office tools, prompts, skills, and tool gating
- Unsaved-document mode without local filesystem access
- Saved-document mode bound to the document folder for AGENTS/skills/file context

## Repo shape

- `apps/taskpane`: Office-hosted React UI and Office bridge executor
- `apps/companion`: local HTTPS service, Pi session runtime, provider auth, WebSocket bridge
- `packages/pi-office-pack`: Office-specific Pi package
- `manifests`: host-specific XML manifests
- `scripts`: validation and local setup helpers

## Quality bar

- Keep manifests valid and side-load friendly
- Prefer native Office edits over export/import hacks when possible
- Make bridge contracts explicit and versioned
- Validate build/typecheck/manifests before handoff
