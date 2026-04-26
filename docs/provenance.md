# Pi-Office Provenance Audit

This file is the release-safety trail for competitor-inspired work. It covers active Pi-Office source and public-release artifacts; archived research captures remain non-release reference material only.

## Policy

- Study competitor add-ins for user expectations, capability gaps, and workflow ideas only.
- Do not copy competitor source code, prompts, UI text, assets, private API contracts, distinctive interaction patterns, or bundled resources.
- Rebuild useful capabilities from public APIs, original Pi-Office UX, and original implementation.
- Treat Microsoft Office.js documentation, web standards, and Pi-Office-owned code as the implementation basis for Office behavior.
- Keep risky inspiration work tied to a backlog item and this provenance file before release.

## Audit Scope

Audited on 2026-04-26:

- `AGENTS.md`, `backlog.md`, `README.md`, `docs/`
- `addin/apps/taskpane/src`, `addin/apps/taskpane/public`, `addin/manifests/`
- `addin/packages/pi-office-pack/src`, `addin/packages/pi-office-pack/skills`
- `companion/src`, `addin/scripts/`

Excluded from the release audit:

- `archive/`, local capture/research folders, generated `dist/`, `node_modules/`, temp/debug outputs, and local certs.
- Historical investigation documents are allowed to mention competitors as analysis records, but they are not implementation source.

## Evidence Commands

The latest audit used:

```bash
rg -n "Claude|ChatGPT|Copilot|Ghostwriter|Anthropic|OpenAI|Codex|OpenCode|inspiration|competitor|provenance|DMCA|copyright|private API|prompt" AGENTS.md backlog.md README.md docs addin companion --glob "!**/node_modules/**" --glob "!**/dist/**"
rg --files addin companion docs | rg "(?i)(logo|icon|asset|png|svg|jpg|jpeg|webp|prompt|skill|md)$"
```

Findings:

- No active source file contains competitor source code, private endpoint contracts, copied prompt bundles, or competitor-bundled assets.
- Competitor references in active source are policy, backlog, provider-label, or audit-context references.
- One code comment that described a taskpane style as "Claude Excel inspired" was replaced with original neutral wording.
- Pi-Office system prompts and skills use original Pi-Office wording centered on Office.js tools, safety boundaries, and document state.

## Provenance Matrix

| Capability | Inspiration / Comparator | Public Or Original Basis | Pi-Office Implementation | Audit Status |
| --- | --- | --- | --- | --- |
| Office host tools for Word, Excel, and PowerPoint | Claude Office add-in, Microsoft Copilot, generic Office assistants | Microsoft Office.js public APIs plus original tool contracts | `addin/packages/pi-office-pack/src/protocol.ts`, `addin/packages/pi-office-pack/src/extension.ts`, `addin/apps/taskpane/src/lib/office-*` | Original implementation; no copied code or private contracts found. |
| Taskpane chat and streaming UX | Claude/ChatGPT/Copilot user expectations | Original React taskpane and Pi agent runtime | `addin/apps/taskpane/src/app`, `addin/apps/taskpane/src/lib/runtime/inprocess-kernel.ts` | Original copy and component structure; no competitor text found. |
| Next-prompt suggestions | ChatGPT/Codex-style follow-up affordances | Original prompt-suggestion parser and Pi-Office context rules | `addin/packages/pi-office-pack/src/prompt-suggestions.ts`, `addin/apps/taskpane/src/app/PromptSuggestionStrip.tsx` | Original parser and UI copy; source tests cover behavior. |
| Permission and autonomy model | Codex/OpenCode-style approval concepts | Original categories in Pi-Office protocol plus user preferences | `TOOL_CATEGORY_MAP`, `AUTONOMY_LEVEL_AUTO_APPROVE`, tool permission popup | Original implementation; no copied permission prompt text. |
| Connector marketplace/setup UX | Connector marketplaces and MCP ecosystem expectations | Original connector catalog, MCP protocol concepts, read-only policy | `addin/apps/taskpane/src/lib/runtime/connector-catalog.ts`, `IntegrationsSection.tsx` | Original setup flow; connector brand assets require separate third-party license/source audit before public packaging. |
| Visual/image generation workflows | ChatGPT image workflows and Office visual-assistant expectations | OpenAI Images API where configured, Office.js insertion paths | `generate_image`, `ImageBlock`, image model catalog | Original implementation; provider support honesty tracked separately in `BUG-006`. |
| Rewind/checkpoint workflow | AI editor undo/review expectations | Original browser checkpoint store and Office snapshot contracts | `BrowserCheckpointStore`, taskpane rewind UI | Original implementation; Excel formula fidelity tracked in `BUG-008`. |
| Companion sandbox direction | Codex/OpenCode safety patterns | Original companion architecture and future sandbox policy | `AGENTS.md`, `SECURITY-006` | No raw shell implementation is present; detailed sandbox design remains a separate P0 task. |
| Workflow packs/playbooks | Claude host playbook density and professional Office tasks | Original Pi-Office workflow-pack registry built around public Office.js tool contracts and Pi-Office review gates | `addin/packages/pi-office-pack/src/workflow-packs.ts`, `addin/packages/pi-office-pack/src/defaults.ts`, `addin/packages/pi-office-pack/skills/office-host.SKILL.md`, `addin/apps/taskpane/src/lib/helpers.ts`, `addin/scripts/office-tests/src/workflow-packs.test.ts` | Original specs and prompt copy; tests cover host coverage, tool references, prompt/skill injection, and taskpane starter prompts. |

## Asset Notes

- Pi-Office brand assets under `addin/apps/taskpane/public/brand` are project-owned/provided assets; generated PNG sizes are derived from the project SVG.
- Connector logos under `addin/apps/taskpane/public/connectors` are vendor-identification assets, not competitor-add-in assets. Their source/license status should be audited before public distribution; this is tracked separately in `SECURITY-007`.
- Do not import icons, screenshots, SVGs, prompt text, or bundled resources from competitor add-ins.

## Contribution Rules

- New features inspired by a competitor must add or update one row in the provenance matrix before merge.
- The row must state the user problem, the public API or original technical basis, the Pi-Office implementation files, and validation evidence.
- Any copied-looking UI text, prompt language, image, icon, or interaction should be rewritten in Pi-Office language before review.
- Do not commit local capture files, downloaded competitor bundles, private endpoint traces, or generated screenshots unless a maintainer explicitly marks them as non-release research material.
- Third-party brand marks need a source/license note or a fallback neutral badge before public release packaging.

## Release Gate

Before public release, every competitor-inspired feature should have:

- A user problem statement.
- A public API or original technical basis.
- Original UI copy and prompt text.
- Original implementation code.
- Validation evidence or a backlog link.
