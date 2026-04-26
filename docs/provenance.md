# Pi-Office Provenance Notes

This file tracks the release-safety rule for competitor-inspired work.

## Policy

- Study competitor add-ins for user expectations, capability gaps, and workflow ideas only.
- Do not copy competitor source code, prompts, UI text, assets, private API contracts, or bundled resources.
- Rebuild useful capabilities from public APIs, original Pi-Office UX, and original implementation.
- Treat Office.js and Microsoft documentation as the platform basis for Office behavior.

## Current Audit Matrix

| Source | Allowed Use | Do Not Use | Pi-Office Direction |
| --- | --- | --- | --- |
| Microsoft Office.js docs | Public API behavior and host capabilities | N/A | Use documented Word, Excel, and PowerPoint APIs as the implementation foundation. |
| Claude Office add-in capture | Capability mapping, host-workflow expectations, validation gaps | Bundle code, prompts, UI text, assets, private endpoint contracts | Re-spec workflow packs and Office tools in original Pi-Office wording and code. |
| ChatGPT/Codex-style add-in captures | UX hardening ideas, CSP/privacy comparison, cross-host workflow expectations | Bundle code, generated assets, private API contracts, prompts | Build original provider-agnostic UX and capability-honesty checks. |
| Microsoft Copilot | Enterprise expectation benchmark | Branding, proprietary behaviors, copied interactions | Compete through transparent provider choice, Office-native edits, and better professional document UX. |
| OpenCode/Codex sandbox designs | Public safety patterns for permissions and sandboxing | Vendored code or unreviewed security claims | Design a companion sandbox with original policy, tests, and platform-specific implementation. |

## Release Gate

Before public release, every competitor-inspired feature should have:

- A user problem statement.
- A public API or original technical basis.
- Original UI copy and prompt text.
- Original implementation code.
- Validation evidence or a backlog link.
