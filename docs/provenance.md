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
| Deferred tool discovery, programmatic batches, and MCP result handles | Public Claude tool-use docs for tool search, programmatic tool calling, and tool-context management: `https://platform.claude.com/docs/en/agents-and-tools/tool-use/overview`, `https://platform.claude.com/docs/en/agents-and-tools/tool-use/programmatic-tool-calling`, `https://platform.claude.com/docs/en/agents-and-tools/tool-use/manage-tool-context`, `https://platform.claude.com/docs/en/agents-and-tools/tool-use/tool-search-tool` | Original Pi-Office registry/dispatcher/batch/result-handle contracts over public Office.js, MCP, and local companion boundaries | `protocol.ts`, `capabilities.ts`, `defaults.ts`, `office/tools/`, `inprocess-kernel.ts`, `batch-executor.ts`, `mcp-result-store.ts`, `companion/src/connector-bridge.ts`, `companion/src/server.ts` | Idea-level inspiration only. `FEATURE-028` through `FEATURE-030` and `BUG-019` record the original Pi-Office implementation and regression evidence; no Anthropic code, private API contracts, prompt text, or UI text was copied. |
| Connector marketplace/setup UX | Connector marketplaces and MCP ecosystem expectations | Original connector catalog, public MCP docs, per-connector setup research, read-only policy | `addin/packages/pi-office-pack/src/connector-catalog.yaml`, `addin/apps/taskpane/src/lib/runtime/connector-catalog-generated.ts`, `IntegrationsSection.tsx`, `companion/src/connector-bridge.ts` | Original profile-based setup flow and tool-safety UX; connector brand assets accepted for the current release path by stakeholder decision. |
| Browser-direct hosted MCP setup | MCP OAuth/DCR and Streamable HTTP ecosystem | First-party MCP docs and original Pi-Office browser MCP client | `connector-catalog.yaml`, `browser-mcp-client.ts`, `browser-connectors.ts`, `IntegrationsSection.tsx` | `FEATURE-008` records first-party provenance per setup profile. `BUG-015` keeps hosted HTTP setup from being companion-gated, marks Slack coming soon because first-party docs require a registered Slack app and no DCR, and treats Granola browser-side DCR as provider-CORS-dependent with fail-closed broker guidance. `BUG-018` refreshes Parallel Search MCP endpoints from first-party docs. Perplexity is official local STDIO/API-key; Obsidian and PostgreSQL Reader are community/reference local-only; Google Drive exact endpoint is planned/unverified until first-party evidence is attached. |
| Companion-brokered connector OAuth | MCP OAuth/DCR, system-browser sign-in, Office taskpane callback limits | First-party connector OAuth docs plus original loopback companion broker | `connector-catalog.yaml`, `protocol.ts`, `companion/src/oauth-broker.ts`, `companion/src/server.ts`, `companion/src/connector-bridge.ts`, `IntegrationsSection.tsx` | `FEATURE-019` adds YAML-driven broker settings, PKCE/DCR/token exchange in the companion, and system-browser launch only for profiles that explicitly opt into `oauth.broker=companion`. Granola is the first enabled brokered profile because Office-webview DCR is blocked by provider CORS. Token storage hardening to OS keychain/platform encryption is tracked in `SECURITY-008`. |
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

## Connector Curation Notes

- Connector setup facts now live in `addin/packages/pi-office-pack/src/connector-catalog.yaml`; the generated TypeScript catalog is checked by `npm --prefix addin run check:connector-catalog`.
- Every setup profile carries officialness, availability, docs/evidence URLs, checked date, browser-direct support, companion requirement, and risk notes.
- Profiles marked `planned` or `setupDisabled` stay visible as roadmap entries but cannot be configured from the wizard.
- Community/reference profiles are intentionally available only behind an extra warning and remain subject to read-only tool classification.
- Hosted HTTP profiles should not surface companion polling as a setup prerequisite. The wizard may attempt direct taskpane verification for hosted endpoints and fail closed on CORS/auth limitations; provider-specific OAuth broker work must be tracked instead of relabeling hosted MCP as local-companion-only.
- Slack MCP is marked planned/coming soon as of 2026-04-27 because Slack documents Streamable HTTP at `https://mcp.slack.com/mcp`, no Dynamic Client Registration, and confidential OAuth backed by a registered Slack app.
- Granola remains an official hosted OAuth MCP, but live Office webviews can hit provider CORS at DCR/token exchange. As of `FEATURE-019`, its profile is companion-brokered: sign-in opens in the system browser, the loopback companion receives `/v1/connectors/oauth/callback`, and the companion injects the stored bearer token during MCP verification/execution.
- Parallel Search MCP was refreshed on 2026-04-27 from first-party docs: anonymous/free search uses `https://search.parallel.ai/mcp`, while OAuth, account attribution, organization controls, and ZDR use `https://search.parallel.ai/mcp-oauth` with Parallel OAuth metadata from `https://platform.parallel.ai/.well-known/oauth-authorization-server`.
- System-browser OAuth is enabled only for profiles with explicit broker metadata. `Office.context.ui.openBrowserWindow` opens the external browser; the companion broker owns state, DCR, PKCE, token exchange, refresh, and callback completion. Browser-direct profiles continue using the taskpane popup/callback path only when provider CORS allows it.

## Release Gate

Before public release, every competitor-inspired feature should have:

- A user problem statement.
- A public API or original technical basis.
- Original UI copy and prompt text.
- Original implementation code.
- Validation evidence or a backlog link.
