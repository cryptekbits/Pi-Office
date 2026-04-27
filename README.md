# Pi-Office

Pi-powered Microsoft Office add-in scaffold for Word, Excel, and PowerPoint.

## What Is In This Repo

- `addin/apps/taskpane`
  Shared React taskpane app served over `https://localhost:3443`
- `companion`
  Optional local HTTPS companion for read-only local file access and read-only MCP execution
- `addin/packages/pi-office-pack`
  Internal Pi package with Office-specific tools, prompts, and skills
- `addin/manifests`
  Separate XML manifests for Word, Excel, and PowerPoint
- `addin/certs`
  Local development certificate artifacts used by local HTTPS services
- `website`
  Placeholder for the future public website
- `docs`
  Detailed project docs, trackers, provenance notes, and design references

## Current Architecture

- Office hosts the taskpane UI and runs `Office.js`
- The taskpane runtime is self-contained for chat, providers, models, Office tools, and browser-safe connectors
- The optional companion is a separate loopback capability provider on `https://localhost:3444`
- Saved documents provide folder context, but local file tools stay disabled until the optional companion connects
- Local stdio, local HTTP, and companion-brokered MCP connectors require the optional companion; browser-compatible hosted HTTP profiles can run directly when their catalog profile and provider CORS allow it
- Multiple Office windows can reuse one machine-local companion while keeping logical taskpane sessions isolated
- Privacy and storage behavior, including local credential limits and clear-data controls, is documented in [`docs/privacy-and-storage.md`](docs/privacy-and-storage.md)

## Local Setup

1. Install add-in dependencies:

```bash
npm install --prefix addin
```

2. Install optional companion dependencies:

```bash
npm install --prefix companion
```

3. Create and trust the local HTTPS certificate:

```bash
npm run prepare:certs
```

4. Start the taskpane dev host for sideload development:

```bash
npm run dev
```

This starts the taskpane web host on `https://localhost:3443`.

Opening `https://localhost:3443` directly in a normal browser starts dev-only browser preview mode. The preview uses a synthetic unsaved Office context so the React taskpane, settings, provider catalog, connector setup UI, and browser-safe runtime surfaces can be debugged without Word/Excel/PowerPoint. Real Office.js document reads, edits, selection refresh, and visual snapshots still require sideloading inside an Office host. Use `?piOfficeHost=excel` or `?piOfficeHost=powerpoint` to preview host-specific chrome, and `?piOfficeBrowserDebug=0` to disable the fallback.

In sideload or localhost development, the main chat header shows a conversation-log icon immediately before **New chat**. It copies the active conversation debug log to the clipboard as JSONL, including visible chat, model-emitted reasoning deltas when the provider sends them, raw bridge/session events, tool calls, tool results, Office state, and runtime diagnostics. Known secret-bearing fields are redacted, but prompts and document snippets are intentionally included, so review the clipboard contents before sharing them for debugging.

5. Optionally start the local companion:

```bash
npm run dev:companion
```

This starts the optional companion on `https://localhost:3444`.

6. Sideload a host manifest:

```bash
npm run sideload:word
```

Or:

```bash
npm run sideload:excel
npm run sideload:powerpoint
```

7. Remove the sideload registration when you are done:

```bash
npm run sideload:stop
```

## Validation

```bash
npm run typecheck
npm run build
npm run validate:manifests
```

## Sideload Flow

Use the XML files in [`addin/manifests`](addin/manifests) for desktop/manual sideload. All three point at the taskpane dev host on `https://localhost:3443`.

- Word: [`word.xml`](addin/manifests/word.xml)
- Excel: [`excel.xml`](addin/manifests/excel.xml)
- PowerPoint: [`powerpoint.xml`](addin/manifests/powerpoint.xml)

If Word launches with a blank document during debugging, open the target saved document in that same Word instance and then open the add-in there. The taskpane session tracks the document it is attached to, not another Word window.

## Companion Notes

- Pi-Office works without the companion
- Without the companion, local files, local stdio/local HTTP MCP execution, and companion-brokered OAuth connectors are unavailable
- With the companion connected, Pi-Office enables read-only `read`, `grep`, `find`, and `ls` for the saved document folder
- With the companion connected, verified read-safe local stdio, local HTTP, companion-required remote HTTP, and brokered OAuth MCP tools can be exposed through the generic `mcp` tool
- Smart Auto routing prefers companion execution for eligible non-Office capabilities when the companion advertises support, while keeping taskpane fallback for browser-supported providers and image generation
- True viewport/window screenshots are companion-native only; taskpane-only visual tools remain Office.js snapshots and metadata
- Shell/bash access is hidden unless the companion shell sandbox reports an available isolation backend and passing destructive probes; there is no raw host shell fallback
- Companion-owned provider auth, inference, memory, and agent sessions are capability-gated advanced-mode work. Browser API-key providers remain taskpane-local unless the user explicitly moves auth to the companion.
- Browser-direct hosted HTTP connectors can run from the taskpane when supported; companion-brokered connectors such as Granola use the companion for system-browser sign-in and MCP execution
- Packaging for `dist/binaries`, zip, and npm distribution is planned later
- Runtime capability boundaries are documented in [`docs/capability-boundaries.md`](docs/capability-boundaries.md)

## Notes

- Pi is consumed as a dependency. This repo does not copy Pi source from `pi-mono`.
- `npm run dev` means "start the local taskpane web host for sideload development."
- The optional companion is intentionally read-only for v1 file and MCP access.
- Contributor originality rules live in [`CONTRIBUTING.md`](CONTRIBUTING.md), with release provenance tracked in [`docs/provenance.md`](docs/provenance.md).
