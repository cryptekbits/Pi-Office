# Pi-Office

Pi-powered Microsoft Office add-in scaffold for Word, Excel, and PowerPoint.

## What Is In This Repo

- `addin/apps/taskpane`
  Shared React taskpane app served over `https://localhost:3443`
- `companion`
  Optional local HTTPS companion for read-only local file access and read-only local MCP execution
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
- The optional companion is a separate capability provider on `https://localhost:3444`
- Saved documents provide folder context, but local file tools stay disabled until the optional companion connects
- Remote HTTP connectors can be configured in browser mode, but agent execution is setup-only until the browser remote-MCP execution path is implemented
- Local stdio connectors require the optional companion and stay read-only
- Multiple Office windows can reuse one machine-local companion while keeping logical taskpane sessions isolated

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

Use the XML files in [`addin/manifests`](/C:/Users/manan/Code/Personal/office-word-addin/addin/manifests) for desktop/manual sideload. All three point at the taskpane dev host on `https://localhost:3443`.

- Word: [`word.xml`](/C:/Users/manan/Code/Personal/office-word-addin/addin/manifests/word.xml)
- Excel: [`excel.xml`](/C:/Users/manan/Code/Personal/office-word-addin/addin/manifests/excel.xml)
- PowerPoint: [`powerpoint.xml`](/C:/Users/manan/Code/Personal/office-word-addin/addin/manifests/powerpoint.xml)

If Word launches with a blank document during debugging, open the target saved document in that same Word instance and then open the add-in there. The taskpane session tracks the document it is attached to, not another Word window.

## Companion Notes

- Pi-Office works without the companion
- Without the companion, local files and local stdio MCP connectors are unavailable
- With the companion connected, Pi-Office enables read-only `read`, `grep`, `find`, and `ls` for the saved document folder
- Shell/bash access is hidden unless the companion shell sandbox reports an available isolation backend and passing destructive probes; there is no raw host shell fallback
- The companion does not host providers, auth, models, or the taskpane
- Remote HTTP connectors are currently setup-only; do not present them as usable agent tools until the runtime exposes verified read-only remote MCP execution
- Packaging for `dist/binaries`, zip, and npm distribution is planned later

## Notes

- Pi is consumed as a dependency. This repo does not copy Pi source from `pi-mono`.
- `npm run dev` means "start the local taskpane web host for sideload development."
- The optional companion is intentionally read-only for v1 local access.
- Contributor originality rules live in [`CONTRIBUTING.md`](/C:/Users/manan/Code/Personal/office-word-addin/CONTRIBUTING.md), with release provenance tracked in [`docs/provenance.md`](/C:/Users/manan/Code/Personal/office-word-addin/docs/provenance.md).
