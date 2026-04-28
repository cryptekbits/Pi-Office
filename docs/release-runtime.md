# Release Runtime Assumptions

This document is the current packaging and runtime contract for the source-available Pi-Office repository. It records what works from this checkout today and what should not be implied until a later packaged release exists.

## Shipped Surfaces

- The Office add-in surface is the shared taskpane app in `addin/apps/taskpane`.
- The internal Office Pi package is `addin/packages/pi-office-pack`; add-in and companion builds compile it first.
- The Office manifests live in `addin/manifests/word.xml`, `addin/manifests/excel.xml`, and `addin/manifests/powerpoint.xml`.
- The optional local companion lives in `companion`; it is a source package, not a published npm package, zip, binary, service installer, or auto-updater yet.
- The repo root is a command router. Add-in dependencies are installed under `addin/`; companion dependencies are installed under `companion/`.

## Development Runtime

Run these commands from the repository root unless noted otherwise.

| Goal | Command | Runtime assumption |
| --- | --- | --- |
| Install add-in dependencies | `npm install --prefix addin` | Uses `addin/package-lock.json`; no companion install is implied. |
| Install companion dependencies | `npm install --prefix companion` | Uses `companion/package-lock.json`; optional for taskpane-only development. |
| Create local HTTPS certs | `npm run prepare:certs` | Generates gitignored files under `addin/certs`. On Windows, the cert must be trusted in `Cert:\CurrentUser\Root`. |
| Start the taskpane | `npm run dev` | Runs pack watch plus Vite taskpane on `https://localhost:3443`; checks certs and port `3443` first. |
| Start only taskpane workspace dev | `npm run dev:taskpane` | Runs the taskpane Vite workspace script and still uses the dev cert path. |
| Start the optional companion | `npm run dev:companion` | Runs the companion on `https://localhost:3444`; not required for basic add-in load. |
| Validate the local sideload host | `npm run preflight:sideload` | Probes the taskpane dev host resources on `https://localhost:3443`; the companion is explicitly not required. |

Opening `https://localhost:3443` directly in a normal browser enters the dev-only browser preview mode. That preview uses a synthetic unsaved Office context and cannot execute real Office.js document reads, writes, selection refresh, or host visual snapshots.

## Sideload Runtime

Desktop/manual sideload uses the XML files in `addin/manifests`. All three manifests currently point to the taskpane dev host:

- `SourceLocation`, `FunctionFile`, `SupportUrl`, and resource URL `Taskpane.Url`: `https://localhost:3443/`
- Ribbon icons: `https://localhost:3443/brand/icon-16.png?v=20260421`, `icon-32.png?v=20260421`, and `icon-80.png?v=20260421`
- `ExtendedOverrides`: `https://localhost:3443/shortcuts.json`
- Manifest version: `1.0.0.2`

Use:

```bash
npm run sideload:word
npm run sideload:excel
npm run sideload:powerpoint
```

The root sideload commands run `preflight:sideload` first. If Office reports that a required resource cannot be downloaded, check `https://localhost:3443/`, `/shortcuts.json`, and the three icon URLs before debugging companion behavior. The companion endpoint on `3444` is optional and is not needed for the taskpane to load.

## Companion Runtime

Pi-Office works without the companion. When connected, the companion can provide guarded non-Office capabilities such as saved-folder read-only file tools, MCP connector execution, companion-brokered connector OAuth, secure companion provider API-key storage where supported, native capture where available, and shell capability only when the sandbox reports available.

The companion does not own model inference yet. Basic and Smart Auto taskpane sessions keep browser-supported inference in the taskpane. Advanced mode is allowed to prefer companion-owned inference only after `CompanionCapabilities.agent` and `CompanionCapabilities.providerAuth` both report available.

Companion sessions receive non-secret settings sync only. Provider API keys, OAuth tokens, and manual connector secrets require explicit setup or migration paths; they are never moved silently through settings sync.

## Build And Validation

CI runs on Node 24 with `actions/checkout@v6` and `actions/setup-node@v6`. The current local validation gates are:

```bash
npm run typecheck
npm run build
npm run check:bundle
npm run validate:manifests
npm run test:office
```

`npm run build` builds the add-in package/taskpane and companion TypeScript outputs for local validation. It does not produce a signed Office Store package, production-hosted manifest, companion installer, npm-distributed companion package, or zip/binary release.

## Packaging Status

The current distributable posture is source-available checkout plus local development/sideload scripts. Public production packaging still needs a separate release plan for:

- Production taskpane hosting and manifest URL replacement.
- Manifest/icon/shortcut cache-busting outside localhost.
- Companion distribution format, update channel, install/uninstall behavior, and secure storage expectations per OS.
- Manual Office desktop validation across the chosen Word, Excel, and PowerPoint release matrix.

Until that plan lands, docs and release notes should describe Pi-Office as locally runnable from source rather than as a packaged Office Store add-in or installed companion app.
