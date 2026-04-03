# Pi-Office

Pi-powered Microsoft Office add-in scaffold for Word, Excel, and PowerPoint. The taskpane and local companion are built around Pi as the runtime core.

## What is in this repo

- `apps/taskpane`
  Shared React taskpane app served over `https://localhost:3443`
- `apps/companion`
  Local HTTPS companion that owns Pi auth, sessions, workspace access, and the Office bridge
- `packages/pi-office-pack`
  Internal Pi package with Office-specific tools, prompts, and skills
- `manifests`
  Separate XML manifests for Word, Excel, and PowerPoint
- `certs`
  Local development certificate artifacts generated for the companion

## Current architecture

- Office hosts the taskpane UI and runs `Office.js`
- The companion runs Pi locally and serves the taskpane on the same origin
- Pi treats Office manipulation as custom tools exposed by `pi-office-pack`
- The taskpane executes those Office tool calls locally and returns results to the companion over WebSocket
- Unsaved documents run in `document-only` mode
- Saved documents switch to `workspace` mode and bind Pi to the document directory so AGENTS/skills/file context become available

## Local setup

1. Install dependencies:

```bash
npm install
```

2. Create and trust the local HTTPS certificate used by the companion:

```bash
npm run prepare:certs
```

3. Start the Pi-Office stack:

```bash
npm run dev
```

The companion listens on `https://localhost:3443`.

4. Sideload a host manifest:

```bash
npm run sideload:word
```

Or:

```bash
npm run sideload:excel
npm run sideload:powerpoint
```

5. Remove the sideload registration when you are done:

```bash
npm run sideload:stop
```

## Validation

```bash
npm run typecheck
npm run build
npm run validate:manifests
```

## Sideload flow

Use the XML files in [`manifests`](/C:/Users/manan/Code/Personal/office-word-addin/manifests) for desktop/manual sideload. All three point at the same local companion origin.

- Word: [`word.xml`](/C:/Users/manan/Code/Personal/office-word-addin/manifests/word.xml)
- Excel: [`excel.xml`](/C:/Users/manan/Code/Personal/office-word-addin/manifests/excel.xml)
- PowerPoint: [`powerpoint.xml`](/C:/Users/manan/Code/Personal/office-word-addin/manifests/powerpoint.xml)

If Word launches with a blank document during debugging, open the target saved document in that same Word instance and then open the add-in there. The taskpane session tracks the document it is attached to, not another Word window.

## Notes

- Pi is consumed as a dependency. This repo does not copy Pi source from `pi-mono`.
- The Office bridge is intentionally local-first and Windows desktop-oriented for v1.
- OAuth support is scaffolded through Pi auth storage, but API key entry is also supported so OpenAI/OpenRouter/Anthropic/Copilot-style provider setups are possible without a single-vendor lock-in.
- No commits have been made.
