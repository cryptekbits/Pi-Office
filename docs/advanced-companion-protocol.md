# Advanced Companion Protocol

`FEATURE-006` splits Pi-Office into two honest runtime modes:

- **Basic:** the Office taskpane owns the Pi agent, provider API-key calls, chat loop, Office.js execution, and browser-supported connector fallback.
- **Advanced:** the companion will own model inference, provider auth/session state, MCP and non-Office tools, memory, and local workspace context. The taskpane remains the presentation layer, Office.js executor, review surface, and permission owner for Office document writes.

The shared protocol descriptor lives in `addin/packages/pi-office-pack/src/protocol.ts` as `TASKPANE_COMPANION_PROTOCOL`. Companion health and taskpane fallback state advertise that descriptor through `CompanionCapabilities.protocol`.

## Contract

| Area | State | Owner | Current route contract |
| --- | --- | --- | --- |
| Capability discovery | Available | Shared | `GET /v1/health`, `POST /v1/sessions/open` |
| Settings sync | Planned | Shared | No route yet; non-secret preferences and connector config should sync explicitly. |
| Chat streaming | Reserved | Companion | `POST /v1/sessions/:sessionId/agent/prompt` exists as an unavailable stub until companion provider auth/session storage lands. |
| Tool requests | Available | Companion | File, MCP, MCP result, shell, and native-capture session routes execute non-Office work under capability gates. |
| Office tool execution | Reserved | Taskpane | `POST /v1/sessions/:sessionId/agent/office-tool-result` is reserved for future companion-agent Office tool result handoff. Office.js execution stays in the taskpane. |
| Auth migration | Planned | Shared | No route yet; provider secrets require explicit user action and secure companion storage. |

## Non-Negotiables

- Office document reads/writes, selection-sensitive actions, and Office permission prompts remain taskpane-owned because only the Office host can safely run Office.js against the active document.
- Taskpane provider secrets must never silently migrate into the companion. Any move to companion provider auth needs explicit user action, clear storage disclosure, and a secure backend.
- Basic mode must stay usable when the companion is absent, stopped, unhealthy, or disconnected.
- Companion-owned inference must not be advertised as available until `CompanionCapabilities.agent` and `CompanionCapabilities.providerAuth` both report available.
- Reserved routes should fail closed with clear unavailable responses until their capability state changes.

## Validation

`addin/scripts/office-tests/src/companion-protocol.test.ts` protects the descriptor shape, the Basic/Advanced ownership rules, the no-silent-secret-migration rule, and the reserved companion-agent route stubs.
