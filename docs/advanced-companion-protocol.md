# Advanced Companion Protocol

`FEATURE-006` splits Pi-Office into honest runtime ownership modes:

- **Basic:** the Office taskpane owns the Pi agent, provider API-key calls, chat loop, Office.js execution, and browser-supported connector fallback. Companion-only model tools stay hidden even if a companion is connected.
- **Smart Auto:** the default. The taskpane remains fully usable and eligible non-Office companion capabilities are published only when the companion advertises them.
- **Advanced:** the companion will own model inference, provider auth/session state, MCP and non-Office tools, memory, and local workspace context. The taskpane remains the presentation layer, Office.js executor, review surface, and permission owner for Office document writes. Companion API-key provider storage is now available where secure companion storage exists, while companion-owned agent/inference still uses honest taskpane fallback until agent sessions land.

The runtime preference lives in `UserPreferences.companionRuntimeMode` (`basic`, `smart_auto`, or `advanced`). The shared protocol descriptor lives in `addin/packages/pi-office-pack/src/protocol.ts` as `TASKPANE_COMPANION_PROTOCOL`. Companion health and taskpane fallback state advertise that descriptor through `CompanionCapabilities.protocol`.

## Contract

| Area | State | Owner | Current route contract |
| --- | --- | --- | --- |
| Capability discovery | Available | Shared | `GET /v1/health`, `POST /v1/sessions/open` |
| Settings sync | Available | Shared | `POST /v1/sessions/:sessionId/settings/sync` syncs non-secret preferences, enabled provider/model choices, and connector summary counts. Provider secrets are excluded. |
| Chat streaming | Reserved | Companion | `POST /v1/sessions/:sessionId/agent/prompt` exists as an unavailable stub until companion-owned Pi agent sessions land. |
| Tool requests | Available | Companion | File, MCP, MCP result, shell, and native-capture session routes execute non-Office work under capability gates. |
| Office tool execution | Reserved | Taskpane | `POST /v1/sessions/:sessionId/agent/office-tool-result` is reserved for future companion-agent Office tool result handoff. Office.js execution stays in the taskpane. |
| Companion provider auth setup | Available | Shared | `GET /v1/provider-auth/status`, `POST /v1/provider-auth/api-key`, and `DELETE /v1/provider-auth` expose explicit API-key setup/clear through secure companion storage. The taskpane Settings provider cards can copy an already stored taskpane API key into companion storage only after a user click and confirmation. Broader OAuth/cloud migration remains planned. |

## Non-Negotiables

- Office document reads/writes, selection-sensitive actions, and Office permission prompts remain taskpane-owned because only the Office host can safely run Office.js against the active document.
- Taskpane provider secrets must never silently migrate into the companion. Any move to companion provider auth needs explicit user action, clear storage disclosure, and a secure backend.
- The taskpane copy-to-companion bridge reads the existing taskpane API key inside the in-process kernel and sends it only to the local companion provider-auth route with `explicitUserAction=true`; React state and settings-sync payloads remain redacted.
- Settings sync is for non-secret state only: preferences, runtime mode, enabled providers/models, defaults, and connector counts. API keys, OAuth tokens, and manual connector secrets are excluded from this route.
- Companion provider API-key setup requires `explicitUserAction=true`, returns only redacted provider state, and fails closed when secure companion storage is unavailable.
- Basic mode must stay usable when the companion is absent, stopped, unhealthy, or disconnected.
- Basic mode must not publish companion-only tools to model turns.
- Companion-owned inference must not be advertised as available until `CompanionCapabilities.agent` and `CompanionCapabilities.providerAuth` both report available.
- Reserved routes should fail closed with clear unavailable responses until their capability state changes.

## Validation

`addin/scripts/office-tests/src/companion-protocol.test.ts` protects the descriptor shape, the Basic/Advanced ownership rules, the no-silent-secret-migration rule, settings-sync routing, companion provider-auth route declarations, and the reserved companion-agent route stubs. `addin/scripts/office-tests/src/companion-provider-auth.test.ts` covers secure API-key storage, explicit user-action gating, fail-closed unsupported storage, clear behavior, and route implementation. `addin/scripts/office-tests/src/capability-routing.test.ts` covers Basic mode hiding companion-only tools and Advanced mode falling back to taskpane inference while companion agent support is unavailable.
