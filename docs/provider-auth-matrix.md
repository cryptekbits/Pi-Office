# Pi-Office Provider And Auth Matrix

Last updated: 2026-04-26

This matrix records what Pi-Office can honestly execute today versus what is only planned for companion-owned auth. The active browser taskpane can store API keys locally and pass them to Pi's provider registry. It cannot start OAuth or safely broker subscription tokens; those paths require the companion/advanced-mode work before they should appear as usable sign-in options.

## Status Legend

| Status | Meaning |
| --- | --- |
| `supported` | Exposed as browser-callable in the taskpane today. Users may store a provider API key and the model can execute from the taskpane. |
| `planned` | A real provider exists in Pi's model registry, but Pi-Office needs more runtime/auth work before exposing it as executable. |
| `blocked` | Not viable for the current architecture until an external requirement changes. |
| `research_only` | Product target or idea is documented, but there is no active runtime contract yet. |

## Current Provider Catalog

| Provider | Status | Current Pi-Office Surface | Auth Methods | Browser Callable | Companion Required | Subscription Backed | Image Generation | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Pi runtime | `supported` | Dependency/runtime layer, not a model provider | N/A | N/A | No | No | No | Pi is consumed as the agent/model orchestration dependency. Pi-Office does not vendor Pi source. |
| OpenAI API | `supported` | Browser taskpane | API key | Yes | No | No | Yes | Chat/text models and the current OpenAI-only image generation path are executable with a stored API key. |
| ChatGPT / OpenAI Codex subscription | `planned` | Companion | OAuth | No | Yes | Yes | No | `openai-codex` models require ChatGPT subscription OAuth/token brokerage. The browser `/v1/auth/start` route intentionally rejects this today. |
| Anthropic API | `supported` | Browser taskpane | API key | Yes | No | No | No | Direct Anthropic API keys are accepted. Claude subscription OAuth is not implemented in the browser. |
| Anthropic subscription OAuth | `planned` | Companion | OAuth | No | Yes | Yes | No | Needs companion-owned OAuth storage/refresh before public UI can advertise it. |
| Google Gemini API | `supported` | Browser taskpane | API key | Yes | No | No | No | Direct Gemini API keys are accepted for text models. |
| Google Vertex AI | `planned` | Companion | API key or cloud identity | No | Yes | No | No | Requires project/location and ADC/API-key handling beyond the current provider key field. |
| Google Gemini CLI / Cloud Code Assist | `planned` | Companion | OAuth | No | Yes | Yes | No | Requires Google OAuth and project/account state. |
| Google Antigravity | `planned` | Companion | OAuth | No | Yes | Yes | No | Requires Google OAuth/token brokerage. |
| GitHub Copilot | `planned` | Companion | OAuth | No | Yes | Yes | No | Requires Copilot subscription OAuth/token brokerage and model enablement checks. |
| OpenCode Zen / OpenCode Go | `supported` | Browser taskpane | API key | Yes | No | No | No | Exposed through Pi's OpenCode providers with `OPENCODE_API_KEY`-style credentials. |
| OpenRouter | `supported` | Browser taskpane | API key | Yes | No | No | No | Text models are API-key backed. Image generation is not implemented for OpenRouter in Pi-Office. |
| Vercel AI Gateway | `supported` | Browser taskpane | API key | Yes | No | No | No | Text models are API-key backed. |
| Cloudflare Workers AI / Gateway | `research_only` | Not implemented | API token or gateway key | No | TBD | No | No | Not present in the active Pi provider registry used by Pi-Office. Needs a provider contract before UI exposure. |
| Azure OpenAI Responses | `planned` | Companion | API key plus endpoint/deployment config | No | Yes | No | No | A plain API-key field is insufficient because Azure deployments are tenant-specific. |
| Amazon Bedrock | `planned` | Companion | AWS credentials or Bedrock bearer token | No | Yes | No | No | Requires AWS credential discovery or token handling outside browser localStorage. |
| Mistral, Groq, Cerebras, xAI, Hugging Face, MiniMax, Kimi, Z.AI | `supported` | Browser taskpane | API key | Yes | No | No | No | Exposed as API-key text providers through Pi's browser-compatible registry. Regional/enterprise-risk UX remains tracked under `IMPROVEMENT-005`. |

## Runtime Enforcement

- `/v1/providers` now carries capability fields for each provider: support status, runtime surface, auth methods, browser-callable state, companion requirement, subscription backing, and image support.
- `/v1/auth/api-key` rejects providers that are not browser-callable API-key providers, so a user cannot create false "configured" state for Codex, Copilot, Gemini CLI, Antigravity, Bedrock, Vertex, or Azure.
- `/v1/auth/start` remains unavailable in browser-only mode and returns a companion-owned OAuth message for providers that need OAuth.
- The settings UI shows companion-required and planned providers as informational instead of presenting them as direct browser setup paths.

## Next Implementation Order

1. Keep browser API-key providers honest and verified through `BUG-005` readiness states.
2. Implement one companion-owned OAuth provider at a time, starting with the provider that has the clearest official token-refresh contract.
3. Move provider secrets to companion/OS keychain storage as part of `FEATURE-006` before advertising subscription-backed sign-in broadly.
4. Revisit image providers only after their runtime execution path exists; the active image catalog remains OpenAI-only.
