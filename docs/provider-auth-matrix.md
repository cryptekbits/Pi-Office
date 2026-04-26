# Pi-Office Provider/Auth And Model Matrix

Last updated: 2026-04-26

This matrix records what Pi-Office can honestly execute today versus what is only planned for companion-owned auth. **Basic** means the browser taskpane owns inference and Office.js execution. **Pro** means the optional companion will own inference, provider auth/session state, MCP, memory, and non-Office tools while the taskpane remains the Office.js executor. **Simple** and **Advanced** are only settings catalog detail levels.

Canonical model curation lives in `addin/packages/pi-office-pack/src/provider-model-preferences.yaml`. Regenerate the typed runtime module with:

```sh
npm --prefix addin run generate:provider-models
npm --prefix addin run check:provider-models
```

The current generated catalog was validated against refreshed `pi-mono` commit `05f79b08` and `@mariozechner/pi-ai@0.70.2`.

## Runtime Dependency

Pi-Office consumes Pi as the orchestration and model-catalog runtime through `pi-mono` / `@mariozechner/pi-ai`. Pi is not a selectable model provider in the provider/auth matrix; the rows below are only user-facing inference providers or provider-like subscription/auth surfaces.

## Status Legend

| Status | Meaning |
| --- | --- |
| `supported` | Exposed as browser-callable in Basic today. Users may store a provider API key and the model can execute from the taskpane. |
| `planned` | A real provider exists in Pi's model registry, but Pi-Office needs Pro/companion runtime or auth work before execution. |
| `blocked` | Not viable for the current architecture until an external requirement changes. |
| `research_only` | Product target or idea is documented, but there is no active runtime contract yet. |

## Provider/Auth/Model Matrix

| Provider | Status | Runtime Surface | Auth Method | Simple Models | Advanced Availability | Source URL | Pi Catalog Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| OpenAI API | `supported` | Basic | API key | `gpt-5.5` default, `gpt-5.4-pro`, `gpt-5.4`, `gpt-4.1` | Full OpenAI Pi catalog in Advanced; un-recommended picks warn. | https://developers.openai.com/api/docs/models | Preferred `gpt-5.5-pro` is absent after refresh; recorded as drift and not added as a local override. |
| ChatGPT / OpenAI Codex OAuth | `planned` | Pro | OAuth/subscription token brokerage | `gpt-5.5` default, `gpt-5.4`, `gpt-5.3-codex` | Full `openai-codex` catalog visible setup-only until Pro auth exists. | https://developers.openai.com/api/docs/models | `/v1/auth/start` rejects browser OAuth today. |
| Anthropic API | `supported` | Basic | API key | `claude-opus-4-7` default, `claude-opus-4-6`, `claude-sonnet-4-6` | Full Anthropic Pi catalog in Advanced; older Claude models warn if selected. | https://www.anthropic.com/claude/opus and https://www.anthropic.com/news/claude-sonnet-4-6 | Direct API-key path is executable; Claude subscription OAuth remains Pro work. |
| Google Gemini API | `supported` | Basic | API key | `gemini-3.1-pro-preview` default, `gemini-3-pro-preview`, `gemini-3-flash-preview` | Full Gemini API catalog in Advanced. | https://ai.google.dev/gemini-api/docs/models | Gemini CLI/Antigravity OAuth paths are separate Pro providers. |
| Google Gemini CLI / Cloud Code Assist | `planned` | Pro | OAuth | Advanced/setup-only for now | Full Pi catalog visible in Advanced only until companion auth exists. | https://ai.google.dev/gemini-api/docs/models | Requires Google OAuth and project/account state. |
| Google Antigravity | `planned` | Pro | OAuth | Advanced/setup-only for now | Full Pi catalog visible in Advanced only until companion auth exists. | https://ai.google.dev/gemini-api/docs/models | Requires Google OAuth/token brokerage. |
| GitHub Copilot OAuth | `planned` | Pro | OAuth/subscription token brokerage | `gpt-5.4` default, `claude-opus-4.6`, `claude-sonnet-4.6`, `gemini-3.1-pro-preview` | Full Copilot catalog visible setup-only until Pro auth exists. | https://docs.github.com/copilot/reference/ai-models/supported-models | Browser API-key setup is rejected. |
| OpenCode Zen | `supported` | Basic | API key | `claude-opus-4-7` default, `gpt-5.4`, `gemini-3.1-pro`, `kimi-k2.6`, `glm-5.1` | Full OpenCode catalog in Advanced. | Pi catalog / OpenCode provider registry | Browser API-key path is executable. |
| OpenCode Go | `supported` | Basic | API key | `kimi-k2.6` default, `kimi-k2.5`, `glm-5.1` | Full OpenCode Go catalog in Advanced. | Pi catalog / OpenCode provider registry | Browser API-key path is executable. |
| OpenRouter | `supported` | Basic | API key | `openai/gpt-5.4` default, `google/gemini-3.1-pro-preview`, `x-ai/grok-4.20`, `moonshotai/kimi-k2.6`, `deepseek/deepseek-v4-pro` | Full OpenRouter catalog in Advanced. | https://openrouter.ai/docs/guides/overview/models | Gateway path keeps many labs Simple without enabling direct regional provider setup by default. |
| Vercel AI Gateway | `supported` | Basic | API key | `openai/gpt-5.4` default, `google/gemini-3.1-pro-preview`, `xai/grok-4.20-reasoning`, `moonshotai/kimi-k2.6`, `zai/glm-5.1` | Full Gateway catalog in Advanced. | https://vercel.com/docs/ai-gateway/models-and-providers | Gateway path remains API-key backed. |
| Groq | `supported` | Basic | API key | `moonshotai/kimi-k2-instruct` default, `openai/gpt-oss-120b` | Full Groq catalog in Advanced. | https://console.groq.com/docs/models | Kimi is exposed through Groq; direct Kimi Coding remains Advanced. |
| xAI | `supported` | Basic | API key | `grok-4.20-0309-reasoning` default, `grok-4.20-0309-non-reasoning` | Full xAI catalog in Advanced. | https://docs.x.ai/overview | Current direct Grok 4.20-compatible slugs came from refreshed Pi. |
| DeepSeek API | `supported` | Basic, Advanced settings only | API key | Not Simple as a direct provider | Direct `deepseek-v4-pro` and `deepseek-v4-flash` are recommended in Advanced; unlisted DeepSeek models warn. | https://api-docs.deepseek.com/quick_start/pricing/ | Simple users can still reach DeepSeek through OpenRouter. |
| Z.AI API | `supported` | Basic, Advanced settings only | API key | Not Simple as a direct provider | Direct `glm-5.1` and `glm-5` are recommended in Advanced; unlisted GLM models warn. | https://docs.z.ai/guides/llm/glm-5 | Simple users can still reach GLM through Vercel/OpenCode. |
| Kimi Coding / Moonshot | `supported` | Basic, Advanced settings only | API key | Not Simple as a direct provider | Direct `kimi-k2-thinking` and `k2p6` are recommended in Advanced. | https://platform.moonshot.ai/ | Simple users can still reach Kimi through OpenRouter, Vercel, Groq, or OpenCode. |
| Azure OpenAI Responses | `planned` | Pro | API key plus endpoint/deployment config | Advanced/setup-only for now | Full Pi catalog visible in Advanced only after tenant-specific config exists. | Microsoft Azure OpenAI docs | A plain API-key field is insufficient because deployments are tenant-specific. |
| Amazon Bedrock | `planned` | Pro | AWS credentials or Bedrock bearer token | Advanced/setup-only for now | Full Bedrock catalog visible in Advanced only after companion credential handling exists. | AWS Bedrock docs | Requires AWS credential discovery or token handling outside browser localStorage. |
| Google Vertex AI | `planned` | Pro | API key or cloud identity | Advanced/setup-only for now | Full Vertex catalog visible in Advanced only after project/location config exists. | Google Vertex AI docs | Requires project/location and ADC/API-key handling beyond the current provider key field. |
| Cloudflare Workers AI / Gateway | `research_only` | Not implemented | API token or gateway key | Not Simple | Not in active Pi provider registry used by Pi-Office. | Cloudflare docs | Needs a provider contract before UI exposure. |
| Cerebras, Fireworks, Hugging Face, MiniMax, Mistral, other Pi providers | `supported` or `research_only` by runtime flag | Basic when browser API-key callable, otherwise Advanced/research | API key where supported | Not Simple unless explicitly curated later | Full Pi catalog in Advanced with un-recommended warnings. | Provider docs / Pi catalog | Kept out of Simple to avoid overwhelming first-time users. |

## Runtime Enforcement

- `/v1/providers` carries capability fields plus curation metadata for each model: `settingsVisibility`, `lab`, `family`, `recommended`, `recommendationReason`, `defaultForProvider`, and `requiresUnrecommendedWarning`.
- Simple settings show only Simple-visible providers and recommended Simple models. Advanced settings expose the full Pi catalog.
- Model selection warns before using any enabled un-recommended model unless the user enables `suppressUnrecommendedModelWarning`.
- `defaultModelByProvider` stores user defaults and is validated against the Pi catalog before persistence.
- `/v1/auth/api-key` rejects providers that are not browser-callable API-key providers, so users cannot create false configured state for Codex, Copilot, Gemini CLI, Antigravity, Bedrock, Vertex, or Azure.
- `/v1/auth/start` remains unavailable in Basic browser-only mode and returns a Pro/companion-owned OAuth message for providers that need OAuth.
- Image generation remains OpenAI-only in the browser taskpane.

## Next Implementation Order

1. Keep browser API-key providers honest and verified through readiness states.
2. Implement one Pro/companion-owned OAuth provider at a time, starting with the provider that has the clearest official token-refresh contract.
3. Move provider secrets to companion/OS keychain storage as part of `FEATURE-006` before advertising subscription-backed sign-in broadly.
4. Revisit non-OpenAI image providers only after their runtime execution paths exist.
