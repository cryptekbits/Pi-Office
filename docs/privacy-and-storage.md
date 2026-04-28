# Privacy And Storage

Pi-Office keeps Office.js execution inside the active Office taskpane. Network calls happen only when a configured AI provider, connector, or optional companion feature is used.

## What Leaves The Machine

- AI provider requests can include the user's prompt, selected Office context, generated-image prompts, model/tool results, and any document snippets the taskpane attaches for the current request.
- Remote HTTP connectors can send connector-specific requests to the configured service endpoint after the optional companion verifies read-safe MCP tools.
- Companion-brokered OAuth connectors open the provider sign-in page in the system browser. The provider redirects back to the loopback companion callback, and the companion exchanges the authorization code with the provider.
- Local stdio connectors also run through the optional companion when connected. The companion is designed for read-only MCP execution.
- The optional companion runs on loopback and receives saved-document context only when a saved document is bound to the session.

## What Stays In Browser Storage

- Provider credentials are stored under the taskpane origin in an AES-GCM envelope in `localStorage`.
- Connector configuration, connector secrets, OAuth credential handoffs, scope overrides, and redacted connector audit logs are stored under the taskpane origin in an AES-GCM envelope in `localStorage`. That envelope includes MCP fields such as service URL, stdio launch settings, inline environment key/value pairs, optional stdio host-env passthrough names, and remote HTTP static headers (header name/value pairs). Header *names* paired with environment-variable references are stored; resolved header values come from the companion process environment at execution time, not from browser storage.
- Saved chat history is stored in `localStorage` and may include prompts, document snippets, model responses, host metadata, and saved-document identifiers.
- User preferences, enabled providers/models, recent models, and persistent permission choices are stored locally.

The AES-GCM envelopes are local obfuscation only. The crypto keys are also stored in `localStorage`, so this does not protect secrets from same-origin script access, a compromised taskpane bundle, or XSS. Stronger storage such as OS keychain-backed companion storage is tracked as future provider/auth work.

## What Stays In Companion Storage

- On Windows companion hosts, companion-brokered connector OAuth tokens are stored under `.pi-office/companion/connector-oauth-tokens.dpapi.json` in a DPAPI `CurrentUser` encrypted envelope. Existing plain `.pi-office/companion/connector-oauth-tokens.json` files are migrated into that encrypted envelope and renamed with a `.migrated` suffix.
- On platforms without an implemented secure backend, companion-brokered connector OAuth tokens still fall back to `.pi-office/companion/connector-oauth-tokens.json` with local file permissions. This keeps provider tokens out of the Office taskpane, but it is not yet OS-keychain-backed storage on those platforms.
- The companion uses those stored tokens to add bearer authentication when verifying or executing the matching MCP connector. Tokens are not exported in connector bundles and are not copied back into taskpane `localStorage`.

`SECURITY-008` tracks the remaining follow-up to add secure macOS/Linux keychain backends and provider revoke controls.

## Clear Data Controls

The taskpane Settings -> Privacy tab can clear:

- All provider credentials and the provider credential encryption key.
- All connector configuration, connector secrets, OAuth state, connector scopes, connector logs, and the connector encryption key. When the optional companion is connected, this also clears companion-held connector OAuth tokens.
- Saved local chat history.

Per-provider auth removal remains available from Settings -> AI Providers, and per-connector removal remains available from Settings -> Integrations. When the optional companion is connected, per-connector removal also clears any companion-held OAuth token for that connector.

Connector records also store the selected setup profile, redacted config metadata, per-tool enablement overrides, and whether the user suppressed the advanced-tool warning for that connector. Exported connector bundles include profile and tool-policy metadata so teams can reproduce safe defaults, but secrets and OAuth tokens are still omitted and imported OAuth connectors must sign in again.

Hosted MCP profiles send connector requests and returned tool data to the named provider endpoint through the optional companion. Local STDIO profiles run a command on the companion machine and send tool results back into the active Pi-Office session. Non-read-only or unknown MCP tools are disabled unless a user explicitly enables them from the connector details view.

## Sideload Debug Export

In sideload or localhost development, the main chat header can copy the active conversation debug log directly to the clipboard as JSONL. The copy action is user-initiated and not uploaded by Pi-Office. It includes visible chat, model-emitted reasoning deltas when available, raw taskpane bridge/session events, tool calls, tool results, Office state, runtime diagnostics, prompts, and document snippets that were part of the session. Known secret-bearing fields and bearer/API-key shaped strings are redacted, but the clipboard contents should still be reviewed before sharing because free-form prompts, document content, and tool results are intentionally preserved for debugging.

## Telemetry Default

Pi-Office does not enable product analytics or telemetry by default. Provider services, connector services, and locally installed connector runtimes may have their own logging and retention policies outside Pi-Office.

## CSP Policy

The taskpane declares a `Content-Security-Policy` meta tag in `addin/apps/taskpane/index.html` because the dev/sideloaded Office taskpane is a static Vite app and may not control deployment headers in every host. Public deployments should emit the same or stricter policy as HTTP headers.

The current policy restricts the default source to the taskpane origin, allows Microsoft-hosted Office.js, allows provider/connector HTTPS calls and loopback companion calls, blocks plugins with `object-src 'none'`, and pins form submissions to `none`.

Two allowances are intentional and should be revisited before public release hardening:

- `'unsafe-inline'` is currently needed by inline Office bootstrap scripts in `index.html` and inline styles used by the app.
- `'unsafe-eval'` preserves the manual-only `office_execute_js` escape hatch and some development/runtime compatibility. Removing it requires replacing or disabling raw snippet execution.
