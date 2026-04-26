# Privacy And Storage

Pi-Office keeps Office.js execution inside the active Office taskpane. Network calls happen only when a configured AI provider, connector, or optional companion feature is used.

## What Leaves The Machine

- AI provider requests can include the user's prompt, selected Office context, generated-image prompts, model/tool results, and any document snippets the taskpane attaches for the current request.
- Remote HTTP connectors can send connector-specific requests to the configured service endpoint. Remote connector execution is setup-only until browser remote-MCP execution is implemented.
- Local stdio connectors run through the optional companion when connected. The companion is designed for read-only local MCP execution.
- The optional companion runs on loopback and receives saved-document context only when a saved document is bound to the session.

## What Stays In Browser Storage

- Provider credentials are stored under the taskpane origin in an AES-GCM envelope in `localStorage`.
- Connector configuration, connector secrets, OAuth credential handoffs, scope overrides, and redacted connector audit logs are stored under the taskpane origin in an AES-GCM envelope in `localStorage`.
- Saved chat history is stored in `localStorage` and may include prompts, document snippets, model responses, host metadata, and saved-document identifiers.
- User preferences, enabled providers/models, recent models, and persistent permission choices are stored locally.

The AES-GCM envelopes are local obfuscation only. The crypto keys are also stored in `localStorage`, so this does not protect secrets from same-origin script access, a compromised taskpane bundle, or XSS. Stronger storage such as OS keychain-backed companion storage is tracked as future provider/auth work.

## Clear Data Controls

The taskpane Settings -> Privacy tab can clear:

- All provider credentials and the provider credential encryption key.
- All connector configuration, connector secrets, OAuth state, connector scopes, connector logs, and the connector encryption key.
- Saved local chat history.

Per-provider auth removal remains available from Settings -> AI Providers, and per-connector removal remains available from Settings -> Integrations.

## Telemetry Default

Pi-Office does not enable product analytics or telemetry by default. Provider services, connector services, and locally installed connector runtimes may have their own logging and retention policies outside Pi-Office.

## CSP Policy

The taskpane declares a `Content-Security-Policy` meta tag in `addin/apps/taskpane/index.html` because the dev/sideloaded Office taskpane is a static Vite app and may not control deployment headers in every host. Public deployments should emit the same or stricter policy as HTTP headers.

The current policy restricts the default source to the taskpane origin, allows Microsoft-hosted Office.js, allows provider/connector HTTPS calls and loopback companion calls, blocks plugins with `object-src 'none'`, and pins form submissions to `none`.

Two allowances are intentional and should be revisited before public release hardening:

- `'unsafe-inline'` is currently needed by inline Office bootstrap scripts in `index.html` and inline styles used by the app.
- `'unsafe-eval'` preserves the manual-only `office_execute_js` escape hatch and some development/runtime compatibility. Removing it requires replacing or disabling raw snippet execution.
