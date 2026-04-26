# Pi-Office Capability Boundaries

Pi-Office uses Smart Auto routing. The taskpane remains fully usable by itself. When a healthy companion is connected, eligible non-Office capabilities prefer companion execution and keep taskpane fallback where the browser runtime can honestly execute the work.

## Runtime Fields

Every resolved capability is described with these fields:

| Field | Meaning |
| --- | --- |
| `available` | Whether the capability can be used in the current host/document/session. |
| `preferredRuntime` | `addin` or `companion`, based on the ownership model and advertised companion capability. |
| `activeRuntime` | The runtime Pi-Office will use right now. |
| `fallbackRuntime` | Runtime available if the preferred runtime disconnects or is not configured. |
| `requiresCompanion` | Whether the capability cannot exist in taskpane-only mode. |
| `requiresSavedDocument` | Whether the document must be saved before the capability can bind to the document folder. |
| `honestyLabel` | Short user/model-facing boundary text. |

The shared registry lives in `addin/packages/pi-office-pack/src/capabilities.ts` and is consumed by the taskpane runtime and Settings UI. The companion reports its machine-side capability facts through `CompanionCapabilities`.

## Ownership Classes

| Class | Capabilities | Runtime rule |
| --- | --- | --- |
| `addin-only` | Office.js reads/writes, selection, navigation, review cards, Office permissions, PowerPoint native slide/shape visual verification | Always execute in the Office-hosted taskpane. |
| `companion-only` | Saved-folder file reads, MCP/web connectors, sandbox shell, true viewport/window screenshots, memory | Hidden/unavailable unless the companion advertises support. |
| `either` | Inference, browser API-key provider calls, image generation, diagnostics, background jobs | Prefer companion only when companion provider auth/agent support is available; otherwise use taskpane fallback. |

Office document mutation is permanently taskpane-owned. A companion-owned agent may request Office tool execution, but the taskpane remains the authoritative Office.js executor and permission surface.

## Provider Auth And Inference

Browser API-key providers remain supported in taskpane-only mode. When the companion eventually has provider auth configured, Smart Auto can prefer companion inference for eligible turns. Pi-Office must not silently move provider secrets from browser storage to the companion; any migration must be an explicit user action.

Current companion capability metadata exposes provider-auth and agent-session contracts but reports them unavailable until real companion auth/session storage exists.

## Visual Capture

| Tool | Runtime | Boundary |
| --- | --- | --- |
| `office_capture_snapshot` | Add-in | Office.js context snapshots and metadata. Not a true OS/window screenshot. |
| `verify_doc_visual` | Add-in | Word viewport/context metadata and selection visuals. Not a true OS/window screenshot. |
| `read_range_image` | Add-in | Excel active-selection Office.js image snapshot. Does not render arbitrary offscreen ranges. |
| `verify_slide_visual` | Add-in | PowerPoint native slide/shape snapshot paths where Office APIs support them. |
| `office_capture_viewport` | Companion only | True viewport/window screenshot. Published to the model only when companion native capture is available for the active host. |

For Word and Excel, true visible-window capture requires the companion native capture helper. For PowerPoint, add-in slide/shape visual verification remains useful without the companion; companion screenshot is only needed for full window/frame capture.

## Disconnect Behavior

When the companion is absent, stopped, unhealthy, or disconnected, new turns use taskpane-supported tools only. Companion-only tools are hidden from the active model tool inventory. In-flight companion-owned calls should fail clearly and can be retried after discovery reconnects or retried through taskpane fallback when the capability supports it.

## MCP Connectors (stdio and remote HTTP)

Connector **configuration** (including user-pasted secrets and static HTTP headers) is saved in the taskpane's encrypted `localStorage` envelope on the device. **Verification and execution** of MCP transports (`local_stdio`, `remote_http`) run through the optional companion: the taskpane does not open stdio child processes or call remote MCP URLs directly. Users can save configuration first and run **Check connection** when the companion is online; the Integrations wizard gates verification when the companion is disconnected for those transports. See `docs/privacy-and-storage.md` for storage boundaries.

Connector catalog entries can now expose multiple setup profiles for the same service. The noob-friendly default chooses hosted HTTP when the companion is absent and chooses the local companion path only when a local command/stdio profile is actually available and the companion is connected. STDIO-only connectors are visibly disabled in the Library while the companion is disconnected because there is no browser-safe way to spawn the required local process.

After verification, the companion returns the full MCP tool inventory plus a smaller execution target list. Tool annotations are honored first, then Pi-Office's curated connector read policy, then conservative name-based fallback classification. Read and sensitive-read tools are enabled by default; write, destructive, and unknown tools are disabled by default. Users can enable an advanced tool from the connector details view after acknowledging the warning, but the tool still remains subject to the normal Pi-Office permission flow and disabled tools fail closed if called directly.
