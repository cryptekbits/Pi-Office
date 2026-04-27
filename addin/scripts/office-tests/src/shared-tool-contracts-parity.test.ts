import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import test from "node:test";

import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import {
  getOfficeToolDefinitionsForHost,
  OFFICE_TOOL_DEFINITIONS,
  officeToolSupportsHost,
  type OfficeToolDefinition,
} from "../../../apps/taskpane/src/lib/office/tools/index.js";
import { createOfficeExtension } from "../../../packages/pi-office-pack/src/extension.js";
import {
  OFFICE_HOSTS,
  OFFICE_TOOL_NAMES,
  TOOL_CATEGORY_MAP,
  type AskUserRequest,
  type AskUserResponse,
  type OfficeHost,
  type OfficeToolRequest,
} from "../../../packages/pi-office-pack/src/protocol.js";

class MemoryStorage {
  private readonly store = new Map<string, string>();

  clear(): void {
    this.store.clear();
  }

  getItem(key: string): string | null {
    return this.store.has(key) ? this.store.get(key)! : null;
  }

  key(index: number): string | null {
    const keys = Array.from(this.store.keys());
    return keys[index] ?? null;
  }

  removeItem(key: string): void {
    this.store.delete(key);
  }

  setItem(key: string, value: string): void {
    this.store.set(key, String(value));
  }

  get length(): number {
    return this.store.size;
  }
}

const RUNTIME_ONLY_AGENT_TOOLS = ["ask_user", "generate_image"] as const;
const FINAL_AGENT_TOOL_INVENTORY = [...OFFICE_TOOL_NAMES, ...RUNTIME_ONLY_AGENT_TOOLS] as const;
const DEFAULT_DEFERRED_WORD_BASELINE = [
  "ask_user",
  "generate_image",
  "office_batch_execute",
  "office_apply_edit",
  "office_capture_snapshot",
  "office_get_context",
  "office_navigate",
  "office_read_section",
  "office_tool_get",
  "office_tool_search",
  "verify_doc",
  "verify_doc_visual",
  "word_search",
] as const;
const REQUIRED_VALIDATION_COMMANDS = ["typecheck", "build", "check:bundle", "validate:manifests", "test:office"] as const;
const BRIDGE_DISPATCH_SOURCE_FILES = [
  "apps/taskpane/src/lib/office/bridge/common-executor.ts",
  "apps/taskpane/src/lib/office/bridge/word.ts",
  "apps/taskpane/src/lib/office/bridge/excel.ts",
  "apps/taskpane/src/lib/office/bridge/powerpoint.ts",
] as const;
const VALID_EXECUTOR_KINDS = new Set<OfficeToolDefinition["executor"]>([
  "office-bridge",
  "reviewable-word-edit",
  "companion-native-capture",
  "runtime-registry",
]);
const RUNTIME_ONLY_OFFICE_TOOLS = new Set<string>([
  "office_batch_execute",
  "mcp_batch_execute",
  "mcp_result_get",
  "mcp_result_summarize",
  "mcp_result_clear",
]);
const EXTENSION_REGISTERED_TOOL_INVENTORY = FINAL_AGENT_TOOL_INVENTORY.filter(
  (toolName) =>
    toolName !== "office_tool_search" &&
    toolName !== "office_tool_get" &&
    toolName !== "mcp_tool_search" &&
    toolName !== "word_search" &&
    toolName !== "word_format_text" &&
    toolName !== "word_list_format" &&
    toolName !== "word_reference_inventory" &&
    toolName !== "word_hyperlink" &&
    toolName !== "word_table" &&
    toolName !== "word_section_layout" &&
    !RUNTIME_ONLY_OFFICE_TOOLS.has(toolName),
);

function sorted(values: Iterable<string>): string[] {
  return Array.from(new Set(values)).sort((a, b) => a.localeCompare(b));
}

function readProjectFile(path: string): string {
  return readFileSync(join(process.cwd(), path), "utf8");
}

function getDefaultTaskpaneAgentToolInventory(host: OfficeHost): readonly string[] {
  return host === "word" ? DEFAULT_DEFERRED_WORD_BASELINE : [
    ...getOfficeToolDefinitionsForHost(host)
      .filter((definition) => definition.executor !== "companion-native-capture")
      .filter((definition) => {
        if (definition.core === true) return true;
        if (definition.deferred === true) return false;
        const visibility = definition.discovery?.visibility ?? definition.discovery?.tier;
        return visibility !== "deferred" && visibility !== "specialized";
      })
      .map((definition) => definition.name),
    ...RUNTIME_ONLY_AGENT_TOOLS,
  ];
}

function getBridgeDispatchedToolNames(): string[] {
  return BRIDGE_DISPATCH_SOURCE_FILES.flatMap((path) => {
    const source = readProjectFile(path);
    return Array.from(source.matchAll(/if\s*\(\s*request\.toolName === "([^"]+)"/g), (match) => match[1]!);
  });
}

function installRuntimePolyfills(): void {
  const globalAny = globalThis as unknown as {
    localStorage?: Storage;
    window?: { location?: { origin?: string }; open?: (...args: unknown[]) => unknown };
    WebSocket?: { CONNECTING: number; OPEN: number; CLOSING: number; CLOSED: number };
    CloseEvent?: typeof CloseEvent;
    MessageEvent?: typeof MessageEvent;
    atob?: (value: string) => string;
    btoa?: (value: string) => string;
  };

  globalAny.localStorage = new MemoryStorage() as unknown as Storage;
  globalAny.window = globalAny.window ?? {};
  globalAny.window.location = globalAny.window.location ?? { origin: "https://localhost:3443" };
  globalAny.window.location.origin = globalAny.window.location.origin ?? "https://localhost:3443";
  globalAny.window.open = globalAny.window.open ?? (() => null);

  if (!globalAny.WebSocket) {
    globalAny.WebSocket = {
      CONNECTING: 0,
      OPEN: 1,
      CLOSING: 2,
      CLOSED: 3,
    };
  }

  if (!globalAny.CloseEvent) {
    class CloseEventPolyfill extends Event {
      readonly code: number;
      readonly reason: string;
      readonly wasClean: boolean;

      constructor(type: string, init?: { code?: number; reason?: string; wasClean?: boolean }) {
        super(type);
        this.code = init?.code ?? 0;
        this.reason = init?.reason ?? "";
        this.wasClean = init?.wasClean ?? true;
      }
    }
    globalAny.CloseEvent = CloseEventPolyfill as unknown as typeof CloseEvent;
  }

  if (!globalAny.MessageEvent) {
    class MessageEventPolyfill<T = unknown> extends Event {
      readonly data: T;

      constructor(type: string, init?: { data?: T }) {
        super(type);
        this.data = init?.data as T;
      }
    }
    globalAny.MessageEvent = MessageEventPolyfill as unknown as typeof MessageEvent;
  }

  if (!globalAny.atob) {
    globalAny.atob = (value: string) => Buffer.from(value, "base64").toString("binary");
  }
  if (!globalAny.btoa) {
    globalAny.btoa = (value: string) => Buffer.from(value, "binary").toString("base64");
  }
}

async function loadKernelModule() {
  installRuntimePolyfills();
  const specifier = `../../../apps/taskpane/src/lib/runtime/inprocess-kernel.js?test=${Date.now()}-${Math.random()}`;
  return import(specifier);
}

function waitForEvent(target: EventTarget, name: string, timeoutMs = 2_000): Promise<Event> {
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      target.removeEventListener(name, onEvent);
      reject(new Error(`Timed out waiting for ${name} event.`));
    }, timeoutMs);

    const onEvent = (event: Event) => {
      clearTimeout(timeout);
      target.removeEventListener(name, onEvent);
      resolve(event);
    };

    target.addEventListener(name, onEvent);
  });
}

async function waitForServerMessage(
  socket: {
    addEventListener: (name: string, listener: (event: Event) => void) => void;
    removeEventListener: (name: string, listener: (event: Event) => void) => void;
  },
  predicate: (payload: Record<string, unknown>) => boolean,
  timeoutMs = 2_000,
): Promise<Record<string, unknown>> {
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      socket.removeEventListener("message", onMessage);
      reject(new Error("Timed out waiting for server message."));
    }, timeoutMs);

    const onMessage = (event: Event) => {
      try {
        const payload = JSON.parse(String((event as MessageEvent).data)) as Record<string, unknown>;
        if (!predicate(payload)) return;
        clearTimeout(timeout);
        socket.removeEventListener("message", onMessage);
        resolve(payload);
      } catch {
        // Ignore parse errors from unrelated events.
      }
    };

    socket.addEventListener("message", onMessage);
  });
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

test("final inventory stays synchronized across protocol exports and tool-category mapping", () => {
  const inventory = sorted(FINAL_AGENT_TOOL_INVENTORY);
  assert.equal(inventory.length, FINAL_AGENT_TOOL_INVENTORY.length);

  for (const toolName of FINAL_AGENT_TOOL_INVENTORY) {
    assert.ok(
      Object.prototype.hasOwnProperty.call(TOOL_CATEGORY_MAP, toolName),
      `${toolName} is missing from TOOL_CATEGORY_MAP.`,
    );
  }

  for (const officeTool of OFFICE_TOOL_NAMES) {
    const category = TOOL_CATEGORY_MAP[officeTool];
    assert.ok(
      category === "read" || category === "write-doc" || category === "escape-hatch" || category === "connector",
      `${officeTool} should map to read/write-doc/escape-hatch/connector but mapped to ${String(category)}.`,
    );
  }

  assert.equal(TOOL_CATEGORY_MAP.ask_user, "interaction");
  assert.equal(TOOL_CATEGORY_MAP.generate_image, "write-doc");
});

test("office tool registry has one definition, category, host rule, and executable owner per protocol tool", () => {
  const protocolToolNames = sorted(OFFICE_TOOL_NAMES);
  const registryToolNames = OFFICE_TOOL_DEFINITIONS.map((definition) => definition.name);
  assert.deepEqual(sorted(registryToolNames), protocolToolNames);
  assert.equal(registryToolNames.length, OFFICE_TOOL_DEFINITIONS.length, "Office tool registry has duplicate names.");

  const validHosts = new Set<string>(OFFICE_HOSTS);
  const bridgeDispatchNames = getBridgeDispatchedToolNames();
  assert.equal(bridgeDispatchNames.length, sorted(bridgeDispatchNames).length, "Bridge dispatch has duplicate tool handlers.");
  for (const runtimeOnlyTool of RUNTIME_ONLY_OFFICE_TOOLS) {
    assert.equal(bridgeDispatchNames.includes(runtimeOnlyTool), false);
  }

  const kernelSource = readProjectFile("apps/taskpane/src/lib/runtime/inprocess-kernel.ts");
  assert.match(kernelSource, /getCoreOfficeToolDefinitionsForHost\(this\.officeState\.host\)/);
  assert.match(kernelSource, /searchOfficeToolDefinitions/);
  assert.match(kernelSource, /definition\.executor === "reviewable-word-edit"/);
  assert.match(kernelSource, /executeCompanionNativeCapture/);

  for (const toolName of OFFICE_TOOL_NAMES) {
    const matchingDefinitions = OFFICE_TOOL_DEFINITIONS.filter((definition) => definition.name === toolName);
    assert.equal(matchingDefinitions.length, 1, `${toolName} should have exactly one registry definition.`);

    const definition = matchingDefinitions[0]!;
    assert.equal(definition.category, TOOL_CATEGORY_MAP[toolName], `${toolName} category drifted from TOOL_CATEGORY_MAP.`);
    assert.ok(definition.label.trim(), `${toolName} is missing a registry label.`);
    assert.ok(definition.description.trim(), `${toolName} is missing a registry description.`);
    assert.ok((definition.compactSummary ?? definition.discovery?.summary ?? definition.description).trim(), `${toolName} is missing deferred-discovery summary.`);
    assert.ok(
      (definition.keywords?.length ?? definition.searchKeywords?.length ?? definition.capabilityTags?.length ?? definition.discovery?.keywords?.length ?? 0) > 0 ||
        definition.description.trim().split(/\s+/).length >= 3,
      `${toolName} is missing deferred-discovery keywords.`,
    );
    assert.ok(definition.parameters, `${toolName} is missing a JSON schema.`);
    assert.ok(VALID_EXECUTOR_KINDS.has(definition.executor), `${toolName} has an unknown executor kind.`);

    if (definition.hosts !== "all") {
      assert.ok(definition.hosts.length > 0, `${toolName} must support at least one host.`);
      for (const host of definition.hosts) {
        assert.ok(validHosts.has(host), `${toolName} has unsupported host rule ${host}.`);
      }
    }

    const supportedHosts = OFFICE_HOSTS.filter((host) => officeToolSupportsHost(definition, host));
    assert.ok(supportedHosts.length > 0, `${toolName} is not supported by any host.`);
  }

  for (const host of OFFICE_HOSTS) {
    assert.deepEqual(
      sorted(getOfficeToolDefinitionsForHost(host).map((definition) => definition.name)),
      sorted(OFFICE_TOOL_DEFINITIONS.filter((definition) => officeToolSupportsHost(definition, host)).map((definition) => definition.name)),
      `${host} host registry helper drifted from host support rules.`,
    );
  }
});

test("extension registration and runtime-published tools stay synchronized with the final inventory", async () => {
  const registeredByExtension = new Set<string>();
  const extensionFactory = createOfficeExtension({
    getHost: () => "word",
    getState: () => undefined,
    invokeTool: async () => ({ ok: true }),
    invokeAskUser: async (request: AskUserRequest): Promise<AskUserResponse> => ({
      requestId: request.requestId,
      answers: [],
    }),
    generateImage: async () => ({
      base64: "ZmFrZQ==",
      mimeType: "image/png",
      width: 1,
      height: 1,
      modelKey: "test-model",
      modelName: "Test Model",
    }),
  });

  extensionFactory({
    registerTool: (tool: { name: string }) => {
      registeredByExtension.add(tool.name);
    },
    on: () => {},
  } as any);

  const runtime = await loadKernelModule();
  const openResponse = await runtime.dispatchKernelRequest("/v1/sessions/open", {
    method: "POST",
    body: JSON.stringify({
      host: "word",
      documentId: `doc-contract-parity-${Date.now()}`,
      saved: true,
      title: "Shared Tool Contracts Parity",
    }),
  }) as {
    sessionId: string;
  };

  const socket = runtime.createLocalBridgeSocket(openResponse.sessionId);
  await waitForEvent(socket, "open");
  const readyMessage = waitForServerMessage(
    socket,
    (payload) => payload.type === "connection_state" && payload.state === "ready",
  );
  socket.send(JSON.stringify({ type: "client_ready" }));
  await readyMessage;

  const session = (socket as unknown as { session: { agent: { state: { tools: Array<{ name: string }> } } } }).session;
  const runtimeTools = session.agent.state.tools.map((tool) => tool.name);
  socket.close();

  assert.deepEqual(sorted(registeredByExtension), sorted(EXTENSION_REGISTERED_TOOL_INVENTORY));
  assert.deepEqual(sorted(runtimeTools), sorted(getDefaultTaskpaneAgentToolInventory("word")));
});

test("taskpane bridge dispatch cases stay in sync with supported Office tools", async () => {
  const dispatchedToolNames = getBridgeDispatchedToolNames();
  assert.deepEqual(sorted(dispatchedToolNames), sorted(OFFICE_TOOL_NAMES.filter(
    (toolName) => !RUNTIME_ONLY_OFFICE_TOOLS.has(String(toolName)),
  )));

  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async () => ({ ok: true }),
    applyHostAction: async () => ({ ok: true }),
    navigateOfficeAnchor: async () => ({ ok: true }),
    readDocumentSection: async () => ({ ok: true }),
    executeOfficeJs: async () => ({ ok: true }),
    proposeEdits: async () => ({ ok: true }),
  });
  const unsupported = await executeOfficeTool({
    requestId: "unsupported-tool",
    toolName: "not_a_supported_tool" as OfficeToolRequest["toolName"],
    host: "word",
    params: {},
  } as OfficeToolRequest);
  assert.equal(unsupported.success, false);
  assert.match(String(unsupported.error), /not supported by the taskpane bridge/);
});

test("automated validation commands remain exposed in package scripts and CI gate", () => {
  const packageJson = JSON.parse(readFileSync(join(process.cwd(), "package.json"), "utf8")) as {
    scripts?: Record<string, string>;
  };
  const scripts = packageJson.scripts ?? {};

  for (const commandName of REQUIRED_VALIDATION_COMMANDS) {
    assert.equal(typeof scripts[commandName], "string", `package.json is missing script ${commandName}.`);
  }

  const ciWorkflow = readFileSync(join(process.cwd(), "..", ".github", "workflows", "ci.yml"), "utf8");
  for (const commandName of REQUIRED_VALIDATION_COMMANDS) {
    assert.match(
      ciWorkflow,
      new RegExp(`npm run\\s+${escapeRegExp(commandName)}`),
      `CI workflow is missing npm run ${commandName}.`,
    );
  }
});
