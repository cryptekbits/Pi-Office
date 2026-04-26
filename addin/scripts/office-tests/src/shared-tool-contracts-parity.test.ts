import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import test from "node:test";

import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import { createOfficeExtension } from "../../../packages/pi-office-pack/src/extension.js";
import {
  OFFICE_TOOL_NAMES,
  TOOL_CATEGORY_MAP,
  type AskUserRequest,
  type AskUserResponse,
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
const DEFAULT_TASKPANE_AGENT_TOOL_INVENTORY = FINAL_AGENT_TOOL_INVENTORY.filter(
  (toolName) => toolName !== "office_capture_viewport",
);
const REQUIRED_VALIDATION_COMMANDS = ["typecheck", "build", "check:bundle", "validate:manifests", "test:office"] as const;

function sorted(values: Iterable<string>): string[] {
  return Array.from(new Set(values)).sort((a, b) => a.localeCompare(b));
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
      category === "read" || category === "write-doc" || category === "escape-hatch",
      `${officeTool} should map to read/write-doc/escape-hatch but mapped to ${String(category)}.`,
    );
  }

  assert.equal(TOOL_CATEGORY_MAP.ask_user, "interaction");
  assert.equal(TOOL_CATEGORY_MAP.generate_image, "write-doc");
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

  assert.deepEqual(sorted(registeredByExtension), sorted(FINAL_AGENT_TOOL_INVENTORY));
  assert.deepEqual(sorted(runtimeTools), sorted(DEFAULT_TASKPANE_AGENT_TOOL_INVENTORY));
});

test("taskpane bridge dispatch cases stay in sync with supported Office tools", async () => {
  const bridgeSource = readFileSync(join(process.cwd(), "apps/taskpane/src/lib/office-bridge.ts"), "utf8");
  const dispatchedToolNames = Array.from(
    bridgeSource.matchAll(/request\.toolName === "([^"]+)"/g),
    (match) => match[1]!,
  );
  assert.deepEqual(sorted(dispatchedToolNames), sorted(OFFICE_TOOL_NAMES));

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
