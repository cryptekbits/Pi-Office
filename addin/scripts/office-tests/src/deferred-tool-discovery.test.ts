import assert from "node:assert/strict";
import test from "node:test";

import { getOfficeToolDefinition, searchOfficeToolDefinitions } from "../../../apps/taskpane/src/lib/office/tools/index.js";
import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import {
  OFFICE_TOOL_NAMES,
  TOOL_CATEGORY_MAP,
  type AskUserRequest,
  type AskUserResponse,
  type OfficeToolRequest,
} from "../../../packages/pi-office-pack/src/protocol.js";
import { createOfficeExtension } from "../../../packages/pi-office-pack/src/extension.js";
import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";

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

test("Office tool registry search ranks host-gated deferred tools with permission and requirement metadata", () => {
  const wordListResults = searchOfficeToolDefinitions({
    host: "word",
    query: "list numbering bullets legal clauses",
    limit: 5,
  });

  const listResult = wordListResults.find((result) => result.toolName === "edit_doc_list");
  assert.ok(listResult, "edit_doc_list should be discoverable for list/numbering queries.");
  assert.equal(listResult.category, "write-doc");
  assert.ok(listResult.keywords.some((keyword) => /number|list|bullet|legal/.test(keyword)));
  assert.equal(listResult.schemaAvailable, true);

  const powerpointResults = searchOfficeToolDefinitions({
    host: "powerpoint",
    query: "slide chart",
    category: "write-doc",
  });
  assert.ok(powerpointResults.some((result) => result.toolName === "edit_slide_chart"));
  assert.ok(!powerpointResults.some((result) => result.toolName === "edit_doc_list"));

  const executeJs = getOfficeToolDefinition("office_execute_js");
  assert.equal(executeJs.category, "escape-hatch");
  assert.match(executeJs.description, /escape hatch/i);
});

test("deferred discovery protocol tools are read-only registry entries", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("office_tool_search"));
  assert.ok(OFFICE_TOOL_NAMES.includes("office_tool_get"));
  assert.ok(OFFICE_TOOL_NAMES.includes("office_tool_call"));
  assert.ok(OFFICE_TOOL_NAMES.includes("mcp_tool_search"));
  assert.equal(TOOL_CATEGORY_MAP.office_tool_search, "read");
  assert.equal(TOOL_CATEGORY_MAP.office_tool_get, "read");
  assert.equal(TOOL_CATEGORY_MAP.office_tool_call, "write-doc");
  assert.equal(TOOL_CATEGORY_MAP.mcp_tool_search, "read");
  assert.match(getOfficeToolDefinition("office_tool_search").description, /registry/i);
  assert.match(getOfficeToolDefinition("office_tool_get").description, /schema|definition/i);
  assert.match(getOfficeToolDefinition("office_tool_call").description, /discovering/i);
  assert.match(getOfficeToolDefinition("mcp_tool_search").description, /connector/i);
});

test("prompt and registry steer Word document generation to structured HTML insertion", () => {
  const applyEditDefinition = getOfficeToolDefinition("office_apply_edit");
  const actionSchema = (applyEditDefinition.parameters as { properties?: { action?: { type?: string } } }).properties?.action;

  assert.match(applyEditDefinition.description, /insertHtml/i);
  assert.match(applyEditDefinition.description, /not a JSON string/i);
  assert.equal(actionSchema?.type, "object");
  assert.match(applyEditDefinition.description, /content/i);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /new Word document generation/i);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /replaceDocumentHtml/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /"kind": "document"/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /JSON-encoded string/i);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /values\[\]/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /word_section_layout/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /word_equation/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /page-break-before/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /parallel tool calls/i);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /Do not claim a page break landed/i);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /do not mix Markdown markers/i);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /fast draft/i);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /before claiming the document is formatted/);
});

test("in-process runtime normalizes Gemini-style JSON string tool arguments", async () => {
  const runtime = await loadKernelModule();
  assert.deepEqual(
    runtime.normalizeToolParams("{\"operation\":\"insertBreak\",\"breakType\":\"page\",\"placement\":\"before\"}"),
    {
      operation: "insertBreak",
      breakType: "page",
      placement: "before",
    },
  );
  assert.deepEqual(runtime.normalizeToolParams("\"not-an-object\""), {});
  assert.deepEqual(runtime.normalizeToolParams("{not json"), {});
});

test("in-process runtime publishes compact core tools and defers long-tail Office schemas", async () => {
  const runtime = await loadKernelModule();
  const openResponse = await runtime.dispatchKernelRequest("/v1/sessions/open", {
    method: "POST",
    body: JSON.stringify({
      host: "word",
      documentId: `doc-deferred-discovery-${Date.now()}`,
      saved: true,
      title: "Deferred Discovery Tool Inventory Test",
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
  const toolNames = session.agent.state.tools.map((tool) => tool.name);
  socket.close();

  assert.ok(toolNames.includes("office_get_context"));
  assert.ok(toolNames.includes("office_tool_search"));
  assert.ok(toolNames.includes("office_tool_get"));
  assert.ok(toolNames.includes("office_tool_call"));
  assert.ok(toolNames.includes("office_apply_edit"));
  assert.ok(toolNames.includes("ask_user"));
  assert.ok(!toolNames.includes("edit_doc_list"), "Word list edit schema should be deferred behind office_tool_search.");
  assert.ok(!toolNames.includes("office_propose_edits"), "Review proposal schema should be deferred behind office_tool_search.");
  assert.ok(!toolNames.includes("office_execute_js"), "Escape hatch schema should be discoverable but not baseline-published.");
});

test("extension can disable deferred Office tools while registry still exposes final inventory", () => {
  const registered = new Set<string>();
  const extensionFactory = createOfficeExtension({
    getHost: () => "word",
    getState: () => undefined,
    isToolDisabled: (toolName) => toolName === "office_tool_search" || toolName === "office_tool_get",
    invokeTool: async () => ({ ok: true }),
    invokeAskUser: async (request: AskUserRequest): Promise<AskUserResponse> => ({
      requestId: request.requestId,
      answers: [],
    }),
  });

  extensionFactory({
    registerTool: (tool: { name: string }) => {
      registered.add(tool.name);
    },
    on: () => {},
  } as any);

  assert.ok(OFFICE_TOOL_NAMES.includes("office_tool_search"));
  assert.ok(!registered.has("office_tool_search"));
  assert.ok(!registered.has("office_tool_get"));
  assert.ok(registered.has("office_get_context"));
});
