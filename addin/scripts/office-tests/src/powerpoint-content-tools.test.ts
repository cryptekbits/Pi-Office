import assert from "node:assert/strict";
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

test("protocol inventory includes first-class PowerPoint content tools", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("insert_slide_element" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("remove_slide_element" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("edit_slide_text" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("edit_slide_xml" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("edit_slide_master" as any));
  assert.equal(TOOL_CATEGORY_MAP.insert_slide_element, "write-doc");
  assert.equal(TOOL_CATEGORY_MAP.remove_slide_element, "write-doc");
  assert.equal(TOOL_CATEGORY_MAP.edit_slide_text, "write-doc");
  assert.equal(TOOL_CATEGORY_MAP.edit_slide_xml, "write-doc");
  assert.equal(TOOL_CATEGORY_MAP.edit_slide_master, "write-doc");
});

test("createOfficeExtension registers first-class PowerPoint content tools and routes calls", async () => {
  const registeredTools = new Map<string, {
    execute: (toolCallId: string, params: Record<string, unknown>) => Promise<{ content: unknown; details: unknown }>;
  }>();
  const invocations: Array<{ toolName: string; params: Record<string, unknown> }> = [];

  const extensionFactory = createOfficeExtension({
    getHost: () => "powerpoint",
    getState: () => undefined,
    invokeTool: async (toolName, params) => {
      invocations.push({ toolName, params });
      return {
        ok: true,
        tool: toolName,
        params,
      };
    },
    invokeAskUser: async (request: AskUserRequest): Promise<AskUserResponse> => ({
      requestId: request.requestId,
      answers: [],
    }),
  });

  extensionFactory({
    registerTool: (tool: { name: string; execute: (toolCallId: string, params: Record<string, unknown>) => Promise<{ content: unknown; details: unknown }> }) => {
      registeredTools.set(tool.name, tool);
    },
    on: () => {},
  } as any);

  const powerPointToolNames = [
    "insert_slide_element",
    "remove_slide_element",
    "edit_slide_text",
    "edit_slide_xml",
    "edit_slide_master",
  ];
  for (const toolName of powerPointToolNames) {
    assert.ok(registeredTools.has(toolName));
  }

  await registeredTools.get("insert_slide_element")!.execute("ppt-insert-element", {
    operation: "add_text_box",
    slideId: "slide-2",
    text: "Agenda",
  });
  await registeredTools.get("remove_slide_element")!.execute("ppt-remove-element", {
    operation: "remove_shape",
    slideId: "slide-2",
    shapeId: "shape-1",
    confirmDestructive: true,
  });
  await registeredTools.get("edit_slide_text")!.execute("ppt-edit-text", {
    operation: "set_shape_text",
    slideId: "slide-2",
    shapeId: "shape-2",
    text: "Updated headline",
  });
  await registeredTools.get("edit_slide_xml")!.execute("ppt-edit-xml", {
    operation: "replace_slide_notes",
    slideId: "slide-2",
    content: "Updated speaker notes",
  });
  await registeredTools.get("edit_slide_master")!.execute("ppt-edit-master", {
    operation: "apply_layout",
    slideId: "slide-2",
    layoutName: "Title and Content",
  });

  assert.equal(invocations[0]?.toolName, "insert_slide_element");
  assert.equal(invocations[1]?.toolName, "remove_slide_element");
  assert.equal(invocations[2]?.toolName, "edit_slide_text");
  assert.equal(invocations[3]?.toolName, "edit_slide_xml");
  assert.equal(invocations[4]?.toolName, "edit_slide_master");
});

test("in-process runtime publishes first-class PowerPoint content tools", async () => {
  const runtime = await loadKernelModule();
  const openResponse = await runtime.dispatchKernelRequest("/v1/sessions/open", {
    method: "POST",
    body: JSON.stringify({
      host: "powerpoint",
      documentId: `doc-ppt-content-${Date.now()}`,
      saved: true,
      title: "PowerPoint Content Tool Inventory Test",
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

  assert.ok(toolNames.includes("insert_slide_element"));
  assert.ok(toolNames.includes("remove_slide_element"));
  assert.ok(toolNames.includes("edit_slide_text"));
  assert.ok(toolNames.includes("edit_slide_xml"));
  assert.ok(toolNames.includes("edit_slide_master"));
  socket.close();
});

test("createOfficeToolExecutor dispatches first-class PowerPoint content tools and enforces PowerPoint-only scope", async () => {
  const calls: Array<{ host: string; action: unknown }> = [];
  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async () => ({ ok: true }),
    applyHostAction: async (host, action) => {
      calls.push({ host, action });
      return { ok: true, host, action };
    },
    navigateOfficeAnchor: async () => ({ ok: true }),
    readDocumentSection: async () => ({ ok: true }),
    executeOfficeJs: async () => ({ ok: true }),
    proposeEdits: async () => ({ ok: true }),
  });

  const insertResult = await executeOfficeTool({
    requestId: "ppt-insert-1",
    toolName: "insert_slide_element" as OfficeToolRequest["toolName"],
    host: "powerpoint",
    params: {
      operation: "add_text_box",
      slideId: "slide-1",
      text: "Executive summary",
      left: 72,
      top: 120,
    },
  } as OfficeToolRequest);
  const removeResult = await executeOfficeTool({
    requestId: "ppt-remove-1",
    toolName: "remove_slide_element" as OfficeToolRequest["toolName"],
    host: "powerpoint",
    params: {
      operation: "remove_shape",
      slideId: "slide-1",
      shapeId: "shape-4",
      confirmDestructive: true,
    },
  } as OfficeToolRequest);
  const textResult = await executeOfficeTool({
    requestId: "ppt-text-1",
    toolName: "edit_slide_text" as OfficeToolRequest["toolName"],
    host: "powerpoint",
    params: {
      operation: "set_shape_text",
      slideId: "slide-1",
      shapeId: "shape-8",
      text: "Updated title",
    },
  } as OfficeToolRequest);
  const xmlResult = await executeOfficeTool({
    requestId: "ppt-xml-1",
    toolName: "edit_slide_xml" as OfficeToolRequest["toolName"],
    host: "powerpoint",
    params: {
      operation: "replace_slide_notes",
      slideId: "slide-1",
      content: "Updated notes body",
    },
  } as OfficeToolRequest);
  const masterResult = await executeOfficeTool({
    requestId: "ppt-master-1",
    toolName: "edit_slide_master" as OfficeToolRequest["toolName"],
    host: "powerpoint",
    params: {
      operation: "apply_layout",
      slideId: "slide-1",
      layoutName: "Title and Content",
    },
  } as OfficeToolRequest);
  const unsupportedResult = await executeOfficeTool({
    requestId: "ppt-insert-unsupported",
    toolName: "insert_slide_element" as OfficeToolRequest["toolName"],
    host: "word",
    params: {},
  } as OfficeToolRequest);

  assert.equal(insertResult.success, true);
  assert.equal(removeResult.success, true);
  assert.equal(textResult.success, true);
  assert.equal(xmlResult.success, true);
  assert.equal(masterResult.success, true);
  assert.equal((calls[0]?.action as { type: string }).type, "addTextBox");
  assert.equal((calls[1]?.action as { type: string }).type, "deleteShape");
  assert.equal((calls[2]?.action as { type: string }).type, "setShapeText");
  assert.equal((calls[3]?.action as { type: string }).type, "replaceSlideNotes");
  assert.equal((calls[4]?.action as { type: string }).type, "applyLayout");
  assert.equal(unsupportedResult.success, false);
  assert.match(String(unsupportedResult.error), /only available for PowerPoint/);
});
