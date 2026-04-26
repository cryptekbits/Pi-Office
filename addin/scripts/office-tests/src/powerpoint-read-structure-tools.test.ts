import assert from "node:assert/strict";
import test from "node:test";

import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import { resolvePowerPointShapeSlideMetadata } from "../../../apps/taskpane/src/lib/office/powerpoint-context.js";
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

test("PowerPoint selected shape metadata prefers the owning parent slide over the first selected slide", () => {
  const firstSelectedSlide = {
    id: "slide-1",
    index: 1,
    layoutName: "Title Slide",
  };
  const owningSlide = {
    id: "slide-3",
    index: 3,
    layoutName: "Chart Detail",
    slideMasterName: "Corporate",
  };
  const slideMetadataById = new Map([
    [firstSelectedSlide.id, firstSelectedSlide],
    [owningSlide.id, owningSlide],
  ]);

  assert.deepEqual(
    resolvePowerPointShapeSlideMetadata(
      { isNullObject: false, id: "slide-3", index: 2 },
      { id: "slide-1", index: 0 },
      slideMetadataById,
    ),
    owningSlide,
  );
  assert.deepEqual(
    resolvePowerPointShapeSlideMetadata(
      undefined,
      { id: "slide-1", index: 0 },
      slideMetadataById,
    ),
    firstSelectedSlide,
  );
});

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

test("protocol inventory includes first-class PowerPoint read/navigation and structure-mutation tools", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("get_presentation_structure" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("get_slide" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("list_slide_shapes" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("modify_presentation_structure" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("duplicate_slide" as any));
  assert.equal(TOOL_CATEGORY_MAP.get_presentation_structure, "read");
  assert.equal(TOOL_CATEGORY_MAP.get_slide, "read");
  assert.equal(TOOL_CATEGORY_MAP.list_slide_shapes, "read");
  assert.equal(TOOL_CATEGORY_MAP.modify_presentation_structure, "write-doc");
  assert.equal(TOOL_CATEGORY_MAP.duplicate_slide, "write-doc");
});

test("createOfficeExtension registers first-class PowerPoint structure tools and routes calls", async () => {
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
    "get_presentation_structure",
    "get_slide",
    "list_slide_shapes",
    "modify_presentation_structure",
    "duplicate_slide",
  ];
  for (const toolName of powerPointToolNames) {
    assert.ok(registeredTools.has(toolName));
  }

  await registeredTools.get("get_presentation_structure")!.execute("ppt-structure", { maxSlides: 10 });
  await registeredTools.get("get_slide")!.execute("ppt-slide", { slideIndex: 2 });
  await registeredTools.get("list_slide_shapes")!.execute("ppt-shapes", { slideId: "slide-2" });
  await registeredTools.get("modify_presentation_structure")!.execute("ppt-modify", {
    operation: "reorder_slides",
    slideIds: ["slide-2", "slide-1"],
  });
  await registeredTools.get("duplicate_slide")!.execute("ppt-duplicate", {
    slideId: "slide-1",
    targetSlideId: "slide-3",
  });

  assert.equal(invocations[0]?.toolName, "get_presentation_structure");
  assert.equal(invocations[1]?.toolName, "get_slide");
  assert.equal(invocations[2]?.toolName, "list_slide_shapes");
  assert.equal(invocations[3]?.toolName, "modify_presentation_structure");
  assert.equal(invocations[4]?.toolName, "duplicate_slide");
});

test("in-process runtime publishes first-class PowerPoint structure tools", async () => {
  const runtime = await loadKernelModule();
  const openResponse = await runtime.dispatchKernelRequest("/v1/sessions/open", {
    method: "POST",
    body: JSON.stringify({
      host: "powerpoint",
      documentId: `doc-ppt-structure-${Date.now()}`,
      saved: true,
      title: "PowerPoint Structure Tool Inventory Test",
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

  assert.ok(toolNames.includes("get_presentation_structure"));
  assert.ok(toolNames.includes("get_slide"));
  assert.ok(toolNames.includes("list_slide_shapes"));
  assert.ok(toolNames.includes("modify_presentation_structure"));
  assert.ok(toolNames.includes("duplicate_slide"));
  socket.close();
});

test("createOfficeToolExecutor dispatches first-class PowerPoint structure tools and enforces PowerPoint-only scope", async () => {
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

  const structureResult = await executeOfficeTool({
    requestId: "ppt-structure-1",
    toolName: "get_presentation_structure" as OfficeToolRequest["toolName"],
    host: "powerpoint",
    params: {
      maxSlides: 10,
    },
  } as OfficeToolRequest);
  const slideResult = await executeOfficeTool({
    requestId: "ppt-slide-1",
    toolName: "get_slide" as OfficeToolRequest["toolName"],
    host: "powerpoint",
    params: {
      slideId: "slide-2",
    },
  } as OfficeToolRequest);
  const shapesResult = await executeOfficeTool({
    requestId: "ppt-shapes-1",
    toolName: "list_slide_shapes" as OfficeToolRequest["toolName"],
    host: "powerpoint",
    params: {
      slideIndex: 3,
    },
  } as OfficeToolRequest);
  const modifyResult = await executeOfficeTool({
    requestId: "ppt-modify-1",
    toolName: "modify_presentation_structure" as OfficeToolRequest["toolName"],
    host: "powerpoint",
    params: {
      operation: "reorder_slides",
      slideIds: ["slide-2", "slide-1"],
      slideIndex: 1,
    },
  } as OfficeToolRequest);
  const duplicateResult = await executeOfficeTool({
    requestId: "ppt-duplicate-1",
    toolName: "duplicate_slide" as OfficeToolRequest["toolName"],
    host: "powerpoint",
    params: {
      slideId: "slide-2",
      targetSlideId: "slide-3",
      formatting: "KeepSourceFormatting",
    },
  } as OfficeToolRequest);
  const unsupportedResult = await executeOfficeTool({
    requestId: "ppt-slide-unsupported",
    toolName: "get_slide" as OfficeToolRequest["toolName"],
    host: "word",
    params: {},
  } as OfficeToolRequest);

  assert.equal(structureResult.success, true);
  assert.equal(slideResult.success, true);
  assert.equal(shapesResult.success, true);
  assert.equal(modifyResult.success, true);
  assert.equal(duplicateResult.success, true);
  assert.equal((calls[0]?.action as { type: string }).type, "getPresentationStructure");
  assert.equal((calls[1]?.action as { type: string }).type, "getSlide");
  assert.equal((calls[2]?.action as { type: string }).type, "listSlideShapes");
  assert.equal((calls[3]?.action as { type: string }).type, "reorderSlides");
  assert.deepEqual((calls[3]?.action as { options?: { slideIds?: string[] } }).options?.slideIds, ["slide-2", "slide-1"]);
  assert.equal((calls[4]?.action as { type: string }).type, "duplicateSlide");
  assert.deepEqual((calls[4]?.action as { options?: { slideIds?: string[] } }).options?.slideIds, ["slide-2"]);
  assert.equal(unsupportedResult.success, false);
  assert.match(String(unsupportedResult.error), /only available for PowerPoint/);
});
