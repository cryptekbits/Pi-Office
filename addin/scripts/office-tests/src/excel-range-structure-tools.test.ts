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

test("protocol inventory includes first-class Excel range/layout and worksheet-structure tools", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("get_cell_ranges" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("set_cell_range" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("clear_cell_range" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("resize_range" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("copy_to" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("modify_sheet_structure" as any));
  assert.equal(TOOL_CATEGORY_MAP.get_cell_ranges, "read");
  assert.equal(TOOL_CATEGORY_MAP.set_cell_range, "write-doc");
  assert.equal(TOOL_CATEGORY_MAP.clear_cell_range, "write-doc");
  assert.equal(TOOL_CATEGORY_MAP.resize_range, "write-doc");
  assert.equal(TOOL_CATEGORY_MAP.copy_to, "write-doc");
  assert.equal(TOOL_CATEGORY_MAP.modify_sheet_structure, "write-doc");
});

test("createOfficeExtension registers first-class Excel range/layout and worksheet-structure tools", async () => {
  const registeredTools = new Map<string, {
    execute: (toolCallId: string, params: Record<string, unknown>) => Promise<{ content: unknown; details: unknown }>;
  }>();
  const invocations: Array<{ toolName: string; params: Record<string, unknown> }> = [];

  const extensionFactory = createOfficeExtension({
    getHost: () => "excel",
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

  const excelToolNames = [
    "get_cell_ranges",
    "set_cell_range",
    "clear_cell_range",
    "resize_range",
    "copy_to",
    "modify_sheet_structure",
  ];
  for (const toolName of excelToolNames) {
    assert.ok(registeredTools.has(toolName));
  }

  await registeredTools.get("get_cell_ranges")!.execute("excel-read", { sheetName: "Budget", address: "A1:B2" });
  await registeredTools.get("set_cell_range")!.execute("excel-write", { sheetName: "Budget", address: "A1:B2", values: [[1, 2]] });
  await registeredTools.get("clear_cell_range")!.execute("excel-clear", { sheetName: "Budget", address: "A1:B2" });
  await registeredTools.get("resize_range")!.execute("excel-resize", { sheetName: "Budget", address: "A1:B2", rowCount: 4, columnCount: 3 });
  await registeredTools.get("copy_to")!.execute("excel-copy", {
    sourceSheetName: "Budget",
    sourceAddress: "A1:B2",
    destinationSheetName: "Summary",
    destinationAddress: "D5:E6",
  });
  await registeredTools.get("modify_sheet_structure")!.execute("excel-sheet-ops", {
    operation: "create_worksheet",
    name: "Forecast",
  });

  assert.equal(invocations[0]?.toolName, "get_cell_ranges");
  assert.equal(invocations[1]?.toolName, "set_cell_range");
  assert.equal(invocations[2]?.toolName, "clear_cell_range");
  assert.equal(invocations[3]?.toolName, "resize_range");
  assert.equal(invocations[4]?.toolName, "copy_to");
  assert.equal(invocations[5]?.toolName, "modify_sheet_structure");
});

test("in-process runtime publishes first-class Excel range/layout and worksheet-structure tools", async () => {
  const runtime = await loadKernelModule();
  const openResponse = await runtime.dispatchKernelRequest("/v1/sessions/open", {
    method: "POST",
    body: JSON.stringify({
      host: "excel",
      documentId: `doc-excel-range-${Date.now()}`,
      saved: true,
      title: "Excel Range Tool Inventory Test",
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

  assert.ok(toolNames.includes("get_cell_ranges"));
  assert.ok(toolNames.includes("set_cell_range"));
  assert.ok(toolNames.includes("clear_cell_range"));
  assert.ok(toolNames.includes("resize_range"));
  assert.ok(toolNames.includes("copy_to"));
  assert.ok(toolNames.includes("modify_sheet_structure"));
  socket.close();
});

test("createOfficeToolExecutor dispatches first-class Excel range/layout and worksheet-structure tools", async () => {
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

  const readResult = await executeOfficeTool({
    requestId: "excel-read-1",
    toolName: "get_cell_ranges" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {
      sheetName: "Budget",
      address: "A1:B2",
      includeFormulas: true,
      includeNumberFormat: true,
    },
  } as OfficeToolRequest);
  const writeResult = await executeOfficeTool({
    requestId: "excel-write-1",
    toolName: "set_cell_range" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {
      sheetName: "Budget",
      address: "A1:B2",
      values: [[1, 2]],
    },
  } as OfficeToolRequest);
  const clearResult = await executeOfficeTool({
    requestId: "excel-clear-1",
    toolName: "clear_cell_range" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {
      sheetName: "Budget",
      address: "A1:B2",
      applyTo: "contents",
      confirmDestructive: true,
    },
  } as OfficeToolRequest);
  const resizeResult = await executeOfficeTool({
    requestId: "excel-resize-1",
    toolName: "resize_range" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {
      sheetName: "Budget",
      address: "A1:B2",
      rowCount: 5,
      columnCount: 4,
      activate: true,
    },
  } as OfficeToolRequest);
  const copyResult = await executeOfficeTool({
    requestId: "excel-copy-1",
    toolName: "copy_to" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {
      sourceSheetName: "Budget",
      sourceAddress: "A1:B2",
      destinationSheetName: "Summary",
      destinationAddress: "D5:E6",
      copyType: "All",
    },
  } as OfficeToolRequest);
  const createSheetResult = await executeOfficeTool({
    requestId: "excel-sheet-create-1",
    toolName: "modify_sheet_structure" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {
      operation: "create_worksheet",
      name: "Forecast",
    },
  } as OfficeToolRequest);
  const renameSheetResult = await executeOfficeTool({
    requestId: "excel-sheet-rename-1",
    toolName: "modify_sheet_structure" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {
      operation: "rename_worksheet",
      sheetName: "Forecast",
      name: "Forecast Q2",
    },
  } as OfficeToolRequest);
  const duplicateSheetResult = await executeOfficeTool({
    requestId: "excel-sheet-duplicate-1",
    toolName: "modify_sheet_structure" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {
      operation: "duplicate_worksheet",
      sheetName: "Forecast Q2",
      name: "Forecast Q2 Copy",
    },
  } as OfficeToolRequest);
  const deleteSheetResult = await executeOfficeTool({
    requestId: "excel-sheet-delete-1",
    toolName: "modify_sheet_structure" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {
      operation: "delete_worksheet",
      sheetName: "Forecast Q2 Copy",
      confirmDestructive: true,
    },
  } as OfficeToolRequest);
  const unsupportedResult = await executeOfficeTool({
    requestId: "excel-read-unsupported",
    toolName: "get_cell_ranges" as OfficeToolRequest["toolName"],
    host: "word",
    params: {},
  } as OfficeToolRequest);

  assert.equal(readResult.success, true);
  assert.equal(writeResult.success, true);
  assert.equal(clearResult.success, true);
  assert.equal(resizeResult.success, true);
  assert.equal(copyResult.success, true);
  assert.equal(createSheetResult.success, true);
  assert.equal(renameSheetResult.success, true);
  assert.equal(duplicateSheetResult.success, true);
  assert.equal(deleteSheetResult.success, true);
  assert.equal((calls[0]?.action as { type: string }).type, "getRangeValues");
  assert.equal((calls[1]?.action as { type: string }).type, "setRangeValues");
  assert.deepEqual((calls[1]?.action as { values?: unknown }).values, [[1, 2]]);
  assert.equal((calls[2]?.action as { type: string }).type, "clearRange");
  assert.equal((calls[3]?.action as { type: string }).type, "resizeRange");
  assert.equal((calls[4]?.action as { type: string }).type, "copyRange");
  assert.equal((calls[5]?.action as { type: string }).type, "createWorksheet");
  assert.equal((calls[6]?.action as { type: string }).type, "renameWorksheet");
  assert.equal((calls[7]?.action as { type: string }).type, "duplicateWorksheet");
  assert.equal((calls[8]?.action as { type: string }).type, "deleteWorksheet");
  assert.equal(unsupportedResult.success, false);
  assert.match(String(unsupportedResult.error), /only available for Excel/);
});
