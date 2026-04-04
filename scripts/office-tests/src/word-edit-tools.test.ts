import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import test from "node:test";

import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
import { createOfficeExtension } from "../../../packages/pi-office-pack/src/extension.js";
import {
  OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH,
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

test("protocol inventory includes first-class Word text/list edit tools with write-doc categories", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("edit_doc_text"));
  assert.ok(OFFICE_TOOL_NAMES.includes("edit_doc_list"));
  assert.equal(TOOL_CATEGORY_MAP.edit_doc_text, "write-doc");
  assert.equal(TOOL_CATEGORY_MAP.edit_doc_list, "write-doc");
});

test("createOfficeExtension registers Word text/list edit tools and routes calls", async () => {
  const registeredTools = new Map<string, {
    execute: (toolCallId: string, params: Record<string, unknown>) => Promise<{ content: unknown; details: unknown }>;
  }>();
  const invocations: Array<{ toolName: string; params: Record<string, unknown> }> = [];

  const extensionFactory = createOfficeExtension({
    getHost: () => "word",
    getState: () => undefined,
    invokeTool: async (toolName, params) => {
      invocations.push({ toolName, params });
      if (toolName === "edit_doc_text") {
        return { ok: true, operation: "insertText" };
      }
      return {
        summary: "List edits proposed",
        edits: [
          {
            kind: "replace",
            searchText: "First bullet",
            newText: "Updated bullet",
          },
        ],
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

  assert.ok(registeredTools.has("edit_doc_text"));
  assert.ok(registeredTools.has("edit_doc_list"));

  const textEditResult = await registeredTools.get("edit_doc_text")!.execute("tool-word-text-edit", {
    operation: "insertText",
    text: "Updated clause",
  });
  assert.equal(invocations[0]?.toolName, "edit_doc_text");
  assert.equal((textEditResult.details as { operation: string }).operation, "insertText");

  const listEditResult = await registeredTools.get("edit_doc_list")!.execute("tool-word-list-edit", {
    summary: "Update list language",
    edits: [
      {
        kind: "replace",
        searchText: "First bullet",
        newText: "Updated bullet",
      },
    ],
  });
  assert.equal(invocations[1]?.toolName, "edit_doc_list");
  assert.match(String((listEditResult.details as { summary: string }).summary), /List edits proposed/i);
});

test("in-process runtime publishes first-class Word text/list edit tools", async () => {
  const runtime = await loadKernelModule();
  const openResponse = await runtime.dispatchKernelRequest("/v1/sessions/open", {
    method: "POST",
    body: JSON.stringify({
      host: "word",
      documentId: `doc-word-edit-${Date.now()}`,
      saved: true,
      title: "Word Edit Tool Inventory Test",
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

  assert.ok(toolNames.includes("edit_doc_text"));
  assert.ok(toolNames.includes("edit_doc_list"));
  socket.close();
});

test("createOfficeToolExecutor dispatches first-class Word text/list tools and enforces Word-only scope", async () => {
  const calls: Array<{ name: string; payload: unknown }> = [];
  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async () => ({ ok: true }),
    applyHostAction: async (host, action) => {
      calls.push({ name: "applyHostAction", payload: { host, action } });
      return { ok: true, host, action };
    },
    navigateOfficeAnchor: async () => ({ ok: true }),
    readDocumentSection: async () => ({ ok: true }),
    executeOfficeJs: async () => ({ ok: true }),
    proposeEdits: async (host, params) => {
      calls.push({ name: "proposeEdits", payload: { host, params } });
      return {
        ok: true,
        summary: "List proposal prepared",
        edits: [{ id: "edit-1", kind: "replace" }],
      };
    },
  });

  const textEditResult = await executeOfficeTool({
    requestId: "word-text-1",
    toolName: "edit_doc_text" as OfficeToolRequest["toolName"],
    host: "word",
    params: {
      operation: "insertText",
      text: "Updated clause",
    },
  } as OfficeToolRequest);
  const listEditResult = await executeOfficeTool({
    requestId: "word-list-1",
    toolName: "edit_doc_list" as OfficeToolRequest["toolName"],
    host: "word",
    params: {
      summary: "Revise bullets",
      edits: [
        {
          kind: "replace",
          searchText: "First bullet",
          newText: "Updated bullet",
        },
      ],
    },
  } as OfficeToolRequest);
  const unsupportedTextResult = await executeOfficeTool({
    requestId: "word-text-unsupported",
    toolName: "edit_doc_text" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {},
  } as OfficeToolRequest);
  const unsupportedListResult = await executeOfficeTool({
    requestId: "word-list-unsupported",
    toolName: "edit_doc_list" as OfficeToolRequest["toolName"],
    host: "powerpoint",
    params: {},
  } as OfficeToolRequest);

  assert.equal(textEditResult.success, true);
  assert.equal(listEditResult.success, true);
  assert.equal(calls[0]?.name, "applyHostAction");
  assert.equal((calls[0]?.payload as { action: { type: string } }).action.type, "insertText");
  assert.equal(calls[1]?.name, "proposeEdits");
  assert.equal((calls[1]?.payload as { host: string }).host, "word");
  assert.equal(unsupportedTextResult.success, false);
  assert.match(String(unsupportedTextResult.error), /only available for Word/);
  assert.equal(unsupportedListResult.success, false);
  assert.match(String(unsupportedListResult.error), /only available for Word/);
});

test("Word guidance aligns first-class review/editing contract with runtime constraints", () => {
  const limitPattern = new RegExp(`under\\s+${OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH}\\s+characters`, "i");
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bedit_doc_text\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bedit_doc_list\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /tracked changes|legal/i);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /paragraphId/i);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\banchor\b/i);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, limitPattern);

  const officeHostSkillPath = join(process.cwd(), "packages", "pi-office-pack", "skills", "office-host.SKILL.md");
  const officeHostSkillText = readFileSync(officeHostSkillPath, "utf8");
  assert.match(officeHostSkillText, /\bedit_doc_text\b/);
  assert.match(officeHostSkillText, /\bedit_doc_list\b/);
  assert.match(officeHostSkillText, /tracked changes|legal/i);
  assert.match(officeHostSkillText, /paragraphId/i);
  assert.match(officeHostSkillText, /\banchor\b/i);
  assert.match(officeHostSkillText, limitPattern);
});

test("investigation artifact includes Word manual checklist scenarios for anchors, structured edits, visual verification, and taskpane stability", () => {
  const investigationPath = join(process.cwd(), "CLAUDE_ADDIN_INVESTIGATION_AND_TRACKING.md");
  const investigationText = readFileSync(investigationPath, "utf8");
  assert.match(investigationText, /Manual Testing Scenarios and Steps/i);
  assert.match(investigationText, /word-review-anchors/);
  assert.match(investigationText, /word-structured-edits/);
  assert.match(investigationText, /word-visual-verification/);
  assert.match(investigationText, /word-taskpane-stability/);
});
