import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import test from "node:test";

import {
  applyAcceptedEdits,
  proposeDocumentEdits,
} from "../../../apps/taskpane/src/lib/office/document-tools.js";
import { createOfficeExtension } from "../../../packages/pi-office-pack/src/extension.js";
import { OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH } from "../../../packages/pi-office-pack/src/protocol.js";

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

function extractSearchTextSchema(parameters: unknown): Record<string, unknown> | undefined {
  if (!parameters || typeof parameters !== "object") {
    return undefined;
  }
  const root = parameters as { properties?: Record<string, unknown> };
  const editsSchema = root.properties?.edits as { items?: { properties?: Record<string, unknown> } } | undefined;
  const itemProps = editsSchema?.items?.properties;
  const searchTextSchema = itemProps?.searchText;
  if (!searchTextSchema || typeof searchTextSchema !== "object") {
    return undefined;
  }
  return searchTextSchema as Record<string, unknown>;
}

async function getInProcessProposeEditsTool(): Promise<{ description?: string; parameters?: unknown }> {
  installRuntimePolyfills();
  const runtime = await import(`../../../apps/taskpane/src/lib/runtime/inprocess-kernel.js?test=${Date.now()}-${Math.random()}`);
  const openResponse = await runtime.dispatchKernelRequest("/v1/sessions/open", {
    method: "POST",
    body: JSON.stringify({
      host: "word",
      documentId: `doc-searchtext-limit-${Date.now()}-${Math.random()}`,
      saved: true,
      title: "SearchText Limit",
    }),
  }) as { sessionId: string };

  const socket = runtime.createLocalBridgeSocket(openResponse.sessionId);
  try {
    const session = (socket as unknown as {
      session?: {
        agent?: {
          state?: {
            tools?: Array<{ name: string; description?: string; parameters?: unknown }>;
          };
        };
      };
    }).session;
    const tools = session?.agent?.state?.tools;
    assert.ok(Array.isArray(tools), "in-process runtime session should expose agent tools");

    const proposeTool = tools.find((tool) => tool.name === "office_propose_edits");
    assert.ok(proposeTool, "in-process runtime should register office_propose_edits");
    return proposeTool;
  } finally {
    socket.close();
  }
}

test("office_propose_edits shares one canonical searchText limit across extension, in-process runtime, and skill guidance", async () => {
  const registeredTools: Array<{ name: string; description?: string; parameters?: unknown }> = [];
  const extension = createOfficeExtension({
    getHost: () => "word",
    getState: () => undefined,
    invokeTool: async () => ({ ok: true }),
    invokeAskUser: async (request) => ({ requestId: request.requestId, answers: [] }),
  });
  extension({
    registerTool: (tool: { name: string; description?: string; parameters?: unknown }) => {
      registeredTools.push(tool);
    },
    on: () => undefined,
  } as unknown as Parameters<ReturnType<typeof createOfficeExtension>>[0]);

  const extensionProposeTool = registeredTools.find((tool) => tool.name === "office_propose_edits");
  assert.ok(extensionProposeTool, "extension should register office_propose_edits");
  assert.match(
    extensionProposeTool.description ?? "",
    new RegExp(`under\\s+${OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH}\\s+characters`, "i"),
  );
  const extensionSearchTextSchema = extractSearchTextSchema(extensionProposeTool.parameters);
  assert.equal(
    extensionSearchTextSchema?.maxLength,
    OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH,
    "extension schema should enforce the canonical searchText maxLength",
  );

  const inProcessProposeTool = await getInProcessProposeEditsTool();
  assert.match(
    inProcessProposeTool.description ?? "",
    new RegExp(`under\\s+${OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH}\\s+characters`, "i"),
  );
  const inProcessSearchTextSchema = extractSearchTextSchema(inProcessProposeTool.parameters);
  assert.equal(
    inProcessSearchTextSchema?.maxLength,
    OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH,
    "in-process runtime schema should enforce the canonical searchText maxLength",
  );

  const officeHostSkillPath = join(process.cwd(), "packages", "pi-office-pack", "skills", "office-host.SKILL.md");
  const officeHostSkillText = readFileSync(officeHostSkillPath, "utf8");
  const searchTextGuidanceLines = officeHostSkillText
    .split(/\r?\n/u)
    .filter((line) => /searchText/i.test(line) && /under/i.test(line));
  assert.ok(searchTextGuidanceLines.length > 0, "office host skill should include searchText limit guidance");
  const searchTextNumbers = [
    ...new Set(
      searchTextGuidanceLines.flatMap((line) =>
        [...line.matchAll(/\b\d+\b/g)].map((match) => Number(match[0])),
      ),
    ),
  ];
  assert.deepEqual(
    searchTextNumbers,
    [OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH],
    "skill guidance should mention only the canonical searchText limit",
  );
});

test("proposeDocumentEdits rejects over-limit searchText inputs explicitly", async () => {
  const overLimitSearchText = "x".repeat(OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH + 1);
  const overLimitCases = [
    { kind: "replace", searchText: overLimitSearchText, newText: "replacement" },
    { kind: "replace", oldText: overLimitSearchText, newText: "replacement" },
  ];

  for (const edit of overLimitCases) {
    const result = await proposeDocumentEdits("word", {
      summary: "Over-limit searchText test",
      edits: [edit],
    }) as { error?: string };
    assert.equal(typeof result.error, "string");
    assert.match(result.error ?? "", /searchText/i);
    assert.match(result.error ?? "", /edit 1/i);
    assert.match(result.error ?? "", new RegExp(`\\b${OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH}\\b`));
  }
});

test("applyAcceptedEdits rejects over-limit searchText inputs explicitly", async () => {
  const overLimitSearchText = "x".repeat(OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH + 1);
  const result = await applyAcceptedEdits([
    { kind: "replace", searchText: overLimitSearchText, newText: "replacement" },
    { kind: "replace", oldText: overLimitSearchText, newText: "replacement" },
  ]);

  assert.equal(result.applied, 0);
  assert.equal(result.failed, 2);
  assert.equal(result.errors.length, 2);
  assert.match(result.errors[0] ?? "", /searchText/i);
  assert.match(result.errors[0] ?? "", /edit 1/i);
  assert.match(result.errors[1] ?? "", /edit 2/i);
  assert.match(result.errors[0] ?? "", new RegExp(`\\b${OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH}\\b`));
});
