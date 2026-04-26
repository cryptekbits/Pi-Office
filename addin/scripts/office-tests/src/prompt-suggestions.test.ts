import assert from "node:assert/strict";
import test from "node:test";

import { fauxAssistantMessage, registerFauxProvider, type Context } from "@mariozechner/pi-ai";
import {
  parsePromptSuggestions,
  shouldClearPromptSuggestionsForUiChange,
  shouldShowPromptSuggestions,
} from "../../../packages/pi-office-pack/src/prompt-suggestions.js";
import type { PromptSuggestionResponse } from "../../../packages/pi-office-pack/src/protocol.js";

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

test("prompt suggestions parser trims, dedupes, and clamps to three", () => {
  const suggestions = parsePromptSuggestions(JSON.stringify({
    suggestions: [
      { text: "  Inspect the current selection  " },
      { text: "Inspect the current selection." },
      { text: "Summarize the selected table" },
      { text: "Apply this rewrite natively" },
      { text: "Verify the document state" },
    ],
  }));

  assert.deepEqual(
    suggestions.map((entry) => entry.text),
    [
      "Inspect the current selection",
      "Summarize the selected table",
      "Apply this rewrite natively",
    ],
  );
});

test("prompt suggestions parser accepts empty and fenced JSON results", () => {
  assert.deepEqual(parsePromptSuggestions('{"suggestions":[]}'), []);

  const suggestions = parsePromptSuggestions('```json\n{"suggestions":["Compare these alternatives"]}\n```');
  assert.equal(suggestions.length, 1);
  assert.equal(suggestions[0]?.text, "Compare these alternatives");
});

test("prompt suggestions parser filters generic, noisy, and unsaved-file suggestions", () => {
  const suggestions = parsePromptSuggestions(JSON.stringify({
    suggestions: [
      "Continue",
      "Tell me more",
      "```json",
      "Scan the local folder for AGENTS.md",
      "Verify the selected paragraph formatting",
    ],
  }), { documentState: "unsaved" });

  assert.deepEqual(
    suggestions.map((entry) => entry.text),
    ["Verify the selected paragraph formatting"],
  );
});

test("prompt suggestion visibility and clearing rules match the taskpane behavior", () => {
  const suggestion = { id: "sug-1", text: "Inspect the current selection" };
  const idle = {
    nextPromptSuggestionsEnabled: true,
    isBusy: false,
    draft: "",
    sessionId: "s1",
    selectionFingerprint: "sel1",
  };

  assert.equal(shouldShowPromptSuggestions([suggestion], idle), true);
  assert.equal(shouldShowPromptSuggestions([suggestion], { ...idle, isBusy: true }), false);
  assert.equal(shouldShowPromptSuggestions([suggestion], { ...idle, draft: "new prompt" }), false);
  assert.equal(shouldShowPromptSuggestions([suggestion], { ...idle, nextPromptSuggestionsEnabled: false }), false);

  assert.equal(shouldClearPromptSuggestionsForUiChange(idle, { ...idle, isBusy: true }), true);
  assert.equal(shouldClearPromptSuggestionsForUiChange(idle, { ...idle, draft: "edit" }), true);
  assert.equal(shouldClearPromptSuggestionsForUiChange(idle, { ...idle, nextPromptSuggestionsEnabled: false }), true);
  assert.equal(shouldClearPromptSuggestionsForUiChange(idle, { ...idle, selectionFingerprint: "sel2" }), true);
  assert.equal(shouldClearPromptSuggestionsForUiChange(idle, { ...idle, sessionId: "s2" }), true);
  assert.equal(shouldClearPromptSuggestionsForUiChange(idle, { ...idle }), false);
});

test("prompt suggestions route grounds model context and returns empty on invalid model JSON", async () => {
  const faux = registerFauxProvider({
    provider: "faux-suggestions",
    models: [{ id: "suggestion-faux", name: "Suggestion Faux" }],
    tokensPerSecond: 0,
  });
  let seenContext: Context | undefined;
  faux.setResponses([
    (context) => {
      seenContext = context;
      return fauxAssistantMessage("not json at all");
    },
  ]);

  try {
    const runtime = await loadKernelModule();
    await runtime.dispatchKernelRequest("/v1/auth/api-key", {
      method: "POST",
      body: JSON.stringify({ provider: "faux-suggestions", apiKey: "test-key" }),
    });

    const openResponse = await runtime.dispatchKernelRequest("/v1/sessions/open", {
      method: "POST",
      body: JSON.stringify({
        host: "word",
        documentId: `doc-suggestions-${Date.now()}`,
        saved: true,
        title: "Quarterly Plan",
        selectionSummary: {
          label: "Body paragraph",
          textPreview: "The selected paragraph describes quarterly priorities.",
        },
      }),
    }) as { sessionId: string };

    const socket = runtime.createLocalBridgeSocket(openResponse.sessionId);
    const internalSession = (socket as unknown as {
      session: { agent: { setModel: (model: unknown) => void } };
    }).session;
    internalSession.agent.setModel(faux.getModel());

    const response = await runtime.dispatchKernelRequest(
      `/v1/sessions/${openResponse.sessionId}/prompt-suggestions`,
      {
        method: "POST",
        body: JSON.stringify({
          generationId: "gen-route",
          latestAssistantText: "I summarized the selected paragraph and found two priority gaps.",
          recentMessages: [
            { role: "user", text: "Summarize the selected paragraph." },
            { role: "assistant", text: "I summarized the selected paragraph." },
          ],
        }),
      },
    ) as PromptSuggestionResponse;

    assert.equal(response.generationId, "gen-route");
    assert.deepEqual(response.suggestions, []);
    assert.equal(faux.state.callCount, 1);
    assert.match(seenContext?.systemPrompt ?? "", /strict JSON/);
    const userContent = String(seenContext?.messages[0]?.content ?? "");
    assert.match(userContent, /Host: Word/);
    assert.match(userContent, /Document title: Quarterly Plan/);
    assert.match(userContent, /Document state: saved/);
    assert.match(userContent, /Selection label: Body paragraph/);
    assert.match(userContent, /The selected paragraph describes quarterly priorities/);
    assert.match(userContent, /Latest assistant response:/);
    socket.close();
  } finally {
    faux.unregister();
  }
});
