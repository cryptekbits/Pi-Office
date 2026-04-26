import assert from "node:assert/strict";
import test from "node:test";

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

async function withFastPermissionTimeout<T>(body: () => Promise<T>): Promise<T> {
  const timers = globalThis as unknown as {
    setTimeout: typeof setTimeout;
  };
  const originalSetTimeout = timers.setTimeout;
  timers.setTimeout = ((handler: TimerHandler, timeout?: number, ...args: unknown[]) =>
    originalSetTimeout(handler, timeout === 120_000 ? 5 : timeout, ...args)) as typeof setTimeout;
  try {
    return await body();
  } finally {
    timers.setTimeout = originalSetTimeout;
  }
}

async function openSessionHarness(label: string) {
  const runtime = await loadKernelModule();
  const documentId = `doc-${label}-${Date.now()}`;
  const openResponse = await runtime.dispatchKernelRequest("/v1/sessions/open", {
    method: "POST",
    body: JSON.stringify({
      host: "word",
      documentId,
      saved: true,
      title: `Doc ${label}`,
    }),
  }) as {
    sessionId: string;
    documentState: string;
    companion: { status: string };
    origin: string;
    eventsPath: string;
  };

  assert.equal(openResponse.documentState, "saved");
  assert.equal(typeof openResponse.companion.status, "string");

  const socket = runtime.createLocalBridgeSocket(openResponse.sessionId);
  await waitForEvent(socket, "open");
  const readyMessage = waitForServerMessage(
    socket,
    (payload) => payload.type === "connection_state" && payload.state === "ready",
  );
  socket.send(JSON.stringify({ type: "client_ready" }));
  await readyMessage;

  return {
    runtime,
    socket,
    session: (socket as unknown as { session: Record<string, unknown> }).session as Record<string, unknown>,
    documentId,
  };
}

test("protocol parity: ask_user request/response roundtrip", async () => {
  const { socket, session } = await openSessionHarness("ask-user");

  const request = {
    requestId: "ask-1",
    questions: [
      {
        id: "q1",
        question: "Proceed?",
        options: [{ title: "Yes" }, { title: "No" }],
      },
    ],
  };

  const outboundPromise = waitForServerMessage(
    socket,
    (payload) =>
      payload.type === "ask_user_request" &&
      (payload.request as { requestId?: string } | undefined)?.requestId === request.requestId,
  );
  const pending = (session as { invokeAskUser: (req: unknown) => Promise<unknown> }).invokeAskUser(request);
  const outbound = await outboundPromise;
  assert.equal((outbound.request as { requestId: string }).requestId, request.requestId);

  socket.send(
    JSON.stringify({
      type: "ask_user_response",
      response: {
        requestId: request.requestId,
        answers: [{ questionId: "q1", selectedOption: "Yes" }],
      },
    }),
  );

  const response = await pending as { requestId: string; answers: Array<{ selectedOption: string | null }> };
  assert.equal(response.requestId, request.requestId);
  assert.equal(response.answers[0]?.selectedOption, "Yes");
  socket.close();
});

test("protocol parity: tool_permission request/response + session approval caching", async () => {
  const { socket, session } = await openSessionHarness("tool-permission");

  const outboundPromise = waitForServerMessage(socket, (payload) => payload.type === "tool_permission_request");
  const pending = (session as {
    requestToolPermission: (toolName: string, params: Record<string, unknown>) => Promise<unknown>;
  }).requestToolPermission("office_apply_edit", { mode: "replaceSelection" });

  const outbound = await outboundPromise;
  const request = outbound.request as { requestId: string; toolName: string };
  assert.equal(request.toolName, "office_apply_edit");

  socket.send(
    JSON.stringify({
      type: "tool_permission_response",
      requestId: request.requestId,
      decision: {
        toolName: "office_apply_edit",
        allowed: true,
        scope: "session",
      },
    }),
  );

  const decision = await pending as { allowed: boolean; scope: string };
  assert.equal(decision.allowed, true);
  assert.equal(decision.scope, "session");
  const isAutoApproved = (session as { shouldAutoApproveTool: (name: string) => boolean }).shouldAutoApproveTool("office_apply_edit");
  assert.equal(isAutoApproved, true);
  socket.close();
});

test("protocol parity: tool permission timeouts deny across gated categories", async () => {
  const { socket, session } = await openSessionHarness("tool-permission-timeout");

  await withFastPermissionTimeout(async () => {
    for (const scenario of [
      { toolName: "office_apply_edit", category: "write-doc" },
      { toolName: "mcp", category: "connector" },
      { toolName: "read", category: "read-external" },
      { toolName: "bash", category: "write-external" },
    ]) {
      const requestMessage = waitForServerMessage(
        socket,
        (payload) =>
          payload.type === "tool_permission_request" &&
          (payload.request as { toolName?: string } | undefined)?.toolName === scenario.toolName,
      );
      const expiredMessage = waitForServerMessage(
        socket,
        (payload) => payload.type === "tool_permission_expired" && payload.toolName === scenario.toolName,
      );
      const decisionPromise = (session as {
        requestToolPermission: (toolName: string, params: Record<string, unknown>) => Promise<unknown>;
      }).requestToolPermission(scenario.toolName, { reason: "test" });

      const outbound = await requestMessage;
      const request = outbound.request as { requestId: string; toolName: string; toolCategory: string };
      assert.equal(request.toolName, scenario.toolName);
      assert.equal(request.toolCategory, scenario.category);

      const expired = await expiredMessage;
      assert.equal(expired.requestId, request.requestId);

      const decision = await decisionPromise as { toolName: string; allowed: boolean; scope: string };
      assert.equal(decision.toolName, scenario.toolName);
      assert.equal(decision.allowed, false);
      assert.equal(decision.scope, "once");
    }
  });

  socket.close();
});

test("protocol parity: office_execute_js requires per-call approval", async () => {
  const { runtime, socket, session } = await openSessionHarness("execute-js-per-call");

  for (const autonomyLevel of ["medium", "high", "extreme"]) {
    await runtime.dispatchKernelRequest("/v1/preferences", {
      method: "POST",
      body: JSON.stringify({ autonomyLevel, toolPermissionOverrides: [] }),
    });
    const isAutoApproved = (session as { shouldAutoApproveTool: (name: string) => boolean }).shouldAutoApproveTool("office_execute_js");
    assert.equal(isAutoApproved, false, `${autonomyLevel} autonomy must not auto-approve office_execute_js.`);
  }

  const deniedMessage = waitForServerMessage(socket, (payload) => payload.type === "tool_permission_request");
  const deniedPromise = (session as {
    requestToolPermission: (toolName: string, params: Record<string, unknown>) => Promise<unknown>;
  }).requestToolPermission("office_execute_js", { code: "return 2 + 2;" });
  const deniedRequest = (await deniedMessage).request as { requestId: string; toolCategory: string };
  assert.equal(deniedRequest.toolCategory, "escape-hatch");
  socket.send(
    JSON.stringify({
      type: "tool_permission_response",
      requestId: deniedRequest.requestId,
      decision: { toolName: "office_execute_js", allowed: false, scope: "once" },
    }),
  );
  const denied = await deniedPromise as { allowed: boolean; scope: string };
  assert.equal(denied.allowed, false);
  assert.equal(denied.scope, "once");

  const approvedMessage = waitForServerMessage(socket, (payload) => payload.type === "tool_permission_request");
  const approvedPromise = (session as {
    requestToolPermission: (toolName: string, params: Record<string, unknown>) => Promise<unknown>;
  }).requestToolPermission("office_execute_js", { code: "return 4;" });
  const approvedRequest = (await approvedMessage).request as { requestId: string; toolCategory: string };
  assert.equal(approvedRequest.toolCategory, "escape-hatch");
  socket.send(
    JSON.stringify({
      type: "tool_permission_response",
      requestId: approvedRequest.requestId,
      decision: { toolName: "office_execute_js", allowed: true, scope: "session" },
    }),
  );
  const approved = await approvedPromise as { allowed: boolean; scope: string };
  assert.equal(approved.allowed, true);
  assert.equal(approved.scope, "once");
  assert.equal(
    (session as { shouldAutoApproveTool: (name: string) => boolean }).shouldAutoApproveTool("office_execute_js"),
    false,
  );

  const nextPrompt = waitForServerMessage(socket, (payload) => payload.type === "tool_permission_request");
  const nextPromise = (session as {
    requestToolPermission: (toolName: string, params: Record<string, unknown>) => Promise<unknown>;
  }).requestToolPermission("office_execute_js", { code: "return 5;" });
  const nextRequest = (await nextPrompt).request as { requestId: string };
  socket.send(
    JSON.stringify({
      type: "tool_permission_response",
      requestId: nextRequest.requestId,
      decision: { toolName: "office_execute_js", allowed: false, scope: "once" },
    }),
  );
  const next = await nextPromise as { allowed: boolean };
  assert.equal(next.allowed, false);
  socket.close();
});

test("protocol parity: edit proposal request/decision roundtrip", async () => {
  const { socket, session } = await openSessionHarness("edit-proposal");
  const proposal = {
    requestId: "proposal-1",
    summary: "Propose one edit",
    edits: [
      {
        id: "edit-1",
        kind: "replace",
        searchText: "Before",
        newText: "After",
      },
    ],
  };

  const outboundPromise = waitForServerMessage(
    socket,
    (payload) =>
      payload.type === "edit_proposal_request" &&
      (payload.proposal as { requestId?: string } | undefined)?.requestId === proposal.requestId,
  );
  const pending = (session as { invokeEditProposal: (proposal: unknown) => Promise<unknown> }).invokeEditProposal(proposal);
  const outbound = await outboundPromise;
  assert.equal((outbound.proposal as { requestId: string }).requestId, proposal.requestId);

  socket.send(
    JSON.stringify({
      type: "edit_proposal_decision",
      decision: {
        requestId: proposal.requestId,
        decisions: [{ editId: "edit-1", accepted: true }],
      },
    }),
  );

  const decision = await pending as { requestId: string; decisions: Array<{ accepted: boolean }> };
  assert.equal(decision.requestId, proposal.requestId);
  assert.equal(decision.decisions[0]?.accepted, true);
  socket.close();
});

test("protocol parity: rewind and checkpoint persist/load events", async () => {
  const { socket, session, documentId } = await openSessionHarness("rewind-checkpoint");
  const agent = (session as { agent: { replaceMessages: (messages: unknown[]) => void; state: { messages: unknown[] } } }).agent;
  agent.replaceMessages([
    { role: "user", content: [{ type: "text", text: "User message" }], timestamp: Date.now() - 1000 },
    { role: "assistant", content: [{ type: "text", text: "Assistant message" }], timestamp: Date.now() },
  ]);

  socket.send(JSON.stringify({ type: "rewind_session", targetMessageCount: 1 }));
  assert.equal(agent.state.messages.length, 1);

  const checkpoint = {
    id: "cp-1",
    timestamp: Date.now(),
    host: "word",
    userPrompt: "Test prompt",
    messageCount: 1,
    ooxml: "<w:document/>",
  };

  const availablePromise = waitForServerMessage(
    socket,
    (payload) =>
      payload.type === "available_checkpoints" &&
      Array.isArray(payload.checkpoints) &&
      payload.checkpoints.some((entry: { id?: string }) => entry.id === checkpoint.id),
  );
  socket.send(
    JSON.stringify({
      type: "persist_checkpoint",
      documentId,
      checkpoint,
    }),
  );
  const available = await availablePromise;
  const checkpoints = available.checkpoints as Array<{ id: string }>;
  assert.equal(checkpoints.some((entry) => entry.id === checkpoint.id), true);

  const loadedPromise = waitForServerMessage(
    socket,
    (payload) =>
      payload.type === "session_event" &&
      (payload.event as { type?: string; checkpoint?: { id?: string } } | undefined)?.type === "checkpoint_data" &&
      (payload.event as { checkpoint?: { id?: string } }).checkpoint?.id === checkpoint.id,
  );
  socket.send(
    JSON.stringify({
      type: "load_checkpoint",
      documentId,
      checkpointId: checkpoint.id,
    }),
  );
  const loaded = await loadedPromise;
  const event = loaded.event as { checkpoint?: { id?: string } };
  assert.equal(event.checkpoint?.id, checkpoint.id);
  socket.close();
});

test("protocol parity: disconnect cancels pending interactive requests", async () => {
  const { socket, session } = await openSessionHarness("disconnect-cancel");

  const askRequestMessage = waitForServerMessage(socket, (payload) => payload.type === "ask_user_request");
  const permissionRequestMessage = waitForServerMessage(socket, (payload) => payload.type === "tool_permission_request");
  const proposalRequestMessage = waitForServerMessage(socket, (payload) => payload.type === "edit_proposal_request");

  const askPromise = (session as { invokeAskUser: (req: unknown) => Promise<unknown> }).invokeAskUser({
    requestId: "ask-disconnect",
    questions: [{ id: "q1", question: "Question", options: [{ title: "A" }] }],
  });
  const permissionPromise = (session as {
    requestToolPermission: (toolName: string, params: Record<string, unknown>) => Promise<unknown>;
  }).requestToolPermission("office_apply_edit", {});
  const proposalPromise = (session as { invokeEditProposal: (proposal: unknown) => Promise<unknown> }).invokeEditProposal({
    requestId: "proposal-disconnect",
    summary: "Proposal",
    edits: [{ id: "e1", kind: "replace", searchText: "A", newText: "B" }],
  });

  await askRequestMessage;
  await permissionRequestMessage;
  await proposalRequestMessage;

  socket.close();

  await assert.rejects(askPromise, /Bridge disconnected before interactive request completed/);
  await assert.rejects(permissionPromise, /Bridge disconnected before interactive request completed/);
  await assert.rejects(proposalPromise, /Bridge disconnected before interactive request completed/);
});

test("protocol parity: session cleanup rejects pending tool permissions", async () => {
  const { runtime, socket, session, documentId } = await openSessionHarness("permission-cleanup");

  const permissionRequestMessage = waitForServerMessage(socket, (payload) => payload.type === "tool_permission_request");
  const permissionPromise = (session as {
    requestToolPermission: (toolName: string, params: Record<string, unknown>) => Promise<unknown>;
  }).requestToolPermission("mcp", { toolName: "connector.read" });

  await permissionRequestMessage;

  const reopened = await runtime.dispatchKernelRequest("/v1/sessions/open", {
    method: "POST",
    body: JSON.stringify({
      host: "word",
      documentId,
      saved: true,
      title: "Doc permission cleanup",
      forceNew: true,
    }),
  }) as { sessionId: string };

  assert.equal(typeof reopened.sessionId, "string");
  await assert.rejects(permissionPromise, /Session disposed/);
});
