import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import test from "node:test";

import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
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

test("protocol inventory includes first-class PowerPoint chart/media/icon and verification tools", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("edit_slide_chart" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("copy_image_between_slides" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("search_icons" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("insert_icon" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("verify_slides" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("verify_slide_visual" as any));

  assert.equal(TOOL_CATEGORY_MAP.edit_slide_chart, "write-doc");
  assert.equal(TOOL_CATEGORY_MAP.copy_image_between_slides, "write-doc");
  assert.equal(TOOL_CATEGORY_MAP.search_icons, "read");
  assert.equal(TOOL_CATEGORY_MAP.insert_icon, "write-doc");
  assert.equal(TOOL_CATEGORY_MAP.verify_slides, "read");
  assert.equal(TOOL_CATEGORY_MAP.verify_slide_visual, "read");
});

test("createOfficeExtension registers first-class PowerPoint chart/media/icon and verification tools", async () => {
  const registeredTools = new Map<string, {
    execute: (toolCallId: string, params: Record<string, unknown>) => Promise<{ content: unknown; details: unknown }>;
  }>();
  const invocations: Array<{ toolName: string; params: Record<string, unknown> }> = [];

  const extensionFactory = createOfficeExtension({
    getHost: () => "powerpoint",
    getState: () => undefined,
    invokeTool: async (toolName, params) => {
      invocations.push({ toolName, params });
      if (toolName === "verify_slides") {
        return {
          summary: "PowerPoint structural verification complete.",
          details: { kind: "powerpoint-slide-verification", mutating: false },
        };
      }
      if (toolName === "verify_slide_visual") {
        return {
          summary: "PowerPoint visual verification complete.",
          visual: { kind: "powerpoint-slide-snapshot" },
          details: { kind: "powerpoint-slide-visual-verification", mutating: false },
          visuals: [{ kind: "slide", data: "ZmFrZQ==", mimeType: "image/png" }],
        };
      }
      return { ok: true, tool: toolName, params };
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
    "edit_slide_chart",
    "copy_image_between_slides",
    "search_icons",
    "insert_icon",
    "verify_slides",
    "verify_slide_visual",
  ];
  for (const toolName of powerPointToolNames) {
    assert.ok(registeredTools.has(toolName));
  }

  await registeredTools.get("edit_slide_chart")!.execute("ppt-chart-tool", {
    operation: "update_slide_chart",
    slideId: "slide-2",
    shapeId: "shape-chart-1",
    title: "Q4 Pipeline",
  });
  await registeredTools.get("copy_image_between_slides")!.execute("ppt-copy-image-tool", {
    sourceSlideId: "slide-1",
    sourceShapeId: "shape-image-1",
    targetSlideId: "slide-2",
    targetShapeId: "shape-image-2",
  });
  await registeredTools.get("search_icons")!.execute("ppt-search-icon-tool", {
    query: "growth",
    maxResults: 5,
  });
  await registeredTools.get("insert_icon")!.execute("ppt-insert-icon-tool", {
    iconId: "trend-up",
    slideId: "slide-2",
  });
  const verifySlidesResult = await registeredTools.get("verify_slides")!.execute("ppt-verify-structural", {
    maxSlides: 8,
  });
  const verifyVisualResult = await registeredTools.get("verify_slide_visual")!.execute("ppt-verify-visual", {
    includeFormatting: true,
    maxImages: 1,
  });

  assert.equal(invocations[0]?.toolName, "edit_slide_chart");
  assert.equal(invocations[1]?.toolName, "copy_image_between_slides");
  assert.equal(invocations[2]?.toolName, "search_icons");
  assert.equal(invocations[3]?.toolName, "insert_icon");
  assert.equal(invocations[4]?.toolName, "verify_slides");
  assert.equal(invocations[5]?.toolName, "verify_slide_visual");
  assert.equal((verifySlidesResult.details as { details: { kind: string } }).details.kind, "powerpoint-slide-verification");
  assert.equal((verifyVisualResult.details as { details: { kind: string } }).details.kind, "powerpoint-slide-visual-verification");
});

test("in-process runtime publishes first-class PowerPoint chart/media/icon and verification tools", async () => {
  const runtime = await loadKernelModule();
  const openResponse = await runtime.dispatchKernelRequest("/v1/sessions/open", {
    method: "POST",
    body: JSON.stringify({
      host: "powerpoint",
      documentId: `doc-ppt-verify-media-${Date.now()}`,
      saved: true,
      title: "PowerPoint Verify/Media Tool Inventory Test",
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

  assert.ok(toolNames.includes("edit_slide_chart"));
  assert.ok(toolNames.includes("copy_image_between_slides"));
  assert.ok(toolNames.includes("search_icons"));
  assert.ok(toolNames.includes("insert_icon"));
  assert.ok(toolNames.includes("verify_slides"));
  assert.ok(toolNames.includes("verify_slide_visual"));
  socket.close();
});

test("createOfficeToolExecutor dispatches PowerPoint chart/media/icon tools and returns structured verification payloads", async () => {
  const actionCalls: Array<{ host: string; action: unknown }> = [];
  const contextCalls: Array<{ host: string; options: unknown }> = [];

  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async (host, options) => {
      contextCalls.push({ host, options });
      return {
        summary: "PowerPoint context captured",
        snippets: {
          selectedSlides: [{ id: "slide-2", index: 2 }],
          selectedShapeDescriptors: [{ id: "shape-3", name: "Revenue Chart" }],
        },
        formatting: {
          selectedSlides: [{ index: 2, layoutName: "Title and Content" }],
        },
        visuals: [{ kind: "slide", data: "ZmFrZQ==", mimeType: "image/png" }],
      };
    },
    applyHostAction: async (host, action) => {
      actionCalls.push({ host, action });
      if ((action as { type?: string }).type === "getPresentationStructure") {
        return {
          ok: true,
          host,
          action: "getPresentationStructure",
          slideCount: 12,
          selectedSlideCount: 1,
          slides: [{ slideId: "slide-2", slideIndex: 2, label: "Slide 2" }],
          slidePreviews: [{ slideId: "slide-2", title: "Revenue Overview" }],
        };
      }
      if ((action as { type?: string }).type === "searchIcons") {
        return {
          ok: true,
          host,
          action: "searchIcons",
          query: "growth",
          icons: [{ id: "trend-up", name: "Trend Up" }],
        };
      }
      return { ok: true, host, action };
    },
    navigateOfficeAnchor: async () => ({ ok: true }),
    readDocumentSection: async () => ({ ok: true }),
    executeOfficeJs: async () => ({ ok: true }),
    proposeEdits: async () => ({ ok: true }),
  });

  const chartResult = await executeOfficeTool({
    requestId: "ppt-chart-dispatch",
    toolName: "edit_slide_chart" as OfficeToolRequest["toolName"],
    host: "powerpoint",
    params: {
      operation: "update_slide_chart",
      slideId: "slide-2",
      shapeId: "shape-chart-1",
      title: "Updated chart title",
    },
  } as OfficeToolRequest);
  const mediaResult = await executeOfficeTool({
    requestId: "ppt-media-dispatch",
    toolName: "copy_image_between_slides" as OfficeToolRequest["toolName"],
    host: "powerpoint",
    params: {
      sourceSlideId: "slide-1",
      sourceShapeId: "shape-image-1",
      targetSlideId: "slide-2",
      targetShapeId: "shape-image-2",
    },
  } as OfficeToolRequest);
  const searchResult = await executeOfficeTool({
    requestId: "ppt-icon-search-dispatch",
    toolName: "search_icons" as OfficeToolRequest["toolName"],
    host: "powerpoint",
    params: {
      query: "growth",
      maxResults: 6,
    },
  } as OfficeToolRequest);
  const insertIconResult = await executeOfficeTool({
    requestId: "ppt-icon-insert-dispatch",
    toolName: "insert_icon" as OfficeToolRequest["toolName"],
    host: "powerpoint",
    params: {
      iconId: "trend-up",
      targetSlideId: "slide-2",
    },
  } as OfficeToolRequest);
  const verifySlidesResult = await executeOfficeTool({
    requestId: "ppt-verify-slides-dispatch",
    toolName: "verify_slides" as OfficeToolRequest["toolName"],
    host: "powerpoint",
    params: {
      maxSlides: 8,
      includeSlideText: true,
    },
  } as OfficeToolRequest);
  const verifyVisualResult = await executeOfficeTool({
    requestId: "ppt-verify-visual-dispatch",
    toolName: "verify_slide_visual" as OfficeToolRequest["toolName"],
    host: "powerpoint",
    params: {
      includeFormatting: true,
      maxImages: 1,
    },
  } as OfficeToolRequest);
  const verifyVisualUnsupported = await executeOfficeTool({
    requestId: "ppt-verify-visual-unsupported",
    toolName: "verify_slide_visual" as OfficeToolRequest["toolName"],
    host: "word",
    params: {},
  } as OfficeToolRequest);

  assert.equal(chartResult.success, true);
  assert.equal(mediaResult.success, true);
  assert.equal(searchResult.success, true);
  assert.equal(insertIconResult.success, true);
  assert.equal(verifySlidesResult.success, true);
  assert.equal(verifyVisualResult.success, true);
  assert.equal((actionCalls[0]?.action as { type: string }).type, "updateSlideChart");
  assert.equal((actionCalls[1]?.action as { type: string }).type, "copyImageBetweenSlides");
  assert.equal((actionCalls[2]?.action as { type: string }).type, "searchIcons");
  assert.equal((actionCalls[3]?.action as { type: string }).type, "insertIcon");
  assert.equal((actionCalls[4]?.action as { type: string }).type, "getPresentationStructure");
  const verifySlidesPayload = verifySlidesResult.content as { details: { kind: string; mutating: boolean } };
  assert.equal(verifySlidesPayload.details.kind, "powerpoint-slide-verification");
  assert.equal(verifySlidesPayload.details.mutating, false);
  assert.deepEqual(contextCalls[0], {
    host: "powerpoint",
    options: {
      includeFormatting: true,
      maxImages: 1,
      scope: "slide",
    },
  });
  const verifyVisualPayload = verifyVisualResult.content as {
    visual: { kind: string; imageCount: number };
    details: { kind: string; mutating: boolean };
    visuals: unknown[];
  };
  assert.equal(verifyVisualPayload.visual.kind, "powerpoint-slide-snapshot");
  assert.equal(verifyVisualPayload.visual.imageCount, 1);
  assert.equal(verifyVisualPayload.details.kind, "powerpoint-slide-visual-verification");
  assert.equal(verifyVisualPayload.details.mutating, false);
  assert.equal(Array.isArray(verifyVisualPayload.visuals), true);
  assert.equal(verifyVisualUnsupported.success, false);
  assert.match(String(verifyVisualUnsupported.error), /only available for PowerPoint/);
});

test("PowerPoint guidance aligns chart/media/icon and verification tools with supported runtime paths", () => {
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bedit_slide_chart\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bcopy_image_between_slides\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bsearch_icons\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\binsert_icon\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bverify_slides\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bverify_slide_visual\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /serialized|XML|OOXML/i);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /layout\/master|layout and master/i);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /slide snapshot|visual verification/i);

  const officeHostSkillPath = join(process.cwd(), "packages", "pi-office-pack", "skills", "office-host.SKILL.md");
  const officeHostSkillText = readFileSync(officeHostSkillPath, "utf8");
  assert.match(officeHostSkillText, /\bedit_slide_chart\b/);
  assert.match(officeHostSkillText, /\bcopy_image_between_slides\b/);
  assert.match(officeHostSkillText, /\bsearch_icons\b/);
  assert.match(officeHostSkillText, /\binsert_icon\b/);
  assert.match(officeHostSkillText, /\bverify_slides\b/);
  assert.match(officeHostSkillText, /\bverify_slide_visual\b/);
  assert.match(officeHostSkillText, /serialized|XML|OOXML/i);
  assert.match(officeHostSkillText, /layout\/master|layout and master/i);
  assert.match(officeHostSkillText, /slide snapshot|visual verification/i);
});

test("investigation artifact includes PowerPoint manual checklist scenarios for structure, authoring, notes/charts, and visual verification coverage", () => {
  const investigationPath = join(process.cwd(), "CLAUDE_ADDIN_INVESTIGATION_AND_TRACKING.md");
  const investigationText = readFileSync(investigationPath, "utf8");
  assert.match(investigationText, /Manual Testing Scenarios and Steps/i);
  assert.match(investigationText, /powerpoint-anchor-navigation/);
  assert.match(investigationText, /powerpoint-slide-and-shape-authoring/);
  assert.match(investigationText, /powerpoint-notes-and-charts/);
  assert.match(investigationText, /powerpoint-taskpane-stability/);
});
