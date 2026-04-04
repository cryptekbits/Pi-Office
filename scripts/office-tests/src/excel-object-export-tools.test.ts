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

test("protocol inventory includes first-class Excel object/data and export/visualization tools", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("modify_object" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("get_all_objects" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("search_data" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("get_range_as_csv" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("read_range_image" as any));
  assert.ok(OFFICE_TOOL_NAMES.includes("extract_chart_xml" as any));

  assert.equal(TOOL_CATEGORY_MAP.modify_object, "write-doc");
  assert.equal(TOOL_CATEGORY_MAP.get_all_objects, "read");
  assert.equal(TOOL_CATEGORY_MAP.search_data, "read");
  assert.equal(TOOL_CATEGORY_MAP.get_range_as_csv, "read");
  assert.equal(TOOL_CATEGORY_MAP.read_range_image, "read");
  assert.equal(TOOL_CATEGORY_MAP.extract_chart_xml, "read");
});

test("createOfficeExtension registers first-class Excel object/data and export/visualization tools", async () => {
  const registeredTools = new Map<string, {
    execute: (toolCallId: string, params: Record<string, unknown>) => Promise<{ content: unknown; details: unknown }>;
  }>();
  const invocations: Array<{ toolName: string; params: Record<string, unknown> }> = [];

  const extensionFactory = createOfficeExtension({
    getHost: () => "excel",
    getState: () => undefined,
    invokeTool: async (toolName, params) => {
      invocations.push({ toolName, params });
      if (toolName === "read_range_image") {
        return {
          summary: "Excel range image captured.",
          visual: { kind: "excel-range-image", imageCount: 1 },
          details: { kind: "excel-range-image-read", mutating: false },
          visuals: [{ kind: "worksheet", data: "ZmFrZQ==", mimeType: "image/png" }],
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

  const excelToolNames = [
    "modify_object",
    "get_all_objects",
    "search_data",
    "get_range_as_csv",
    "read_range_image",
    "extract_chart_xml",
  ];
  for (const toolName of excelToolNames) {
    assert.ok(registeredTools.has(toolName));
  }

  await registeredTools.get("modify_object")!.execute("excel-modify-object", {
    operation: "update_chart",
    chartName: "Revenue",
    title: "Q4 Revenue",
  });
  await registeredTools.get("get_all_objects")!.execute("excel-get-objects", { scope: "workbook" });
  await registeredTools.get("search_data")!.execute("excel-search-data", { query: "Revenue", objectTypes: ["chart", "table"] });
  await registeredTools.get("get_range_as_csv")!.execute("excel-export-csv", { sheetName: "Budget", address: "A1:B3" });
  const rangeImageResult = await registeredTools.get("read_range_image")!.execute("excel-range-image", { sheetName: "Budget", address: "A1:B3" });
  await registeredTools.get("extract_chart_xml")!.execute("excel-chart-xml", { sheetName: "Summary", chartName: "Revenue" });

  assert.equal(invocations[0]?.toolName, "modify_object");
  assert.equal(invocations[1]?.toolName, "get_all_objects");
  assert.equal(invocations[2]?.toolName, "search_data");
  assert.equal(invocations[3]?.toolName, "get_range_as_csv");
  assert.equal(invocations[4]?.toolName, "read_range_image");
  assert.equal(invocations[5]?.toolName, "extract_chart_xml");
  assert.equal((rangeImageResult.details as { details: { kind: string } }).details.kind, "excel-range-image-read");
});

test("in-process runtime publishes first-class Excel object/data and export/visualization tools", async () => {
  const runtime = await loadKernelModule();
  const openResponse = await runtime.dispatchKernelRequest("/v1/sessions/open", {
    method: "POST",
    body: JSON.stringify({
      host: "excel",
      documentId: `doc-excel-object-export-${Date.now()}`,
      saved: true,
      title: "Excel Object/Export Tool Inventory Test",
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

  assert.ok(toolNames.includes("modify_object"));
  assert.ok(toolNames.includes("get_all_objects"));
  assert.ok(toolNames.includes("search_data"));
  assert.ok(toolNames.includes("get_range_as_csv"));
  assert.ok(toolNames.includes("read_range_image"));
  assert.ok(toolNames.includes("extract_chart_xml"));
  socket.close();
});

test("createOfficeToolExecutor dispatches first-class Excel object/data and export/visualization tools", async () => {
  const actionCalls: Array<{ host: string; action: unknown }> = [];
  const contextCalls: Array<{ host: string; options: unknown }> = [];

  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async (host, options) => {
      contextCalls.push({ host, options });
      return {
        summary: "Excel context captured.",
        snippets: {
          workbookSheets: [
            {
              name: "Summary",
              tableCount: 1,
              chartCount: 1,
              pivotTableCount: 1,
              tables: ["RevenueTable"],
              charts: ["RevenueChart"],
              pivotTables: ["RevenuePivot"],
            },
          ],
          workbookObjects: {
            tables: [{ sheetName: "Summary", name: "RevenueTable", id: "tbl-1" }],
            charts: [{ sheetName: "Summary", name: "RevenueChart", id: "chart-1" }],
            pivotTables: [{ sheetName: "Summary", name: "RevenuePivot", id: "pivot-1" }],
          },
          namedItems: [{ name: "RevenueTarget", type: "Range" }],
          selectionCellCitations: [
            {
              label: "Summary!A1",
              text: "Revenue",
              formula: "=SUM(B2:B10)",
              anchor: {
                kind: "cell",
                sheetName: "Summary",
                address: "A1",
              },
            },
          ],
        },
        visuals: [{ kind: "worksheet", data: "ZmFrZQ==", mimeType: "image/png" }],
      };
    },
    applyHostAction: async (host, action) => {
      actionCalls.push({ host, action });
      if ((action as { type?: string }).type === "getRangeValues") {
        return {
          ok: true,
          host,
          action: "getRangeValues",
          sheetName: "Summary",
          address: "A1:B3",
          values: [
            ["Region", "Sales"],
            ["North", 120],
            ["South", 95],
          ],
        };
      }
      if ((action as { type?: string }).type === "extractChartXml") {
        return {
          ok: true,
          host,
          action: "extractChartXml",
          chartName: "RevenueChart",
          chartXml: "<chart name=\"RevenueChart\" chartType=\"ColumnClustered\" />",
        };
      }
      return { ok: true, host, action };
    },
    navigateOfficeAnchor: async () => ({ ok: true }),
    readDocumentSection: async () => ({ ok: true }),
    executeOfficeJs: async () => ({ ok: true }),
    proposeEdits: async () => ({ ok: true }),
  });

  const modifyObjectResult = await executeOfficeTool({
    requestId: "excel-modify-object-dispatch",
    toolName: "modify_object" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {
      operation: "update_chart",
      chartName: "RevenueChart",
      title: "Updated Revenue",
    },
  } as OfficeToolRequest);
  const getObjectsResult = await executeOfficeTool({
    requestId: "excel-get-objects-dispatch",
    toolName: "get_all_objects" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {
      scope: "workbook",
    },
  } as OfficeToolRequest);
  const searchDataResult = await executeOfficeTool({
    requestId: "excel-search-data-dispatch",
    toolName: "search_data" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {
      query: "revenue",
      scope: "workbook",
      objectTypes: ["chart", "table", "pivotTable", "namedItem", "cell"],
    },
  } as OfficeToolRequest);
  const exportCsvResult = await executeOfficeTool({
    requestId: "excel-export-csv-dispatch",
    toolName: "get_range_as_csv" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {
      sheetName: "Summary",
      address: "A1:B3",
    },
  } as OfficeToolRequest);
  const rangeImageResult = await executeOfficeTool({
    requestId: "excel-read-image-dispatch",
    toolName: "read_range_image" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {
      scope: "selection",
      includeFormatting: true,
      maxImages: 1,
    },
  } as OfficeToolRequest);
  const extractChartXmlResult = await executeOfficeTool({
    requestId: "excel-chart-xml-dispatch",
    toolName: "extract_chart_xml" as OfficeToolRequest["toolName"],
    host: "excel",
    params: {
      chartName: "RevenueChart",
      sheetName: "Summary",
    },
  } as OfficeToolRequest);
  const unsupportedResult = await executeOfficeTool({
    requestId: "excel-object-unsupported",
    toolName: "modify_object" as OfficeToolRequest["toolName"],
    host: "word",
    params: {},
  } as OfficeToolRequest);

  assert.equal(modifyObjectResult.success, true);
  assert.equal(getObjectsResult.success, true);
  assert.equal(searchDataResult.success, true);
  assert.equal(exportCsvResult.success, true);
  assert.equal(rangeImageResult.success, true);
  assert.equal(extractChartXmlResult.success, true);
  assert.equal((actionCalls[0]?.action as { type: string }).type, "updateChart");
  assert.equal((actionCalls[1]?.action as { type: string }).type, "getRangeValues");
  assert.equal((actionCalls[2]?.action as { type: string }).type, "extractChartXml");

  const objectPayload = getObjectsResult.content as { details: { kind: string; mutating: boolean; matchCount: number } };
  assert.equal(objectPayload.details.kind, "excel-object-inventory");
  assert.equal(objectPayload.details.mutating, false);
  assert.equal(objectPayload.details.matchCount >= 3, true);

  const searchPayload = searchDataResult.content as { details: { kind: string; query: string; mutating: boolean; matchCount: number } };
  assert.equal(searchPayload.details.kind, "excel-data-search");
  assert.equal(searchPayload.details.query, "revenue");
  assert.equal(searchPayload.details.mutating, false);
  assert.equal(searchPayload.details.matchCount >= 1, true);

  const csvPayload = exportCsvResult.content as { csv: string; details: { kind: string; mutating: boolean } };
  assert.equal(csvPayload.details.kind, "excel-range-csv-export");
  assert.equal(csvPayload.details.mutating, false);
  assert.match(csvPayload.csv, /Region,Sales/);
  assert.match(csvPayload.csv, /North,120/);

  const imagePayload = rangeImageResult.content as {
    visual: { kind: string; imageCount: number };
    details: { kind: string; mutating: boolean };
    visuals: unknown[];
  };
  assert.equal(imagePayload.visual.kind, "excel-range-image");
  assert.equal(imagePayload.visual.imageCount, 1);
  assert.equal(imagePayload.details.kind, "excel-range-image-read");
  assert.equal(imagePayload.details.mutating, false);
  assert.equal(Array.isArray(imagePayload.visuals), true);
  assert.deepEqual(contextCalls[0], {
    host: "excel",
    options: {
      includeFormatting: true,
      maxImages: 0,
      scope: "workbook",
    },
  });
  assert.deepEqual(contextCalls[1], {
    host: "excel",
    options: {
      includeFormatting: false,
      maxImages: 0,
      scope: "workbook",
    },
  });
  assert.deepEqual(contextCalls[2], {
    host: "excel",
    options: {
      includeFormatting: true,
      maxImages: 1,
      scope: "selection",
    },
  });

  const xmlPayload = extractChartXmlResult.content as {
    xml: string;
    details: { kind: string; mutating: boolean; chartName?: string };
  };
  assert.equal(xmlPayload.details.kind, "excel-chart-xml");
  assert.equal(xmlPayload.details.mutating, false);
  assert.equal(xmlPayload.details.chartName, "RevenueChart");
  assert.match(xmlPayload.xml, /RevenueChart/);

  assert.equal(unsupportedResult.success, false);
  assert.match(String(unsupportedResult.error), /only available for Excel/);
});

test("Excel guidance aligns formula-first, auditable-cell expectations with runtime tools", () => {
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bmodify_object\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bget_all_objects\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bsearch_data\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bget_range_as_csv\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bread_range_image\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bextract_chart_xml\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /formula-first/i);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /auditable[-\s]cell/i);

  const officeHostSkillPath = join(process.cwd(), "packages", "pi-office-pack", "skills", "office-host.SKILL.md");
  const officeHostSkillText = readFileSync(officeHostSkillPath, "utf8");
  assert.match(officeHostSkillText, /\bmodify_object\b/);
  assert.match(officeHostSkillText, /\bget_all_objects\b/);
  assert.match(officeHostSkillText, /\bsearch_data\b/);
  assert.match(officeHostSkillText, /\bget_range_as_csv\b/);
  assert.match(officeHostSkillText, /\bread_range_image\b/);
  assert.match(officeHostSkillText, /\bextract_chart_xml\b/);
  assert.match(officeHostSkillText, /formula-first/i);
  assert.match(officeHostSkillText, /auditable[-\s]cell/i);
});

test("investigation artifact includes Excel manual checklist scenarios for workbook context, formatting/tables, charts/pivots, and taskpane stability", () => {
  const investigationPath = join(process.cwd(), "CLAUDE_ADDIN_INVESTIGATION_AND_TRACKING.md");
  const investigationText = readFileSync(investigationPath, "utf8");
  assert.match(investigationText, /Manual Testing Scenarios and Steps/i);
  assert.match(investigationText, /excel-context-citations/);
  assert.match(investigationText, /excel-formatting-and-tables/);
  assert.match(investigationText, /excel-charts-and-pivots/);
  assert.match(investigationText, /excel-taskpane-stability/);
});
