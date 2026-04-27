import type {
  McpResultClearRequest,
  McpResultClearResponse,
  McpResultHandle,
  McpResultPageRequest,
  McpResultPageResponse,
  McpResultSummarizeRequest,
  McpResultSummarizeResponse,
} from "@pi-office/pi-office-pack/protocol";

const DEFAULT_PAGE_SIZE = 8_000;
const DEFAULT_TTL_MS = 30 * 60 * 1000;

interface StoredMcpResult {
  handle: McpResultHandle;
  text: string;
}

function toText(value: unknown): string {
  if (typeof value === "string") return value;
  try {
    return JSON.stringify(value, null, 2);
  } catch {
    return String(value);
  }
}

function summarize(text: string, maxChars = 700): string {
  const compact = text.replace(/\s+/g, " ").trim();
  return compact.length > maxChars ? `${compact.slice(0, maxChars)}…` : compact;
}

function byteSize(text: string): number {
  return new TextEncoder().encode(text).byteLength;
}

export class McpResultStore {
  private readonly results = new Map<string, StoredMcpResult>();

  store(input: {
    source: "browser" | "companion";
    connectorId?: string | undefined;
    connectorName?: string | undefined;
    toolName?: string | undefined;
    value: unknown;
    pageSizeBytes?: number | undefined;
  }): { handle: McpResultHandle; content: unknown } {
    const text = toText(input.value);
    const sizeBytes = byteSize(text);
    const pageSizeBytes = Math.max(1024, input.pageSizeBytes ?? DEFAULT_PAGE_SIZE);
    if (sizeBytes <= pageSizeBytes) {
      return { handle: this.createHandle(input, text, sizeBytes, pageSizeBytes, 1), content: input.value };
    }

    const handle = this.createHandle(
      input,
      text,
      sizeBytes,
      pageSizeBytes,
      Math.max(1, Math.ceil(text.length / pageSizeBytes)),
    );
    this.results.set(handle.handleId, { handle, text });
    return {
      handle,
      content: {
        summary: handle.summary,
        resultHandle: handle,
        note: "Large MCP result cached locally. Use mcp_result_get for pages or mcp_result_summarize for targeted extracts.",
      },
    };
  }

  getPage(request: McpResultPageRequest): McpResultPageResponse {
    const stored = this.getStored(request.handleId);
    const page = Math.max(0, Math.trunc(request.page ?? 0));
    const start = page * stored.handle.pageSizeBytes;
    const end = start + stored.handle.pageSizeBytes;
    return {
      ok: true,
      handle: stored.handle,
      page,
      content: stored.text.slice(start, end),
      hasNextPage: end < stored.text.length,
    };
  }

  summarize(request: McpResultSummarizeRequest): McpResultSummarizeResponse {
    const stored = this.getStored(request.handleId);
    const maxChars = Math.max(200, Math.min(4_000, request.maxChars ?? 1_000));
    const query = request.query?.trim().toLowerCase();
    const source = query
      ? stored.text
          .split(/\r?\n/)
          .filter((line) => line.toLowerCase().includes(query))
          .join("\n") || stored.text
      : stored.text;
    return {
      ok: true,
      handle: stored.handle,
      summary: summarize(source, maxChars),
    };
  }

  clear(request: McpResultClearRequest = {}): McpResultClearResponse {
    if (request.handleId) {
      const cleared = this.results.delete(request.handleId) ? 1 : 0;
      return { ok: true, cleared };
    }
    const cleared = this.results.size;
    this.results.clear();
    return { ok: true, cleared };
  }

  private createHandle(
    input: {
      source: "browser" | "companion";
      connectorId?: string | undefined;
      connectorName?: string | undefined;
      toolName?: string | undefined;
    },
    text: string,
    sizeBytes: number,
    pageSizeBytes: number,
    pageCount: number,
  ): McpResultHandle {
    const now = Date.now();
    return {
      handleId: `mcp-result-${now}-${Math.random().toString(36).slice(2, 10)}`,
      source: input.source,
      connectorId: input.connectorId,
      connectorName: input.connectorName,
      toolName: input.toolName,
      createdAt: new Date(now).toISOString(),
      expiresAt: new Date(now + DEFAULT_TTL_MS).toISOString(),
      sizeBytes,
      pageSizeBytes,
      pageCount,
      summary: summarize(text),
      redacted: false,
    };
  }

  private getStored(handleId: string): StoredMcpResult {
    const stored = this.results.get(handleId);
    if (!stored) {
      throw new Error(`Unknown or expired MCP result handle: ${handleId}`);
    }
    return stored;
  }
}
