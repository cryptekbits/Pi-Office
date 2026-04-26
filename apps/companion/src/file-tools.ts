import { randomUUID } from "node:crypto";
import { resolve, relative, sep } from "node:path";
import {
  createFindTool,
  createGrepTool,
  createLsTool,
  createReadTool,
} from "@mariozechner/pi-coding-agent";

type FileToolName = "read" | "grep" | "find" | "ls";

function normalizePathForTool(rootDir: string, candidate: unknown): string | undefined {
  if (typeof candidate !== "string" || !candidate.trim()) {
    return undefined;
  }

  const requested = candidate.trim();
  const resolvedRoot = resolve(rootDir);
  const resolvedPath = resolve(resolvedRoot, requested);
  const relativePath = relative(resolvedRoot, resolvedPath);
  if (
    relativePath.startsWith("..") ||
    relativePath.includes(`..${sep}`) ||
    relativePath === ".." ||
    resolve(resolvedRoot, relativePath) !== resolvedPath
  ) {
    throw new Error("Requested path must stay inside the saved document folder.");
  }

  return relativePath === "" ? "." : relativePath;
}

function sanitizeParams(rootDir: string, toolName: FileToolName, raw: Record<string, unknown>): Record<string, unknown> {
  const next = { ...raw };
  if ("path" in next) {
    const normalized = normalizePathForTool(rootDir, next.path);
    if (normalized) {
      next.path = normalized;
    } else {
      delete next.path;
    }
  }

  if (toolName === "read" && typeof next.path !== "string") {
    throw new Error("read requires a path inside the saved document folder.");
  }

  return next;
}

export async function executeFileTool(
  rootDir: string | undefined,
  toolName: FileToolName,
  params: Record<string, unknown>,
): Promise<unknown> {
  if (!rootDir) {
    throw new Error("Read-only local file access is only available for saved documents.");
  }

  const tool =
    toolName === "read"
      ? createReadTool(rootDir)
      : toolName === "grep"
        ? createGrepTool(rootDir)
        : toolName === "find"
          ? createFindTool(rootDir)
          : createLsTool(rootDir);

  return tool.execute(randomUUID(), sanitizeParams(rootDir, toolName, params) as never);
}
