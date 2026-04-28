import { randomUUID } from "node:crypto";
import { realpath } from "node:fs/promises";
import { isAbsolute, resolve, relative, sep } from "node:path";
import {
  createFindTool,
  createGrepTool,
  createLsTool,
  createReadTool,
} from "@mariozechner/pi-coding-agent";

type FileToolName = "read" | "grep" | "find" | "ls";

function isInsideRoot(rootDir: string, candidate: string): boolean {
  const relativePath = relative(rootDir, candidate);
  return (
    relativePath === "" ||
    (relativePath !== ".." && !relativePath.startsWith(`..${sep}`) && !isAbsolute(relativePath))
  );
}

async function normalizePathForTool(rootDir: string, candidate: unknown): Promise<string | undefined> {
  if (typeof candidate !== "string" || !candidate.trim()) {
    return undefined;
  }

  const requested = candidate.trim();
  const resolvedRoot = await realpath(rootDir);
  const resolvedPath = resolve(resolvedRoot, requested);
  if (!isInsideRoot(resolvedRoot, resolvedPath)) {
    throw new Error("Requested path must stay inside the saved document folder.");
  }

  let realTarget: string;
  try {
    realTarget = await realpath(resolvedPath);
  } catch {
    throw new Error("Requested path must exist inside the saved document folder.");
  }
  if (!isInsideRoot(resolvedRoot, realTarget)) {
    throw new Error("Requested path must stay inside the saved document folder after resolving symlinks.");
  }

  const relativePath = relative(resolvedRoot, realTarget);
  return relativePath === "" ? "." : relativePath.split(sep).join("/");
}

async function sanitizeParams(
  rootDir: string,
  toolName: FileToolName,
  raw: Record<string, unknown>,
): Promise<Record<string, unknown>> {
  const next = { ...raw };
  const requestedPath = "path" in next ? next.path : next.file_path;
  if (requestedPath !== undefined) {
    const normalized = await normalizePathForTool(rootDir, requestedPath);
    if (normalized) {
      next.path = normalized;
      delete next.file_path;
    } else {
      delete next.path;
      delete next.file_path;
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

  return tool.execute(randomUUID(), await sanitizeParams(rootDir, toolName, params) as never);
}
