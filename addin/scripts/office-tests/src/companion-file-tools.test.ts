import assert from "node:assert/strict";
import { mkdtemp, mkdir, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import test from "node:test";
import { executeFileTool } from "../../../../companion/src/file-tools.js";

async function withWorkspace<T>(callback: (paths: {
  rootDir: string;
  outsideDir: string;
  insideFile: string;
  dotDotNamedFile: string;
  outsideFile: string;
}) => Promise<T>): Promise<T> {
  const baseDir = await mkdtemp(join(tmpdir(), "pi-office-file-tools-"));
  const rootDir = join(baseDir, "document-folder");
  const outsideDir = join(baseDir, "outside");
  const insideFile = join(rootDir, "notes.txt");
  const dotDotNamedFile = join(rootDir, "..not-secret.txt");
  const outsideFile = join(outsideDir, "secret.txt");
  await mkdir(rootDir, { recursive: true });
  await mkdir(outsideDir, { recursive: true });
  await writeFile(insideFile, "inside document folder\n", "utf8");
  await writeFile(dotDotNamedFile, "odd but valid name\n", "utf8");
  await writeFile(outsideFile, "outside secret\n", "utf8");
  try {
    return await callback({ rootDir, outsideDir, insideFile, dotDotNamedFile, outsideFile });
  } finally {
    await rm(baseDir, { recursive: true, force: true });
  }
}

test("companion read-only file tools stay bound to the saved document folder", async () => {
  await withWorkspace(async ({ rootDir, insideFile, outsideFile }) => {
    const inside = await executeFileTool(rootDir, "read", { path: "notes.txt" });
    assert.match(JSON.stringify(inside), /inside document folder/);

    const absoluteInside = await executeFileTool(rootDir, "read", { path: insideFile });
    assert.match(JSON.stringify(absoluteInside), /inside document folder/);

    await assert.rejects(
      executeFileTool(rootDir, "read", { path: "../outside/secret.txt" }),
      /inside the saved document folder/i,
    );
    await assert.rejects(
      executeFileTool(rootDir, "read", { path: outsideFile }),
      /inside the saved document folder/i,
    );
    await assert.rejects(
      executeFileTool(undefined, "ls", {}),
      /only available for saved documents/i,
    );
  });
});

test("companion file path guard allows valid names that only look like traversal", async () => {
  await withWorkspace(async ({ rootDir, dotDotNamedFile }) => {
    const relativeResult = await executeFileTool(rootDir, "read", { path: "..not-secret.txt" });
    assert.match(JSON.stringify(relativeResult), /odd but valid name/);

    const absoluteResult = await executeFileTool(rootDir, "read", { file_path: dotDotNamedFile });
    assert.match(JSON.stringify(absoluteResult), /odd but valid name/);
  });
});

test("companion file path guard rejects symlink or junction escapes when the platform supports them", async (t) => {
  await withWorkspace(async ({ rootDir, outsideDir, outsideFile }) => {
    const linkPath = resolve(rootDir, "linked-secret.txt");
    try {
      await symlink(outsideFile, linkPath, "file");
    } catch (error) {
      const junctionPath = resolve(rootDir, "linked-outside");
      try {
        await symlink(outsideDir, junctionPath, "junction");
      } catch {
        t.skip(`Symlink and junction escapes are unavailable in this test environment: ${error instanceof Error ? error.message : String(error)}`);
        return;
      }
      await assert.rejects(
        executeFileTool(rootDir, "read", { path: "linked-outside/secret.txt" }),
        /resolving symlinks/i,
      );
      return;
    }

    await assert.rejects(
      executeFileTool(rootDir, "read", { path: "linked-secret.txt" }),
      /resolving symlinks/i,
    );
  });
});
