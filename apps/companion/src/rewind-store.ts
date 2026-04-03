import { mkdirSync, readFileSync, writeFileSync, readdirSync, rmSync, existsSync } from "node:fs";
import { join } from "node:path";
import type { CompanionConfig } from "./config.js";
import type { DocumentCheckpointPayload, CheckpointMetadata } from "@pi-office/pi-office-pack/protocol";

const MAX_CHECKPOINT_AGE_MS = 3 * 24 * 60 * 60 * 1000; // 3 days
const MAX_CHECKPOINTS_PER_DOC = 50;

function snapshotDir(config: CompanionConfig, documentId: string): string {
  return join(config.scratchDir, "rewind-snapshots", sanitizeId(documentId));
}

function sanitizeId(id: string): string {
  return id.replace(/[^a-zA-Z0-9_-]/g, "_").slice(0, 120);
}

function manifestPath(dir: string): string {
  return join(dir, "manifest.json");
}

interface Manifest {
  documentId: string;
  documentTitle?: string;
  host: string;
  checkpoints: CheckpointMetadata[];
}

function readManifest(dir: string): Manifest | null {
  const p = manifestPath(dir);
  if (!existsSync(p)) return null;
  try {
    return JSON.parse(readFileSync(p, "utf-8"));
  } catch {
    return null;
  }
}

function writeManifest(dir: string, manifest: Manifest): void {
  writeFileSync(manifestPath(dir), JSON.stringify(manifest, null, 2), "utf-8");
}

export function saveCheckpointToDisk(
  config: CompanionConfig,
  documentId: string,
  checkpoint: DocumentCheckpointPayload,
): void {
  const dir = snapshotDir(config, documentId);
  mkdirSync(dir, { recursive: true });

  const fileName = `cp-${checkpoint.timestamp}.json`;
  writeFileSync(join(dir, fileName), JSON.stringify(checkpoint), "utf-8");

  const manifest = readManifest(dir) ?? {
    documentId,
    host: checkpoint.host,
    checkpoints: [],
  };
  manifest.checkpoints.push({
    id: checkpoint.id,
    timestamp: checkpoint.timestamp,
    userPrompt: checkpoint.userPrompt,
    file: fileName,
  });
  writeManifest(dir, manifest);
}

export function loadCheckpointsFromDisk(
  config: CompanionConfig,
  documentId: string,
): CheckpointMetadata[] {
  const dir = snapshotDir(config, documentId);
  const manifest = readManifest(dir);
  return manifest?.checkpoints ?? [];
}

export function loadCheckpointData(
  config: CompanionConfig,
  documentId: string,
  checkpointId: string,
): DocumentCheckpointPayload | null {
  const dir = snapshotDir(config, documentId);
  const manifest = readManifest(dir);
  if (!manifest) return null;
  const entry = manifest.checkpoints.find((cp) => cp.id === checkpointId);
  if (!entry) return null;
  try {
    return JSON.parse(readFileSync(join(dir, entry.file), "utf-8"));
  } catch {
    return null;
  }
}

export function pruneOldCheckpoints(config: CompanionConfig, documentId: string): number {
  const dir = snapshotDir(config, documentId);
  const manifest = readManifest(dir);
  if (!manifest) return 0;

  const now = Date.now();
  const before = manifest.checkpoints.length;
  manifest.checkpoints = manifest.checkpoints
    .filter((cp) => now - cp.timestamp < MAX_CHECKPOINT_AGE_MS)
    .slice(-MAX_CHECKPOINTS_PER_DOC);
  const removed = before - manifest.checkpoints.length;

  if (removed > 0) {
    const keepFiles = new Set(manifest.checkpoints.map((cp) => cp.file));
    keepFiles.add("manifest.json");
    try {
      for (const f of readdirSync(dir)) {
        if (!keepFiles.has(f)) {
          try { rmSync(join(dir, f)); } catch { /* ignore */ }
        }
      }
    } catch { /* ignore */ }
    writeManifest(dir, manifest);
  }

  return removed;
}
