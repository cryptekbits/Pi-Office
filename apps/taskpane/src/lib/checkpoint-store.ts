import type { DocumentCheckpointPayload } from "@pi-office/pi-office-pack/protocol";

export interface DocumentCheckpoint extends DocumentCheckpointPayload {
  hasDocumentData: boolean;
}

const MAX_CHECKPOINTS = 20;
const MAX_PPT_CHECKPOINTS = 5;

const store = new Map<string, DocumentCheckpoint>();

export function addCheckpoint(cp: DocumentCheckpoint): void {
  store.set(cp.id, cp);
  pruneStore(cp.host);
}

export function getCheckpoint(id: string): DocumentCheckpoint | undefined {
  return store.get(id);
}

export function hasCheckpoint(id: string): boolean {
  return store.has(id);
}

export function clearCheckpoints(): void {
  store.clear();
}

export function getAllCheckpoints(): DocumentCheckpoint[] {
  return [...store.values()].sort((a, b) => b.timestamp - a.timestamp);
}

function pruneStore(host: string): void {
  const max = host === "powerpoint" ? MAX_PPT_CHECKPOINTS : MAX_CHECKPOINTS;
  if (store.size <= max) return;
  const sorted = [...store.entries()].sort(([, a], [, b]) => a.timestamp - b.timestamp);
  while (sorted.length > max) {
    const oldest = sorted.shift();
    if (oldest) store.delete(oldest[0]);
  }
}
