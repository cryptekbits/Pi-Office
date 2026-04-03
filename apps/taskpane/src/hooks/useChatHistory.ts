import { useCallback, useState } from "react";
import type { OfficeStateUpdate } from "@pi-office/pi-office-pack/protocol";
import type { ChatEntry } from "../lib/helpers";

const STORAGE_KEY = "pi-office-chat-history";
const MAX_ENTRIES = 50;
const MAX_MESSAGES_PER_CHAT = 100;

export type HistoryScope = "document" | "workspace" | "global";

export interface ChatHistoryEntry {
  chatId: string;
  subject: string | undefined;
  messages: ChatEntry[];
  documentId: string;
  documentTitle: string;
  host: string;
  workspaceDir: string | undefined;
  createdAt: string;
  updatedAt: string;
}

function loadHistory(): ChatHistoryEntry[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return [];
    return JSON.parse(raw) as ChatHistoryEntry[];
  } catch {
    return [];
  }
}

function persistHistory(entries: ChatHistoryEntry[]): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(entries));
  } catch {
    // quota exceeded
  }
}

function matchesScope(
  entry: ChatHistoryEntry,
  scope: HistoryScope,
  officeState: OfficeStateUpdate | undefined,
): boolean {
  if (scope === "global") return true;
  if (!officeState) return false;
  if (scope === "document") {
    return entry.host === officeState.host && entry.documentId === officeState.document.id;
  }
  if (scope === "workspace") {
    return Boolean(
      officeState.document.workspaceDir &&
        entry.workspaceDir === officeState.document.workspaceDir,
    );
  }
  return false;
}

export function useChatHistory() {
  const [entries, setEntries] = useState<ChatHistoryEntry[]>(loadHistory);

  const saveChat = useCallback(
    (
      chatId: string,
      messages: ChatEntry[],
      subject: string | undefined,
      officeState: OfficeStateUpdate | undefined,
    ): string => {
      const now = new Date().toISOString();
      setEntries((current) => {
        const existingIndex = current.findIndex((e) => e.chatId === chatId);
        const trimmedMessages = messages.slice(-MAX_MESSAGES_PER_CHAT);
        const entry: ChatHistoryEntry = {
          chatId,
          subject,
          messages: trimmedMessages,
          documentId: officeState?.document.id ?? "",
          documentTitle: officeState?.document.title ?? "Untitled",
          host: officeState?.host ?? "word",
          workspaceDir: officeState?.document.workspaceDir,
          createdAt: existingIndex >= 0 ? current[existingIndex]!.createdAt : now,
          updatedAt: now,
        };

        let next: ChatHistoryEntry[];
        if (existingIndex >= 0) {
          next = current.map((e, i) => (i === existingIndex ? entry : e));
        } else {
          next = [entry, ...current];
        }

        if (next.length > MAX_ENTRIES) {
          next = next.slice(0, MAX_ENTRIES);
        }

        persistHistory(next);
        return next;
      });
      return chatId;
    },
    [],
  );

  const loadChat = useCallback(
    (chatId: string): ChatHistoryEntry | undefined => {
      return entries.find((e) => e.chatId === chatId);
    },
    [entries],
  );

  const deleteChat = useCallback((chatId: string) => {
    setEntries((current) => {
      const next = current.filter((e) => e.chatId !== chatId);
      persistHistory(next);
      return next;
    });
  }, []);

  const listChats = useCallback(
    (scope: HistoryScope, officeState: OfficeStateUpdate | undefined): ChatHistoryEntry[] => {
      return entries
        .filter((e) => matchesScope(e, scope, officeState))
        .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
    },
    [entries],
  );

  const updateSubject = useCallback((chatId: string, subject: string) => {
    setEntries((current) => {
      const next = current.map((e) =>
        e.chatId === chatId ? { ...e, subject, updatedAt: new Date().toISOString() } : e,
      );
      persistHistory(next);
      return next;
    });
  }, []);

  return { entries, saveChat, loadChat, deleteChat, listChats, updateSubject } as const;
}
