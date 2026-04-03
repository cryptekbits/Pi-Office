import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import type {
  AuthStatusResponse,
  BridgeServerMessage,
  ConnectorAuditPreference,
  ConnectorAuditPreferenceResponse,
  ConnectorCatalogResponse,
  ConnectorCatalogItem,
  ConnectorDiagnosticsResponse,
  ConnectorExportBundle,
  ConnectorFavoriteRequest,
  ConnectorImportApplyResponse,
  ConnectorImportPreviewResponse,
  ConnectorLogResponse,
  ConnectorPrepareResponse,
  ConnectorScopeContext,
  ConnectorScopeUpdateRequest,
  ConnectorSetupRequest,
  ConnectorSetupResponse,
  ConnectorStatus,
  ConnectorStatusResponse,
  ConnectorTestResponse,
  OfficeMode,
  OfficeSessionOpenResponse,
  OfficeStateUpdate,
  PromptImagePayload,
  ThinkingCapabilities,
  ThinkingLevel,
  AskUserRequest,
  AskUserQuestionAnswer,
  ProviderDescriptor,
  SessionStatsResponse,
} from "@pi-office/pi-office-pack/protocol";
import { HOST_LABELS } from "@pi-office/pi-office-pack/defaults";
import { deleteJson, fetchJson, postJson } from "../lib/api";
import {
  buildConnectorScopeContext,
  buildOpenRequest,
  captureDocumentSnapshot,
  capturePromptVisuals,
  collectOfficeState,
  readOfficeTheme,
  restoreDocumentSnapshot,
  subscribeToOfficeChanges,
  waitForOfficeReady,
  type OfficeThemeSnapshot,
} from "../lib/office";
import { executeOfficeTool } from "../lib/office-tools";
import { addCheckpoint, clearCheckpoints, getCheckpoint, hasCheckpoint } from "../lib/checkpoint-store";
import type { DocumentCheckpoint } from "../lib/checkpoint-store";
import type { RewindMode } from "./components/RewindDialog";
import { buildThemeStyle } from "../lib/theme";
import {
  createEntry,
  extractMessageImages,
  extractMessageText,
  hasNonTextMessageContent,
  parseApiError,
  shouldAutoAttachVisuals,
  type ChatEntry,
  type LocalQueueEntry,
  type RemoteQueueEntry,
  type ThinkingSegment,
  type ToolCallEntry,
} from "../lib/helpers";
import { usePreferences, useEnabledModels, pushRecentModel } from "../hooks/usePreferences";
import { useChatHistory } from "../hooks/useChatHistory";
import { Header } from "./components/Header";
import { ContextBar } from "./components/ContextBar";
import { ChatView } from "./components/ChatView";
import { Composer } from "./components/Composer";
import { QueueStrip } from "./components/QueueStrip";
import { SettingsPage } from "./components/SettingsPage";
import { HistoryDropdown } from "./components/HistoryDropdown";
import { AskUserPopup } from "./components/AskUserPopup";
import { ToolPermissionPopup } from "./components/ToolPermissionPopup";
import { EditProposalPopup } from "./components/EditProposalPopup";
import type { ConfiguredModelEntry } from "./components/ModelSelector";
import type {
  OfficeEditProposal,
  OfficeEditProposalDecision,
  ToolPermissionRequest,
  ToolPermissionDecision,
} from "@pi-office/pi-office-pack/protocol";
import { useToolPermissions } from "../hooks/useToolPermissions";
import { applyAcceptedEdits } from "../lib/office";

function buildSelectionFingerprint(selection: OfficeStateUpdate["selection"] | undefined): string {
  if (!selection) return "";
  return JSON.stringify({
    label: selection.label,
    kind: selection.kind,
    textPreview: selection.textPreview,
    imageCount: selection.imageCount ?? 0,
    objectCount: selection.objectCount ?? 0,
    details: selection.details ?? [],
  });
}

function buildPromptImagesFingerprint(images: PromptImagePayload[] | undefined): string | undefined {
  if (!images?.length) return undefined;
  return images
    .map((image) =>
      [
        image.kind ?? "",
        image.label ?? "",
        image.mimeType,
        image.width ?? "",
        image.height ?? "",
        image.data.length,
        image.data.slice(0, 48),
      ].join(":"),
    )
    .join("|");
}

function upsertToolCall(toolCalls: ToolCallEntry[] | undefined, nextToolCall: ToolCallEntry): ToolCallEntry[] {
  const current = toolCalls ?? [];
  let matched = false;
  const next = current.map((toolCall) => {
    if (toolCall.toolCallId !== nextToolCall.toolCallId) return toolCall;
    matched = true;
    return { ...toolCall, ...nextToolCall };
  });
  return matched ? next : [...current, nextToolCall];
}

function updateToolCallStatus(
  toolCalls: ToolCallEntry[] | undefined,
  toolCallId: string | undefined,
  toolName: string,
  status: ToolCallEntry["status"],
  errorMessage?: string | undefined,
  result?: unknown,
): ToolCallEntry[] {
  const patch = { toolName, status, errorMessage, result };
  const current = toolCalls ?? [];
  if (!current.length) {
    return [{ toolCallId: toolCallId ?? crypto.randomUUID(), ...patch }];
  }

  if (toolCallId) {
    let matched = false;
    const next = current.map((toolCall) => {
      if (toolCall.toolCallId !== toolCallId) return toolCall;
      matched = true;
      return { ...toolCall, ...patch };
    });
    if (matched) return next;
  }

  for (let index = current.length - 1; index >= 0; index -= 1) {
    const toolCall = current[index];
    if (!toolCall || toolCall.toolName !== toolName || toolCall.status !== "running") continue;
    return current.map((entry, entryIndex) => (entryIndex === index ? { ...entry, ...patch } : entry));
  }

  return [...current, { toolCallId: toolCallId ?? crypto.randomUUID(), ...patch }];
}

export function App() {
  const [officeState, setOfficeState] = useState<OfficeStateUpdate>();
  const [officeTheme, setOfficeTheme] = useState<OfficeThemeSnapshot>();
  const [mode, setMode] = useState<OfficeMode>("document-only");
  const [providers, setProviders] = useState<ProviderDescriptor[]>([]);
  const [authStatus, setAuthStatus] = useState<AuthStatusResponse>();
  const [connectors, setConnectors] = useState<ConnectorCatalogItem[]>([]);
  const [connectorStatuses, setConnectorStatuses] = useState<ConnectorStatus[]>([]);
  const [connectorDiagnostics, setConnectorDiagnostics] = useState<ConnectorDiagnosticsResponse>();
  const [connectorAuditPreference, setConnectorAuditPreference] = useState<ConnectorAuditPreference>();
  const [messages, setMessages] = useState<ChatEntry[]>([]);
  const [draft, setDraft] = useState("");
  const [sessionId, setSessionId] = useState<string>();
  const [connectionState, setConnectionState] = useState("connecting");
  const [selectedModelKey, setSelectedModelKey] = useState("");
  const [thinkingLevel, setThinkingLevel] = useState<ThinkingLevel>("high");
  const [availableThinkingLevels, setAvailableThinkingLevels] = useState<ThinkingLevel[]>([]);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [isBusy, setIsBusy] = useState(false);
  const [activeToolName, setActiveToolName] = useState<string>();
  const [localQueue, setLocalQueue] = useState<LocalQueueEntry[]>([]);
  const [remoteQueue, setRemoteQueue] = useState<RemoteQueueEntry[]>([]);
  const [selectedLocalQueueId, setSelectedLocalQueueId] = useState<string>();
  const [sessionStats, setSessionStats] = useState<SessionStatsResponse>();
  const [chatSubject, setChatSubject] = useState<string>();
  const [historyOpen, setHistoryOpen] = useState(false);
  const [activeChatId, setActiveChatId] = useState<string>();
  const [askUserRequest, setAskUserRequest] = useState<AskUserRequest | null>(null);
  const [toolPermissionRequest, setToolPermissionRequest] = useState<ToolPermissionRequest | null>(null);
  const [editProposal, setEditProposal] = useState<OfficeEditProposal | null>(null);

  const { preferences, updatePreferences } = usePreferences();
  const { enabledModels, enabledProviders, toggleModel, toggleProvider, isModelEnabled } = useEnabledModels();
  const { grantPermission } = useToolPermissions(
    preferences.autonomyLevel,
    preferences.toolPermissionOverrides,
    officeState?.document.workspaceDir,
  );
  const { saveChat, loadChat, deleteChat, listChats, updateSubject } = useChatHistory();

  const sessionIdRef = useRef<string | undefined>(undefined);
  const bridgeRef = useRef<WebSocket | null>(null);
  const assistantEntryIdRef = useRef<string | undefined>(undefined);
  const toolCountRef = useRef(0);
  const pendingUserEchoesRef = useRef<string[]>([]);
  const drainingQueueRef = useRef(false);
  const modelPinnedRef = useRef(false);
  const bridgeHandlerRef = useRef<((payload: BridgeServerMessage) => Promise<void>) | undefined>(undefined);
  const lastSentVisualFingerprintRef = useRef<string | undefined>(undefined);
  const lastSelectionFingerprintRef = useRef<string>("");
  const assistantEndCountRef = useRef(0);
  const subjectDerivedRef = useRef(false);
  const disconnectBridgeRef = useRef<(() => void) | undefined>(undefined);
  const officeStateRef = useRef<OfficeStateUpdate | undefined>(undefined);
  const pendingAskUserSummaryRef = useRef<ChatEntry | null>(null);
  const thinkingSegmentOpenRef = useRef(false);

  officeStateRef.current = officeState;

  const readyProviders = useMemo(
    () => providers.filter((p) => p.configured && p.models.some((m) => m.configured)),
    [providers],
  );

  const allConfiguredModels: ConfiguredModelEntry[] = useMemo(
    () =>
      readyProviders.flatMap((provider) =>
        provider.models
          .filter((m) => m.configured)
          .map((model) => ({
            provider: { provider: provider.provider, label: provider.label },
            model,
            key: `${model.provider}::${model.modelId}`,
          })),
      ),
    [readyProviders],
  );

  const configuredModels: ConfiguredModelEntry[] = useMemo(
    () => allConfiguredModels.filter((e) => isModelEnabled(e.model.provider, e.model.modelId)),
    [allConfiguredModels, isModelEnabled],
  );

  const themeStyle = useMemo(
    () => buildThemeStyle(officeTheme, officeState?.host),
    [officeState?.host, officeTheme],
  );
  const connectorScopeContext = useMemo(
    () => buildConnectorScopeContext(officeState),
    [
      officeState?.host,
      officeState?.document.id,
      officeState?.document.title,
      officeState?.document.saved,
      officeState?.document.documentUrl,
      officeState?.document.workspaceDir,
    ],
  );

  const hasConversation = useMemo(
    () => messages.some((e) => e.role === "user" || e.role === "assistant"),
    [messages],
  );

  const userMessages = useMemo(
    () => messages.filter((e) => e.role === "user").map((e) => e.text),
    [messages],
  );

  const shouldAttachDraftVisuals = useMemo(
    () => preferences.autoAttachVisuals && shouldAutoAttachVisuals(draft, officeState),
    [draft, officeState, preferences.autoAttachVisuals],
  );
  const selectionFingerprint = useMemo(
    () => buildSelectionFingerprint(officeState?.selection),
    [
      officeState?.selection.label,
      officeState?.selection.kind,
      officeState?.selection.textPreview,
      officeState?.selection.imageCount,
      officeState?.selection.objectCount,
      officeState?.selection.details?.join("\u0000"),
    ],
  );

  const pushSystemMessage = useCallback((text: string) => {
    setMessages((c) => [...c, createEntry("system", text)]);
  }, []);

  const pushErrorMessage = useCallback((text: string) => {
    setMessages((c) => [...c, createEntry("error", text)]);
  }, []);

  const refreshProviderState = useCallback(async () => {
    const [catalog, nextAuth] = await Promise.all([
      fetchJson<{ providers: ProviderDescriptor[] }>("/v1/providers"),
      fetchJson<AuthStatusResponse>("/v1/auth/status"),
    ]);
    setProviders(catalog.providers);
    setAuthStatus(nextAuth);
  }, []);

  const refreshConnectorState = useCallback(async (scopeContext?: ConnectorScopeContext) => {
    const [catalog, nextStatus, nextDiagnostics, nextAudit] = await Promise.all([
      fetchJson<ConnectorCatalogResponse>("/v1/connectors/catalog"),
      postJson<ConnectorStatusResponse>("/v1/connectors/status", { scopeContext }),
      fetchJson<ConnectorDiagnosticsResponse>("/v1/connectors/diagnostics"),
      fetchJson<ConnectorAuditPreferenceResponse>("/v1/connectors/audit"),
    ]);
    setConnectors(catalog.connectors);
    setConnectorStatuses(nextStatus.connectors);
    setConnectorDiagnostics(nextDiagnostics);
    setConnectorAuditPreference(nextAudit.preference);
  }, []);

  const refreshSessionStats = useCallback(async (targetId?: string) => {
    const sid = targetId ?? sessionIdRef.current;
    if (!sid) { setSessionStats(undefined); return; }
    try {
      const stats = await fetchJson<SessionStatsResponse>(`/v1/sessions/${sid}/stats`);
      if (sessionIdRef.current === sid) setSessionStats(stats);
    } catch { /* non-critical */ }
  }, []);

  const refreshThinkingCapabilities = useCallback(async (targetId?: string) => {
    const sid = targetId ?? sessionIdRef.current;
    if (!sid) return;
    try {
      const caps = await fetchJson<ThinkingCapabilities>(`/v1/sessions/${sid}/thinking`);
      if (sessionIdRef.current === sid) {
        setAvailableThinkingLevels(caps.availableLevels as ThinkingLevel[]);
        setThinkingLevel(caps.currentLevel as ThinkingLevel);
      }
    } catch { /* non-critical */ }
  }, []);

  const handleSetThinkingLevel = useCallback(async (level: ThinkingLevel) => {
    const sid = sessionIdRef.current;
    if (!sid) return;
    setThinkingLevel(level);
    try {
      await postJson<{ ok: true }>(`/v1/sessions/${sid}/thinking-level`, { level });
      await refreshThinkingCapabilities(sid);
    } catch (error) {
      pushErrorMessage(`Thinking level change failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  }, [pushErrorMessage, refreshThinkingCapabilities]);

  const consumeOptimisticUserEcho = useCallback((text: string) => {
    const i = pendingUserEchoesRef.current.findIndex((e) => e === text);
    if (i === -1) return false;
    pendingUserEchoesRef.current.splice(i, 1);
    return true;
  }, []);

  const markQueueItem = useCallback((id: string, status: LocalQueueEntry["status"]) => {
    setLocalQueue((c) => c.map((e) => (e.id === id ? { ...e, status } : e)));
  }, []);

  const removeQueueItem = useCallback((id: string) => {
    setLocalQueue((c) => c.filter((e) => e.id !== id));
    setSelectedLocalQueueId((c) => (c === id ? undefined : c));
  }, []);

  const ensureAssistantEntry = useCallback(
    (seed?: Partial<Pick<ChatEntry, "text" | "thinking" | "thinkingSegments" | "toolCalls">>) => {
      if (assistantEntryIdRef.current) return assistantEntryIdRef.current;
      const entry = createEntry("assistant", seed?.text ?? "", true, {
        thinking: seed?.thinking,
        toolCalls: seed?.toolCalls,
      });
      assistantEntryIdRef.current = entry.id;
      setMessages((c) => [...c, entry]);
      return entry.id;
    },
    [],
  );

  const updateAssistantEntry = useCallback(
    (
      updater: (entry: ChatEntry) => ChatEntry,
      seed?: Partial<Pick<ChatEntry, "text" | "thinking" | "thinkingSegments" | "toolCalls">>,
    ) => {
      const entryId = ensureAssistantEntry(seed);
      setMessages((c) => c.map((entry) => (entry.id === entryId ? updater(entry) : entry)));
      return entryId;
    },
    [ensureAssistantEntry],
  );

  const buildPromptImages = useCallback(
    async (text: string): Promise<PromptImagePayload[] | undefined> => {
      if (!officeState || !preferences.autoAttachVisuals || !shouldAutoAttachVisuals(text, officeState)) {
        return undefined;
      }
      try {
        const visuals = await capturePromptVisuals(officeState.host, officeState.host === "powerpoint" ? 2 : 3);
        const images = visuals.length
          ? visuals.map((v) => ({
              data: v.data, mimeType: v.mimeType, label: v.label,
              width: v.width, height: v.height, kind: v.kind,
            }))
          : undefined;
        const fingerprint = buildPromptImagesFingerprint(images);
        if (fingerprint && fingerprint === lastSentVisualFingerprintRef.current) {
          return undefined;
        }
        return images;
      } catch (error) {
        pushSystemMessage(`Visual capture unavailable: ${error instanceof Error ? error.message : String(error)}`);
        return undefined;
      }
    },
    [officeState, preferences.autoAttachVisuals, pushSystemMessage],
  );

  const captureCheckpointBeforePrompt = useCallback(
    async (userMessageId: string, promptText: string) => {
      if (!officeState?.host) return;
      try {
        const currentMessages = await new Promise<ChatEntry[]>((resolve) => {
          setMessages((c) => { resolve(c); return c; });
        });
        const data = await captureDocumentSnapshot(officeState.host);
        const cp: DocumentCheckpoint = {
          id: userMessageId,
          timestamp: Date.now(),
          host: officeState.host,
          userPrompt: promptText.slice(0, 80),
          messageCount: currentMessages.length,
          hasDocumentData: Boolean(data.ooxml || data.sheets?.length || data.presentationBase64),
          ...data,
        };
        addCheckpoint(cp);

        if (preferences.experimentalRewindSnapshots && officeState.document.id) {
          bridgeRef.current?.send(JSON.stringify({
            type: "persist_checkpoint",
            documentId: officeState.document.id,
            checkpoint: { ...cp, hasDocumentData: undefined },
          }));
        }
      } catch (error) {
        console.warn("[rewind] Checkpoint capture failed:", error);
      }
    },
    [officeState, preferences.experimentalRewindSnapshots],
  );

  const postPrompt = useCallback(
    async (
      text: string,
      modeName: "prompt" | "steer" | "followUp",
      images?: PromptImagePayload[],
      optimisticUser = false,
    ) => {
      const sid = sessionIdRef.current;
      if (!sid) throw new Error("Pi session is not connected yet.");

      const userMessageId = crypto.randomUUID();

      if (modeName === "prompt") {
        await captureCheckpointBeforePrompt(userMessageId, text);
      }

      const route =
        modeName === "prompt"
          ? `/v1/sessions/${sid}/prompt`
          : modeName === "steer"
            ? `/v1/sessions/${sid}/steer`
            : `/v1/sessions/${sid}/follow-up`;

      if (optimisticUser) {
        pendingUserEchoesRef.current.push(text);
        setMessages((c) => [...c, { ...createEntry("user", text), id: userMessageId }]);
      }
      if (modeName === "prompt") setIsBusy(true);
      await postJson<{ ok: true }>(route, { text, mode: modeName, images });
      const imageFingerprint = buildPromptImagesFingerprint(images);
      if (imageFingerprint) {
        lastSentVisualFingerprintRef.current = imageFingerprint;
      }
    },
    [captureCheckpointBeforePrompt],
  );

  const handleRewind = useCallback(
    async (messageId: string, mode: RewindMode) => {
      if (mode === "cancel") return;
      const cp = getCheckpoint(messageId);
      if (!cp) { pushErrorMessage("No checkpoint found for this message."); return; }

      const messageIndex = messages.findIndex((m) => m.id === messageId);
      if (messageIndex < 0) return;

      if (mode === "conversation_and_document" && cp.hasDocumentData && officeState?.host) {
        try {
          const snapshotData: Parameters<typeof restoreDocumentSnapshot>[1] = {};
          if (cp.ooxml) snapshotData.ooxml = cp.ooxml;
          if (cp.sheets?.length) snapshotData.sheets = cp.sheets;
          if (cp.presentationBase64) snapshotData.presentationBase64 = cp.presentationBase64;
          await restoreDocumentSnapshot(officeState.host, snapshotData);
        } catch (error) {
          pushErrorMessage(`Document restore failed: ${error instanceof Error ? error.message : String(error)}`);
        }
      }

      setMessages((c) => c.slice(0, messageIndex));
      bridgeRef.current?.send(JSON.stringify({
        type: "rewind_session",
        targetMessageCount: cp.messageCount,
      }));

      assistantEntryIdRef.current = undefined;
      setIsBusy(false);

      const label = mode === "conversation_and_document" ? "conversation and document" : "conversation";
      const prompt = cp.userPrompt.length > 50 ? cp.userPrompt.slice(0, 47) + "..." : cp.userPrompt;
      pushSystemMessage(`Rewound ${label} to before: "${prompt}"`);
    },
    [messages, officeState, pushErrorMessage, pushSystemMessage],
  );

  const sendDraftPrompt = useCallback(async () => {
    const text = draft.trim();
    if (!text || !sessionIdRef.current) return;
    const images = await buildPromptImages(text);
    setDraft("");
    setSelectedLocalQueueId(undefined);
    try {
      await postPrompt(text, "prompt", images, true);
    } catch (error) {
      pushErrorMessage(`Prompt failed: ${parseApiError(error instanceof Error ? error.message : String(error)) || "Unknown error"}`);
      setIsBusy(false);
    }
  }, [buildPromptImages, draft, postPrompt, pushErrorMessage]);

  const queueCurrentDraft = useCallback(async () => {
    const text = draft.trim();
    if (!text) return;
    const images = await buildPromptImages(text);
    setLocalQueue((c) => [...c, { id: crypto.randomUUID(), text, images, status: "queued" }]);
    setDraft("");
    setSelectedLocalQueueId(undefined);
  }, [buildPromptImages, draft]);

  const sendSelectedQueueAsSteer = useCallback(async () => {
    if (!selectedLocalQueueId) return;
    const target = localQueue.find((e) => e.id === selectedLocalQueueId);
    if (!target) { setSelectedLocalQueueId(undefined); return; }
    const text = draft.trim() || target.text;
    const images = target.images ?? (await buildPromptImages(text));
    try {
      await postPrompt(text, "steer", images, false);
      removeQueueItem(target.id);
      setDraft("");
    } catch (error) {
      pushErrorMessage(`Steer failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  }, [buildPromptImages, draft, localQueue, postPrompt, pushErrorMessage, removeQueueItem, selectedLocalQueueId]);

  const abortTurn = useCallback(async () => {
    const sid = sessionIdRef.current;
    if (!sid) return;
    try {
      await postJson<{ ok: true }>(`/v1/sessions/${sid}/abort`, {});
    } catch (error) {
      pushErrorMessage(`Abort failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  }, [pushErrorMessage]);

  const handleAskUserSubmit = useCallback(
    (answers: AskUserQuestionAnswer[]) => {
      const req = askUserRequest;
      if (!req) return;
      bridgeRef.current?.send(JSON.stringify({ type: "ask_user_response", response: { requestId: req.requestId, answers } }));
      const summaryLines = answers.map((a) => {
        const q = req.questions.find((q) => q.id === a.questionId);
        const answer = a.selectedOption ?? "None of the above";
        const notes = a.notes ? `\nNotes: ${a.notes}` : "";
        return `Q: ${q?.question ?? "?"}\nA: ${answer}${notes}`;
      });
      pendingAskUserSummaryRef.current = createEntry("system", `Pi asked:\n\n${summaryLines.join("\n\n")}`);
      setAskUserRequest(null);
    },
    [askUserRequest],
  );

  const handleToolPermissionDecision = useCallback(
    (requestId: string, allowed: boolean, scope: "once" | "session" | "workspace" | "always") => {
      const req = toolPermissionRequest;
      if (!req || req.requestId !== requestId) return;
      const decision: ToolPermissionDecision = {
        toolName: req.toolName,
        allowed,
        scope,
      };
      bridgeRef.current?.send(JSON.stringify({
        type: "tool_permission_response",
        requestId,
        decision,
      }));
      if (allowed && scope !== "once") {
        grantPermission(req.toolName, scope);
      }
      setToolPermissionRequest(null);
    },
    [toolPermissionRequest, grantPermission],
  );

  const handleSetAutonomyLevel = useCallback(
    (level: import("@pi-office/pi-office-pack/protocol").AutonomyLevel) => {
      updatePreferences({ autonomyLevel: level });
    },
    [updatePreferences],
  );

  const handleEditProposalDecision = useCallback(
    async (decision: OfficeEditProposalDecision) => {
      const proposal = editProposal;
      if (!proposal) return;

      const accepted = decision.decisions.filter((d) => d.accepted);
      if (accepted.length > 0) {
        const editsToApply = accepted
          .map((d) => {
            const edit = proposal.edits.find((e) => e.id === d.editId);
            if (!edit) return null;
            return {
              searchText: edit.searchText ?? edit.oldText ?? "",
              newText: d.modifiedText ?? edit.newText ?? "",
              kind: edit.kind,
            };
          })
          .filter((e): e is NonNullable<typeof e> => e !== null && Boolean(e.searchText));

        try {
          const result = await applyAcceptedEdits(editsToApply);
          decision.applicationResult = {
            applied: result.applied,
            failed: result.failed,
            errors: result.errors,
          };
        } catch (error) {
          decision.applicationResult = {
            applied: 0,
            failed: editsToApply.length,
            errors: [error instanceof Error ? error.message : String(error)],
          };
        }
      }

      bridgeRef.current?.send(JSON.stringify({ type: "edit_proposal_decision", decision }));
      setEditProposal(null);
    },
    [editProposal],
  );

  const handleSessionEvent = useCallback(
    (event: unknown) => {
      if (!event || typeof event !== "object") return;
      const r = event as Record<string, unknown>;
      const type = String(r.type ?? "");
      if (type === "turn_start") { setIsBusy(true); return; }

      if (type === "turn_end" || type === "agent_end") {
        toolCountRef.current = 0;
        if (type === "agent_end") {
          thinkingSegmentOpenRef.current = false;
          // Safety net: finalize any pending entry that wasn't closed by message_end
          const tid = assistantEntryIdRef.current;
          if (tid) {
            setMessages((c) => c.map((e) => {
              if (e.id !== tid) return e;
              let segs = e.thinkingSegments;
              if (segs?.length) {
                const last = segs[segs.length - 1]!;
                if (!last.durationMs) {
                  segs = [...segs.slice(0, -1), { ...last, durationMs: Date.now() - last.startedAt }];
                }
              }
              return { ...e, pending: false, thinkingSegments: segs };
            }));
          }
          assistantEntryIdRef.current = undefined;
        }
        setActiveToolName(undefined);
        setIsBusy(false);
        void refreshSessionStats();
        // Surface API / model errors
        const turnMsg = r.message as Record<string, unknown> | undefined;
        const stopReason = String(turnMsg?.stopReason ?? "");
        if (stopReason === "error" || stopReason.startsWith("err")) {
          const raw = turnMsg?.error ?? turnMsg?.errorMessage ?? "";
          const detail = parseApiError(raw);
          pushErrorMessage(detail || "The model returned an error. Check companion logs for details.");
        }
        return;
      }

      if (type === "queue_update") {
        const steering = Array.isArray(r.steering) ? r.steering : [];
        const followUp = Array.isArray(r.followUp) ? r.followUp : [];
        setRemoteQueue([
          ...steering.map((t, i) => ({ id: `steer-${i}-${String(t)}`, text: String(t), mode: "steer" as const })),
          ...followUp.map((t, i) => ({ id: `follow-${i}-${String(t)}`, text: String(t), mode: "followUp" as const })),
        ]);
        return;
      }

      if (type === "tool_execution_start") {
        toolCountRef.current += 1;
        const toolName = String(r.toolName ?? "tool");
        const toolCallId = String(r.toolCallId ?? crypto.randomUUID());
        const toolInput = r.input ?? r.args ?? r.params ?? undefined;
        updateAssistantEntry(
          (entry) => ({
            ...entry,
            pending: true,
            toolCalls: upsertToolCall(entry.toolCalls, {
              toolCallId,
              toolName,
              status: "running",
              input: toolInput,
            }),
          }),
          {
            toolCalls: [{
              toolCallId,
              toolName,
              status: "running",
              input: toolInput,
            }],
          },
        );
        setActiveToolName(toolName);
        setIsBusy(true);
        return;
      }

      if (type === "tool_execution_end") {
        const toolName = String(r.toolName ?? "tool");
        const toolCallId = typeof r.toolCallId === "string" ? r.toolCallId : undefined;
        const status: ToolCallEntry["status"] = r.isError ? "error" : "done";
        const errorMessage = r.isError
          ? String(
              r.error ??
              r.errorMessage ??
              r.message ??
              ((() => {
                const result = r.result as Record<string, unknown> | undefined;
                if (!result) return "";
                const content = result.content;
                if (Array.isArray(content)) {
                  const textPart = content.find(
                    (p: unknown) => p && typeof p === "object" && (p as { type?: string }).type === "text",
                  ) as { text?: string } | undefined;
                  if (textPart?.text) return textPart.text;
                }
                return result.error ?? "";
              })()) ??
              "",
            )
          : undefined;
        if (r.isError) {
          console.error(`[tool-error] ${toolName}`, {
            toolCallId,
            errorMessage,
            rawEvent: JSON.parse(JSON.stringify(r)),
          });
        }
        const toolResult = r.result ?? undefined;
        updateAssistantEntry((entry) => ({
          ...entry,
          toolCalls: updateToolCallStatus(entry.toolCalls, toolCallId, toolName, status, errorMessage, toolResult),
        }));
        toolCountRef.current = Math.max(0, toolCountRef.current - 1);
        if (toolCountRef.current === 0) setActiveToolName(undefined);

        if (toolName === "ask_user" && pendingAskUserSummaryRef.current) {
          const summary = pendingAskUserSummaryRef.current;
          pendingAskUserSummaryRef.current = null;
          const currentEntryId = assistantEntryIdRef.current;
          if (currentEntryId) {
            setMessages((c) => c.map((e) => (e.id === currentEntryId ? { ...e, pending: false } : e)));
          }
          assistantEntryIdRef.current = undefined;
          setMessages((c) => [...c, summary]);
        }
        return;
      }

      if (type === "message_start") {
        const msg = r.message as Record<string, unknown> | undefined;
        const role = String(msg?.role ?? "");
        if (role === "assistant") {
          ensureAssistantEntry();
          setIsBusy(true);
          return;
        }
        if (role === "user") {
          const text = extractMessageText(msg?.content);
          if (!text || consumeOptimisticUserEcho(text)) return;
          setMessages((c) => [...c, createEntry("user", text)]);
        }
        return;
      }

      if (type === "message_update") {
        const ame = r.assistantMessageEvent as Record<string, unknown> | undefined;
        if (!ame) return;
        const assistantMessageType = String(ame.type ?? "");
        if (assistantMessageType === "text_delta") {
          const delta = String(ame.delta ?? "");
          if (!delta) return;
          updateAssistantEntry(
            (entry) => ({ ...entry, text: `${entry.text}${delta}`, pending: true }),
            { text: delta },
          );
          return;
        }
        if (assistantMessageType === "thinking_delta") {
          if (!preferences.showThinkingTraces) return;
          const delta = String(ame.delta ?? "");
          if (!delta) return;
          const needNewSegment = !thinkingSegmentOpenRef.current;
          thinkingSegmentOpenRef.current = true;
          updateAssistantEntry(
            (entry) => {
              const segments = entry.thinkingSegments ?? [];
              let nextSegments: ThinkingSegment[];
              if (needNewSegment || segments.length === 0) {
                nextSegments = [...segments, { text: delta, startedAt: Date.now() }];
              } else {
                const last = segments[segments.length - 1]!;
                nextSegments = [...segments.slice(0, -1), { ...last, text: last.text + delta }];
              }
              return {
                ...entry,
                thinking: `${entry.thinking ?? ""}${delta}`,
                thinkingSegments: nextSegments,
                pending: true,
              };
            },
            { thinking: delta, thinkingSegments: [{ text: delta, startedAt: Date.now() }] },
          );
        }
        return;
      }

      if (type === "message_end") {
        const msg = r.message as Record<string, unknown> | undefined;
        if (String(msg?.role ?? "") === "assistant") {
          const finalText = extractMessageText(msg?.content);
          const finalImages = extractMessageImages(msg?.content);
          const hasNonTextContent = hasNonTextMessageContent(msg?.content);
          const stopReason = String(msg?.stopReason ?? "");
          const isError = stopReason === "error" || stopReason.startsWith("err");

          // Intermediate message (model will call tools next) — keep accumulating into same entry
          if (stopReason === "tool_use") {
            // Close the current thinking segment so the next one starts fresh
            thinkingSegmentOpenRef.current = false;
            updateAssistantEntry((entry) => {
              const segments = entry.thinkingSegments;
              let nextSegments = segments;
              if (segments?.length) {
                const last = segments[segments.length - 1]!;
                if (!last.durationMs) {
                  nextSegments = [
                    ...segments.slice(0, -1),
                    { ...last, durationMs: Date.now() - last.startedAt },
                  ];
                }
              }
              return {
                ...entry,
                text: finalText || entry.text,
                thinkingSegments: nextSegments,
              };
            });
            return;
          }

          thinkingSegmentOpenRef.current = false;
          const tid = assistantEntryIdRef.current;
          assistantEntryIdRef.current = undefined;
          if (!tid) {
            if (finalText) setMessages((c) => [...c, createEntry("assistant", finalText)]);
            return;
          }

          // Determine renderability from the event data directly rather than
          // a mutable variable inside setMessages. React 19 automatic batching
          // defers updaters, so a `let` set inside the updater may still be
          // at its initial value when checked synchronously after setMessages.
          const hasKnownContent = Boolean(
            finalText?.trim() || hasNonTextContent || finalImages?.length,
          );

          setMessages((c) => {
            const entry = c.find((message) => message.id === tid);
            if (!entry) {
              return finalText ? [...c, createEntry("assistant", finalText)] : c;
            }
            let closedSegments = entry.thinkingSegments;
            if (closedSegments?.length) {
              const last = closedSegments[closedSegments.length - 1]!;
              if (!last.durationMs) {
                closedSegments = [
                  ...closedSegments.slice(0, -1),
                  { ...last, durationMs: Date.now() - last.startedAt },
                ];
              }
            }
            const nextEntry = {
              ...entry,
              text: finalText || entry.text,
              pending: false,
              images: finalImages ?? entry.images,
              thinkingSegments: closedSegments,
            };
            const renderable = Boolean(
              nextEntry.text.trim() ||
              nextEntry.thinking?.trim() ||
              nextEntry.toolCalls?.length ||
              nextEntry.images?.length,
            );
            return renderable
              ? c.map((e) => (e.id === tid ? nextEntry : e))
              : c.filter((e) => e.id !== tid);
          });

          if (!hasKnownContent && !isError) {
            console.error("[empty-response] message_end with no renderable content", {
              entryId: tid,
              finalText,
              stopReason,
              hasNonTextContent,
              messageContent: msg?.content,
              rawEvent: JSON.parse(JSON.stringify(r)),
            });
            pushSystemMessage("The model returned an empty response.");
          }
          void refreshSessionStats();

          // Derive chat subject after the 2nd assistant response
          assistantEndCountRef.current += 1;
          if (assistantEndCountRef.current >= 2 && !subjectDerivedRef.current && sessionIdRef.current) {
            subjectDerivedRef.current = true;
            void (async () => {
              try {
                const currentMessages = (await new Promise<ChatEntry[]>((resolve) => {
                  setMessages((c) => { resolve(c); return c; });
                }));
                const summaryMessages = currentMessages
                  .filter((m) => m.role === "user" || m.role === "assistant")
                  .slice(0, 6)
                  .map((m) => ({ role: m.role, text: m.text.slice(0, 200) }));
                if (!summaryMessages.length) return;
                const resp = await postJson<{ subject: string }>(
                  `/v1/sessions/${sessionIdRef.current}/derive-subject`,
                  { messages: summaryMessages },
                );
                if (resp.subject) {
                  setChatSubject(resp.subject);
                }
              } catch { /* non-critical */ }
            })();
          }
        }
      }
    },
    [
      consumeOptimisticUserEcho,
      ensureAssistantEntry,
      preferences.showThinkingTraces,
      pushErrorMessage,
      pushSystemMessage,
      refreshSessionStats,
      updateAssistantEntry,
    ],
  );

  const handleBridgeMessage = useCallback(
    async (payload: BridgeServerMessage) => {
      if (payload.type === "connection_state") { setConnectionState(payload.state); return; }
      if (payload.type === "error") { pushErrorMessage(payload.message); return; }
      if (payload.type === "office_tool_call") {
        try {
          const result = await executeOfficeTool(payload.request);
          bridgeRef.current?.send(JSON.stringify({ type: "office_tool_result", result }));
        } catch (error) {
          pushErrorMessage(`Office tool failed: ${error instanceof Error ? error.message : String(error)}`);
        }
        return;
      }
      if (payload.type === "ask_user_request") {
        setAskUserRequest(payload.request);
        return;
      }
      if (payload.type === "tool_permission_request") {
        setToolPermissionRequest(payload.request);
        return;
      }
      if (payload.type === "edit_proposal_request") {
        setEditProposal(payload.proposal);
        return;
      }
      if (payload.type === "session_event") {
        handleSessionEvent(payload.event);
        return;
      }
    },
    [handleSessionEvent, pushErrorMessage],
  );

  // Keep a stable ref so the WebSocket listener always calls the latest handler.
  bridgeHandlerRef.current = handleBridgeMessage;

  const openBridge = useCallback(
    (response: OfficeSessionOpenResponse) => {
      const socketOrigin = response.origin.replace(/^https:/i, "wss:");
      const socket = new WebSocket(`${socketOrigin}${response.eventsPath}`);
      bridgeRef.current = socket;

      socket.addEventListener("open", () => {
        socket.send(JSON.stringify({ type: "client_ready" }));
      });
      socket.addEventListener("message", (msg) => {
        try {
          const payload = JSON.parse(String(msg.data)) as BridgeServerMessage;
          void bridgeHandlerRef.current?.(payload);
        } catch (error) {
          console.error("[bridge] message error:", error);
        }
      });
      socket.addEventListener("close", () => {
        if (bridgeRef.current === socket) { bridgeRef.current = null; setConnectionState("offline"); }
      });
      socket.addEventListener("error", () => { setConnectionState("offline"); });

      return () => { if (bridgeRef.current === socket) bridgeRef.current = null; socket.close(); };
    },
    [],
  );

  const openSession = useCallback(async (stateOverride?: OfficeStateUpdate, forceNew?: boolean) => {
    const currentState = stateOverride ?? officeStateRef.current;
    if (!currentState) throw new Error("Office state not available");

    disconnectBridgeRef.current?.();
    disconnectBridgeRef.current = undefined;

    setConnectionState("connecting");
    const session = await postJson<OfficeSessionOpenResponse>("/v1/sessions/open", buildOpenRequest(currentState, forceNew));
    setSessionId(session.sessionId);
    sessionIdRef.current = session.sessionId;
    setMode(session.mode);
    disconnectBridgeRef.current = openBridge(session);

    const next = await postJson<{ ok: true; mode: OfficeMode }>(
      `/v1/sessions/${session.sessionId}/office-state`, currentState,
    );
    setMode(next.mode);
    void refreshSessionStats(session.sessionId);
    void refreshThinkingCapabilities(session.sessionId);
  }, [openBridge, refreshSessionStats, refreshThinkingCapabilities]);

  // Bootstrap
  useEffect(() => {
    let active = true;
    let unsubOffice: (() => void) | undefined;

    const boot = async () => {
      try {
        setConnectionState("connecting");
        await Promise.all([refreshProviderState(), refreshConnectorState(undefined)]);
        if (!active) return;

        const host = await waitForOfficeReady();
        if (!active) return;
        setOfficeTheme(readOfficeTheme());
        const initialState = await collectOfficeState(host);
        if (!active) return;
        setOfficeState(initialState);

        if (!active) return;
        await openSession(initialState);

        unsubOffice = subscribeToOfficeChanges(async () => {
          if (!sessionIdRef.current) return;
          try {
            setOfficeTheme(readOfficeTheme());
            const s = await collectOfficeState(host);
            if (!active) return;
            setOfficeState(s);
            const r = await postJson<{ ok: true; mode: OfficeMode }>(
              `/v1/sessions/${sessionIdRef.current}/office-state`, s,
            );
            if (!active) return;
            setMode(r.mode);
          } catch (error) {
            if (active) pushErrorMessage(`Office state refresh failed: ${error instanceof Error ? error.message : String(error)}`);
          }
        });
      } catch (error) {
        if (active) {
          pushErrorMessage(`Connection failed: ${error instanceof Error ? error.message : String(error)}`);
          setConnectionState("offline");
        }
      }
    };

    void boot();
    return () => {
      active = false;
      unsubOffice?.();
      disconnectBridgeRef.current?.();
      disconnectBridgeRef.current = undefined;
      sessionIdRef.current = undefined;
      toolCountRef.current = 0;
      assistantEntryIdRef.current = undefined;
      setSessionStats(undefined);
    };
  }, [openSession, pushErrorMessage, refreshConnectorState, refreshProviderState]);

  useEffect(() => {
    if (lastSelectionFingerprintRef.current && lastSelectionFingerprintRef.current !== selectionFingerprint) {
      lastSentVisualFingerprintRef.current = undefined;
    }
    lastSelectionFingerprintRef.current = selectionFingerprint;
  }, [selectionFingerprint]);

  useEffect(() => {
    if (!connectorScopeContext) return;
    void refreshConnectorState(connectorScopeContext);
  }, [connectorScopeContext, refreshConnectorState]);

  // Persist chat subject + messages to history when subject changes
  useEffect(() => {
    if (!chatSubject || !messages.some((m) => m.role === "user")) return;
    const chatId = activeChatId ?? crypto.randomUUID();
    if (!activeChatId) setActiveChatId(chatId);
    saveChat(chatId, messages, chatSubject, officeState);
  }, [chatSubject]); // eslint-disable-line react-hooks/exhaustive-deps

  // Provider auto-select
  useEffect(() => {
    if (!providers.length) return;
  }, [providers]);

  // Model sync to session
  useEffect(() => {
    const exists = selectedModelKey ? configuredModels.some((e) => e.key === selectedModelKey) : true;
    if (!exists) { setSelectedModelKey(""); modelPinnedRef.current = false; }
  }, [configuredModels, selectedModelKey]);

  useEffect(() => {
    if (!sessionId || !selectedModelKey) return;
    const model = configuredModels.find((e) => e.key === selectedModelKey);
    if (!model) return;

    let active = true;
    modelPinnedRef.current = true;

    void (async () => {
      try {
        await postJson<{ ok: true }>(`/v1/sessions/${sessionId}/model`, {
          provider: model.model.provider,
          modelId: model.model.modelId,
        });
        if (active) {
          await refreshSessionStats(sessionId);
          await refreshThinkingCapabilities(sessionId);
        }
      } catch (error) {
        if (active) pushErrorMessage(`Model change failed: ${error instanceof Error ? error.message : String(error)}`);
      }
    })();

    return () => { active = false; };
  }, [configuredModels, pushErrorMessage, refreshSessionStats, refreshThinkingCapabilities, selectedModelKey, sessionId]);

  // Queue drain
  useEffect(() => {
    if (isBusy || !sessionIdRef.current || drainingQueueRef.current || localQueue.length === 0) return;
    const next = localQueue[0];
    if (!next) return;

    drainingQueueRef.current = true;
    markQueueItem(next.id, "sending");

    void (async () => {
      try {
        await postPrompt(next.text, "prompt", next.images, true);
        removeQueueItem(next.id);
        setDraft((c) => (c.trim() === next.text.trim() ? "" : c));
      } catch (error) {
        markQueueItem(next.id, "queued");
        pushErrorMessage(`Queued prompt failed: ${error instanceof Error ? error.message : String(error)}`);
        setIsBusy(false);
      } finally {
        drainingQueueRef.current = false;
      }
    })();
  }, [isBusy, localQueue, markQueueItem, postPrompt, pushErrorMessage, removeQueueItem]);

  // Handlers for settings actions
  const handleSaveApiKey = useCallback(async (provider: string, key: string) => {
    try {
      await postJson<{ ok: true }>("/v1/auth/api-key", { provider, apiKey: key });
      await refreshProviderState();
      pushSystemMessage(`Stored API key for ${provider}.`);
    } catch (error) {
      pushErrorMessage(`API key save failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  }, [pushErrorMessage, pushSystemMessage, refreshProviderState]);

  const handleClearAuth = useCallback(async (provider: string) => {
    try {
      await deleteJson<{ ok: true }>(`/v1/auth/${provider}`);
      await refreshProviderState();
      pushSystemMessage(`Removed stored auth for ${provider}.`);
    } catch (error) {
      pushErrorMessage(`Auth removal failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  }, [pushErrorMessage, pushSystemMessage, refreshProviderState]);

  const handleStartOAuth = useCallback(async (provider: string) => {
    try {
      await postJson<{ ok: true }>("/v1/auth/start", { providerId: provider });
      await refreshProviderState();
      pushSystemMessage(`OAuth flow started for ${provider}.`);
    } catch (error) {
      pushErrorMessage(`OAuth start failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  }, [pushErrorMessage, pushSystemMessage, refreshProviderState]);

  const handlePrepareConnector = useCallback(
    async (connectorId: string, scopeContext?: ConnectorScopeContext) =>
      postJson<ConnectorPrepareResponse>("/v1/connectors/setup/prepare", { connectorId, scopeContext }),
    [],
  );

  const handleConnectConnector = useCallback(async (request: ConnectorSetupRequest) => {
    try {
      const response = await postJson<ConnectorSetupResponse>("/v1/connectors/setup/connect", request);
      await refreshConnectorState(connectorScopeContext);
      pushSystemMessage(`Saved connector: ${response.status.name}.`);
      return response;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      pushErrorMessage(`Connector save failed: ${message}`);
      throw error;
    }
  }, [connectorScopeContext, pushErrorMessage, pushSystemMessage, refreshConnectorState]);

  const handleTestConnector = useCallback(async (request: ConnectorSetupRequest) => {
    try {
      return await postJson<ConnectorTestResponse>("/v1/connectors/setup/test", request);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      pushErrorMessage(`Connector test failed: ${message}`);
      throw error;
    }
  }, [pushErrorMessage]);

  const handleReverifyConnector = useCallback(async (connectorId: string, scopeContext?: ConnectorScopeContext) => {
    try {
      const response = await postJson<ConnectorTestResponse>("/v1/connectors/reverify", { connectorId, scopeContext });
      await refreshConnectorState(connectorScopeContext);
      return response;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      pushErrorMessage(`Connector re-verification failed: ${message}`);
      throw error;
    }
  }, [connectorScopeContext, pushErrorMessage, refreshConnectorState]);

  const handleStartConnectorOAuth = useCallback(async (connectorId: string) => {
    try {
      await postJson<{ ok: true; url?: string }>("/v1/connectors/oauth/start", { connectorId });
      pushSystemMessage("Connector sign-in opened in your browser.");
      await refreshConnectorState(connectorScopeContext);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      pushErrorMessage(`Connector sign-in failed: ${message}`);
      throw error;
    }
  }, [connectorScopeContext, pushErrorMessage, pushSystemMessage, refreshConnectorState]);

  const handleRemoveConnector = useCallback(async (storedConnectorId: string) => {
    try {
      await deleteJson<{ ok: true }>(`/v1/connectors/${storedConnectorId}`);
      await refreshConnectorState(connectorScopeContext);
      pushSystemMessage("Connector removed.");
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      pushErrorMessage(`Connector removal failed: ${message}`);
      throw error;
    }
  }, [connectorScopeContext, pushErrorMessage, pushSystemMessage, refreshConnectorState]);

  const handleSetConnectorFavorite = useCallback(async (request: ConnectorFavoriteRequest) => {
    try {
      await postJson<{ ok: true; status: ConnectorStatus }>("/v1/connectors/favorite", request);
      await refreshConnectorState(connectorScopeContext);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      pushErrorMessage(`Favorite update failed: ${message}`);
      throw error;
    }
  }, [connectorScopeContext, pushErrorMessage, refreshConnectorState]);

  const handleUpdateConnectorScope = useCallback(async (request: ConnectorScopeUpdateRequest) => {
    try {
      await postJson<{ ok: true; status: ConnectorStatus }>("/v1/connectors/scope", request);
      await refreshConnectorState(connectorScopeContext);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      pushErrorMessage(`Scope update failed: ${message}`);
      throw error;
    }
  }, [connectorScopeContext, pushErrorMessage, refreshConnectorState]);

  const handleLoadConnectorLogs = useCallback(async (connectorId: string) => {
    try {
      return await fetchJson<ConnectorLogResponse>(`/v1/connectors/${connectorId}/logs`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      pushErrorMessage(`Connector logs failed to load: ${message}`);
      throw error;
    }
  }, [pushErrorMessage]);

  const handleExportConnectors = useCallback(async () => {
    try {
      return await fetchJson<ConnectorExportBundle>("/v1/connectors/export");
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      pushErrorMessage(`Connector export failed: ${message}`);
      throw error;
    }
  }, [pushErrorMessage]);

  const handlePreviewConnectorImport = useCallback(async (bundle: ConnectorExportBundle) => {
    try {
      return await postJson<ConnectorImportPreviewResponse>("/v1/connectors/import/preview", { bundle });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      pushErrorMessage(`Connector import preview failed: ${message}`);
      throw error;
    }
  }, [pushErrorMessage]);

  const handleApplyConnectorImport = useCallback(async (bundle: ConnectorExportBundle, resolutions?: Record<string, "skip" | "replace">) => {
    try {
      const response = await postJson<ConnectorImportApplyResponse>("/v1/connectors/import/apply", { bundle, resolutions });
      await refreshConnectorState(connectorScopeContext);
      return response;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      pushErrorMessage(`Connector import failed: ${message}`);
      throw error;
    }
  }, [connectorScopeContext, pushErrorMessage, refreshConnectorState]);

  const handleSetConnectorAuditPreference = useCallback(async (preference: ConnectorAuditPreference) => {
    try {
      const response = await postJson<ConnectorAuditPreferenceResponse>("/v1/connectors/audit", preference);
      setConnectorAuditPreference(response.preference);
      return response.preference;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      pushErrorMessage(`Audit preference update failed: ${message}`);
      throw error;
    }
  }, [pushErrorMessage]);

  const handleSelectModel = useCallback((key: string) => {
    if (!key && modelPinnedRef.current) {
      pushSystemMessage("This session keeps its current model until you reopen the document.");
    }
    setSelectedModelKey(key);
    if (key) pushRecentModel(key);
  }, [pushSystemMessage]);

  const handleQueueSelection = useCallback((entry: LocalQueueEntry) => {
    setSelectedLocalQueueId((c) => (c === entry.id ? undefined : entry.id));
    setDraft(entry.text);
  }, []);

  // Settings page replaces the entire chat view
  if (settingsOpen) {
    return (
      <div className="shell" data-host={officeState?.host ?? "word"} style={themeStyle}>
        <SettingsPage
          officeState={officeState}
          mode={mode}
          providers={providers}
          authStatus={authStatus}
          connectors={connectors}
          connectorStatuses={connectorStatuses}
          connectorDiagnostics={connectorDiagnostics}
          connectorAuditPreference={connectorAuditPreference}
          connectorScopeContext={connectorScopeContext}
          sessionStats={sessionStats}
          preferences={preferences}
          enabledModels={enabledModels}
          enabledProviders={enabledProviders}
          onClose={() => setSettingsOpen(false)}
          onUpdatePreferences={updatePreferences}
          onToggleModel={toggleModel}
          onToggleProvider={toggleProvider}
          onSaveApiKey={handleSaveApiKey}
          onStartOAuth={handleStartOAuth}
          onClearAuth={handleClearAuth}
          onPrepareConnector={handlePrepareConnector}
          onConnectConnector={handleConnectConnector}
          onTestConnector={handleTestConnector}
          onReverifyConnector={handleReverifyConnector}
          onStartConnectorOAuth={handleStartConnectorOAuth}
          onRemoveConnector={handleRemoveConnector}
          onSetConnectorFavorite={handleSetConnectorFavorite}
          onUpdateConnectorScope={handleUpdateConnectorScope}
          onLoadConnectorLogs={handleLoadConnectorLogs}
          onExportConnectors={handleExportConnectors}
          onPreviewConnectorImport={handlePreviewConnectorImport}
          onApplyConnectorImport={handleApplyConnectorImport}
          onSetConnectorAuditPreference={handleSetConnectorAuditPreference}
        />
      </div>
    );
  }

  return (
    <div
      className={`shell ${preferences.compactMessages ? "shell-compact" : ""}`}
      data-host={officeState?.host ?? "word"}
      style={themeStyle}
    >
      <Header
        host={officeState?.host}
        connectionState={connectionState}
        chatSubject={chatSubject}
        historyOpen={historyOpen}
        onSettingsClick={() => setSettingsOpen(true)}
        onNewChat={() => {
          if (messages.some((m) => m.role === "user" || m.role === "assistant")) {
            const chatId = activeChatId ?? crypto.randomUUID();
            saveChat(chatId, messages, chatSubject, officeState);
          }
          setMessages([]);
          clearCheckpoints();
          setDraft("");
          setChatSubject(undefined);
          setActiveChatId(undefined);
          assistantEndCountRef.current = 0;
          subjectDerivedRef.current = false;
          setHistoryOpen(false);
          void openSession(undefined, true).catch((error) => {
            pushErrorMessage(`Session reset failed: ${error instanceof Error ? error.message : String(error)}`);
          });
        }}
        onHistoryToggle={() => setHistoryOpen((c) => !c)}
      />

      {historyOpen && (
        <HistoryDropdown
          officeState={officeState}
          listChats={listChats}
          onDeleteChat={(chatId) => {
            deleteChat(chatId);
            if (activeChatId === chatId) {
              setMessages([]);
              clearCheckpoints();
              setDraft("");
              setChatSubject(undefined);
              setActiveChatId(undefined);
              assistantEndCountRef.current = 0;
              subjectDerivedRef.current = false;
              void openSession(undefined, true).catch((error) => {
                pushErrorMessage(`Session reset failed: ${error instanceof Error ? error.message : String(error)}`);
              });
            }
          }}
          onSelectChat={(chatId) => {
            const entry = loadChat(chatId);
            if (!entry) return;
            // Save current chat before switching
            if (messages.some((m) => m.role === "user" || m.role === "assistant")) {
              const currentId = activeChatId ?? crypto.randomUUID();
              saveChat(currentId, messages, chatSubject, officeState);
            }
            setMessages(entry.messages);
            setChatSubject(entry.subject);
            setActiveChatId(entry.chatId);
            assistantEndCountRef.current = entry.messages.filter((m) => m.role === "assistant").length;
            subjectDerivedRef.current = Boolean(entry.subject);
            setHistoryOpen(false);
            void openSession(undefined, true).catch((error) => {
              pushErrorMessage(`Session restore failed: ${error instanceof Error ? error.message : String(error)}`);
            });
          }}
          onClose={() => setHistoryOpen(false)}
        />
      )}

      <ContextBar
        officeState={officeState}
        mode={mode}
        shouldAttachDraftVisuals={shouldAttachDraftVisuals}
      />

      <ChatView
        messages={messages}
        hasConversation={hasConversation}
        isBusy={isBusy}
        showThinkingTraces={preferences.showThinkingTraces}
        officeTheme={officeTheme}
        officeState={officeState}
        onQuickPrompt={setDraft}
        onRewind={handleRewind}
        hasCheckpoint={hasCheckpoint}
        scrollDeps={[isBusy, localQueue, messages, remoteQueue]}
      />

      <QueueStrip
        localQueue={localQueue}
        remoteQueue={remoteQueue}
        selectedLocalQueueId={selectedLocalQueueId}
        onQueueSelection={handleQueueSelection}
      />

      <Composer
        draft={draft}
        setDraft={setDraft}
        userMessages={userMessages}
        sessionId={sessionId}
        isBusy={isBusy}
        activeToolName={activeToolName}
        officeState={officeState}
        shouldAttachDraftVisuals={shouldAttachDraftVisuals}
        selectedLocalQueueId={selectedLocalQueueId}
        configuredModels={configuredModels}
        selectedModelKey={selectedModelKey}
        thinkingLevel={thinkingLevel}
        availableThinkingLevels={availableThinkingLevels}
        sessionStats={sessionStats}
        autonomyLevel={preferences.autonomyLevel}
        onSend={sendDraftPrompt}
        onAbort={abortTurn}
        onQueue={queueCurrentDraft}
        onSteer={sendSelectedQueueAsSteer}
        onSelectModel={handleSelectModel}
        onSetThinkingLevel={handleSetThinkingLevel}
        onSetAutonomyLevel={handleSetAutonomyLevel}
      />

      {askUserRequest && (
        <AskUserPopup
          request={askUserRequest}
          onSubmit={handleAskUserSubmit}
        />
      )}

      {toolPermissionRequest && (
        <ToolPermissionPopup
          request={toolPermissionRequest}
          onDecision={handleToolPermissionDecision}
        />
      )}

      {editProposal && (
        <EditProposalPopup
          proposal={editProposal}
          onDecision={handleEditProposalDecision}
        />
      )}
    </div>
  );
}
