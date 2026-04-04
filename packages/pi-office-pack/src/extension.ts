import type { ExtensionFactory } from "@mariozechner/pi-coding-agent";
import { Type } from "@sinclair/typebox";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { HOST_LABELS } from "./defaults.js";
import type {
  AskUserQuestion,
  AskUserQuestionAnswer,
  AskUserRequest,
  AskUserResponse,
  ImageReasoningEffort,
  OfficeApplyEditParams,
  OfficeContextPayload,
  OfficeHost,
  OfficeNavigateParams,
  OfficeStateUpdate,
  OfficeToolName,
  OfficeVisualSnapshot,
} from "./protocol.js";

export interface ImageGenToolResult {
  base64: string;
  mimeType: string;
  width: number;
  height: number;
  modelKey: string;
  modelName: string;
}

export interface ImageGenToolParams {
  prompt: string;
  aspectRatio?: string | undefined;
  size?: string | undefined;
  quality?: string | undefined;
  insert?: boolean | undefined;
}

export interface OfficeExtensionOptions {
  getHost(): OfficeHost;
  getState(): OfficeStateUpdate | undefined;
  isToolDisabled?(toolName: string): boolean;
  invokeTool(toolName: OfficeToolName, params: Record<string, unknown>): Promise<unknown>;
  invokeAskUser(request: AskUserRequest): Promise<AskUserResponse>;
  invokeEditProposal?(proposal: import("./protocol.js").OfficeEditProposal): Promise<import("./protocol.js").OfficeEditProposalDecision>;
  isImageGenerationEnabled?(): boolean;
  getImageReasoningEffort?(): ImageReasoningEffort;
  getDefaultImageModel?(): string;
  generateImage?(params: ImageGenToolParams): Promise<ImageGenToolResult>;
}

function stripBinaryData(result: unknown): unknown {
  if (!result || typeof result !== "object") {
    return result;
  }

  if (Array.isArray(result)) {
    return result.map((entry) => stripBinaryData(entry));
  }

  const next = { ...(result as Record<string, unknown>) };
  if (Array.isArray(next.visuals)) {
    next.visuals = next.visuals.map((visual) => {
      if (!visual || typeof visual !== "object") {
        return visual;
      }

      const copy = { ...(visual as Record<string, unknown>) };
      if (typeof copy.data === "string") {
        copy.data = `[base64 ${copy.data.length} chars]`;
      }
      return copy;
    });
  }

  return Object.fromEntries(Object.entries(next).map(([key, value]) => [key, stripBinaryData(value)]));
}

function isContextPayload(result: unknown): result is OfficeContextPayload {
  return Boolean(result && typeof result === "object" && "summary" in result && "state" in result);
}

function hasVisuals(result: unknown): result is { visuals: OfficeVisualSnapshot[] } {
  return Boolean(
    result &&
      typeof result === "object" &&
      Array.isArray((result as { visuals?: unknown }).visuals) &&
      (result as { visuals: unknown[] }).visuals.every(
        (visual) =>
          visual &&
          typeof visual === "object" &&
          typeof (visual as OfficeVisualSnapshot).data === "string" &&
          typeof (visual as OfficeVisualSnapshot).mimeType === "string",
      ),
  );
}

function buildReadableParagraphMap(paragraphs: unknown): string | undefined {
  if (!Array.isArray(paragraphs) || !paragraphs.length) return undefined;
  const lines: string[] = [];
  for (const p of paragraphs.slice(0, 40)) {
    if (!p || typeof p !== "object") continue;
    const record = p as Record<string, unknown>;
    const text = typeof record.text === "string" ? record.text.trim() : "";
    const style = typeof record.style === "string" ? record.style : "";
    const isHeading = record.isHeading === true;
    if (!text && !style) continue;
    const prefix = isHeading ? (style.match(/\d/) ? "#".repeat(Number(style.match(/\d/)?.[0]) || 1) + " " : "# ") : "";
    lines.push(`${prefix}${text || "(empty)"}  [${style || "Normal"}]`);
  }
  return lines.length ? lines.join("\n") : undefined;
}

function toToolText(result: unknown): string {
  if (typeof result === "string") {
    return result;
  }

  if (isContextPayload(result)) {
    const snippets = (result as { snippets?: Record<string, unknown> }).snippets;
    const paragraphPreview = snippets ? buildReadableParagraphMap(snippets.paragraphs) : undefined;
    const structured = JSON.stringify(stripBinaryData(result), null, 2);
    return [
      result.summary,
      paragraphPreview ? `\nDocument structure:\n${paragraphPreview}` : undefined,
      `\nStructured data:\n${structured}`,
    ].filter(Boolean).join("\n");
  }

  return JSON.stringify(stripBinaryData(result), null, 2);
}

function toToolContent(result: unknown) {
  const content: Array<{ type: "text"; text: string } | { type: "image"; data: string; mimeType: string }> = [
    { type: "text", text: toToolText(result) },
  ];

  if (hasVisuals(result)) {
    for (const visual of result.visuals) {
      content.push({
        type: "image",
        data: visual.data,
        mimeType: visual.mimeType,
      });
    }
  }

  return content;
}

const getContextParams = Type.Object({
  scope: Type.Optional(
    Type.String({
      description:
        "Optional context hint (selection, document, worksheet, workbook, slide, presentation). Scope filtering is currently strongest for Excel and may be treated as a hint for Word/PowerPoint.",
    }),
  )
});

const applyEditParams = Type.Object({
  mode: Type.Optional(Type.String({ description: "Legacy edit mode such as replaceSelection, insertAfterSelection, or setRangeValues." })),
  content: Type.Optional(Type.String({ description: "Legacy text, HTML, or JSON matrix payload to insert into Office." })),
  text: Type.Optional(Type.String({ description: "Alias for legacy text content. Use when operation/type is insertText." })),
  html: Type.Optional(Type.String({ description: "Alias for legacy HTML content. Use when operation/type is insertHtml." })),
  format: Type.Optional(Type.String({ description: "Legacy content format such as text, html, or matrix." })),
  operation: Type.Optional(Type.String({ description: "Top-level action type alias, e.g., insertText, insertHtml, setRangeValues." })),
  type: Type.Optional(Type.String({ description: "Top-level action type alias when not wrapping with action.type." })),
  values: Type.Optional(Type.Any({ description: "2D array of values for setRangeValues actions." })),
  action: Type.Optional(
    Type.Any({
      description:
        "Structured host action. Prefer this over legacy mode/content for workbook, slide, comment, shape image, chart, table, and navigation-aware edits. Destructive actions should set action.options.confirmDestructive=true so the host can distinguish intentional deletes/clears from accidental ones.",
    }),
  ),
});

const navigateParams = Type.Object({
  target: Type.Optional(Type.String({ description: "Legacy anchor label or text to navigate to." })),
  kind: Type.Optional(
    Type.String({
      description:
        "heading, comment, revision, footnote, endnote, paragraph, field, contentControl, cell, range, sheet, workbook, slide, notesRegion, layout, slideMaster, shape, chart, or pivotTable.",
    }),
  ),
  anchor: Type.Optional(
    Type.Any({
      description: "Structured navigation anchor. Prefer this when sheet names, range addresses, slide IDs, or paragraph IDs are known.",
    }),
  ),
});

const captureSnapshotParams = Type.Object({
  scope: Type.Optional(
    Type.String({
      description:
        "Optional context hint (selection, document, worksheet, workbook, slide, shape). Scope is currently most effective for Excel and may be treated as a hint in Word/PowerPoint.",
    }),
  ),
  includeFormatting: Type.Optional(Type.Boolean({ description: "Include formatting and layout metadata alongside the visuals." })),
  maxImages: Type.Optional(Type.Number({ minimum: 0, maximum: 4, description: "Maximum number of visual snapshots to include." })),
});

const captureViewportParams = Type.Object({
  includeFormatting: Type.Optional(
    Type.Boolean({ description: "Include Word viewport metadata such as visible pages, scroll position, and view mode." }),
  ),
  includeWindowFrame: Type.Optional(
    Type.Boolean({
      description:
        "Reserved for future native capture support. In browser-only runtime this is acknowledged but cannot capture the full OS window frame.",
    }),
  ),
});

const readSectionParams = Type.Object({
  startIndex: Type.Optional(Type.Number({ description: "Zero-based paragraph start index (inclusive)." })),
  endIndex: Type.Optional(Type.Number({ description: "Zero-based paragraph end index (exclusive). Defaults to startIndex + 20." })),
  start: Type.Optional(Type.Number({ description: "Alias for startIndex." })),
  end: Type.Optional(Type.Number({ description: "Alias for endIndex." })),
  includeStyles: Type.Optional(Type.Boolean({ description: "Include paragraph styles and heading levels. Defaults to true." })),
});

const executeJsParams = Type.Object({
  code: Type.Optional(
    Type.String({
      description:
        "Office.js code to execute. Must use the active host run function (Word.run, Excel.run, or PowerPoint.run) and return a JSON-serializable value. The runtime enforces a best-effort restricted subset (regex checks only, not an isolated sandbox) and blocks network, storage, eval, and system-access patterns.",
    }),
  ),
  script: Type.Optional(Type.String({ description: "Alias for code." })),
});

const proposeEditsParams = Type.Object({
  edits: Type.Array(
    Type.Object({
      kind: Type.String({ description: "insert, replace, or delete." }),
      searchText: Type.Optional(Type.String({ description: "Text to locate in the document for replace/delete." })),
      oldText: Type.Optional(Type.String({ description: "Expected existing text (for replace/delete verification)." })),
      newText: Type.Optional(Type.String({ description: "Replacement text (for insert/replace)." })),
      anchor: Type.Optional(Type.String({ description: "Heading or paragraph label to scope the search." })),
      paragraphId: Type.Optional(Type.String({ description: "Paragraph unique ID for precise targeting." })),
      explanation: Type.Optional(Type.String({ description: "Brief rationale for this edit." })),
    }),
    { description: "Ordered list of proposed edits." },
  ),
  summary: Type.String({ description: "One-sentence summary of all proposed changes." }),
});

const askUserOptionSchema = Type.Union([
  Type.String(),
  Type.Object({
    title: Type.String(),
    description: Type.Optional(Type.String()),
  }),
]);

const askUserQuestionSchema = Type.Object({
  question: Type.String({ description: "The question to ask." }),
  context: Type.Optional(Type.String({ description: "Relevant context summary shown before the question." })),
  options: Type.Optional(Type.Array(askUserOptionSchema, { description: "Multiple-choice options." })),
});

const askUserParams = Type.Object({
  question: Type.Optional(Type.String({ description: "A single question to ask the user." })),
  context: Type.Optional(Type.String({ description: "Context for a single question." })),
  options: Type.Optional(Type.Array(askUserOptionSchema, { description: "Options for a single question." })),
  questions: Type.Optional(Type.Array(askUserQuestionSchema, { description: "Multiple questions to ask in a carousel." })),
});

const generateImageParams = Type.Object({
  prompt: Type.String({ description: "Description of the image to generate." }),
  aspectRatio: Type.Optional(
    Type.String({ description: "Aspect ratio: 1:1, 16:9, 9:16, 4:3, 3:4, 3:2, 2:3. Defaults based on document type." }),
  ),
  size: Type.Optional(Type.String({ description: "Resolution tier: 1K, 2K, or 4K. Auto-resolved based on document context." })),
  quality: Type.Optional(Type.String({ description: "Quality: auto, low, medium, high. Defaults to auto." })),
  insert: Type.Optional(
    Type.Boolean({ description: "Whether to insert the generated image into the document. Defaults to true." }),
  ),
});

export function getOfficeSkillPaths(): string[] {
  const distDir = dirname(fileURLToPath(import.meta.url));
  return [
    join(distDir, "../skills/office-host.SKILL.md"),
    join(distDir, "../skills/workspace-handoff.SKILL.md"),
  ];
}

export function createOfficeExtension(options: OfficeExtensionOptions): ExtensionFactory {
  return (pi) => {
    const isDisabled = options.isToolDisabled ?? (() => false);

    if (!isDisabled("office_get_context"))
    pi.registerTool({
      name: "office_get_context",
      label: "Office Context",
      description: "Read the current Office document or selection context from the active host.",
      parameters: getContextParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("office_get_context", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("office_apply_edit"))
    pi.registerTool({
      name: "office_apply_edit",
      label: "Office Edit",
      description: "Apply a native edit to the active Office document, selection, range, worksheet, or slide.",
      parameters: applyEditParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("office_apply_edit", params as OfficeApplyEditParams as Record<string, unknown>);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("office_navigate"))
    pi.registerTool({
      name: "office_navigate",
      label: "Office Navigate",
      description: "Move the user back to a document anchor such as a heading, comment, slide, or range when supported.",
      parameters: navigateParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("office_navigate", params as OfficeNavigateParams as Record<string, unknown>);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("office_capture_snapshot"))
    pi.registerTool({
      name: "office_capture_snapshot",
      label: "Office Snapshot",
      description: "Capture visual snapshots and formatting metadata for the current Office selection when layout fidelity matters.",
      parameters: captureSnapshotParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("office_capture_snapshot", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("office_capture_viewport"))
    pi.registerTool({
      name: "office_capture_viewport",
      label: "Office Viewport",
      description:
        "Capture Word viewport metadata (visible pages, scroll position, and view state) from Office.js context. Use this for layout-sensitive troubleshooting. This is not a pixel-perfect OS window screenshot and does not capture off-screen document content.",
      parameters: captureViewportParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("office_capture_viewport", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("office_read_section"))
    pi.registerTool({
      name: "office_read_section",
      label: "Read Document Section",
      description:
        "Read a range of paragraphs from the active Word document by paragraph index. Use when the document is too large to fit in context and you need to page through content. Returns paragraph text with styles. Start from index 0 and advance by 20 to page through the full document.",
      parameters: readSectionParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("office_read_section", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("office_execute_js"))
    pi.registerTool({
      name: "office_execute_js",
      label: "Execute Office.js",
      description:
        "Execute Office.js code directly in the active host as an escape hatch when structured tools cannot achieve the desired result. The code MUST use Word.run, Excel.run, or PowerPoint.run and return a JSON-serializable result. This tool is a best-effort restricted subset enforced by regex checks (not an isolated sandbox) and blocks network, storage, eval, and system-access patterns. Do not use it for simple edits that office_apply_edit can handle.",
      parameters: executeJsParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("office_execute_js", params);
        return {
          content: toToolContent(result),
          details: result,
        };
      },
    });

    if (!isDisabled("office_propose_edits"))
    pi.registerTool({
      name: "office_propose_edits",
      label: "Propose Document Edits",
      description:
        "Propose a batch of text edits to the active Word document for the user to review before applying. Each edit specifies a kind (insert/replace/delete), the text to find (searchText), and the replacement. The user sees a reviewable card for each edit and can accept, modify, or reject individually. CRITICAL: each edit's searchText MUST be under 200 characters. Break large paragraph rewrites into multiple small, targeted edits — one per sentence or distinct phrase. Never use a full paragraph as searchText. For example, instead of one edit replacing a 3-sentence paragraph, create 3 separate edits each targeting one sentence. Use this instead of office_apply_edit when making multi-paragraph changes or when the user should verify changes first.",
      parameters: proposeEditsParams,
      execute: async (_toolCallId, params) => {
        const result = await options.invokeTool("office_propose_edits", params);
        const resultObj = result as Record<string, unknown> | undefined;
        if (!resultObj || resultObj.error || !Array.isArray(resultObj.edits) || !options.invokeEditProposal) {
          return { content: toToolContent(result), details: result };
        }

        const proposal: import("./protocol.js").OfficeEditProposal = {
          requestId: crypto.randomUUID(),
          edits: resultObj.edits as import("./protocol.js").OfficeProposedEdit[],
          summary: String(resultObj.summary ?? "Proposed edits"),
        };

        const decision = await options.invokeEditProposal(proposal);
        const accepted = decision.decisions.filter((d) => d.accepted);
        const rejected = decision.decisions.filter((d) => !d.accepted);

        const reasonLabels: Record<string, string> = {
          keep_original: "Keep original",
          rewrite_differently: "Rewrite differently",
          not_relevant: "Not relevant",
        };

        const appResult = decision.applicationResult;
        let applicationLine: string;
        if (appResult) {
          if (appResult.failed > 0) {
            applicationLine = `Application result: ${appResult.applied} applied, ${appResult.failed} failed. Errors: ${appResult.errors.join("; ")}`;
          } else if (appResult.applied > 0) {
            applicationLine = `All ${appResult.applied} accepted edit${appResult.applied !== 1 ? "s" : ""} applied successfully.`;
          } else {
            applicationLine = "No edits were applied.";
          }
        } else {
          applicationLine = accepted.length ? "Accepted edits were applied to the document." : "No edits were applied.";
        }

        const lines: string[] = [
          `Edit proposal reviewed: ${accepted.length} accepted, ${rejected.length} rejected out of ${proposal.edits.length} edits.`,
          applicationLine,
        ];

        if (rejected.length) {
          lines.push("", "Rejected edits:");
          for (const d of rejected) {
            const edit = proposal.edits.find((e) => e.id === d.editId);
            const preview = (edit?.searchText ?? edit?.oldText ?? "").slice(0, 60);
            const reason = d.rejectReason ? reasonLabels[d.rejectReason] ?? d.rejectReason : "No reason given";
            const note = d.rejectNote ? ` (note: "${d.rejectNote}")` : "";
            lines.push(`- "${preview}${preview.length >= 60 ? "..." : ""}": ${reason}${note}`);
          }
        }

        if (decision.globalFeedback) {
          lines.push("", `User feedback: "${decision.globalFeedback}"`);
        }

        return {
          content: [{ type: "text" as const, text: lines.join("\n") }],
          details: {
            proposalId: proposal.requestId,
            summary: proposal.summary,
            totalEdits: proposal.edits.length,
            accepted: accepted.length,
            rejected: rejected.length,
            decisions: decision.decisions,
            globalFeedback: decision.globalFeedback,
          },
        };
      },
    });

    if (!isDisabled("ask_user"))
    pi.registerTool({
      name: "ask_user",
      label: "Ask User",
      description:
        "Ask the user one or more questions with selectable options. Use this proactively when the task involves subjective choices, ambiguous intent, or assumptions that would materially change the result (e.g. tone, audience, format, scope, style). After receiving the user's answers, immediately proceed to execute the original task using those answers — do not stop at acknowledging the choices. Supports single or multiple questions in one call.",
      parameters: askUserParams,
      execute: async (_toolCallId, params) => {
        const typed = params as {
          question?: string;
          context?: string;
          options?: (string | { title: string; description?: string })[];
          questions?: { question: string; context?: string; options?: (string | { title: string; description?: string })[] }[];
        };

        const normalizeOptions = (raw?: (string | { title: string; description?: string })[]) =>
          (raw ?? []).map((o) => (typeof o === "string" ? { title: o } : o));

        const questions: AskUserQuestion[] = [];
        if (typed.questions?.length) {
          for (const q of typed.questions) {
            questions.push({
              id: crypto.randomUUID(),
              question: q.question,
              context: q.context,
              options: normalizeOptions(q.options),
            });
          }
        } else if (typed.question) {
          questions.push({
            id: crypto.randomUUID(),
            question: typed.question,
            context: typed.context,
            options: normalizeOptions(typed.options),
          });
        } else {
          return {
            content: [{ type: "text" as const, text: "No question provided." }],
            details: { error: "missing_question" } as Record<string, unknown>,
          };
        }

        const request: AskUserRequest = { requestId: crypto.randomUUID(), questions };
        const response = await options.invokeAskUser(request);

        const lines = response.answers.map((a) => {
          const q = questions.find((q) => q.id === a.questionId);
          const answer = a.selectedOption ?? "None of the above";
          const notes = a.notes ? ` (notes: ${a.notes})` : "";
          return `Q: ${q?.question ?? "?"}\nA: ${answer}${notes}`;
        });

        return {
          content: [{ type: "text" as const, text: lines.join("\n\n") }],
          details: { questions, answers: response.answers } as Record<string, unknown>,
        };
      },
    });

    if (options.generateImage && !isDisabled("generate_image")) {
      const generateImage = options.generateImage;
      const isEnabled = options.isImageGenerationEnabled ?? (() => true);
      pi.registerTool({
        name: "generate_image",
        label: "Generate Image",
        description:
          "Generate an AI image from a text prompt and optionally insert it into the active Office document. Use for photos, illustrations, graphics, and visual content. Do not use for diagrams or charts.",
        parameters: generateImageParams,
        execute: async (_toolCallId, params) => {
          if (!isEnabled()) {
            return {
              content: [{ type: "text" as const, text: "Image generation is disabled. The user can enable it in Settings > Preferences > Image Generation." }],
              details: { enabled: false } as Record<string, unknown>,
            };
          }

          const typedParams = params as {
            prompt: string;
            aspectRatio?: string;
            size?: string;
            quality?: string;
            insert?: boolean;
          };

          const result = await generateImage({
            prompt: typedParams.prompt,
            aspectRatio: typedParams.aspectRatio,
            size: typedParams.size,
            quality: typedParams.quality,
            insert: typedParams.insert,
          });

          const shouldInsert = typedParams.insert !== false;
          if (shouldInsert) {
            try {
              await options.invokeTool("office_apply_edit", {
                action: {
                  type: "insertInlinePicture",
                  content: result.base64,
                  placement: "after",
                  options: { altText: typedParams.prompt.slice(0, 120) },
                },
              } as Record<string, unknown>);
            } catch (insertError) {
              const msg = insertError instanceof Error ? insertError.message : String(insertError);
              return {
                content: [
                  { type: "text" as const, text: `Image generated with ${result.modelName} (${result.width}x${result.height}) but insertion failed: ${msg}. The image is shown below.` },
                  { type: "image" as const, data: result.base64, mimeType: result.mimeType },
                ],
                details: { ...result, inserted: false, insertError: msg } as Record<string, unknown>,
              };
            }
          }

          return {
            content: [
              {
                type: "text" as const,
                text: `Image generated with ${result.modelName} (${result.width}x${result.height}).${shouldInsert ? " Inserted into document." : ""}`,
              },
              { type: "image" as const, data: result.base64, mimeType: result.mimeType },
            ],
            details: { ...result, inserted: shouldInsert } as Record<string, unknown>,
          };
        },
      });
    }

    pi.on("session_start", async (_event, ctx) => {
      const state = options.getState();
      const label = HOST_LABELS[options.getHost()];
      const suffix = state?.document.saved ? "workspace mode" : "document-only mode";
      ctx.ui.setStatus("office-host", `${label} · ${suffix}`);
    });
  };
}
