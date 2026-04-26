import { useEffect, useRef, useState } from "react";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import remarkMath from "remark-math";
import rehypeKatex from "rehype-katex";
import type { OfficeStateUpdate } from "@pi-office/pi-office-pack/protocol";
import type { ChatEntry } from "../../lib/helpers";
import { getQuickPrompts } from "../../lib/helpers";
import type { OfficeThemeSnapshot } from "../../lib/office";
import { DiagramBlock } from "./DiagramBlock";
import { CodeBlock } from "./CodeBlock";
import { ImageBlock } from "./ImageBlock";
import { ThinkingBlock } from "./ThinkingBlock";
import { ToolCallBlock } from "./ToolCallBlock";
import { RewindDialog, type RewindMode } from "./RewindDialog";

const DRAWIO_MARKERS = ["<mxfile", "<mxGraphModel", "<diagram"];
const markdownComponents = {
  a: (props: any) => {
    const { node: _node, ...rest } = props;
    return <a {...rest} target="_blank" rel="noreferrer" />;
  },
};

type MarkdownSegment =
  | { type: "markdown"; text: string; key: string }
  | { type: "code"; code: string; lang: string | undefined; key: string };

function isDiagramLanguage(lang: string | undefined): "mermaid" | "drawio" | null {
  if (!lang) return null;
  const lower = lang.toLowerCase();
  if (lower === "mermaid") return "mermaid";
  if (lower === "drawio" || lower === "draw.io" || lower === "xml-drawio") return "drawio";
  return null;
}

function looksLikeDrawio(code: string): boolean {
  return DRAWIO_MARKERS.some((marker) => code.trimStart().startsWith(marker));
}

function splitMarkdownSegments(text: string): MarkdownSegment[] {
  const segments: MarkdownSegment[] = [];
  const fencePattern = /```([^\n`]*)\n([\s\S]*?)\n```(?=\n|$)/g;
  let lastIndex = 0;
  let match: RegExpExecArray | null;

  while ((match = fencePattern.exec(text)) !== null) {
    if (match.index > lastIndex) {
      segments.push({
        type: "markdown",
        text: text.slice(lastIndex, match.index),
        key: `md-${lastIndex}`,
      });
    }

    segments.push({
      type: "code",
      lang: match[1]?.trim() || undefined,
      code: match[2] ?? "",
      key: `code-${match.index}`,
    });
    lastIndex = fencePattern.lastIndex;
  }

  if (lastIndex < text.length) {
    segments.push({
      type: "markdown",
      text: text.slice(lastIndex),
      key: `md-${lastIndex}`,
    });
  }

  return segments.length ? segments : [{ type: "markdown", text, key: "md-0" }];
}

function MarkdownBody({
  text,
  officeTheme,
  officeState,
}: {
  text: string;
  officeTheme: OfficeThemeSnapshot | undefined;
  officeState: OfficeStateUpdate | undefined;
}) {
  const segments = splitMarkdownSegments(text);

  return (
    <div className="markdown-body">
      {segments.map((segment) => {
        if (segment.type === "markdown") {
          return (
            <ReactMarkdown
              key={segment.key}
              remarkPlugins={[remarkGfm, remarkMath]}
              rehypePlugins={[rehypeKatex]}
              components={markdownComponents}
            >
              {segment.text}
            </ReactMarkdown>
          );
        }

        const diagramType =
          isDiagramLanguage(segment.lang) ??
          (segment.lang === "xml" && looksLikeDrawio(segment.code) ? "drawio" : null);

        if (diagramType) {
          return (
            <DiagramBlock
              key={segment.key}
              code={segment.code}
              type={diagramType}
              officeTheme={officeTheme}
              officeState={officeState}
            />
          );
        }

        if (segment.lang) {
          return (
            <CodeBlock
              key={segment.key}
              code={segment.code.replace(/\n$/, "")}
              language={segment.lang}
              officeTheme={officeTheme}
            />
          );
        }

        return (
          <pre key={segment.key}>
            <code>{segment.code}</code>
          </pre>
        );
      })}
    </div>
  );
}

interface ChatViewProps {
  messages: ChatEntry[];
  hasConversation: boolean;
  isBusy: boolean;
  showThinkingTraces: boolean;
  officeTheme: OfficeThemeSnapshot | undefined;
  officeState: OfficeStateUpdate | undefined;
  onQuickPrompt: (text: string) => void;
  onRewind?: (messageId: string, mode: RewindMode) => void;
  hasCheckpoint?: (messageId: string) => boolean;
  scrollDeps: unknown[];
}

export function ChatView({
  messages,
  hasConversation,
  isBusy,
  showThinkingTraces,
  officeTheme,
  officeState,
  onQuickPrompt,
  onRewind,
  hasCheckpoint,
  scrollDeps,
}: ChatViewProps) {
  const scrollRef = useRef<HTMLDivElement | null>(null);
  const autoScrollRef = useRef(true);
  const [rewindTargetId, setRewindTargetId] = useState<string | null>(null);
  const quickPrompts = getQuickPrompts(officeState);

  const isNearBottom = (element: HTMLDivElement) =>
    element.scrollHeight - element.scrollTop - element.clientHeight <= 60;

  const focusTaskpane = () => {
    window.focus();
  };

  const handleScroll = () => {
    const element = scrollRef.current;
    if (!element) return;
    autoScrollRef.current = isNearBottom(element);
  };

  useEffect(() => {
    const element = scrollRef.current;
    if (!element || !autoScrollRef.current) return;
    element.scrollTo({
      top: element.scrollHeight,
      behavior: "smooth",
    });
  }, scrollDeps);

  return (
    <main className="chat-stage">
      <div
        ref={scrollRef}
        className="chat-scroll"
        tabIndex={0}
        onMouseDown={focusTaskpane}
        onMouseEnter={focusTaskpane}
        onScroll={handleScroll}
        onWheelCapture={(event) => event.stopPropagation()}
        onKeyDown={(event) => event.stopPropagation()}
      >
        {!hasConversation && (
          <section className="empty-state">
            <p className="empty-copy">
              Ask Pi to edit the document natively, inspect the current selection, or pull in
              folder context once the file has been saved.
            </p>
            <div className="prompt-grid">
              {quickPrompts.map((prompt) => (
                <button
                  key={prompt}
                  type="button"
                  className="prompt-chip"
                  onClick={() => onQuickPrompt(prompt)}
                >
                  {prompt}
                </button>
              ))}
            </div>
          </section>
        )}

        {messages.map((message) => {
          const roleClass =
            message.role === "user" ? "user"
            : message.role === "assistant" ? "assistant"
            : message.role === "error" ? "error"
            : "system";
          const hasVisibleThinking = showThinkingTraces && Boolean(message.thinking?.trim());
          const hasTextContent = Boolean(message.text?.trim());
          const showMetaExpanded = Boolean(message.pending) && !hasTextContent;
          const isThinkingActive = showMetaExpanded && hasVisibleThinking;
          const showAssistantBody =
            message.role === "assistant" &&
            (Boolean(message.text) || (!hasVisibleThinking && !(message.toolCalls?.length)));
          const showTailCursor = Boolean(message.pending) && !isThinkingActive && hasTextContent;
          return (
            <article key={message.id} className={`message-row message-row-${roleClass}`}>
              {message.role === "error" ? (
                <div className="message message-error">
                  <svg viewBox="0 0 20 20" width="14" height="14" className="message-error-icon" aria-hidden="true">
                    <path d="M10 0C4.48 0 0 4.48 0 10s4.48 10 10 10 10-4.48 10-10S15.52 0 10 0zm1 15H9v-2h2v2zm0-4H9V5h2v6z" fill="currentColor" />
                  </svg>
                  <div className="message-body">{message.text}</div>
                </div>
              ) : (
                <>
                  <span className="message-role">
                    {message.role === "assistant" ? "Pi" : message.role}
                    {message.role === "user" && !isBusy && hasCheckpoint?.(message.id) && (
                      <button
                        type="button"
                        className="rewind-button"
                        aria-label="Rewind to before this message"
                        title="Rewind to before this message"
                        onClick={() => setRewindTargetId(message.id)}
                      >
                        <svg viewBox="0 0 16 16" width="12" height="12" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
                          <path d="M2 8a6 6 0 1 1 1.76 4.24" />
                          <polyline points="2 4 2 8 6 8" />
                        </svg>
                      </button>
                    )}
                  </span>
                  {rewindTargetId === message.id && (
                    <RewindDialog
                      messageText={message.text}
                      host={officeState?.host}
                      onSelect={(mode) => {
                        setRewindTargetId(null);
                        onRewind?.(message.id, mode);
                      }}
                      onClose={() => setRewindTargetId(null)}
                    />
                  )}
                  <div className={`message message-${roleClass}`}>
                    {message.role === "assistant" ? (
                      <>
                        {showThinkingTraces ? (
                          <ThinkingBlock
                            thinking={message.thinking ?? ""}
                            segments={message.thinkingSegments ?? []}
                            pending={showMetaExpanded}
                          />
                        ) : null}
                        <ToolCallBlock toolCalls={message.toolCalls ?? []} pending={showMetaExpanded} />
                        {showAssistantBody ? (
                          <>
                            <MarkdownBody text={message.text || " "} officeTheme={officeTheme} officeState={officeState} />
                            {showTailCursor && <span className="block-cursor" aria-hidden="true" />}
                          </>
                        ) : null}
                        {message.images?.map((img, idx) => (
                          <ImageBlock
                            key={`img-${message.id}-${idx}`}
                            base64={img.data}
                            mimeType={img.mimeType}
                            width={0}
                            height={0}
                            modelName=""
                            prompt={message.text.slice(0, 120)}
                            inserted={message.text.includes("Inserted into document")}
                            officeState={officeState}
                          />
                        ))}
                      </>
                    ) : (
                      <div className="message-body">{message.text}</div>
                    )}
                    {message.pending && !hasTextContent && !isThinkingActive && (
                      <span className="block-cursor" aria-hidden="true" />
                    )}
                  </div>
                </>
              )}
            </article>
          );
        })}
      </div>
    </main>
  );
}
