import { useEffect, useRef, useState } from "react";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import type { ThinkingSegment } from "../../lib/helpers";

interface ThinkingBlockProps {
  thinking: string;
  segments: ThinkingSegment[];
  pending: boolean;
}

function formatDuration(ms: number): string {
  const seconds = Math.round(ms / 1000);
  if (seconds < 1) return "<1s";
  return `${seconds}s`;
}

function SegmentSummary({ segment, index }: { segment: ThinkingSegment; index: number }) {
  const label = segment.durationMs
    ? `Thought for ${formatDuration(segment.durationMs)}`
    : `Thought ${index + 1}`;
  return (
    <details className="thinking-segment-row">
      <summary className="thinking-segment-summary">
        <span className="thinking-segment-label">{label}</span>
        <span className="thinking-segment-chevron" aria-hidden="true" />
      </summary>
      <div className="thinking-segment-body">
        <ReactMarkdown remarkPlugins={[remarkGfm]}>{segment.text}</ReactMarkdown>
      </div>
    </details>
  );
}

export function ThinkingBlock({ thinking, segments, pending }: ThinkingBlockProps) {
  const content = thinking.trim();
  const prevSegmentCountRef = useRef(segments.length);
  const [erasing, setErasing] = useState(false);
  const [displayText, setDisplayText] = useState("");
  const eraseTimerRef = useRef<number | undefined>(undefined);

  // Detect when a new segment starts (previous one was closed) to trigger erase animation
  useEffect(() => {
    if (!pending) {
      prevSegmentCountRef.current = segments.length;
      return;
    }

    const prevCount = prevSegmentCountRef.current;
    const currCount = segments.length;
    prevSegmentCountRef.current = currCount;

    // New segment just started while previous had text — trigger erase
    if (currCount > prevCount && prevCount > 0) {
      const prevSegment = segments[currCount - 2];
      if (prevSegment?.text?.trim()) {
        setErasing(true);
        setDisplayText(prevSegment.text.trim());

        // Erase the text character-by-character
        let text = prevSegment.text.trim();
        const eraseStep = () => {
          if (text.length <= 0) {
            setErasing(false);
            setDisplayText("");
            return;
          }
          // Erase in chunks for speed
          const chunkSize = Math.max(1, Math.ceil(text.length / 12));
          text = text.slice(0, -chunkSize);
          setDisplayText(text);
          eraseTimerRef.current = window.setTimeout(eraseStep, 25);
        };
        eraseTimerRef.current = window.setTimeout(eraseStep, 100);
      }
    }

    return () => {
      if (eraseTimerRef.current) window.clearTimeout(eraseTimerRef.current);
    };
  }, [pending, segments]);

  if (!content && !segments.length) return null;

  if (pending) {
    // During erase animation, show the shrinking text
    if (erasing) {
      const lastChar = displayText.length > 0 ? displayText[displayText.length - 1] : "";
      return (
        <div className="thinking-stream">
          <div className="thinking-stream-text thinking-stream-erasing">
            {displayText.slice(0, -1)}
            <span className="block-cursor-char" aria-hidden="true">{lastChar || "\u00A0"}</span>
          </div>
        </div>
      );
    }

    const activeSegment = segments.length ? segments[segments.length - 1] : null;
    const activeText = activeSegment?.text?.trim() || content;
    if (!activeText) return null;

    const lastChar = activeText[activeText.length - 1] ?? "";
    const textBeforeCursor = activeText.slice(0, -1);

    return (
      <div className="thinking-stream">
        <div className="thinking-stream-text">
          <ReactMarkdown remarkPlugins={[remarkGfm]}>{textBeforeCursor}</ReactMarkdown>
          <span className="block-cursor-char" aria-hidden="true">{lastChar}</span>
        </div>
      </div>
    );
  }

  if (!segments.length) {
    return (
      <details className="thinking-segment-row">
        <summary className="thinking-segment-summary">
          <span className="thinking-segment-label">Reasoning</span>
          <span className="thinking-segment-chevron" aria-hidden="true" />
        </summary>
        <div className="thinking-segment-body">
          <ReactMarkdown remarkPlugins={[remarkGfm]}>{content}</ReactMarkdown>
        </div>
      </details>
    );
  }

  return (
    <div className="thinking-segments">
      {segments.map((segment, i) => (
        <SegmentSummary key={`seg-${i}-${segment.startedAt}`} segment={segment} index={i} />
      ))}
    </div>
  );
}
