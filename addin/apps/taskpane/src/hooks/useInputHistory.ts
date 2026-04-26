import { useCallback, useRef } from "react";
import type { KeyboardEvent } from "react";

function isOnFirstLine(textarea: HTMLTextAreaElement): boolean {
  return textarea.value.lastIndexOf("\n", textarea.selectionStart - 1) === -1;
}

function isOnLastLine(textarea: HTMLTextAreaElement): boolean {
  return textarea.value.indexOf("\n", textarea.selectionStart) === -1;
}

export function useInputHistory(
  userMessages: string[],
  setDraft: (value: string) => void,
) {
  // -1 = current draft; 0 = most recent user message; increases toward oldest
  const historyIndexRef = useRef(-1);
  const savedDraftRef = useRef("");

  const handleHistoryNav = useCallback(
    (event: KeyboardEvent<HTMLTextAreaElement>) => {
      if (event.key !== "ArrowUp" && event.key !== "ArrowDown") return false;

      const textarea = event.currentTarget;
      const hasSelection = textarea.selectionStart !== textarea.selectionEnd;
      if (hasSelection) return false;

      if (event.key === "ArrowUp") {
        if (!isOnFirstLine(textarea)) return false;
        if (userMessages.length === 0) return false;

        let nextIndex: number;
        if (historyIndexRef.current === -1) {
          savedDraftRef.current = textarea.value;
          nextIndex = 0;
        } else if (historyIndexRef.current < userMessages.length - 1) {
          nextIndex = historyIndexRef.current + 1;
        } else {
          // Already at oldest — consume the event to avoid focus leaving textarea
          event.preventDefault();
          return true;
        }

        historyIndexRef.current = nextIndex;
        // userMessages is oldest-first; index 0 = most recent
        const msg = userMessages[userMessages.length - 1 - nextIndex] ?? "";
        event.preventDefault();
        setDraft(msg);
        requestAnimationFrame(() => {
          textarea.setSelectionRange(msg.length, msg.length);
        });
        return true;
      }

      // ArrowDown
      if (!isOnLastLine(textarea)) return false;
      if (historyIndexRef.current === -1) return false;

      let nextIndex: number;
      let nextText: string;
      if (historyIndexRef.current === 0) {
        nextIndex = -1;
        nextText = savedDraftRef.current;
      } else {
        nextIndex = historyIndexRef.current - 1;
        nextText = userMessages[userMessages.length - 1 - nextIndex] ?? "";
      }

      historyIndexRef.current = nextIndex;
      event.preventDefault();
      setDraft(nextText);
      requestAnimationFrame(() => {
        textarea.setSelectionRange(nextText.length, nextText.length);
      });
      return true;
    },
    [userMessages, setDraft],
  );

  const resetHistory = useCallback(() => {
    historyIndexRef.current = -1;
    savedDraftRef.current = "";
  }, []);

  return { handleHistoryNav, resetHistory };
}
