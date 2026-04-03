import { useCallback, useState } from "react";
import type {
  AskUserRequest,
  AskUserQuestionAnswer,
} from "@pi-office/pi-office-pack/protocol";

interface AskUserPopupProps {
  request: AskUserRequest;
  onSubmit: (answers: AskUserQuestionAnswer[]) => void;
}

export function AskUserPopup({ request, onSubmit }: AskUserPopupProps) {
  const { questions } = request;
  const [currentIndex, setCurrentIndex] = useState(0);
  const [answers, setAnswers] = useState<Map<string, AskUserQuestionAnswer>>(() => new Map());
  const [notesExpanded, setNotesExpanded] = useState<Set<string>>(() => new Set());

  const question = questions[currentIndex];
  if (!question) return null;

  const currentAnswer = answers.get(question.id);
  const allAnswered = questions.every((q) => answers.has(q.id));

  const updateAnswer = useCallback(
    (questionId: string, update: Partial<AskUserQuestionAnswer>) => {
      setAnswers((prev) => {
        const next = new Map(prev);
        const existing = next.get(questionId) ?? { questionId, selectedOption: null };
        next.set(questionId, { ...existing, ...update });
        return next;
      });
    },
    [],
  );

  const toggleNotes = useCallback((questionId: string) => {
    setNotesExpanded((prev) => {
      const next = new Set(prev);
      if (next.has(questionId)) next.delete(questionId);
      else next.add(questionId);
      return next;
    });
  }, []);

  const handleSubmit = useCallback(() => {
    updateAnswer(question.id, currentAnswer?.selectedOption !== undefined ? {} : { selectedOption: null });
    const finalAnswers = new Map(answers);
    if (!finalAnswers.has(question.id)) {
      finalAnswers.set(question.id, { questionId: question.id, selectedOption: null });
    }

    if (questions.length === 1 || (allAnswered || questions.every((q) => finalAnswers.has(q.id)))) {
      onSubmit(Array.from(finalAnswers.values()));
    } else {
      const nextUnanswered = questions.findIndex((q, i) => i > currentIndex && !finalAnswers.has(q.id));
      if (nextUnanswered !== -1) {
        setCurrentIndex(nextUnanswered);
      } else {
        const firstUnanswered = questions.findIndex((q) => !finalAnswers.has(q.id));
        if (firstUnanswered !== -1) setCurrentIndex(firstUnanswered);
        else onSubmit(Array.from(finalAnswers.values()));
      }
    }
  }, [answers, allAnswered, currentAnswer, currentIndex, onSubmit, question, questions, updateAnswer]);

  const isShowingNotes = notesExpanded.has(question.id);

  return (
    <div className="ask-user-overlay">
      <div className="ask-user-popup">
        <div className="ask-user-header">
          <span className="ask-user-title">Pi needs your input</span>
          {questions.length > 1 && (
            <span className="ask-user-counter">
              {currentIndex + 1} / {questions.length}
            </span>
          )}
        </div>

        {question.context && (
          <div className="ask-user-context">{question.context}</div>
        )}

        <div className="ask-user-question">{question.question}</div>

        <div className="ask-user-options">
          {question.options.map((opt) => {
            const isSelected = currentAnswer?.selectedOption === opt.title;
            return (
              <button
                key={opt.title}
                type="button"
                className={`ask-user-option ${isSelected ? "ask-user-option-selected" : ""}`}
                onClick={() => updateAnswer(question.id, { questionId: question.id, selectedOption: opt.title })}
              >
                <span className="ask-user-option-radio">{isSelected ? "\u25C9" : "\u25CB"}</span>
                <span className="ask-user-option-content">
                  <span className="ask-user-option-title">{opt.title}</span>
                  {opt.description && (
                    <span className="ask-user-option-desc">{opt.description}</span>
                  )}
                </span>
              </button>
            );
          })}
          <button
            type="button"
            className={`ask-user-option ${currentAnswer?.selectedOption === null && answers.has(question.id) ? "ask-user-option-selected" : ""}`}
            onClick={() => updateAnswer(question.id, { questionId: question.id, selectedOption: null })}
          >
            <span className="ask-user-option-radio">
              {currentAnswer?.selectedOption === null && answers.has(question.id) ? "\u25C9" : "\u25CB"}
            </span>
            <span className="ask-user-option-content">
              <span className="ask-user-option-title">None of the above</span>
            </span>
          </button>
        </div>

        <div className="ask-user-notes-section">
          <button
            type="button"
            className="ask-user-notes-toggle"
            onClick={() => toggleNotes(question.id)}
          >
            <svg viewBox="0 0 16 16" width="12" height="12" className={`ask-user-chevron ${isShowingNotes ? "ask-user-chevron-open" : ""}`}>
              <path d="M6 4l4 4-4 4" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
            Add notes
          </button>
          {isShowingNotes && (
            <textarea
              className="ask-user-notes-area"
              placeholder="Add any additional notes..."
              value={currentAnswer?.notes ?? ""}
              onChange={(e) => updateAnswer(question.id, { questionId: question.id, notes: e.target.value || undefined })}
              rows={3}
            />
          )}
        </div>

        <div className="ask-user-footer">
          {questions.length > 1 && (
            <div className="ask-user-nav">
              <button
                type="button"
                className="ask-user-nav-btn"
                disabled={currentIndex === 0}
                onClick={() => setCurrentIndex((i) => i - 1)}
                aria-label="Previous question"
              >
                <svg viewBox="0 0 16 16" width="14" height="14">
                  <path d="M10 12L6 8l4-4" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                </svg>
              </button>
              <div className="ask-user-progress">
                {questions.map((q, i) => (
                  <span
                    key={q.id}
                    className={`ask-user-dot ${i === currentIndex ? "ask-user-dot-active" : ""} ${answers.has(q.id) ? "ask-user-dot-answered" : ""}`}
                    onClick={() => setCurrentIndex(i)}
                  />
                ))}
              </div>
              <button
                type="button"
                className="ask-user-nav-btn"
                disabled={currentIndex === questions.length - 1}
                onClick={() => setCurrentIndex((i) => i + 1)}
                aria-label="Next question"
              >
                <svg viewBox="0 0 16 16" width="14" height="14">
                  <path d="M6 4l4 4-4 4" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                </svg>
              </button>
            </div>
          )}
          <button
            type="button"
            className="ask-user-submit"
            onClick={handleSubmit}
          >
            {allAnswered ? "Submit All" : "Submit"}
          </button>
        </div>
      </div>
    </div>
  );
}
