import type { PromptSuggestion } from "@pi-office/pi-office-pack/protocol";

interface PromptSuggestionStripProps {
  suggestions: PromptSuggestion[];
  onSelect: (suggestion: PromptSuggestion) => void;
}

export function PromptSuggestionStrip({ suggestions, onSelect }: PromptSuggestionStripProps) {
  if (!suggestions.length) return null;

  return (
    <div className="prompt-suggestion-strip" aria-label="Suggested next prompts">
      {suggestions.map((suggestion) => (
        <button
          key={suggestion.id}
          type="button"
          className="next-prompt-chip"
          onClick={() => onSelect(suggestion)}
          title={suggestion.text}
        >
          {suggestion.text}
        </button>
      ))}
    </div>
  );
}
