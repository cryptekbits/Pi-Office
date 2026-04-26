import { useCallback, useEffect, useState } from "react";
import type { OfficeThemeSnapshot } from "../../lib/office";
import { highlightCode } from "../../lib/code-highlighter";

interface CodeBlockProps {
  code: string;
  language: string;
  officeTheme: OfficeThemeSnapshot | undefined;
}

export function CodeBlock({ code, language, officeTheme }: CodeBlockProps) {
  const [html, setHtml] = useState<string>("");
  const [copied, setCopied] = useState(false);
  const isDark = officeTheme?.isDarkTheme ?? false;

  useEffect(() => {
    let cancelled = false;
    highlightCode(code, language, isDark).then((result) => {
      if (!cancelled) setHtml(result);
    });
    return () => { cancelled = true; };
  }, [code, language, isDark]);

  const handleCopy = useCallback(async () => {
    try {
      await navigator.clipboard.writeText(code);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch { /* clipboard may not be available */ }
  }, [code]);

  const displayLang = language || "text";

  return (
    <div className="code-block">
      <div className="code-block-header">
        <span className="code-block-lang">{displayLang}</span>
        <button type="button" className="code-block-copy" onClick={handleCopy}>
          {copied ? "Copied!" : "Copy"}
        </button>
      </div>
      {html ? (
        <div className="code-block-body" dangerouslySetInnerHTML={{ __html: html }} />
      ) : (
        <pre className="code-block-body">
          <code>{code}</code>
        </pre>
      )}
    </div>
  );
}
