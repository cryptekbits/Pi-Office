import type { Highlighter, BundledLanguage } from "shiki";

const BUNDLED_LANGS: BundledLanguage[] = [
  "typescript",
  "javascript",
  "python",
  "json",
  "html",
  "css",
  "sql",
  "xml",
  "bash",
  "csharp",
  "yaml",
  "markdown",
  "powershell",
  "java",
  "cpp",
  "go",
  "rust",
  "php",
];

const LIGHT_THEME = "github-light";
const DARK_THEME = "github-dark";

let highlighterPromise: Promise<Highlighter> | undefined;

function getHighlighter(): Promise<Highlighter> {
  if (!highlighterPromise) {
    highlighterPromise = import("shiki").then((shiki) =>
      shiki.createHighlighter({
        themes: [LIGHT_THEME, DARK_THEME],
        langs: BUNDLED_LANGS,
      }),
    );
  }
  return highlighterPromise;
}

export async function highlightCode(
  code: string,
  lang: string,
  isDark: boolean,
): Promise<string> {
  try {
    const highlighter = await getHighlighter();
    const loadedLangs = highlighter.getLoadedLanguages();
    const resolvedLang = loadedLangs.includes(lang as BundledLanguage) ? lang : "text";
    return highlighter.codeToHtml(code, {
      lang: resolvedLang,
      theme: isDark ? DARK_THEME : LIGHT_THEME,
    });
  } catch {
    const escaped = code
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
    return `<pre><code>${escaped}</code></pre>`;
  }
}
