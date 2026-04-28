const WORDPROCESSINGML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const OFFICEMATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";

type MathNode =
  | { kind: "text"; text: string }
  | { kind: "group"; children: MathNode[] }
  | { kind: "fraction"; numerator: MathNode[]; denominator: MathNode[] }
  | { kind: "radical"; body: MathNode[] }
  | { kind: "matrix"; rows: MathNode[][][]; leftDelimiter?: string | undefined; rightDelimiter?: string | undefined }
  | { kind: "script"; base: MathNode; subscript?: MathNode[] | undefined; superscript?: MathNode[] | undefined };

const COMMAND_SYMBOLS: Record<string, string> = {
  alpha: "α",
  beta: "β",
  gamma: "γ",
  delta: "δ",
  epsilon: "ε",
  theta: "θ",
  lambda: "λ",
  mu: "μ",
  pi: "π",
  rho: "ρ",
  sigma: "σ",
  tau: "τ",
  phi: "φ",
  omega: "ω",
  Gamma: "Γ",
  Delta: "Δ",
  Theta: "Θ",
  Lambda: "Λ",
  Pi: "Π",
  Sigma: "Σ",
  Phi: "Φ",
  Omega: "Ω",
  infty: "∞",
  sum: "∑",
  prod: "∏",
  int: "∫",
  partial: "∂",
  nabla: "∇",
  times: "×",
  cdot: "⋅",
  pm: "±",
  le: "≤",
  leq: "≤",
  ge: "≥",
  geq: "≥",
  neq: "≠",
  approx: "≈",
  to: "→",
  rightarrow: "→",
  leftarrow: "←",
  degree: "°",
};

const NAMED_FUNCTIONS = new Set(["sin", "cos", "tan", "log", "ln", "lim", "min", "max", "exp"]);

export interface WordEquationOoxmlResult {
  ooxml: string;
  normalizedLatex: string;
  display: "block" | "inline";
  numbering?: string | undefined;
  caption?: string | undefined;
  warnings: string[];
  unsupportedCommands: string[];
  evidence: {
    containsOfficeMath: true;
    containsFraction: boolean;
    containsRadical: boolean;
    containsScript: boolean;
    containsMatrix: boolean;
    containsEquationNumber: boolean;
    containsCaption: boolean;
  };
}

export interface WordEquationOoxmlOptions {
  numbering?: string | number | undefined;
  caption?: string | undefined;
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function stripMathDelimiters(value: string): string {
  const trimmed = value.trim();
  const delimited =
    trimmed.match(/^\$\$([\s\S]*?)\$\$$/) ??
    trimmed.match(/^\\\[([\s\S]*?)\\\]$/) ??
    trimmed.match(/^\\\(([\s\S]*?)\\\)$/) ??
    trimmed.match(/^\$([\s\S]*?)\$$/);
  return (delimited?.[1] ?? trimmed).trim();
}

class LatexMathParser {
  private index = 0;
  readonly warnings = new Set<string>();
  readonly unsupportedCommands = new Set<string>();

  constructor(private readonly input: string) {}

  parse(): MathNode[] {
    return this.parseNodes();
  }

  private parseNodes(stopAtBrace = false): MathNode[] {
    const nodes: MathNode[] = [];
    while (this.index < this.input.length) {
      if (stopAtBrace && this.peek() === "}") {
        break;
      }
      if (this.peek() === "}") {
        this.warnings.add("Ignored an unmatched closing brace in the LaTeX equation.");
        this.index += 1;
        continue;
      }
      const atom = this.parseAtom();
      if (!atom) {
        continue;
      }
      nodes.push(this.parseScripts(atom));
    }
    return nodes;
  }

  private parseAtom(): MathNode | undefined {
    const char = this.peek();
    if (!char) {
      return undefined;
    }
    if (char === "{") {
      return { kind: "group", children: this.parseGroup() };
    }
    if (char === "\\") {
      return this.parseCommand();
    }
    if (char === "^" || char === "_") {
      this.index += 1;
      return { kind: "text", text: char };
    }
    this.index += 1;
    return { kind: "text", text: char };
  }

  private parseCommand(): MathNode {
    this.index += 1;
    const start = this.index;
    while (/[A-Za-z]/.test(this.peek() ?? "")) {
      this.index += 1;
    }
    const command = this.input.slice(start, this.index);

    if (!command) {
      const escaped = this.peek() ?? "";
      this.index += escaped ? 1 : 0;
      return { kind: "text", text: escaped };
    }

    if (command === "left" || command === "right") {
      const delimiter = this.peek();
      if (delimiter && delimiter !== "\\") {
        this.index += 1;
        return { kind: "text", text: delimiter === "." ? "" : delimiter };
      }
      return { kind: "text", text: "" };
    }

    if (command === "frac") {
      return {
        kind: "fraction",
        numerator: this.parseRequiredGroup("fraction numerator"),
        denominator: this.parseRequiredGroup("fraction denominator"),
      };
    }

    if (command === "sqrt") {
      this.skipOptionalBracketGroup();
      return { kind: "radical", body: this.parseRequiredGroup("square-root body") };
    }

    if (command === "begin") {
      return this.parseEnvironment();
    }

    if (command === "text" || command === "mathrm" || command === "operatorname") {
      return { kind: "group", children: this.parseRequiredGroup(`${command} body`) };
    }

    const symbol = COMMAND_SYMBOLS[command];
    if (symbol) {
      return { kind: "text", text: symbol };
    }
    if (NAMED_FUNCTIONS.has(command)) {
      return { kind: "text", text: command };
    }

    this.unsupportedCommands.add(command);
    this.warnings.add(`Unsupported LaTeX command \\${command} was preserved as plain OfficeMath text.`);
    return { kind: "text", text: `\\${command}` };
  }

  private parseScripts(base: MathNode): MathNode {
    let current = base;
    let subscript: MathNode[] | undefined;
    let superscript: MathNode[] | undefined;

    while (this.peek() === "^" || this.peek() === "_") {
      const marker = this.peek();
      this.index += 1;
      const value = this.parseScriptArgument(marker === "^" ? "superscript" : "subscript");
      if (marker === "^") {
        superscript = value;
      } else {
        subscript = value;
      }
      current = { kind: "script", base, subscript, superscript };
    }

    return current;
  }

  private parseScriptArgument(label: string): MathNode[] {
    this.skipSpaces();
    if (this.peek() === "{") {
      return this.parseGroup();
    }
    const atom = this.parseAtom();
    if (!atom) {
      this.warnings.add(`Missing ${label} argument in the LaTeX equation.`);
      return [{ kind: "text", text: "" }];
    }
    return [atom];
  }

  private parseRequiredGroup(label: string): MathNode[] {
    this.skipSpaces();
    if (this.peek() === "{") {
      return this.parseGroup();
    }
    this.warnings.add(`Expected a braced ${label}; parsed the next token instead.`);
    const atom = this.parseAtom();
    return atom ? [atom] : [{ kind: "text", text: "" }];
  }

  private parseEnvironment(): MathNode {
    const environmentName = this.readRequiredGroupText("environment name");
    if (!environmentName) {
      return { kind: "text", text: "\\begin" };
    }

    const rawContent = this.readUntilEndEnvironment(environmentName);
    if (!MATRIX_ENVIRONMENTS.has(environmentName)) {
      this.unsupportedCommands.add(`begin{${environmentName}}`);
      this.warnings.add(`Unsupported LaTeX environment ${environmentName} was preserved as plain OfficeMath text.`);
      return { kind: "text", text: `\\begin{${environmentName}}${rawContent}\\end{${environmentName}}` };
    }

    const rows = rawContent
      .split(/\\\\/)
      .map((row) => row.trim())
      .filter(Boolean)
      .map((row) =>
        row.split("&").map((cell) => {
          const parser = new LatexMathParser(cell.trim());
          const parsed = parser.parse();
          for (const warning of parser.warnings) this.warnings.add(warning);
          for (const unsupported of parser.unsupportedCommands) this.unsupportedCommands.add(unsupported);
          const emptyCell: MathNode[] = [{ kind: "text", text: "" }];
          return parsed.length ? parsed : emptyCell;
        }),
      );

    if (!rows.length) {
      this.warnings.add(`LaTeX ${environmentName} environment did not include any matrix rows.`);
      rows.push([[{ kind: "text", text: "" }]]);
    }

    const delimiters = MATRIX_DELIMITERS[environmentName] ?? {};
    return { kind: "matrix", rows, leftDelimiter: delimiters.left, rightDelimiter: delimiters.right };
  }

  private readRequiredGroupText(label: string): string {
    this.skipSpaces();
    if (this.peek() !== "{") {
      this.warnings.add(`Expected a braced ${label}.`);
      return "";
    }
    this.index += 1;
    const start = this.index;
    while (this.index < this.input.length && this.peek() !== "}") {
      this.index += 1;
    }
    const value = this.input.slice(start, this.index);
    if (this.peek() === "}") {
      this.index += 1;
    } else {
      this.warnings.add(`Missing a closing brace for ${label}.`);
    }
    return value.trim();
  }

  private readUntilEndEnvironment(environmentName: string): string {
    const endMarker = `\\end{${environmentName}}`;
    const start = this.index;
    const endIndex = this.input.indexOf(endMarker, this.index);
    if (endIndex < 0) {
      this.index = this.input.length;
      this.warnings.add(`Missing ${endMarker} in the LaTeX equation.`);
      return this.input.slice(start);
    }
    this.index = endIndex + endMarker.length;
    return this.input.slice(start, endIndex);
  }

  private parseGroup(): MathNode[] {
    if (this.peek() !== "{") {
      return [];
    }
    this.index += 1;
    const children = this.parseNodes(true);
    if (this.peek() === "}") {
      this.index += 1;
    } else {
      this.warnings.add("Missing a closing brace in the LaTeX equation.");
    }
    return children;
  }

  private skipOptionalBracketGroup(): void {
    this.skipSpaces();
    if (this.peek() !== "[") {
      return;
    }
    this.warnings.add("Indexed radicals are inserted as square roots; verify the equation visually.");
    let depth = 0;
    while (this.index < this.input.length) {
      const char = this.peek();
      this.index += 1;
      if (char === "[") depth += 1;
      if (char === "]") {
        depth -= 1;
        if (depth <= 0) break;
      }
    }
  }

  private skipSpaces(): void {
    while (/\s/.test(this.peek() ?? "")) {
      this.index += 1;
    }
  }

  private peek(): string | undefined {
    return this.input[this.index];
  }
}

const MATRIX_ENVIRONMENTS = new Set(["matrix", "pmatrix", "bmatrix", "Bmatrix", "vmatrix", "Vmatrix"]);

const MATRIX_DELIMITERS: Record<string, { left?: string; right?: string }> = {
  pmatrix: { left: "(", right: ")" },
  bmatrix: { left: "[", right: "]" },
  Bmatrix: { left: "{", right: "}" },
  vmatrix: { left: "|", right: "|" },
  Vmatrix: { left: "‖", right: "‖" },
};

function renderNodes(nodes: MathNode[]): string {
  return nodes.map(renderNode).join("");
}

function renderSlot(nodes: MathNode[] | undefined): string {
  const content = nodes && nodes.length ? renderNodes(nodes) : renderText("");
  return content;
}

function renderText(text: string): string {
  return `<m:r><m:t xml:space="preserve">${escapeXml(text)}</m:t></m:r>`;
}

function renderMatrix(node: Extract<MathNode, { kind: "matrix" }>): string {
  const matrixRows = node.rows
    .map((row) => `<m:mr>${row.map((cell) => `<m:e>${renderSlot(cell)}</m:e>`).join("")}</m:mr>`)
    .join("");
  const matrix = `<m:m><m:mPr/><m:mrPr/>${matrixRows}</m:m>`;
  if (!node.leftDelimiter && !node.rightDelimiter) {
    return matrix;
  }
  return `<m:d><m:dPr>${node.leftDelimiter ? `<m:begChr m:val="${escapeXml(node.leftDelimiter)}"/>` : ""}${node.rightDelimiter ? `<m:endChr m:val="${escapeXml(node.rightDelimiter)}"/>` : ""}</m:dPr><m:e>${matrix}</m:e></m:d>`;
}

function renderNode(node: MathNode): string {
  if (node.kind === "text") {
    return renderText(node.text);
  }
  if (node.kind === "group") {
    return renderNodes(node.children);
  }
  if (node.kind === "fraction") {
    return `<m:f><m:fPr><m:type m:val="bar"/></m:fPr><m:num>${renderSlot(node.numerator)}</m:num><m:den>${renderSlot(node.denominator)}</m:den></m:f>`;
  }
  if (node.kind === "radical") {
    return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${renderSlot(node.body)}</m:e></m:rad>`;
  }
  if (node.kind === "matrix") {
    return renderMatrix(node);
  }
  if (node.kind === "script") {
    if (node.subscript && node.superscript) {
      return `<m:sSubSup><m:e>${renderNode(node.base)}</m:e><m:sub>${renderSlot(node.subscript)}</m:sub><m:sup>${renderSlot(node.superscript)}</m:sup></m:sSubSup>`;
    }
    if (node.subscript) {
      return `<m:sSub><m:e>${renderNode(node.base)}</m:e><m:sub>${renderSlot(node.subscript)}</m:sub></m:sSub>`;
    }
    return `<m:sSup><m:e>${renderNode(node.base)}</m:e><m:sup>${renderSlot(node.superscript)}</m:sup></m:sSup>`;
  }
  return "";
}

function normalizeEquationNumber(value: string | number | undefined): string | undefined {
  if (typeof value === "number" && Number.isFinite(value)) {
    return `(${Math.trunc(value)})`;
  }
  const trimmed = typeof value === "string" ? value.trim() : "";
  if (!trimmed) {
    return undefined;
  }
  return /^\(.+\)$/.test(trimmed) ? trimmed : `(${trimmed})`;
}

function renderParagraphText(text: string, alignment?: "center" | "right"): string {
  const paragraphProperties = alignment ? `<w:pPr><w:jc w:val="${alignment}"/></w:pPr>` : "";
  return `<w:p xmlns:w="${WORDPROCESSINGML_NS}">${paragraphProperties}<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

export function countWordOoxmlMathObjects(ooxml: string): number {
  return (ooxml.match(/<m:oMath\b/gi) ?? []).length;
}

export function createWordEquationOoxml(
  latex: string,
  display: "block" | "inline" = "block",
  options: WordEquationOoxmlOptions = {},
): WordEquationOoxmlResult {
  const normalizedLatex = stripMathDelimiters(latex);
  if (!normalizedLatex) {
    throw new Error("word_equation requires non-empty LaTeX math content.");
  }

  const parser = new LatexMathParser(normalizedLatex);
  const nodes = parser.parse();
  const math = renderNodes(nodes);
  const displayMode = display === "inline" ? "inline" : "block";
  const oMath = `<m:oMath xmlns:m="${OFFICEMATH_NS}">${math}</m:oMath>`;
  const numbering = normalizeEquationNumber(options.numbering);
  const caption = typeof options.caption === "string" ? options.caption.trim() || undefined : undefined;
  let ooxml = displayMode === "inline"
    ? oMath
    : `<w:p xmlns:w="${WORDPROCESSINGML_NS}" xmlns:m="${OFFICEMATH_NS}"><m:oMathPara>${oMath}</m:oMathPara></w:p>`;
  if (displayMode === "inline" && (numbering || caption)) {
    parser.warnings.add("Equation numbering and captions are only inserted for block equations; inline metadata was returned but not written visibly.");
  }
  if (displayMode === "block") {
    if (numbering) {
      ooxml += renderParagraphText(numbering, "right");
    }
    if (caption) {
      ooxml += renderParagraphText(caption, "center");
    }
  }

  return {
    ooxml,
    normalizedLatex,
    display: displayMode,
    numbering,
    caption,
    warnings: Array.from(parser.warnings),
    unsupportedCommands: Array.from(parser.unsupportedCommands),
    evidence: {
      containsOfficeMath: true,
      containsFraction: /<m:f\b/.test(ooxml),
      containsRadical: /<m:rad\b/.test(ooxml),
      containsScript: /<m:s(?:Sub|Sup|SubSup)\b/.test(ooxml),
      containsMatrix: /<m:m\b/.test(ooxml),
      containsEquationNumber: Boolean(numbering && displayMode === "block"),
      containsCaption: Boolean(caption && displayMode === "block"),
    },
  };
}
