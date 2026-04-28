import assert from "node:assert/strict";
import test from "node:test";

import { createOfficeToolExecutor } from "../../../apps/taskpane/src/lib/office-bridge.js";
import { createWordEquationOoxml, countWordOoxmlMathObjects } from "../../../apps/taskpane/src/lib/office/word-equations.js";
import { WORD_OFFICE_TOOL_DEFINITIONS } from "../../../apps/taskpane/src/lib/office/tools/word.js";
import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
import { OFFICE_TOOL_NAMES, TOOL_CATEGORY_MAP, type OfficeToolRequest } from "../../../packages/pi-office-pack/src/protocol.js";

test("Word equation protocol and registry expose native OfficeMath insertion", () => {
  assert.ok(OFFICE_TOOL_NAMES.includes("word_equation"));
  assert.equal(TOOL_CATEGORY_MAP.word_equation, "write-doc");
  const definition = WORD_OFFICE_TOOL_DEFINITIONS.find((tool) => tool.name === "word_equation");
  assert.ok(definition);
  assert.match(definition.description, /OfficeMath/i);
  assert.ok(definition.discovery?.capabilityIds?.includes("word.equations"));
});

test("Word equation bridge dispatches LaTeX as an insertEquation host action", async () => {
  const calls: Array<{ host: string; action: Record<string, unknown> }> = [];
  const executeOfficeTool = createOfficeToolExecutor({
    collectOfficeContext: async () => ({ ok: true }),
    applyHostAction: async (host, action) => {
      calls.push({ host, action });
      return { ok: true, action: action.type, verification: { persisted: true, afterMathCount: 1 } };
    },
    navigateOfficeAnchor: async () => ({ ok: true }),
    readDocumentSection: async () => ({ ok: true }),
    executeOfficeJs: async () => ({ ok: true }),
    proposeEdits: async () => ({ ok: true }),
  });

  const result = await executeOfficeTool({
    requestId: "word-equation",
    toolName: "word_equation" as OfficeToolRequest["toolName"],
    host: "word",
    params: {
      latex: "$$E=mc^2+\\frac{a}{b}$$",
      display: "block",
      placement: "after",
      numbering: "1",
      caption: "Einstein mass-energy relation",
      target: { kind: "heading", text: "Methods" },
    },
  } as OfficeToolRequest);

  assert.equal(result.success, true);
  assert.equal(calls[0]?.host, "word");
  assert.equal(calls[0]?.action.type, "insertEquation");
  assert.equal(calls[0]?.action.content, "$$E=mc^2+\\frac{a}{b}$$");
  assert.equal(calls[0]?.action.placement, "after");
  assert.equal((calls[0]?.action.target as { kind?: string }).kind, "heading");
  assert.equal((calls[0]?.action.options as Record<string, unknown>).display, "block");
  assert.equal((calls[0]?.action.options as Record<string, unknown>).numbering, "1");
  assert.equal((calls[0]?.action.options as Record<string, unknown>).caption, "Einstein mass-energy relation");
});

test("LaTeX converter emits persisted Word OfficeMath OOXML evidence", () => {
  const equation = createWordEquationOoxml("\\frac{a_1}{b^2}+\\sqrt{x}", "block");

  assert.match(equation.ooxml, /<m:oMath\b/);
  assert.match(equation.ooxml, /<m:f\b/);
  assert.match(equation.ooxml, /<m:rad\b/);
  assert.match(equation.ooxml, /<m:sSub\b/);
  assert.match(equation.ooxml, /<m:sSup\b/);
  assert.equal(equation.normalizedLatex, "\\frac{a_1}{b^2}+\\sqrt{x}");
  assert.equal(countWordOoxmlMathObjects(equation.ooxml), 1);
  assert.equal(equation.evidence.containsOfficeMath, true);
  assert.equal(equation.evidence.containsFraction, true);

  const fftEquation = createWordEquationOoxml("X_k = \\sum_{n=0}^{N-1} x_n e^{-i2\\pi kn/N}", "block");
  assert.match(fftEquation.ooxml, /<m:sSubSup\b/);
  assert.match(fftEquation.ooxml, /π/);
  assert.equal(fftEquation.unsupportedCommands.length, 0);
});

test("LaTeX converter supports matrix environments and block equation metadata", () => {
  const equation = createWordEquationOoxml(
    "\\begin{bmatrix} a & b \\\\ c & d \\end{bmatrix}",
    "block",
    { numbering: "2.1", caption: "State transition matrix" },
  );

  assert.match(equation.ooxml, /<m:m\b/);
  assert.match(equation.ooxml, /<m:mr>/);
  assert.match(equation.ooxml, /<m:d\b/);
  assert.match(equation.ooxml, /\(2\.1\)/);
  assert.match(equation.ooxml, /State transition matrix/);
  assert.equal(equation.numbering, "(2.1)");
  assert.equal(equation.caption, "State transition matrix");
  assert.equal(equation.evidence.containsMatrix, true);
  assert.equal(equation.evidence.containsEquationNumber, true);
  assert.equal(equation.evidence.containsCaption, true);
  assert.equal(equation.unsupportedCommands.length, 0);
});

test("inline equation metadata is reported but not inserted visibly", () => {
  const equation = createWordEquationOoxml("x+y", "inline", { numbering: 3, caption: "Inline note" });

  assert.equal(equation.numbering, "(3)");
  assert.equal(equation.caption, "Inline note");
  assert.equal(equation.evidence.containsEquationNumber, false);
  assert.equal(equation.evidence.containsCaption, false);
  assert.match(equation.warnings.join("\n"), /only inserted for block equations/i);
});

test("Word equation guidance prevents raw LaTeX HTML insertion claims", () => {
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\bword_equation\b/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /matrices/i);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /numbering\/caption/i);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /\$\$\.\.\.\$\$/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /m:oMath verification/i);
});
