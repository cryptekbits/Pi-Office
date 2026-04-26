import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import test from "node:test";

import { executeOfficeJs } from "../../../apps/taskpane/src/lib/office/document-tools.js";
import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
import { createOfficeExtension } from "../../../packages/pi-office-pack/src/extension.js";
import {
  AUTONOMY_LEVELS,
  AUTONOMY_LEVEL_AUTO_APPROVE,
  TOOL_CATEGORY_MAP,
} from "../../../packages/pi-office-pack/src/protocol.js";

const SAFETY_CONTRACT_REGEXES = [
  /\bbest[- ]effort restricted subset\b/i,
  /\bnetwork\b/i,
  /\bstorage\b/i,
  /\beval\b/i,
  /\bsystem(?:[- ]access| resources?)\b/i,
  /\bnot (?:an?\s+)?(?:isolated|sandbox)\b|\bno isolation guarantee\b/i,
];

function assertSafetyContractText(text: string): void {
  for (const expected of SAFETY_CONTRACT_REGEXES) {
    assert.match(text, expected);
  }
}

test("executeOfficeJs blocks documented network/storage/eval/system-access categories", async () => {
  const blockedSnippets = [
    { code: "return fetch('https://example.com')", category: /\bnetwork\b/i },
    { code: "return localStorage.getItem('token')", category: /\bstorage\b/i },
    { code: "return eval('2 + 2')", category: /\beval\b/i },
    { code: "return process.env.PATH", category: /\bsystem(?:[- ]access| resources?)\b/i },
  ];

  for (const snippet of blockedSnippets) {
    const result = await executeOfficeJs("word", snippet.code) as { ok?: boolean; error?: string };
    assert.notEqual(result.ok, true);
    assert.equal(typeof result.error, "string");
    assert.match(result.error ?? "", snippet.category);
    assertSafetyContractText(result.error ?? "");
  }
});

test("executeOfficeJs still allows safe Office.js snippets", async () => {
  const result = await executeOfficeJs("word", "return 2 + 2;") as { ok?: boolean; result?: unknown };
  assert.equal(result.ok, true);
  assert.equal(result.result, 4);
});

test("execute-js prompt and tool descriptions use the best-effort restricted subset contract", () => {
  const registeredTools: Array<{ name: string; description?: string }> = [];
  const extension = createOfficeExtension({
    getHost: () => "word",
    getState: () => undefined,
    invokeTool: async () => ({ ok: true }),
    invokeAskUser: async (request) => ({ requestId: request.requestId, answers: [] }),
  });

  extension({
    registerTool: (tool: { name: string; description?: string }) => {
      registeredTools.push(tool);
    },
    on: () => undefined,
  } as unknown as Parameters<ReturnType<typeof createOfficeExtension>>[0]);

  const executeJsTool = registeredTools.find((tool) => tool.name === "office_execute_js");
  assert.ok(executeJsTool?.description, "office_execute_js tool must be registered with a description");
  assertSafetyContractText(executeJsTool.description ?? "");

  assertSafetyContractText(OFFICE_APPEND_SYSTEM_PROMPT);

  const officeHostSkillPath = join(process.cwd(), "packages", "pi-office-pack", "skills", "office-host.SKILL.md");
  const officeHostSkillText = readFileSync(officeHostSkillPath, "utf8");
  assertSafetyContractText(officeHostSkillText);
});

test("office_execute_js is categorized as a manual-only escape hatch", () => {
  assert.equal(TOOL_CATEGORY_MAP.office_execute_js, "escape-hatch");

  for (const level of AUTONOMY_LEVELS) {
    assert.equal(
      AUTONOMY_LEVEL_AUTO_APPROVE[level].has("escape-hatch"),
      false,
      `${level} autonomy must not auto-approve office_execute_js.`,
    );
  }
});
