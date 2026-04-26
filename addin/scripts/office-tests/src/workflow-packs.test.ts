import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import test from "node:test";

import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";
import { OFFICE_TOOL_NAMES, type OfficeHost } from "../../../packages/pi-office-pack/src/protocol.js";
import {
  formatWorkflowPackGuidance,
  getWorkflowPacksForHost,
  getWorkflowQuickPrompts,
  OFFICE_WORKFLOW_PACKS,
} from "../../../packages/pi-office-pack/src/workflow-packs.js";

const expectedPackIdsByHost: Record<OfficeHost, string[]> = {
  word: [
    "word-research-paper",
    "word-resume-polish",
    "word-spec-review",
    "word-legal-professional-review",
    "word-business-user-stories",
  ],
  excel: [
    "excel-dcf-review",
    "excel-formula-audit",
    "excel-table-chart-improvement",
    "excel-narrative-export",
  ],
  powerpoint: [
    "powerpoint-pitch-deck-outline",
    "powerpoint-slide-polish",
    "powerpoint-visual-consistency",
    "powerpoint-speaker-notes",
    "powerpoint-data-backed-slides",
  ],
};

test("professional workflow packs cover Word, Excel, and PowerPoint artifact families", () => {
  const toolNames = new Set<string>(OFFICE_TOOL_NAMES);

  for (const [host, expectedIds] of Object.entries(expectedPackIdsByHost) as Array<[OfficeHost, string[]]>) {
    const hostPacks = getWorkflowPacksForHost(host);
    assert.deepEqual(hostPacks.map((pack) => pack.id), expectedIds);

    for (const pack of hostPacks) {
      assert.equal(pack.host, host);
      assert.ok(pack.title.length > 4, `${pack.id} should have a user-facing title.`);
      assert.ok(pack.userPrompt.endsWith("."), `${pack.id} should have a complete starter prompt.`);
      assert.ok(pack.intent.length > 30, `${pack.id} should describe artifact intent.`);
      assert.ok(pack.requiredContext.length >= 3, `${pack.id} should describe required context.`);
      assert.ok(pack.preferredTools.length >= 3, `${pack.id} should name expected Office tools.`);
      assert.ok(pack.reviewGates.length >= 2, `${pack.id} should include review gates.`);
      assert.ok(pack.completionChecks.length >= 2, `${pack.id} should include completion checks.`);

      for (const toolName of pack.preferredTools) {
        assert.ok(toolNames.has(toolName), `${pack.id} references unknown tool ${toolName}.`);
      }
    }
  }
});

test("workflow guidance is injected into prompts and packaged skill text", () => {
  const guidance = formatWorkflowPackGuidance();
  assert.match(guidance, /Research Paper Review/);
  assert.match(guidance, /DCF Review/);
  assert.match(guidance, /Pitch Deck Outline/);
  assert.match(guidance, /\boffice_propose_edits\b/);
  assert.match(guidance, /\bget_range_as_csv\b/);
  assert.match(guidance, /\bverify_slide_visual\b/);
  assert.match(guidance, /review gates=/);
  assert.match(guidance, /completion checks=/);

  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /Professional Workflow Packs/);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /business user stories/i);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /formula audit/i);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /speaker notes/i);

  const officeHostSkillPath = join(process.cwd(), "packages", "pi-office-pack", "skills", "office-host.SKILL.md");
  const officeHostSkillText = readFileSync(officeHostSkillPath, "utf8");
  assert.match(officeHostSkillText, /workflow packs cover research paper review/i);
  assert.match(officeHostSkillText, /workflow packs cover DCF review/i);
  assert.match(officeHostSkillText, /workflow packs cover pitch-deck outline/i);
});

test("taskpane starter prompts expose host-specific workflow packs", () => {
  assert.deepEqual(getWorkflowQuickPrompts(undefined), []);
  assert.deepEqual(getWorkflowQuickPrompts("word"), [
    "Run the research paper workflow on this document.",
    "Polish this resume for a sharper professional story.",
    "Review this spec and turn gaps into action items.",
  ]);
  assert.deepEqual(getWorkflowQuickPrompts("excel"), [
    "Run a DCF review on the active workbook.",
    "Audit formulas and dependencies in the current sheet.",
    "Improve the selected table or chart for presentation.",
  ]);
  assert.deepEqual(getWorkflowQuickPrompts("powerpoint"), [
    "Run the pitch deck workflow for this presentation.",
    "Polish the current slide for clarity and executive tone.",
    "Check this deck for visual consistency issues.",
  ]);

  const helperText = readFileSync(join(process.cwd(), "apps", "taskpane", "src", "lib", "helpers.ts"), "utf8");
  const chatViewText = readFileSync(join(process.cwd(), "apps", "taskpane", "src", "app", "components", "ChatView.tsx"), "utf8");
  assert.match(helperText, /getWorkflowQuickPrompts\(officeState\?\.host,\s*3\)/);
  assert.match(chatViewText, /getQuickPrompts\(officeState\)/);
  assert.equal(OFFICE_WORKFLOW_PACKS.length, 14);
});
