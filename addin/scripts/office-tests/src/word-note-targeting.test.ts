import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import test from "node:test";

import { OFFICE_APPEND_SYSTEM_PROMPT } from "../../../packages/pi-office-pack/src/defaults.js";

function readSource(path: string): string {
  return readFileSync(join(process.cwd(), path), "utf8");
}

test("Word note anchors distinguish body targeting from reference-marker targeting", () => {
  const protocol = readSource("packages/pi-office-pack/src/protocol.ts");
  assert.match(protocol, /noteTarget\?: "reference" \| "body"/);

  const actionSource = readSource("apps/taskpane/src/lib/office/word-actions.ts");
  assert.ok(actionSource.includes('noteTarget === "reference"'));
  assert.ok(actionSource.includes('note.note.body.getRange("Content")'));
  assert.ok(actionSource.includes("note.note.reference"));

  const navigateSource = readSource("apps/taskpane/src/lib/office/word-navigate.ts");
  assert.ok(navigateSource.includes('anchor.noteTarget === "reference"'));
  assert.ok(navigateSource.includes('note.body.getRange("Content").select()'));
  assert.ok(navigateSource.includes("noteTarget,"));
});

test("Word guidance tells agents that footnote and endnote body anchors are separate", () => {
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /footnote or endnote edits/i);
  assert.match(OFFICE_APPEND_SYSTEM_PROMPT, /noteTarget="reference"/);
});
