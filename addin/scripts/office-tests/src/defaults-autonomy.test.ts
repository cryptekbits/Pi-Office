import assert from "node:assert/strict";
import test from "node:test";

import { composeAutonomyPrompt } from "../../../packages/pi-office-pack/src/defaults.js";
import { DEFAULT_USER_PREFERENCES } from "../../../packages/pi-office-pack/src/protocol.js";

test("composeAutonomyPrompt honors explicitly provided available tool names", () => {
  const prompt = composeAutonomyPrompt(
    { ...DEFAULT_USER_PREFERENCES, autonomyLevel: "high" },
    true,
    ["office_get_context", "ask_user", "generate_image"],
  );

  assert.match(prompt, /office_get_context/);
  assert.match(prompt, /ask_user/);
  assert.match(prompt, /generate_image/);
  assert.doesNotMatch(prompt, /workspace-read|workspace-write/);
});

test("composeAutonomyPrompt falls back to workspace tool inventory when explicit list is not provided", () => {
  const prompt = composeAutonomyPrompt({ ...DEFAULT_USER_PREFERENCES, autonomyLevel: "high" }, true);
  assert.match(prompt, /\bread\b/);
  assert.match(prompt, /\boffice_get_context\b/);
});

test("composeAutonomyPrompt only lists workspace tools when the session exposes them", () => {
  const unsavedOrNoCompanionPrompt = composeAutonomyPrompt(
    { ...DEFAULT_USER_PREFERENCES, autonomyLevel: "high" },
    false,
  );
  const savedWithCompanionPrompt = composeAutonomyPrompt(
    { ...DEFAULT_USER_PREFERENCES, autonomyLevel: "high" },
    true,
  );

  assert.doesNotMatch(unsavedOrNoCompanionPrompt, /workspace-read/);
  assert.doesNotMatch(unsavedOrNoCompanionPrompt, /read \(workspace-read\)/);
  assert.match(savedWithCompanionPrompt, /read \(workspace-read\)/);
  assert.match(savedWithCompanionPrompt, /grep \(workspace-read\)/);
  assert.match(savedWithCompanionPrompt, /ls \(workspace-read\)/);
});

test("composeAutonomyPrompt keeps raw Office.js execution out of auto-approved tools", () => {
  const prompt = composeAutonomyPrompt(
    { ...DEFAULT_USER_PREFERENCES, autonomyLevel: "extreme" },
    false,
    ["office_execute_js"],
  );

  assert.match(prompt, /Require user approval: office_execute_js \(manual-escape-hatch\)/);
  assert.doesNotMatch(prompt, /Auto-approved tools: office_execute_js/);
});

test("composeAutonomyPrompt exposes artifact drafting preference guidance", () => {
  const draftNowPrompt = composeAutonomyPrompt(
    { ...DEFAULT_USER_PREFERENCES, artifactClarificationMode: "draft_now" },
    false,
    ["ask_user"],
  );
  const askFirstPrompt = composeAutonomyPrompt(
    { ...DEFAULT_USER_PREFERENCES, artifactClarificationMode: "ask_first" },
    false,
    ["ask_user"],
  );

  assert.match(draftNowPrompt, /Artifact Drafting Preference/);
  assert.match(draftNowPrompt, /Clarification mode: Draft now/);
  assert.match(draftNowPrompt, /Prefer making a strong first draft/);
  assert.match(askFirstPrompt, /Clarification mode: Ask me first/);
  assert.match(askFirstPrompt, /Before broad professional artifact generation/);
});
