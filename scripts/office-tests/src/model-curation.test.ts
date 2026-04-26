import assert from "node:assert/strict";
import test from "node:test";

import { DEFAULT_ENABLED_MODELS } from "../../../apps/taskpane/src/hooks/usePreferences.js";

test("default model shortlist avoids advanced-only regional and legacy provider noise", () => {
  const defaultProviders = Object.keys(DEFAULT_ENABLED_MODELS).sort();
  assert.deepEqual(defaultProviders, ["anthropic", "google", "openai"]);

  const defaultModelKeys = Object.entries(DEFAULT_ENABLED_MODELS)
    .flatMap(([provider, modelIds]) => modelIds.map((modelId) => `${provider}::${modelId}`));

  for (const blockedProvider of ["kimi-coding", "minimax", "minimax-cn", "zai"]) {
    assert.equal(
      defaultProviders.includes(blockedProvider),
      false,
      `${blockedProvider} should stay in the advanced catalog instead of the default shortlist.`,
    );
  }

  for (const legacyModel of ["openai::gpt-4o", "openai::gpt-4.1", "openai::gpt-5.2"]) {
    assert.equal(
      defaultModelKeys.includes(legacyModel),
      false,
      `${legacyModel} should not be enabled by default when newer project-preferred models exist.`,
    );
  }
});
