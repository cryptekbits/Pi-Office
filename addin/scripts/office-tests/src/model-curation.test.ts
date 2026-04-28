import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import test from "node:test";
import { getModels } from "@mariozechner/pi-ai";

import {
  DEFAULT_ENABLED_MODELS_BY_PROVIDER,
  PROVIDER_MODEL_PREFERENCE_SOURCE,
  PROVIDER_MODEL_PREFERENCES,
  SIMPLE_RECOMMENDED_MODELS_BY_PROVIDER,
  SIMPLE_VISIBLE_PROVIDERS,
  getProviderDefaultModel,
  getProviderSettingsPreference,
} from "../../../packages/pi-office-pack/src/provider-model-preferences.generated.js";
import { DEFAULT_ENABLED_MODELS } from "../../../apps/taskpane/src/hooks/usePreferences.js";

const MODEL_PREFERENCES = PROVIDER_MODEL_PREFERENCES as Record<
  string,
  Record<
    string,
    {
      settingsVisibility: string;
      recommended: boolean;
      defaultForProvider: boolean;
      requiresUnrecommendedWarning: boolean;
    }
  >
>;

function hasPiModel(provider: string, modelId: string): boolean {
  return getModels(provider as never).some((model) => model.id === modelId);
}

test("generated Simple recommended models exist in the refreshed Pi catalog", () => {
  for (const [provider, modelIds] of Object.entries(SIMPLE_RECOMMENDED_MODELS_BY_PROVIDER)) {
    assert.ok(modelIds.length > 0, `${provider} should have at least one Simple recommendation.`);
    for (const modelId of modelIds) {
      assert.equal(hasPiModel(provider, modelId), true, `${provider}/${modelId} must exist in Pi.`);
      const preference = MODEL_PREFERENCES[provider]?.[modelId];
      assert.ok(preference, `${provider}/${modelId} should have generated preference metadata.`);
      assert.equal(preference.settingsVisibility, "simple");
      assert.equal(preference.recommended, true);
      assert.equal(preference.requiresUnrecommendedWarning, false);
    }
  }
});

test("each Simple-visible provider has exactly one default model", () => {
  assert.deepEqual(
    [...SIMPLE_VISIBLE_PROVIDERS].sort(),
    [
      "anthropic",
      "github-copilot",
      "google",
      "groq",
      "openai",
      "openai-codex",
      "opencode",
      "opencode-go",
      "openrouter",
      "vercel-ai-gateway",
      "xai",
    ],
  );

  for (const provider of SIMPLE_VISIBLE_PROVIDERS) {
    const defaultModel = getProviderDefaultModel(provider);
    assert.ok(defaultModel, `${provider} should have a default model.`);
    assert.equal(hasPiModel(provider, defaultModel), true, `${provider}/${defaultModel} default must exist in Pi.`);
    const generatedDefaultCount = Object.values(MODEL_PREFERENCES[provider] ?? {})
      .filter((model) => model.defaultForProvider).length;
    assert.equal(generatedDefaultCount, 1, `${provider} should have exactly one generated default.`);
  }
});

test("default enabled shortlist is generated and keeps direct regional providers advanced-only", () => {
  assert.deepEqual(DEFAULT_ENABLED_MODELS, DEFAULT_ENABLED_MODELS_BY_PROVIDER);

  const defaultProviders = Object.keys(DEFAULT_ENABLED_MODELS).sort();
  assert.deepEqual(defaultProviders, ["anthropic", "google", "openai"]);

  const defaultModelKeys = Object.entries(DEFAULT_ENABLED_MODELS)
    .flatMap(([provider, modelIds]) => modelIds.map((modelId) => `${provider}::${modelId}`));

  assert.equal(defaultModelKeys.includes("openai::gpt-5.5"), true);
  assert.equal(defaultModelKeys.includes("anthropic::claude-opus-4-7"), true);
  assert.equal(defaultModelKeys.includes("google::gemini-3.1-pro-preview"), true);
  assert.equal(defaultModelKeys.includes("openai::gpt-4.1"), true);

  for (const advancedProvider of ["deepseek", "kimi-coding", "zai"]) {
    assert.equal(
      getProviderSettingsPreference(advancedProvider).settingsVisibility,
      "advanced",
      `${advancedProvider} should stay behind the Advanced catalog as a direct provider.`,
    );
    assert.equal(
      SIMPLE_VISIBLE_PROVIDERS.includes(advancedProvider as never),
      false,
      `${advancedProvider} should not leak into Simple provider settings.`,
    );
    assert.equal(
      defaultProviders.includes(advancedProvider),
      false,
      `${advancedProvider} should not be enabled by default.`,
    );
  }
});

test("missing preferred Pi slugs are recorded instead of hidden silently", () => {
  const missing = PROVIDER_MODEL_PREFERENCE_SOURCE.catalogDrift.missingPreferredModels;
  assert.equal(missing.some((entry) => entry.provider === "openai" && entry.modelId === "gpt-5.5-pro"), true);
  assert.equal(hasPiModel("openai", "gpt-5.5-pro"), false);
});

test("settings and model selector expose Simple/Advanced filtering and warning controls", () => {
  const settingsSource = readFileSync(
    join(process.cwd(), "apps", "taskpane", "src", "app", "components", "SettingsPage.tsx"),
    "utf8",
  );
  const appSource = readFileSync(
    join(process.cwd(), "apps", "taskpane", "src", "app", "App.tsx"),
    "utf8",
  );
  const modelSelectorSource = readFileSync(
    join(process.cwd(), "apps", "taskpane", "src", "app", "components", "ModelSelector.tsx"),
    "utf8",
  );

  assert.match(settingsSource, /SettingsDetailToggle/);
  assert.match(settingsSource, /model\.settingsVisibility === "simple" && model\.recommended/);
  assert.match(settingsSource, /defaultModelByProvider/);
  assert.match(settingsSource, /Drafting Style/);
  assert.match(settingsSource, /artifactClarificationMode/);
  assert.match(appSource, /pendingUnrecommendedModel/);
  assert.match(appSource, /suppressUnrecommendedModelWarning/);
  assert.match(appSource, /requiresUnrecommendedWarning/);
  assert.match(modelSelectorSource, /Advanced catalog model/);
});
