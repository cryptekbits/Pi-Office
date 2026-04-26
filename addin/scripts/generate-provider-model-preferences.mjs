import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { dirname, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { parse } from "yaml";
import { getModels, getProviders } from "@mariozechner/pi-ai";

const __dirname = dirname(fileURLToPath(import.meta.url));
const addinRoot = resolve(__dirname, "..");
const repoRoot = resolve(addinRoot, "..");
const yamlPath = resolve(addinRoot, "packages/pi-office-pack/src/provider-model-preferences.yaml");
const outputPath = resolve(addinRoot, "packages/pi-office-pack/src/provider-model-preferences.generated.ts");
const defaultPiMonoModelsPath = resolve(repoRoot, "../pi-mono/packages/ai/src/models.generated.ts");
const piMonoModelsPath = process.env.PI_MONO_MODELS_PATH
  ? resolve(process.env.PI_MONO_MODELS_PATH)
  : defaultPiMonoModelsPath;
const checkOnly = process.argv.includes("--check");

const VALID_VISIBILITY = new Set(["simple", "advanced"]);

function asObject(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error(`${label} must be an object.`);
  }
  return value;
}

function asString(value, label) {
  if (typeof value !== "string" || !value.trim()) {
    throw new Error(`${label} must be a non-empty string.`);
  }
  return value;
}

function asVisibility(value, label) {
  const visibility = asString(value, label);
  if (!VALID_VISIBILITY.has(visibility)) {
    throw new Error(`${label} must be "simple" or "advanced".`);
  }
  return visibility;
}

function loadCatalogFromPackage() {
  const providers = new Set(getProviders().map((provider) => String(provider)));
  const models = new Map();
  for (const provider of providers) {
    models.set(provider, new Set(getModels(provider).map((model) => String(model.id))));
  }
  return {
    providers,
    models,
    hasModel(provider, modelId) {
      return Boolean(models.get(provider)?.has(modelId));
    },
  };
}

function findObjectBlock(source, keyText) {
  const start = source.indexOf(keyText);
  if (start === -1) return "";
  const open = source.indexOf("{", start);
  if (open === -1) return "";
  let depth = 0;
  let inString = false;
  let escaped = false;
  for (let index = open; index < source.length; index += 1) {
    const char = source[index];
    if (inString) {
      if (escaped) {
        escaped = false;
      } else if (char === "\\") {
        escaped = true;
      } else if (char === "\"") {
        inString = false;
      }
      continue;
    }
    if (char === "\"") {
      inString = true;
      continue;
    }
    if (char === "{") depth += 1;
    if (char === "}") {
      depth -= 1;
      if (depth === 0) return source.slice(open, index + 1);
    }
  }
  return "";
}

function loadPiMonoSourceCatalog() {
  if (!existsSync(piMonoModelsPath)) return undefined;
  const source = readFileSync(piMonoModelsPath, "utf8");
  const providerBlocks = new Map();
  return {
    path: piMonoModelsPath,
    hasModel(provider, modelId) {
      if (!providerBlocks.has(provider)) {
        providerBlocks.set(provider, findObjectBlock(source, `\"${provider}\": {`));
      }
      const block = providerBlocks.get(provider);
      return block.includes(`\"${modelId}\": {`);
    },
  };
}

function sortedObject(object) {
  return Object.fromEntries(Object.entries(object).sort(([left], [right]) => left.localeCompare(right)));
}

function literal(value) {
  return JSON.stringify(value, null, 2);
}

function normalizeConfig(raw) {
  const root = asObject(raw, "provider model preference YAML");
  const providers = asObject(root.providers, "providers");
  const packageCatalog = loadCatalogFromPackage();
  const piMonoCatalog = loadPiMonoSourceCatalog();
  const providerPreferences = {};
  const modelPreferences = {};
  const defaultEnabledModelsByProvider = {};
  const simpleRecommendedModelsByProvider = {};
  const simpleVisibleProviders = [];
  const validationErrors = [];

  for (const [provider, providerConfigRaw] of Object.entries(providers)) {
    const providerConfig = asObject(providerConfigRaw, `providers.${provider}`);
    const settingsVisibility = asVisibility(providerConfig.settingsVisibility ?? "advanced", `providers.${provider}.settingsVisibility`);
    const label = providerConfig.label ? asString(providerConfig.label, `providers.${provider}.label`) : provider;
    const lab = providerConfig.lab ? asString(providerConfig.lab, `providers.${provider}.lab`) : label;
    const defaultModel = providerConfig.defaultModel
      ? asString(providerConfig.defaultModel, `providers.${provider}.defaultModel`)
      : undefined;
    const defaultEnabled = Boolean(providerConfig.defaultEnabled);
    const models = Array.isArray(providerConfig.models) ? providerConfig.models : [];

    if (settingsVisibility === "simple") {
      simpleVisibleProviders.push(provider);
      if (!defaultModel) {
        validationErrors.push(`${provider} is Simple-visible but has no defaultModel.`);
      }
    }
    if (!packageCatalog.providers.has(provider)) {
      validationErrors.push(`${provider} is not present in @mariozechner/pi-ai getProviders().`);
    }
    if (defaultModel && !packageCatalog.hasModel(provider, defaultModel)) {
      validationErrors.push(`${provider}/${defaultModel} defaultModel is missing from @mariozechner/pi-ai.`);
    }
    if (defaultModel && piMonoCatalog && !piMonoCatalog.hasModel(provider, defaultModel)) {
      validationErrors.push(`${provider}/${defaultModel} defaultModel is missing from refreshed pi-mono models.generated.ts.`);
    }

    providerPreferences[provider] = {
      provider,
      label,
      lab,
      settingsVisibility,
      defaultEnabled,
      ...(defaultModel ? { defaultModel } : {}),
    };
    modelPreferences[provider] = {};

    let defaultCount = 0;
    for (const modelConfigRaw of models) {
      const modelConfig = asObject(modelConfigRaw, `providers.${provider}.models[]`);
      const modelId = asString(modelConfig.modelId, `providers.${provider}.models[].modelId`);
      const modelVisibility = asVisibility(
        modelConfig.settingsVisibility ?? settingsVisibility,
        `providers.${provider}.models.${modelId}.settingsVisibility`,
      );
      const recommended = Boolean(modelConfig.recommended);
      const defaultForProvider = Boolean(modelConfig.defaultForProvider || modelId === defaultModel);
      if (defaultForProvider) defaultCount += 1;

      const existsInPackage = packageCatalog.hasModel(provider, modelId);
      const existsInPiMono = piMonoCatalog ? piMonoCatalog.hasModel(provider, modelId) : true;
      if (!existsInPackage) {
        validationErrors.push(`${provider}/${modelId} is missing from @mariozechner/pi-ai.`);
      }
      if (!existsInPiMono) {
        validationErrors.push(`${provider}/${modelId} is missing from refreshed pi-mono models.generated.ts.`);
      }
      if (modelVisibility === "simple" && recommended && (!existsInPackage || !existsInPiMono)) {
        validationErrors.push(`${provider}/${modelId} cannot be Simple recommended because it is not in the refreshed Pi catalog.`);
      }

      const preference = {
        provider,
        modelId,
        settingsVisibility: modelVisibility,
        lab: modelConfig.lab ? asString(modelConfig.lab, `providers.${provider}.models.${modelId}.lab`) : lab,
        family: modelConfig.family ? asString(modelConfig.family, `providers.${provider}.models.${modelId}.family`) : undefined,
        recommended,
        recommendationReason: modelConfig.recommendationReason
          ? asString(modelConfig.recommendationReason, `providers.${provider}.models.${modelId}.recommendationReason`)
          : undefined,
        defaultForProvider,
        requiresUnrecommendedWarning: !recommended,
      };
      modelPreferences[provider][modelId] = Object.fromEntries(
        Object.entries(preference).filter(([, value]) => value !== undefined),
      );

      if (modelVisibility === "simple" && recommended) {
        simpleRecommendedModelsByProvider[provider] ??= [];
        simpleRecommendedModelsByProvider[provider].push(modelId);
        if (defaultEnabled) {
          defaultEnabledModelsByProvider[provider] ??= [];
          defaultEnabledModelsByProvider[provider].push(modelId);
        }
      }
    }

    if (settingsVisibility === "simple" && defaultCount !== 1) {
      validationErrors.push(`${provider} must have exactly one Simple provider default; found ${defaultCount}.`);
    }
  }

  const missingPreferredModels = Array.isArray(root.catalogDrift?.missingPreferredModels)
    ? root.catalogDrift.missingPreferredModels.map((entry, index) => {
        const item = asObject(entry, `catalogDrift.missingPreferredModels[${index}]`);
        const provider = asString(item.provider, `catalogDrift.missingPreferredModels[${index}].provider`);
        const modelId = asString(item.modelId, `catalogDrift.missingPreferredModels[${index}].modelId`);
        if (packageCatalog.hasModel(provider, modelId)) {
          validationErrors.push(`${provider}/${modelId} is listed as missing but exists in @mariozechner/pi-ai.`);
        }
        if (piMonoCatalog?.hasModel(provider, modelId)) {
          validationErrors.push(`${provider}/${modelId} is listed as missing but exists in refreshed pi-mono.`);
        }
        return {
          provider,
          modelId,
          requestedLabel: item.requestedLabel ? String(item.requestedLabel) : modelId,
          evidence: item.evidence ? String(item.evidence) : "Missing from the validated Pi catalog.",
        };
      })
    : [];

  if (validationErrors.length) {
    throw new Error(`Provider model preference validation failed:\n- ${validationErrors.join("\n- ")}`);
  }

  return {
    catalogEvidence: {
      ...(root.catalogEvidence ?? {}),
      packageCatalog: "@mariozechner/pi-ai",
      piMonoModelsPath: piMonoCatalog ? relative(repoRoot, piMonoCatalog.path).replace(/\\/g, "/") : undefined,
    },
    catalogDrift: { missingPreferredModels },
    providerPreferences: sortedObject(providerPreferences),
    modelPreferences: sortedObject(
      Object.fromEntries(
        Object.entries(modelPreferences).map(([provider, models]) => [provider, sortedObject(models)]),
      ),
    ),
    defaultEnabledModelsByProvider: sortedObject(defaultEnabledModelsByProvider),
    simpleRecommendedModelsByProvider: sortedObject(simpleRecommendedModelsByProvider),
    simpleVisibleProviders: simpleVisibleProviders.sort(),
  };
}

function generateSource(config) {
  return `// Generated by scripts/generate-provider-model-preferences.mjs from provider-model-preferences.yaml.
// Do not edit this file directly.

export type ProviderModelSettingsVisibility = "simple" | "advanced";

export interface ProviderSettingsPreference {
  provider: string;
  label: string;
  lab: string;
  settingsVisibility: ProviderModelSettingsVisibility;
  defaultEnabled: boolean;
  defaultModel?: string | undefined;
}

export interface ProviderModelPreference {
  provider: string;
  modelId: string;
  settingsVisibility: ProviderModelSettingsVisibility;
  lab: string;
  family?: string | undefined;
  recommended: boolean;
  recommendationReason?: string | undefined;
  defaultForProvider: boolean;
  requiresUnrecommendedWarning: boolean;
}

export const PROVIDER_MODEL_PREFERENCE_SOURCE = ${literal({
  catalogEvidence: config.catalogEvidence,
  catalogDrift: config.catalogDrift,
})} as const;

export const PROVIDER_SETTINGS_PREFERENCES = ${literal(config.providerPreferences)} as const satisfies Record<string, ProviderSettingsPreference>;

export const PROVIDER_MODEL_PREFERENCES = ${literal(config.modelPreferences)} as const satisfies Record<string, Record<string, ProviderModelPreference>>;

export const DEFAULT_ENABLED_MODELS_BY_PROVIDER = ${literal(config.defaultEnabledModelsByProvider)} as const satisfies Record<string, readonly string[]>;

export const SIMPLE_RECOMMENDED_MODELS_BY_PROVIDER = ${literal(config.simpleRecommendedModelsByProvider)} as const satisfies Record<string, readonly string[]>;

export const SIMPLE_VISIBLE_PROVIDERS = ${literal(config.simpleVisibleProviders)} as const satisfies readonly string[];

const PROVIDER_SETTINGS_LOOKUP = PROVIDER_SETTINGS_PREFERENCES as Record<string, ProviderSettingsPreference>;
const PROVIDER_MODEL_LOOKUP = PROVIDER_MODEL_PREFERENCES as Record<string, Record<string, ProviderModelPreference>>;

export function getProviderSettingsPreference(provider: string): ProviderSettingsPreference {
  return PROVIDER_SETTINGS_LOOKUP[provider] ?? {
    provider,
    label: provider,
    lab: provider,
    settingsVisibility: "advanced",
    defaultEnabled: false,
  };
}

export function getProviderModelPreference(provider: string, modelId: string): ProviderModelPreference | undefined {
  return PROVIDER_MODEL_LOOKUP[provider]?.[modelId];
}

export function getProviderDefaultModel(provider: string): string | undefined {
  return getProviderSettingsPreference(provider).defaultModel;
}

export function isSimpleVisibleProvider(provider: string): boolean {
  return getProviderSettingsPreference(provider).settingsVisibility === "simple";
}
`;
}

const raw = parse(readFileSync(yamlPath, "utf8"));
const normalized = normalizeConfig(raw);
const generated = generateSource(normalized);

if (checkOnly) {
  const current = existsSync(outputPath) ? readFileSync(outputPath, "utf8") : "";
  if (current !== generated) {
    throw new Error(`${relative(repoRoot, outputPath)} is out of date. Run npm --prefix addin run generate:provider-models.`);
  }
} else {
  writeFileSync(outputPath, generated);
  console.log(`Generated ${relative(repoRoot, outputPath).replace(/\\/g, "/")}`);
}
