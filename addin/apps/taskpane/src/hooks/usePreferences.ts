import { useCallback, useEffect, useState } from "react";
import { DEFAULT_USER_PREFERENCES, type UserPreferences } from "@pi-office/pi-office-pack/protocol";
import { DEFAULT_ENABLED_MODELS_BY_PROVIDER } from "@pi-office/pi-office-pack/provider-model-preferences";
import { syncKernelPreferences, syncKernelProviderSelection } from "../lib/runtime/inprocess-kernel";

const STORAGE_KEY = "pi-office-preferences";
const ENABLED_MODELS_KEY = "pi-office-enabled-models";
const ENABLED_PROVIDERS_KEY = "pi-office-enabled-providers";

/**
 * Default-enabled model IDs keyed by provider.
 * Models not in this list are hidden by default but can be enabled by the user.
 */
export const DEFAULT_ENABLED_MODELS: Record<string, string[]> = Object.fromEntries(
  Object.entries(DEFAULT_ENABLED_MODELS_BY_PROVIDER).map(([provider, modelIds]) => [provider, [...modelIds]]),
);

function buildDefaultEnabledSet(): Set<string> {
  const set = new Set<string>();
  for (const [provider, ids] of Object.entries(DEFAULT_ENABLED_MODELS)) {
    for (const id of ids) set.add(`${provider}::${id}`);
  }
  return set;
}

function loadEnabledModels(): Set<string> {
  try {
    const raw = localStorage.getItem(ENABLED_MODELS_KEY);
    if (!raw) return buildDefaultEnabledSet();
    return new Set<string>(JSON.parse(raw));
  } catch {
    return buildDefaultEnabledSet();
  }
}

function saveEnabledModels(set: Set<string>): void {
  try {
    localStorage.setItem(ENABLED_MODELS_KEY, JSON.stringify([...set]));
  } catch {}
}

function loadEnabledProviders(): Set<string> {
  try {
    const raw = localStorage.getItem(ENABLED_PROVIDERS_KEY);
    if (!raw) return new Set(Object.keys(DEFAULT_ENABLED_MODELS));
    return new Set<string>(JSON.parse(raw));
  } catch {
    return new Set(Object.keys(DEFAULT_ENABLED_MODELS));
  }
}

function saveEnabledProviders(set: Set<string>): void {
  try {
    localStorage.setItem(ENABLED_PROVIDERS_KEY, JSON.stringify([...set]));
  } catch {}
}

export function useEnabledModels() {
  const [enabledModels, setEnabledModels] = useState<Set<string>>(loadEnabledModels);
  const [enabledProviders, setEnabledProviders] = useState<Set<string>>(loadEnabledProviders);

  useEffect(() => {
    syncKernelProviderSelection({
      enabledModels: [...enabledModels],
      enabledProviders: [...enabledProviders],
    });
  }, [enabledModels, enabledProviders]);

  const toggleModel = useCallback((key: string) => {
    setEnabledModels((current) => {
      const next = new Set(current);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      saveEnabledModels(next);
      return next;
    });
  }, []);

  const toggleProvider = useCallback((provider: string) => {
    setEnabledProviders((current) => {
      const next = new Set(current);
      if (next.has(provider)) next.delete(provider);
      else next.add(provider);
      saveEnabledProviders(next);
      return next;
    });
  }, []);

  const isModelEnabled = useCallback(
    (providerKey: string, modelId: string) =>
      enabledProviders.has(providerKey) && enabledModels.has(`${providerKey}::${modelId}`),
    [enabledModels, enabledProviders],
  );

  return { enabledModels, enabledProviders, toggleModel, toggleProvider, isModelEnabled } as const;
}

function loadPreferences(): UserPreferences {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { ...DEFAULT_USER_PREFERENCES };
    return { ...DEFAULT_USER_PREFERENCES, ...JSON.parse(raw) };
  } catch {
    return { ...DEFAULT_USER_PREFERENCES };
  }
}

function savePreferences(prefs: UserPreferences): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(prefs));
  } catch {
    // localStorage quota exceeded or unavailable
  }
}

function syncPreferencesToCompanion(prefs: UserPreferences): void {
  try {
    syncKernelPreferences(prefs);
  } catch {
    // non-critical in browser sandbox
  }
}

export function usePreferences() {
  const [preferences, setPreferencesState] = useState<UserPreferences>(() => {
    const prefs = loadPreferences();
    syncPreferencesToCompanion(prefs);
    return prefs;
  });

  const updatePreferences = useCallback((patch: Partial<UserPreferences>) => {
    setPreferencesState((current) => {
      const next = { ...current, ...patch };
      savePreferences(next);
      syncPreferencesToCompanion(next);
      return next;
    });
  }, []);

  return { preferences, updatePreferences } as const;
}

const RECENT_MODELS_KEY = "pi-office-recent-models";
const MAX_RECENT = 5;

export function getRecentModels(): string[] {
  try {
    const raw = localStorage.getItem(RECENT_MODELS_KEY);
    if (!raw) return [];
    return JSON.parse(raw);
  } catch {
    return [];
  }
}

export function pushRecentModel(modelKey: string): void {
  try {
    const list = getRecentModels().filter((k) => k !== modelKey);
    list.unshift(modelKey);
    localStorage.setItem(RECENT_MODELS_KEY, JSON.stringify(list.slice(0, MAX_RECENT)));
  } catch {
    // ignore
  }
}
