import type {
  CompanionProviderApiKeyRequest,
  CompanionProviderAuthCapability,
  CompanionProviderAuthClearRequest,
  CompanionProviderAuthStatusResponse,
  ProviderAuthDescriptor,
  ProviderAuthMethod,
} from "@pi-office/pi-office-pack/protocol";
import type { CompanionConfig } from "./config.js";
import {
  createCompanionSecretStore,
  type OAuthTokenPersistence,
} from "./oauth-token-store.js";

interface StoredProviderAuthRecord {
  provider: string;
  authMethod: "api_key";
  secret: string;
  storedAt: string;
  verifiedAt?: string | undefined;
  lastVerificationAttemptAt?: string | undefined;
  lastVerificationError?: string | undefined;
}

interface StoredProviderAuthEnvelope {
  version: 1;
  providers: Record<string, StoredProviderAuthRecord>;
}

interface CreateProviderAuthStoreOptions {
  persistence?: OAuthTokenPersistence | undefined;
}

const SUPPORTED_COMPANION_AUTH_METHODS: ProviderAuthMethod[] = ["api_key"];

function emptyEnvelope(): StoredProviderAuthEnvelope {
  return { version: 1, providers: {} };
}

function normalizeProvider(provider: unknown): string {
  return typeof provider === "string" ? provider.trim() : "";
}

function toDescriptor(record: StoredProviderAuthRecord | undefined, provider: string): ProviderAuthDescriptor {
  if (!record) {
    return {
      provider,
      state: "not_configured",
      credentialStored: false,
      verifiedUsable: false,
    };
  }
  return {
    provider,
    state: record.lastVerificationError
      ? "verification_failed"
      : record.verifiedAt
        ? "verified_usable"
        : "credential_stored",
    credentialStored: true,
    verifiedUsable: Boolean(record.verifiedAt) && !record.lastVerificationError,
    verifiedAt: record.verifiedAt,
    lastVerificationAttemptAt: record.lastVerificationAttemptAt,
    lastVerificationError: record.lastVerificationError,
  };
}

export class CompanionProviderAuthStore {
  constructor(private readonly persistence: OAuthTokenPersistence) {}

  static create(config: CompanionConfig, options: CreateProviderAuthStoreOptions = {}): CompanionProviderAuthStore {
    return new CompanionProviderAuthStore(
      options.persistence ?? createCompanionSecretStore(config, {
        legacyFileName: "provider-auth.json",
        encryptedFileName: "provider-auth.dpapi.json",
      }),
    );
  }

  getCapability(): CompanionProviderAuthCapability {
    let configuredProviderCount = 0;
    let loadError: string | undefined;
    try {
      configuredProviderCount = this.status().storedProviders.length;
    } catch (error) {
      loadError = error instanceof Error ? error.message : String(error);
    }

    if (!this.persistence.secure) {
      return {
        state: "unavailable",
        available: false,
        version: "companion-provider-auth-v1",
        explicitMigrationRequired: true,
        supportedAuthMethods: SUPPORTED_COMPANION_AUTH_METHODS,
        secureStorage: false,
        storageKind: this.persistence.storageKind,
        configuredProviderCount,
        reason: "Secure companion provider storage is unavailable on this platform; non-Windows keychain support is tracked separately.",
      };
    }
    if (loadError) {
      return {
        state: "degraded",
        available: false,
        version: "companion-provider-auth-v1",
        explicitMigrationRequired: true,
        supportedAuthMethods: SUPPORTED_COMPANION_AUTH_METHODS,
        secureStorage: true,
        storageKind: this.persistence.storageKind,
        configuredProviderCount: 0,
        reason: `Companion provider auth storage is unreadable and has been disabled: ${loadError}`,
      };
    }
    return {
      state: "available",
      available: true,
      version: "companion-provider-auth-v1",
      explicitMigrationRequired: true,
      supportedAuthMethods: SUPPORTED_COMPANION_AUTH_METHODS,
      secureStorage: true,
      storageKind: this.persistence.storageKind,
      configuredProviderCount,
      reason: "Companion API-key provider storage is available. Taskpane secrets still require explicit user action before moving.",
    };
  }

  status(providerIds?: string[]): CompanionProviderAuthStatusResponse {
    const envelope = this.loadEnvelope();
    const ids = Array.from(new Set([
      ...Object.keys(envelope.providers),
      ...(providerIds ?? []).map((id) => id.trim()).filter(Boolean),
    ])).sort();
    const providerStates = ids.map((provider) => toDescriptor(envelope.providers[provider], provider));
    const storedProviders = providerStates.filter((entry) => entry.credentialStored).map((entry) => entry.provider);
    return {
      ok: true,
      storageKind: this.persistence.storageKind,
      secureStorage: this.persistence.secure,
      explicitMigrationRequired: true,
      supportedAuthMethods: SUPPORTED_COMPANION_AUTH_METHODS,
      storedProviders,
      configuredProviders: providerStates.filter((entry) => entry.verifiedUsable).map((entry) => entry.provider),
      verifiedProviders: providerStates.filter((entry) => entry.verifiedUsable).map((entry) => entry.provider),
      unverifiedProviders: providerStates
        .filter((entry) => entry.credentialStored && !entry.verifiedUsable && entry.state !== "verification_failed")
        .map((entry) => entry.provider),
      verificationFailedProviders: providerStates
        .filter((entry) => entry.state === "verification_failed")
        .map((entry) => entry.provider),
      providerStates,
    };
  }

  getApiKey(provider: string): string | undefined {
    const normalized = normalizeProvider(provider);
    if (!normalized) return undefined;
    return this.loadEnvelope().providers[normalized]?.secret;
  }

  markVerificationSuccess(provider: string): CompanionProviderAuthStatusResponse {
    const normalized = normalizeProvider(provider);
    if (!normalized) {
      throw new Error("provider is required.");
    }
    const envelope = this.loadEnvelope();
    const record = envelope.providers[normalized];
    if (!record) return this.status([normalized]);
    const verifiedAt = new Date().toISOString();
    envelope.providers[normalized] = {
      ...record,
      verifiedAt,
      lastVerificationAttemptAt: verifiedAt,
      lastVerificationError: undefined,
    };
    this.saveEnvelope(envelope);
    return this.status([normalized]);
  }

  markVerificationFailure(provider: string, error: unknown): CompanionProviderAuthStatusResponse {
    const normalized = normalizeProvider(provider);
    if (!normalized) {
      throw new Error("provider is required.");
    }
    const envelope = this.loadEnvelope();
    const record = envelope.providers[normalized];
    if (!record) return this.status([normalized]);
    envelope.providers[normalized] = {
      ...record,
      verifiedAt: undefined,
      lastVerificationAttemptAt: new Date().toISOString(),
      lastVerificationError: error instanceof Error ? error.message : String(error),
    };
    this.saveEnvelope(envelope);
    return this.status([normalized]);
  }

  setApiKey(request: CompanionProviderApiKeyRequest): CompanionProviderAuthStatusResponse {
    if (!this.persistence.secure) {
      throw new Error("Secure companion provider auth storage is unavailable on this platform.");
    }
    const provider = normalizeProvider(request?.provider);
    if (!provider) {
      throw new Error("provider is required.");
    }
    if (request.explicitUserAction !== true) {
      throw new Error("explicitUserAction=true is required before storing provider secrets in the companion.");
    }
    if (typeof request.apiKey !== "string" || !request.apiKey.trim()) {
      throw new Error("apiKey is required.");
    }

    const envelope = this.loadEnvelope();
    envelope.providers[provider] = {
      provider,
      authMethod: "api_key",
      secret: request.apiKey,
      storedAt: new Date().toISOString(),
    };
    this.saveEnvelope(envelope);
    return this.status([provider]);
  }

  clear(request: CompanionProviderAuthClearRequest | undefined): CompanionProviderAuthStatusResponse {
    const provider = normalizeProvider(request?.provider);
    if (!provider) {
      this.persistence.clear();
      return this.status();
    }

    const envelope = this.loadEnvelope();
    delete envelope.providers[provider];
    if (Object.keys(envelope.providers).length === 0) {
      this.persistence.clear();
    } else {
      this.saveEnvelope(envelope);
    }
    return this.status([provider]);
  }

  private loadEnvelope(): StoredProviderAuthEnvelope {
    const raw = this.persistence.load();
    if (!raw) return emptyEnvelope();
    try {
      const parsed = JSON.parse(raw) as Partial<StoredProviderAuthEnvelope>;
      if (parsed.version !== 1 || !parsed.providers || typeof parsed.providers !== "object") {
        throw new Error("unsupported schema");
      }
      const providers: Record<string, StoredProviderAuthRecord> = {};
      for (const [provider, record] of Object.entries(parsed.providers)) {
        if (!record || typeof record !== "object") continue;
        const normalized = normalizeProvider(record.provider || provider);
        if (!normalized || record.authMethod !== "api_key" || typeof record.secret !== "string") continue;
        providers[normalized] = {
          provider: normalized,
          authMethod: "api_key",
          secret: record.secret,
          storedAt: typeof record.storedAt === "string" ? record.storedAt : new Date(0).toISOString(),
          verifiedAt: typeof record.verifiedAt === "string" ? record.verifiedAt : undefined,
          lastVerificationAttemptAt: typeof record.lastVerificationAttemptAt === "string"
            ? record.lastVerificationAttemptAt
            : undefined,
          lastVerificationError: typeof record.lastVerificationError === "string"
            ? record.lastVerificationError
            : undefined,
        };
      }
      return { version: 1, providers };
    } catch (error) {
      throw new Error(`Companion provider auth store is unreadable: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  private saveEnvelope(envelope: StoredProviderAuthEnvelope): void {
    this.persistence.save(`${JSON.stringify(envelope, null, 2)}\n`);
  }
}
