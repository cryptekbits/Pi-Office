import assert from "node:assert/strict";
import test from "node:test";

class MemoryStorage {
  private readonly store = new Map<string, string>();

  clear(): void {
    this.store.clear();
  }

  getItem(key: string): string | null {
    return this.store.has(key) ? this.store.get(key)! : null;
  }

  key(index: number): string | null {
    const keys = Array.from(this.store.keys());
    return keys[index] ?? null;
  }

  removeItem(key: string): void {
    this.store.delete(key);
  }

  setItem(key: string, value: string): void {
    this.store.set(key, String(value));
  }

  get length(): number {
    return this.store.size;
  }
}

function installRuntimePolyfills(): Storage {
  const globalAny = globalThis as unknown as {
    localStorage?: Storage;
    window?: { location?: { origin?: string }; open?: (...args: unknown[]) => unknown };
    WebSocket?: { CONNECTING: number; OPEN: number; CLOSING: number; CLOSED: number };
    CloseEvent?: typeof CloseEvent;
    MessageEvent?: typeof MessageEvent;
    atob?: (value: string) => string;
    btoa?: (value: string) => string;
  };

  globalAny.localStorage = new MemoryStorage() as unknown as Storage;
  globalAny.window = globalAny.window ?? {};
  globalAny.window.location = globalAny.window.location ?? { origin: "https://localhost:3443" };
  globalAny.window.location.origin = globalAny.window.location.origin ?? "https://localhost:3443";
  globalAny.window.open = globalAny.window.open ?? (() => null);

  if (!globalAny.WebSocket) {
    globalAny.WebSocket = {
      CONNECTING: 0,
      OPEN: 1,
      CLOSING: 2,
      CLOSED: 3,
    };
  }

  if (!globalAny.CloseEvent) {
    class CloseEventPolyfill extends Event {
      readonly code: number;
      readonly reason: string;
      readonly wasClean: boolean;

      constructor(type: string, init?: { code?: number; reason?: string; wasClean?: boolean }) {
        super(type);
        this.code = init?.code ?? 0;
        this.reason = init?.reason ?? "";
        this.wasClean = init?.wasClean ?? true;
      }
    }
    globalAny.CloseEvent = CloseEventPolyfill as unknown as typeof CloseEvent;
  }

  if (!globalAny.MessageEvent) {
    class MessageEventPolyfill<T = unknown> extends Event {
      readonly data: T;

      constructor(type: string, init?: { data?: T }) {
        super(type);
        this.data = init?.data as T;
      }
    }
    globalAny.MessageEvent = MessageEventPolyfill as unknown as typeof MessageEvent;
  }

  if (!globalAny.atob) {
    globalAny.atob = (value: string) => Buffer.from(value, "base64").toString("binary");
  }
  if (!globalAny.btoa) {
    globalAny.btoa = (value: string) => Buffer.from(value, "binary").toString("base64");
  }

  return globalAny.localStorage;
}

async function loadKernelModule(seed?: (storage: Storage) => void) {
  const storage = installRuntimePolyfills();
  seed?.(storage);
  const specifier = `../../../apps/taskpane/src/lib/runtime/inprocess-kernel.js?test=${Date.now()}-${Math.random()}`;
  return import(specifier);
}

function waitForEvent(target: EventTarget, name: string, timeoutMs = 2_000): Promise<Event> {
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      target.removeEventListener(name, onEvent);
      reject(new Error(`Timed out waiting for ${name} event.`));
    }, timeoutMs);

    const onEvent = (event: Event) => {
      clearTimeout(timeout);
      target.removeEventListener(name, onEvent);
      resolve(event);
    };

    target.addEventListener(name, onEvent);
  });
}

async function openSession(runtime: {
  dispatchKernelRequest: (path: string, init?: RequestInit) => Promise<unknown>;
  createLocalBridgeSocket: (sessionId: string) => EventTarget;
}, label: string) {
  const openResponse = await runtime.dispatchKernelRequest("/v1/sessions/open", {
    method: "POST",
    body: JSON.stringify({
      host: "word",
      documentId: `doc-provider-auth-${label}`,
      saved: true,
      title: `Provider Auth ${label}`,
    }),
  }) as { sessionId: string };

  const socket = runtime.createLocalBridgeSocket(openResponse.sessionId);
  await waitForEvent(socket, "open");
  return (socket as unknown as { session: { agent: { state: { tools: Array<{ name: string; execute: (toolCallId: string, params: Record<string, unknown>) => Promise<unknown> }> } } } }).session;
}

async function providerState(runtime: { dispatchKernelRequest: (path: string, init?: RequestInit) => Promise<unknown> }, provider: string) {
  const authStatus = await runtime.dispatchKernelRequest("/v1/auth/status") as {
    providerStates: Array<{
      provider: string;
      state: string;
      credentialStored: boolean;
      verifiedUsable: boolean;
      lastVerificationError?: string;
    }>;
  };
  const state = authStatus.providerStates.find((entry) => entry.provider === provider);
  assert.ok(state, `Expected ${provider} in providerStates.`);
  return state;
}

test("provider auth readiness distinguishes stored credentials from verified usability", async () => {
  const runtime = await loadKernelModule();
  const initialCatalog = await runtime.dispatchKernelRequest("/v1/providers") as {
    providers: Array<{
      provider: string;
      authState: string;
      credentialStored: boolean;
      verifiedUsable: boolean;
      configured: boolean;
      models: Array<{
        authState: string;
        credentialStored: boolean;
        verifiedUsable: boolean;
        configured: boolean;
      }>;
    }>;
  };

  const provider = initialCatalog.providers.find((entry) => entry.provider === "openai") ?? initialCatalog.providers[0];
  assert.ok(provider, "Expected at least one browser-supported provider.");
  assert.equal(provider.authState, "not_configured");
  assert.equal(provider.credentialStored, false);
  assert.equal(provider.verifiedUsable, false);
  assert.equal(provider.configured, false);

  await runtime.dispatchKernelRequest("/v1/auth/api-key", {
    method: "POST",
    body: JSON.stringify({ provider: provider.provider, apiKey: "test-provider-key" }),
  });

  const authStatus = await runtime.dispatchKernelRequest("/v1/auth/status") as {
    storedProviders: string[];
    configuredProviders: string[];
    verifiedProviders: string[];
    unverifiedProviders: string[];
    verificationFailedProviders: string[];
    providerStates: Array<{
      provider: string;
      state: string;
      credentialStored: boolean;
      verifiedUsable: boolean;
    }>;
  };
  const state = authStatus.providerStates.find((entry) => entry.provider === provider.provider);
  assert.ok(state, `Expected ${provider.provider} in providerStates.`);
  assert.equal(state.state, "credential_stored");
  assert.equal(state.credentialStored, true);
  assert.equal(state.verifiedUsable, false);
  assert.equal(authStatus.storedProviders.includes(provider.provider), true);
  assert.equal(authStatus.unverifiedProviders.includes(provider.provider), true);
  assert.equal(authStatus.verifiedProviders.includes(provider.provider), false);
  assert.equal(authStatus.configuredProviders.includes(provider.provider), false);
  assert.equal(authStatus.verificationFailedProviders.includes(provider.provider), false);

  const updatedCatalog = await runtime.dispatchKernelRequest("/v1/providers") as typeof initialCatalog;
  const updatedProvider = updatedCatalog.providers.find((entry) => entry.provider === provider.provider);
  assert.ok(updatedProvider, `Expected ${provider.provider} in refreshed catalog.`);
  assert.equal(updatedProvider.authState, "credential_stored");
  assert.equal(updatedProvider.credentialStored, true);
  assert.equal(updatedProvider.verifiedUsable, false);
  assert.equal(updatedProvider.configured, true, "legacy configured remains a credential-present/executable flag.");
  assert.ok(updatedProvider.models.length > 0);
  for (const model of updatedProvider.models) {
    assert.equal(model.authState, "credential_stored");
    assert.equal(model.credentialStored, true);
    assert.equal(model.verifiedUsable, false);
    assert.equal(model.configured, true);
  }
});

test("provider catalog distinguishes browser API-key providers from companion OAuth providers", async () => {
  const runtime = await loadKernelModule();
  const catalog = await runtime.dispatchKernelRequest("/v1/providers") as {
    providers: Array<{
      provider: string;
      supportStatus: string;
      runtimeSurface: string;
      authMethods: string[];
      apiKeySupported: boolean;
      oauthSupported: boolean;
      browserCallable: boolean;
      companionRequired: boolean;
      subscriptionBacked: boolean;
      imageGenerationSupported: boolean;
      models: Array<{
        usesApiKey: boolean;
        configured: boolean;
        browserCallable: boolean;
        companionRequired: boolean;
      }>;
    }>;
  };

  const provider = (id: string) => {
    const entry = catalog.providers.find((item) => item.provider === id);
    assert.ok(entry, `Expected ${id} in provider catalog.`);
    return entry;
  };

  const openAi = provider("openai");
  assert.equal(openAi.supportStatus, "supported");
  assert.equal(openAi.runtimeSurface, "browser_taskpane");
  assert.equal(openAi.apiKeySupported, true);
  assert.equal(openAi.browserCallable, true);
  assert.equal(openAi.oauthSupported, false);
  assert.equal(openAi.companionRequired, false);
  assert.equal(openAi.imageGenerationSupported, true);
  assert.deepEqual(openAi.authMethods, ["api_key"]);
  assert.ok(openAi.models.length > 0);
  assert.equal(openAi.models.every((model) => model.usesApiKey && model.browserCallable), true);

  for (const id of ["openai-codex", "github-copilot", "google-gemini-cli", "google-antigravity"]) {
    const entry = provider(id);
    assert.equal(entry.supportStatus, "planned", `${id} should be planned until companion auth exists.`);
    assert.equal(entry.runtimeSurface, "companion");
    assert.equal(entry.apiKeySupported, false);
    assert.equal(entry.oauthSupported, false, `${id} must not show a usable browser OAuth button.`);
    assert.equal(entry.browserCallable, false);
    assert.equal(entry.companionRequired, true);
    assert.equal(entry.subscriptionBacked, true);
    assert.equal(entry.authMethods.includes("oauth"), true);
    assert.equal(entry.models.every((model) => !model.usesApiKey && !model.configured && !model.browserCallable), true);
  }

  const bedrock = provider("amazon-bedrock");
  assert.equal(bedrock.browserCallable, false);
  assert.equal(bedrock.companionRequired, true);
  assert.equal(bedrock.authMethods.includes("aws_credentials"), true);
});

test("provider auth routes reject unavailable browser auth methods", async () => {
  const runtime = await loadKernelModule();

  await assert.rejects(
    () => runtime.dispatchKernelRequest("/v1/auth/api-key", {
      method: "POST",
      body: JSON.stringify({ provider: "openai-codex", apiKey: "not-a-real-token" }),
    }),
    /does not accept browser-stored API keys/i,
  );

  await assert.rejects(
    () => runtime.dispatchKernelRequest("/v1/auth/start", {
      method: "POST",
      body: JSON.stringify({ providerId: "openai-codex" }),
    }),
    /requires companion-owned OAuth/i,
  );

  await assert.rejects(
    () => runtime.dispatchKernelRequest("/v1/auth/start", {
      method: "POST",
      body: JSON.stringify({ providerId: "openai" }),
    }),
    /OAuth sign-in is unavailable/i,
  );
});

test("image model catalog is OpenAI-only until additional browser execution paths exist", async () => {
  const runtime = await loadKernelModule();
  const catalog = await runtime.dispatchKernelRequest("/v1/image-models") as {
    models: Array<{ provider: string; apiType: string; configured: boolean }>;
    defaultModelKey: string;
  };

  assert.ok(catalog.models.length > 0, "Expected at least one image model.");
  assert.equal(catalog.models.every((model) => model.provider === "openai"), true);
  assert.equal(catalog.models.every((model) => model.apiType === "openai-images"), true);
  assert.match(catalog.defaultModelKey, /^openai::/);

  await assert.rejects(
    () => runtime.dispatchKernelRequest("/v1/preferences", {
      method: "POST",
      body: JSON.stringify({ defaultImageModel: "openrouter::imaginary-image-model" }),
    }),
    /not available in the browser taskpane image catalog/i,
  );
});

test("legacy stored auth records migrate to unverified instead of ready", async () => {
  const runtime = await loadKernelModule((storage) => {
    storage.setItem("pi-office-auth", JSON.stringify([{ provider: "openai", apiKey: "legacy-key" }]));
  });

  const authStatus = await runtime.dispatchKernelRequest("/v1/auth/status") as {
    storedProviders: string[];
    configuredProviders: string[];
    unverifiedProviders: string[];
    providerStates: Array<{
      provider: string;
      state: string;
      credentialStored: boolean;
      verifiedUsable: boolean;
    }>;
  };
  const openAi = authStatus.providerStates.find((entry) => entry.provider === "openai");
  assert.ok(openAi, "Expected OpenAI provider auth state.");
  assert.equal(openAi.state, "credential_stored");
  assert.equal(openAi.credentialStored, true);
  assert.equal(openAi.verifiedUsable, false);
  assert.equal(authStatus.storedProviders.includes("openai"), true);
  assert.equal(authStatus.unverifiedProviders.includes("openai"), true);
  assert.equal(authStatus.configuredProviders.includes("openai"), false);
});

test("provider auth failures demote readiness and later successful use recovers it", async () => {
  const runtime = await loadKernelModule();
  await runtime.dispatchKernelRequest("/v1/auth/api-key", {
    method: "POST",
    body: JSON.stringify({ provider: "openai", apiKey: "test-provider-key" }),
  });
  await runtime.dispatchKernelRequest("/v1/preferences", {
    method: "POST",
    body: JSON.stringify({
      imageGenerationEnabled: true,
      defaultImageModel: "openai::gpt-image-1",
    }),
  });

  const session = await openSession(runtime, "recover");
  const imageTool = session.agent.state.tools.find((tool) => tool.name === "generate_image");
  assert.ok(imageTool, "Expected generate_image tool in session.");

  const originalFetch = globalThis.fetch;
  const globalAny = globalThis as unknown as { fetch: typeof fetch };
  try {
    globalAny.fetch = (async () => new Response("unauthorized", { status: 401, statusText: "Unauthorized" })) as typeof fetch;
    await assert.rejects(
      () => imageTool.execute("image-auth-failure", { prompt: "small chart icon", insert: false }),
      /unauthorized/i,
    );
    const failed = await providerState(runtime, "openai");
    assert.equal(failed.state, "verification_failed");
    assert.equal(failed.credentialStored, true);
    assert.equal(failed.verifiedUsable, false);
    assert.match(failed.lastVerificationError ?? "", /unauthorized/i);

    globalAny.fetch = (async () => new Response(
      JSON.stringify({ data: [{ b64_json: Buffer.from("fake-png").toString("base64") }] }),
      { status: 200, headers: { "Content-Type": "application/json" } },
    )) as typeof fetch;
    await imageTool.execute("image-auth-success", { prompt: "small chart icon", insert: false });

    const recovered = await providerState(runtime, "openai");
    assert.equal(recovered.state, "verified_usable");
    assert.equal(recovered.credentialStored, true);
    assert.equal(recovered.verifiedUsable, true);
    assert.equal(recovered.lastVerificationError, undefined);
  } finally {
    globalAny.fetch = originalFetch;
  }
});
