import { createHash, randomBytes } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import type {
  CompanionConnectorDefinition,
  CompanionConnectorOAuthStatusRequest,
  CompanionConnectorOAuthStatusResponse,
  ConnectorOAuthStartResponse,
} from "@pi-office/pi-office-pack/protocol";
import type { CompanionConfig } from "./config.js";

interface OAuthServerMetadata {
  authorization_endpoint?: string;
  token_endpoint?: string;
  registration_endpoint?: string;
  scopes_supported?: string[];
}

interface OAuthTokenResponse {
  access_token?: string;
  refresh_token?: string;
  token_type?: string;
  scope?: string;
  expires_in?: number;
}

interface OAuthFlowRecord {
  connectorId: string;
  connectorCatalogId: string;
  setupProfileId?: string | undefined;
  state: string;
  authorizationUrl: string;
  tokenEndpoint: string;
  redirectUri: string;
  clientId: string;
  clientSecret?: string | undefined;
  codeVerifier: string;
  scope?: string | undefined;
  createdAt: string;
  expiresAt: string;
  completedAt?: string | undefined;
  error?: string | undefined;
}

interface OAuthTokenRecord {
  connectorId: string;
  connectorCatalogId: string;
  setupProfileId?: string | undefined;
  accessToken: string;
  refreshToken?: string | undefined;
  tokenType?: string | undefined;
  scope?: string | undefined;
  expiresAt?: string | undefined;
  tokenEndpoint: string;
  clientId: string;
  clientSecret?: string | undefined;
  updatedAt: string;
}

interface OAuthTokenStore {
  version: 1;
  tokens: OAuthTokenRecord[];
}

function trimString(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed ? trimmed : undefined;
}

function base64Url(buffer: Buffer): string {
  return buffer.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function randomUrlToken(bytes = 32): string {
  return base64Url(randomBytes(bytes));
}

function pkceChallenge(verifier: string): string {
  return base64Url(createHash("sha256").update(verifier).digest());
}

function metadataUrlForEndpoint(endpoint: string | undefined): string | undefined {
  const normalized = trimString(endpoint);
  if (!normalized) return undefined;
  try {
    const parsed = new URL(normalized);
    if (parsed.hostname === "search.parallel.ai" && parsed.pathname.startsWith("/mcp-oauth")) {
      return "https://platform.parallel.ai/.well-known/oauth-authorization-server";
    }
    return `${parsed.origin}/.well-known/oauth-authorization-server`;
  } catch {
    return undefined;
  }
}

function addSeconds(iso: string, seconds: number): string {
  return new Date(Date.parse(iso) + seconds * 1000).toISOString();
}

function resolveExpiresAt(now: string, expiresIn: number | undefined): string | undefined {
  if (!Number.isFinite(expiresIn)) return undefined;
  const seconds = Math.max(60, Math.min(60 * 60 * 24 * 30, Math.floor(expiresIn as number)));
  return addSeconds(now, seconds);
}

function htmlPage(title: string, message: string): string {
  const safeTitle = title.replace(/[<>&"]/g, "");
  const safeMessage = message.replace(/[<>&"]/g, "");
  return [
    "<!doctype html>",
    "<html>",
    "<head><meta charset=\"utf-8\"><title>Pi-Office connector sign-in</title></head>",
    "<body style=\"font-family:system-ui,-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;margin:40px;color:#211d19;line-height:1.5;\">",
    `<h1 style=\"font-size:22px;margin:0 0 10px;\">${safeTitle}</h1>`,
    `<p style=\"font-size:15px;margin:0;color:#5f5851;max-width:620px;\">${safeMessage}</p>`,
    "</body>",
    "</html>",
  ].join("");
}

export class CompanionOAuthBroker {
  private readonly storagePath: string;
  private readonly flows = new Map<string, OAuthFlowRecord>();
  private tokens: OAuthTokenRecord[] = [];

  constructor(private readonly config: CompanionConfig) {
    this.storagePath = join(config.dataDir, "connector-oauth-tokens.json");
    this.load();
  }

  async start(definition: CompanionConnectorDefinition): Promise<ConnectorOAuthStartResponse> {
    if (definition.authMethod !== "oauth" || definition.oauth?.broker !== "companion") {
      throw new Error("This connector is not configured for companion-brokered OAuth.");
    }
    const endpoint = trimString(definition.url);
    if (!endpoint) {
      throw new Error("Connector URL is required before OAuth sign-in can start.");
    }

    const metadataUrl = trimString(definition.oauth.metadataUrl) ?? metadataUrlForEndpoint(endpoint);
    const metadata = metadataUrl ? await this.fetchMetadata(metadataUrl) : {};
    const authorizationEndpoint = trimString(definition.oauth.authorizationUrl) ?? trimString(metadata.authorization_endpoint);
    const tokenEndpoint = trimString(definition.oauth.tokenUrl) ?? trimString(metadata.token_endpoint);
    const registrationEndpoint = trimString(definition.oauth.registrationUrl) ?? trimString(metadata.registration_endpoint);
    if (!authorizationEndpoint || !tokenEndpoint) {
      throw new Error("OAuth metadata did not include authorization and token endpoints.");
    }
    if (!registrationEndpoint) {
      throw new Error("OAuth metadata did not provide Dynamic Client Registration for Pi-Office.");
    }

    const redirectPath = trimString(definition.oauth.redirectPath) ?? "/v1/connectors/oauth/callback";
    const redirectUri = `${this.config.endpoint}${redirectPath.startsWith("/") ? redirectPath : `/${redirectPath}`}`;
    const registration = await this.registerClient(registrationEndpoint, definition.oauth.clientName ?? "Pi-Office", redirectUri);
    const clientId = trimString(registration.client_id);
    if (!clientId) {
      throw new Error("OAuth dynamic client registration did not return a client id.");
    }

    const now = new Date().toISOString();
    const state = randomUrlToken();
    const codeVerifier = randomUrlToken(64);
    const scope = (definition.oauth.scopes ?? []).map(trimString).filter(Boolean).join(" ") || undefined;
    const expiresAt = addSeconds(now, 10 * 60);
    const authUrl = new URL(authorizationEndpoint);
    authUrl.searchParams.set("response_type", "code");
    authUrl.searchParams.set("client_id", clientId);
    authUrl.searchParams.set("redirect_uri", redirectUri);
    authUrl.searchParams.set("state", state);
    authUrl.searchParams.set("code_challenge", pkceChallenge(codeVerifier));
    authUrl.searchParams.set("code_challenge_method", "S256");
    if (scope) authUrl.searchParams.set("scope", scope);

    this.flows.set(state, {
      connectorId: definition.id,
      connectorCatalogId: definition.connectorId,
      setupProfileId: definition.setupProfileId,
      state,
      authorizationUrl: authUrl.toString(),
      tokenEndpoint,
      redirectUri,
      clientId,
      clientSecret: trimString(registration.client_secret),
      codeVerifier,
      scope,
      createdAt: now,
      expiresAt,
    });

    return {
      ok: true,
      connectorId: definition.id,
      url: authUrl.toString(),
      callbackUrl: redirectUri,
      state,
      expiresAt,
      broker: "companion",
      openMode: definition.oauth.launchMode ?? "system_browser",
    };
  }

  async completeCallback(query: Record<string, unknown>): Promise<{ statusCode: number; html: string }> {
    const state = trimString(query.state);
    if (!state) {
      return {
        statusCode: 400,
        html: htmlPage("Connector sign-in failed", "The provider callback did not include OAuth state. Return to Pi-Office and start sign-in again."),
      };
    }

    const flow = this.flows.get(state);
    if (!flow) {
      return {
        statusCode: 400,
        html: htmlPage("Connector sign-in failed", "This sign-in request was not found or has already expired. Return to Pi-Office and start sign-in again."),
      };
    }
    if (flow.completedAt) {
      return {
        statusCode: 200,
        html: htmlPage("Connector sign-in complete", "You can close this browser window and return to Pi-Office."),
      };
    }
    if (Date.parse(flow.expiresAt) <= Date.now()) {
      flow.error = "OAuth callback state expired. Start sign-in again.";
      return {
        statusCode: 400,
        html: htmlPage("Connector sign-in expired", "Return to Pi-Office and start sign-in again."),
      };
    }

    const callbackError = trimString(query.error_description) ?? trimString(query.error);
    if (callbackError) {
      flow.error = callbackError;
      return {
        statusCode: 400,
        html: htmlPage("Connector sign-in was cancelled", "Return to Pi-Office when you are ready to try again."),
      };
    }

    const code = trimString(query.code);
    if (!code) {
      flow.error = "OAuth callback did not include an authorization code.";
      return {
        statusCode: 400,
        html: htmlPage("Connector sign-in failed", "The provider did not return an authorization code. Return to Pi-Office and try again."),
      };
    }

    try {
      const token = await this.exchangeCode(flow, code);
      this.saveToken(flow, token);
      flow.completedAt = new Date().toISOString();
      flow.error = undefined;
      return {
        statusCode: 200,
        html: htmlPage("Connector sign-in complete", "You can close this browser window and return to Pi-Office. Then run Check connection to load the connector tools."),
      };
    } catch (error) {
      flow.error = error instanceof Error ? error.message : String(error);
      return {
        statusCode: 400,
        html: htmlPage("Connector sign-in failed", "Pi-Office could not finish the token exchange. Return to Pi-Office and try again."),
      };
    }
  }

  status(request: CompanionConnectorOAuthStatusRequest): CompanionConnectorOAuthStatusResponse {
    const state = trimString(request.state);
    const connectorId = trimString(request.connectorId);
    if (!state && !connectorId) {
      return {
        ok: true,
        connected: false,
        pending: false,
        error: "connectorId or state is required.",
      };
    }
    const flow = state ? this.flows.get(state) : undefined;
    const token = this.tokens.find((entry) =>
      (!connectorId || entry.connectorId === connectorId) &&
      (!flow || entry.connectorId === flow.connectorId)
    );
    return {
      ok: true,
      connectorId: flow?.connectorId ?? connectorId,
      state,
      connected: Boolean(token?.accessToken),
      pending: Boolean(flow && !flow.completedAt && !flow.error && Date.parse(flow.expiresAt) > Date.now()),
      expiresAt: flow?.expiresAt ?? token?.expiresAt,
      error: flow?.error,
    };
  }

  async getAccessToken(definition: CompanionConnectorDefinition): Promise<string | undefined> {
    const token = this.tokens.find((entry) =>
      entry.connectorId === definition.id &&
      entry.connectorCatalogId === definition.connectorId &&
      (!definition.setupProfileId || !entry.setupProfileId || entry.setupProfileId === definition.setupProfileId)
    );
    if (!token) return undefined;
    if (!token.expiresAt || Date.parse(token.expiresAt) > Date.now() + 60_000) {
      return token.accessToken;
    }
    if (!token.refreshToken) return undefined;
    try {
      const refreshed = await this.refreshToken(token);
      const updated = this.mergeToken(token, refreshed);
      this.tokens = this.tokens.map((entry) => entry === token ? updated : entry);
      this.persist();
      return updated.accessToken;
    } catch {
      return undefined;
    }
  }

  private async fetchMetadata(metadataUrl: string): Promise<OAuthServerMetadata> {
    const response = await fetch(metadataUrl, { headers: { accept: "application/json" } });
    if (!response.ok) {
      throw new Error(`OAuth metadata discovery failed with HTTP ${response.status}.`);
    }
    return await response.json() as OAuthServerMetadata;
  }

  private async registerClient(registrationEndpoint: string, clientName: string, redirectUri: string): Promise<{ client_id?: string; client_secret?: string }> {
    const response = await fetch(registrationEndpoint, {
      method: "POST",
      headers: {
        "content-type": "application/json",
        accept: "application/json",
      },
      body: JSON.stringify({
        client_name: clientName,
        redirect_uris: [redirectUri],
        grant_types: ["authorization_code", "refresh_token"],
        response_types: ["code"],
        token_endpoint_auth_method: "none",
      }),
    });
    if (!response.ok) {
      const body = await response.text().catch(() => "");
      throw new Error(body.trim() || `OAuth dynamic client registration failed with HTTP ${response.status}.`);
    }
    return await response.json() as { client_id?: string; client_secret?: string };
  }

  private async exchangeCode(flow: OAuthFlowRecord, code: string): Promise<OAuthTokenResponse> {
    const form = new URLSearchParams();
    form.set("grant_type", "authorization_code");
    form.set("code", code);
    form.set("redirect_uri", flow.redirectUri);
    form.set("client_id", flow.clientId);
    form.set("code_verifier", flow.codeVerifier);
    if (flow.clientSecret) form.set("client_secret", flow.clientSecret);
    return this.postToken(flow.tokenEndpoint, form);
  }

  private async refreshToken(token: OAuthTokenRecord): Promise<OAuthTokenResponse> {
    const form = new URLSearchParams();
    form.set("grant_type", "refresh_token");
    form.set("refresh_token", token.refreshToken!);
    form.set("client_id", token.clientId);
    if (token.clientSecret) form.set("client_secret", token.clientSecret);
    return this.postToken(token.tokenEndpoint, form);
  }

  private async postToken(tokenEndpoint: string, form: URLSearchParams): Promise<OAuthTokenResponse> {
    const response = await fetch(tokenEndpoint, {
      method: "POST",
      headers: {
        "content-type": "application/x-www-form-urlencoded",
        accept: "application/json",
      },
      body: form.toString(),
    });
    if (!response.ok) {
      const body = await response.text().catch(() => "");
      throw new Error(body.trim() || `OAuth token exchange failed with HTTP ${response.status}.`);
    }
    return await response.json() as OAuthTokenResponse;
  }

  private saveToken(flow: OAuthFlowRecord, token: OAuthTokenResponse): void {
    const accessToken = trimString(token.access_token);
    if (!accessToken) {
      throw new Error("OAuth token exchange did not return an access token.");
    }
    const record = this.mergeToken({
      connectorId: flow.connectorId,
      connectorCatalogId: flow.connectorCatalogId,
      setupProfileId: flow.setupProfileId,
      accessToken,
      refreshToken: trimString(token.refresh_token),
      tokenType: trimString(token.token_type),
      scope: trimString(token.scope) ?? flow.scope,
      tokenEndpoint: flow.tokenEndpoint,
      clientId: flow.clientId,
      clientSecret: flow.clientSecret,
      updatedAt: new Date().toISOString(),
    }, token);
    this.tokens = [
      ...this.tokens.filter((entry) => entry.connectorId !== flow.connectorId),
      record,
    ];
    this.persist();
  }

  private mergeToken(base: OAuthTokenRecord, token: OAuthTokenResponse): OAuthTokenRecord {
    const now = new Date().toISOString();
    return {
      ...base,
      accessToken: trimString(token.access_token) ?? base.accessToken,
      refreshToken: trimString(token.refresh_token) ?? base.refreshToken,
      tokenType: trimString(token.token_type) ?? base.tokenType,
      scope: trimString(token.scope) ?? base.scope,
      expiresAt: resolveExpiresAt(now, token.expires_in) ?? base.expiresAt,
      updatedAt: now,
    };
  }

  private load(): void {
    try {
      if (!existsSync(this.storagePath)) return;
      const parsed = JSON.parse(readFileSync(this.storagePath, "utf8")) as Partial<OAuthTokenStore>;
      this.tokens = Array.isArray(parsed.tokens)
        ? parsed.tokens.filter((entry): entry is OAuthTokenRecord => Boolean(entry?.connectorId && entry.accessToken))
        : [];
    } catch {
      this.tokens = [];
    }
  }

  private persist(): void {
    mkdirSync(this.config.dataDir, { recursive: true });
    const body: OAuthTokenStore = { version: 1, tokens: this.tokens };
    writeFileSync(this.storagePath, `${JSON.stringify(body, null, 2)}\n`, { mode: 0o600 });
  }
}
