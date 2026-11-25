/**
 * OAuth 2.1 authorization server for the ChatGPT MCP connector.
 *
 * Pure logic only: storage and the portal-user check are injected, so this file
 * has no database import and can be unit-tested directly.
 *
 * Design:
 *  - Public clients (ChatGPT) register dynamically and use PKCE (S256 only).
 *  - The person approving must log in with a Production Portal admin account.
 *  - Access tokens live 1 hour; refresh tokens rotate on every use and slide
 *    forward 365 days, so the connection stays connected while in use.
 *  - Only SHA-256 hashes of codes and tokens are stored.
 *  - Re-using an already-rotated refresh token revokes the whole connection
 *    (a stolen-token signal).
 */
import { createHash, randomBytes, randomUUID, timingSafeEqual } from "node:crypto";

export const SCOPES = ["portal:read", "portal:write"] as const;
export const DEFAULT_SCOPE = SCOPES.join(" ");
export const ACCESS_TOKEN_TTL_MS = 60 * 60 * 1000;
export const REFRESH_TOKEN_TTL_MS = 365 * 24 * 60 * 60 * 1000;
export const AUTH_CODE_TTL_MS = 5 * 60 * 1000;
const MAX_REDIRECT_URIS = 5;

export type OAuthClient = { id: string; name: string; redirectUris: string[] };
export type OAuthCode = {
  codeHash: string;
  clientId: string;
  portalUserId: string;
  redirectUri: string;
  codeChallenge: string;
  scope: string;
  resource: string | null;
  expiresAt: Date;
  usedAt: Date | null;
};
export type OAuthToken = {
  id: number;
  grantId: string;
  clientId: string;
  portalUserId: string;
  scope: string;
  accessTokenHash: string;
  refreshTokenHash: string;
  accessExpiresAt: Date;
  refreshExpiresAt: Date;
  rotatedAt: Date | null;
  revokedAt: Date | null;
};

export interface OAuthStore {
  createClient(client: OAuthClient): Promise<void>;
  getClient(id: string): Promise<OAuthClient | null>;
  touchClient(id: string): Promise<void>;
  saveCode(code: OAuthCode): Promise<void>;
  /** Atomically mark an unused code as used. Returns the code only if this call claimed it. */
  claimCode(codeHash: string, now: Date): Promise<OAuthCode | null>;
  createToken(token: Omit<OAuthToken, "id" | "rotatedAt" | "revokedAt">): Promise<void>;
  findTokenByAccessHash(hash: string): Promise<OAuthToken | null>;
  findTokenByRefreshHash(hash: string): Promise<OAuthToken | null>;
  /** Atomically mark a token rotated. Returns false if it was already rotated/revoked. */
  markRotated(id: number, now: Date): Promise<boolean>;
  revokeGrant(grantId: string, now: Date): Promise<void>;
  touchToken(id: number, now: Date): Promise<void>;
}

/** Who may approve a connection, and who a token still acts for. */
export type PortalApprover = { id: string; name: string; isAdmin: boolean; active: boolean };
export interface PortalUserDirectory {
  verifyLogin(username: string, password: string): Promise<PortalApprover | null>;
  getUser(id: string): Promise<PortalApprover | null>;
}

export class OAuthError extends Error {
  code: string;
  description: string;
  status: number;
  constructor(code: string, description: string, status = 400) {
    super(description);
    this.code = code;
    this.description = description;
    this.status = status;
  }
  toJSON() {
    return { error: this.code, error_description: this.description };
  }
}

export const sha256 = (value: string) => createHash("sha256").update(value).digest("hex");
const randomToken = (prefix: string) => `${prefix}_${randomBytes(32).toString("base64url")}`;

export function safeEqual(a: string, b: string) {
  const x = Buffer.from(a);
  const y = Buffer.from(b);
  return x.length === y.length && timingSafeEqual(x, y);
}

export function pkceChallenge(verifier: string) {
  return createHash("sha256").update(verifier).digest("base64url");
}

/** Hosts ChatGPT may redirect back to. Overridable with MCP_ALLOWED_REDIRECT_HOSTS. */
export function allowedRedirectHosts(env: Record<string, string | undefined> = process.env): string[] {
  const raw = env.MCP_ALLOWED_REDIRECT_HOSTS ?? "chatgpt.com,chat.openai.com";
  return raw.split(",").map((h) => h.trim().toLowerCase()).filter(Boolean);
}

export function validateRedirectUri(uri: string, hosts: string[]): boolean {
  let url: URL;
  try {
    url = new URL(uri);
  } catch {
    return false;
  }
  if (url.protocol !== "https:") return false;
  if (url.hash) return false;
  return hosts.includes(url.hostname.toLowerCase());
}

export function normalizeScope(requested: string | null | undefined): string {
  const asked = (requested ?? "").split(/\s+/).filter(Boolean);
  const granted = asked.length ? asked.filter((s) => (SCOPES as readonly string[]).includes(s)) : [...SCOPES];
  if (!granted.length) throw new OAuthError("invalid_scope", "No supported scope requested");
  // Write implies read.
  if (granted.includes("portal:write") && !granted.includes("portal:read")) granted.unshift("portal:read");
  return Array.from(new Set(granted)).join(" ");
}

// ─── Dynamic client registration (RFC 7591) ─────────────────────────────────

export async function registerClient(
  store: OAuthStore,
  body: unknown,
  hosts: string[],
): Promise<Record<string, unknown>> {
  const meta = (body && typeof body === "object" ? body : {}) as Record<string, unknown>;
  const uris = Array.isArray(meta.redirect_uris) ? meta.redirect_uris.map(String) : [];
  if (!uris.length || uris.length > MAX_REDIRECT_URIS) {
    throw new OAuthError("invalid_redirect_uri", "Provide between 1 and 5 redirect_uris");
  }
  for (const uri of uris) {
    if (!validateRedirectUri(uri, hosts)) {
      throw new OAuthError("invalid_redirect_uri", `Redirect URI not allowed: ${uri}`);
    }
  }
  const authMethod = String(meta.token_endpoint_auth_method ?? "none");
  if (authMethod !== "none") {
    throw new OAuthError("invalid_client_metadata", "Only public clients (token_endpoint_auth_method=none) are supported");
  }
  const name = String(meta.client_name ?? "MCP client").slice(0, 100);
  const id = `mcp_${randomUUID()}`;
  await store.createClient({ id, name, redirectUris: uris });
  return {
    client_id: id,
    client_name: name,
    redirect_uris: uris,
    token_endpoint_auth_method: "none",
    grant_types: ["authorization_code", "refresh_token"],
    response_types: ["code"],
    client_id_issued_at: Math.floor(Date.now() / 1000),
  };
}

// ─── Authorization request ──────────────────────────────────────────────────

export type AuthorizeParams = {
  clientId: string;
  redirectUri: string;
  codeChallenge: string;
  state: string | null;
  scope: string;
  resource: string | null;
};

/**
 * Validate an /authorize request. Errors about the client or redirect URI must
 * be shown to the user, never redirected (open-redirect protection); `redirectable`
 * says whether an error may be sent back to the client instead.
 */
export async function validateAuthorizeRequest(
  store: OAuthStore,
  q: URLSearchParams,
): Promise<{ client: OAuthClient; params: AuthorizeParams }> {
  const clientId = q.get("client_id") ?? "";
  const client = clientId ? await store.getClient(clientId) : null;
  if (!client) throw new OAuthError("invalid_client", "Unknown client");
  const redirectUri = q.get("redirect_uri") ?? (client.redirectUris.length === 1 ? client.redirectUris[0] : "");
  if (!client.redirectUris.includes(redirectUri)) throw new OAuthError("invalid_request", "redirect_uri does not match the registered value");
  if (q.get("response_type") !== "code") throw new OAuthError("unsupported_response_type", "Only response_type=code is supported");
  const codeChallenge = q.get("code_challenge") ?? "";
  if (!/^[A-Za-z0-9_-]{43,128}$/.test(codeChallenge)) throw new OAuthError("invalid_request", "A PKCE code_challenge is required");
  if ((q.get("code_challenge_method") ?? "") !== "S256") throw new OAuthError("invalid_request", "code_challenge_method must be S256");
  const scope = normalizeScope(q.get("scope"));
  return {
    client,
    params: { clientId, redirectUri, codeChallenge, state: q.get("state"), scope, resource: q.get("resource") },
  };
}

/** Called after the portal admin logged in and approved. Returns the redirect URL. */
export async function issueAuthorizationCode(
  store: OAuthStore,
  params: AuthorizeParams,
  approver: PortalApprover,
  now = new Date(),
): Promise<string> {
  if (!approver.active || !approver.isAdmin) {
    throw new OAuthError("access_denied", "Only an active Production Portal admin can connect ChatGPT", 403);
  }
  const code = randomToken("kec");
  await store.saveCode({
    codeHash: sha256(code),
    clientId: params.clientId,
    portalUserId: approver.id,
    redirectUri: params.redirectUri,
    codeChallenge: params.codeChallenge,
    scope: params.scope,
    resource: params.resource,
    expiresAt: new Date(now.getTime() + AUTH_CODE_TTL_MS),
    usedAt: null,
  });
  const url = new URL(params.redirectUri);
  url.searchParams.set("code", code);
  if (params.state) url.searchParams.set("state", params.state);
  return url.toString();
}

// ─── Token endpoint ─────────────────────────────────────────────────────────

export type TokenResponse = {
  access_token: string;
  token_type: "Bearer";
  expires_in: number;
  refresh_token: string;
  scope: string;
};

async function mintTokens(
  store: OAuthStore,
  grant: { grantId: string; clientId: string; portalUserId: string; scope: string },
  now: Date,
): Promise<TokenResponse> {
  const access = randomToken("kea");
  const refresh = randomToken("ker");
  await store.createToken({
    ...grant,
    accessTokenHash: sha256(access),
    refreshTokenHash: sha256(refresh),
    accessExpiresAt: new Date(now.getTime() + ACCESS_TOKEN_TTL_MS),
    refreshExpiresAt: new Date(now.getTime() + REFRESH_TOKEN_TTL_MS),
  });
  return {
    access_token: access,
    token_type: "Bearer",
    expires_in: Math.floor(ACCESS_TOKEN_TTL_MS / 1000),
    refresh_token: refresh,
    scope: grant.scope,
  };
}

export async function exchangeToken(
  store: OAuthStore,
  users: PortalUserDirectory,
  form: URLSearchParams,
  now = new Date(),
): Promise<TokenResponse> {
  const grantType = form.get("grant_type");
  const clientId = form.get("client_id") ?? "";
  const client = clientId ? await store.getClient(clientId) : null;
  if (!client) throw new OAuthError("invalid_client", "Unknown client", 401);

  if (grantType === "authorization_code") {
    const code = form.get("code") ?? "";
    const verifier = form.get("code_verifier") ?? "";
    if (!code || !verifier) throw new OAuthError("invalid_request", "code and code_verifier are required");
    const claimed = await store.claimCode(sha256(code), now);
    if (!claimed) throw new OAuthError("invalid_grant", "Code is invalid, expired or already used");
    if (claimed.clientId !== client.id) throw new OAuthError("invalid_grant", "Code was issued to another client");
    if (claimed.expiresAt.getTime() < now.getTime()) throw new OAuthError("invalid_grant", "Code expired");
    const redirectUri = form.get("redirect_uri");
    if (redirectUri && redirectUri !== claimed.redirectUri) throw new OAuthError("invalid_grant", "redirect_uri mismatch");
    if (!safeEqual(pkceChallenge(verifier), claimed.codeChallenge)) throw new OAuthError("invalid_grant", "PKCE verification failed");
    const user = await users.getUser(claimed.portalUserId);
    if (!user?.active || !user.isAdmin) throw new OAuthError("invalid_grant", "Approving user is no longer an active admin");
    await store.touchClient(client.id);
    return mintTokens(store, { grantId: randomUUID(), clientId: client.id, portalUserId: claimed.portalUserId, scope: claimed.scope }, now);
  }

  if (grantType === "refresh_token") {
    const refresh = form.get("refresh_token") ?? "";
    const existing = refresh ? await store.findTokenByRefreshHash(sha256(refresh)) : null;
    if (!existing || existing.clientId !== client.id) throw new OAuthError("invalid_grant", "Unknown refresh token");
    if (existing.revokedAt) throw new OAuthError("invalid_grant", "Connection was revoked");
    if (existing.rotatedAt) {
      // A refresh token that was already exchanged is being replayed: someone
      // else may hold a copy. Kill the whole connection.
      await store.revokeGrant(existing.grantId, now);
      throw new OAuthError("invalid_grant", "Refresh token reuse detected; connection revoked");
    }
    if (existing.refreshExpiresAt.getTime() < now.getTime()) throw new OAuthError("invalid_grant", "Refresh token expired");
    const user = await users.getUser(existing.portalUserId);
    if (!user?.active || !user.isAdmin) {
      await store.revokeGrant(existing.grantId, now);
      throw new OAuthError("invalid_grant", "Approving user is no longer an active admin");
    }
    if (!(await store.markRotated(existing.id, now))) {
      await store.revokeGrant(existing.grantId, now);
      throw new OAuthError("invalid_grant", "Refresh token reuse detected; connection revoked");
    }
    const scope = form.get("scope") ? normalizeScope(form.get("scope")) : existing.scope;
    const narrowed = scope.split(" ").every((s) => existing.scope.split(" ").includes(s)) ? scope : existing.scope;
    return mintTokens(store, { grantId: existing.grantId, clientId: client.id, portalUserId: existing.portalUserId, scope: narrowed }, now);
  }

  throw new OAuthError("unsupported_grant_type", "Use authorization_code or refresh_token");
}

export async function revokeToken(store: OAuthStore, token: string, now = new Date()): Promise<void> {
  if (!token) return;
  const hash = sha256(token);
  const found = (await store.findTokenByAccessHash(hash)) ?? (await store.findTokenByRefreshHash(hash));
  if (found) await store.revokeGrant(found.grantId, now);
}

// ─── Protecting /mcp ────────────────────────────────────────────────────────

export type McpPrincipal = {
  grantId: string;
  clientId: string;
  portalUserId: string;
  portalUserName: string;
  scopes: string[];
};

export async function authenticateBearer(
  store: OAuthStore,
  users: PortalUserDirectory,
  authorization: string | null,
  now = new Date(),
): Promise<McpPrincipal> {
  const token = authorization?.startsWith("Bearer ") ? authorization.slice(7).trim() : "";
  if (!token) throw new OAuthError("invalid_token", "Missing bearer token", 401);
  const found = await store.findTokenByAccessHash(sha256(token));
  if (!found || found.revokedAt) throw new OAuthError("invalid_token", "Unknown or revoked token", 401);
  if (found.accessExpiresAt.getTime() < now.getTime()) throw new OAuthError("invalid_token", "Token expired", 401);
  const user = await users.getUser(found.portalUserId);
  if (!user?.active || !user.isAdmin) throw new OAuthError("invalid_token", "Approving user is no longer an active admin", 401);
  await store.touchToken(found.id, now);
  return {
    grantId: found.grantId,
    clientId: found.clientId,
    portalUserId: user.id,
    portalUserName: user.name,
    scopes: found.scope.split(" "),
  };
}

// ─── Discovery metadata ─────────────────────────────────────────────────────

export function authorizationServerMetadata(base: string) {
  return {
    issuer: base,
    authorization_endpoint: `${base}/oauth/authorize`,
    token_endpoint: `${base}/oauth/token`,
    registration_endpoint: `${base}/oauth/register`,
    revocation_endpoint: `${base}/oauth/revoke`,
    response_types_supported: ["code"],
    grant_types_supported: ["authorization_code", "refresh_token"],
    code_challenge_methods_supported: ["S256"],
    token_endpoint_auth_methods_supported: ["none"],
    scopes_supported: [...SCOPES],
  };
}

export function protectedResourceMetadata(base: string) {
  return {
    resource: `${base}/mcp`,
    authorization_servers: [base],
    scopes_supported: [...SCOPES],
    bearer_methods_supported: ["header"],
    resource_name: "Karma East Production Portal",
  };
}
