// node --experimental-strip-types --test app/mcp/mcp.test.ts
import { test, describe } from "node:test";
import assert from "node:assert/strict";
import { createHash, randomBytes } from "node:crypto";
import {
  authenticateBearer,
  exchangeToken,
  issueAuthorizationCode,
  OAuthError,
  registerClient,
  revokeToken,
  validateAuthorizeRequest,
  validateRedirectUri,
  type OAuthClient,
  type OAuthCode,
  type OAuthStore,
  type OAuthToken,
  type PortalApprover,
  type PortalUserDirectory,
} from "./oauth-core.ts";
import { handleJsonRpc } from "./protocol.ts";
import { appendNote, compactForModel, findPotentialDuplicates, matchSize, matchStaffRecipients, resolveNewOrderQuantities, resolveQuantityChanges, stableStringify } from "./mcp-logic.ts";

// ─── In-memory OAuth store ──────────────────────────────────────────────────

function memoryStore() {
  const clients = new Map<string, OAuthClient>();
  const codes = new Map<string, OAuthCode>();
  const tokens: OAuthToken[] = [];
  let nextId = 1;
  const store: OAuthStore = {
    async createClient(c) { clients.set(c.id, c); },
    async getClient(id) { return clients.get(id) ?? null; },
    async touchClient() {},
    async saveCode(c) { codes.set(c.codeHash, { ...c }); },
    async claimCode(hash, now) {
      const c = codes.get(hash);
      if (!c || c.usedAt) return null;
      c.usedAt = now;
      return { ...c };
    },
    async createToken(t) { tokens.push({ ...t, id: nextId++, rotatedAt: null, revokedAt: null }); },
    async findTokenByAccessHash(h) { return tokens.find((t) => t.accessTokenHash === h) ?? null; },
    async findTokenByRefreshHash(h) { return tokens.find((t) => t.refreshTokenHash === h) ?? null; },
    async markRotated(id, now) {
      const t = tokens.find((x) => x.id === id);
      if (!t || t.rotatedAt || t.revokedAt) return false;
      t.rotatedAt = now;
      return true;
    },
    async revokeGrant(grantId, now) { for (const t of tokens) if (t.grantId === grantId && !t.revokedAt) t.revokedAt = now; },
    async touchToken() {},
  };
  return { store, tokens, codes };
}

const admin: PortalApprover = { id: "u-koku", name: "Koku", isAdmin: true, active: true };
const staff: PortalApprover = { id: "u-staff", name: "Staff", isAdmin: false, active: true };
function directory(users: PortalApprover[] = [admin, staff]): PortalUserDirectory & { users: PortalApprover[] } {
  return {
    users,
    async verifyLogin(name) { return users.find((u) => u.name.toLowerCase() === name.toLowerCase()) ?? null; },
    async getUser(id) { return users.find((u) => u.id === id) ?? null; },
  };
}

const HOSTS = ["chatgpt.com"];
const REDIRECT = "https://chatgpt.com/connector_platform_oauth_redirect";

function pkce() {
  const verifier = randomBytes(32).toString("base64url");
  const challenge = createHash("sha256").update(verifier).digest("base64url");
  return { verifier, challenge };
}

async function connect(store: OAuthStore, users: PortalUserDirectory, approver = admin) {
  const reg = await registerClient(store, { redirect_uris: [REDIRECT], client_name: "ChatGPT", token_endpoint_auth_method: "none" }, HOSTS);
  const clientId = String(reg.client_id);
  const { verifier, challenge } = pkce();
  const q = new URLSearchParams({ response_type: "code", client_id: clientId, redirect_uri: REDIRECT, code_challenge: challenge, code_challenge_method: "S256", state: "xyz", scope: "portal:read portal:write" });
  const { params } = await validateAuthorizeRequest(store, q);
  const location = new URL(await issueAuthorizationCode(store, params, approver));
  const code = location.searchParams.get("code")!;
  assert.equal(location.searchParams.get("state"), "xyz");
  const tokens = await exchangeToken(store, users, new URLSearchParams({ grant_type: "authorization_code", client_id: clientId, code, code_verifier: verifier, redirect_uri: REDIRECT }));
  return { clientId, code, verifier, tokens };
}

describe("OAuth", () => {
  test("full connect flow issues working tokens", async () => {
    const { store } = memoryStore();
    const users = directory();
    const { tokens } = await connect(store, users);
    assert.match(tokens.access_token, /^kea_/);
    assert.equal(tokens.scope, "portal:read portal:write");
    const p = await authenticateBearer(store, users, `Bearer ${tokens.access_token}`);
    assert.equal(p.portalUserName, "Koku");
    assert.deepEqual(p.scopes, ["portal:read", "portal:write"]);
  });

  test("only SHA-256 hashes of tokens are stored", async () => {
    const { store, tokens: stored } = memoryStore();
    const { tokens } = await connect(store, directory());
    const raw = JSON.stringify(stored);
    assert.equal(raw.includes(tokens.access_token), false);
    assert.equal(raw.includes(tokens.refresh_token), false);
  });

  test("redirect URIs must be https on an allowed host", () => {
    assert.equal(validateRedirectUri(REDIRECT, HOSTS), true);
    assert.equal(validateRedirectUri("http://chatgpt.com/cb", HOSTS), false);
    assert.equal(validateRedirectUri("https://evil.example/cb", HOSTS), false);
    assert.equal(validateRedirectUri("https://chatgpt.com.evil.example/cb", HOSTS), false);
  });

  test("registration rejects foreign redirect hosts", async () => {
    const { store } = memoryStore();
    await assert.rejects(registerClient(store, { redirect_uris: ["https://evil.example/cb"] }, HOSTS), OAuthError);
  });

  test("authorize rejects a redirect_uri that wasn't registered", async () => {
    const { store } = memoryStore();
    const reg = await registerClient(store, { redirect_uris: [REDIRECT] }, HOSTS);
    const { challenge } = pkce();
    const q = new URLSearchParams({ response_type: "code", client_id: String(reg.client_id), redirect_uri: "https://chatgpt.com/other", code_challenge: challenge, code_challenge_method: "S256" });
    await assert.rejects(validateAuthorizeRequest(store, q), /redirect_uri/);
  });

  test("PKCE is mandatory and must be S256", async () => {
    const { store } = memoryStore();
    const reg = await registerClient(store, { redirect_uris: [REDIRECT] }, HOSTS);
    const base = { response_type: "code", client_id: String(reg.client_id), redirect_uri: REDIRECT };
    await assert.rejects(validateAuthorizeRequest(store, new URLSearchParams(base)), /code_challenge/);
    await assert.rejects(validateAuthorizeRequest(store, new URLSearchParams({ ...base, code_challenge: pkce().challenge, code_challenge_method: "plain" })), /S256/);
  });

  test("a non-admin portal user cannot approve", async () => {
    const { store } = memoryStore();
    await assert.rejects(connect(store, directory(), staff), /admin/);
  });

  test("wrong code_verifier fails", async () => {
    const { store } = memoryStore();
    const users = directory();
    const reg = await registerClient(store, { redirect_uris: [REDIRECT] }, HOSTS);
    const { challenge } = pkce();
    const { params } = await validateAuthorizeRequest(store, new URLSearchParams({ response_type: "code", client_id: String(reg.client_id), redirect_uri: REDIRECT, code_challenge: challenge, code_challenge_method: "S256" }));
    const code = new URL(await issueAuthorizationCode(store, params, admin)).searchParams.get("code")!;
    await assert.rejects(exchangeToken(store, users, new URLSearchParams({ grant_type: "authorization_code", client_id: String(reg.client_id), code, code_verifier: pkce().verifier })), /PKCE/);
  });

  test("an authorization code works once only", async () => {
    const { store } = memoryStore();
    const users = directory();
    const { clientId, code, verifier } = await connect(store, users);
    await assert.rejects(exchangeToken(store, users, new URLSearchParams({ grant_type: "authorization_code", client_id: clientId, code, code_verifier: verifier })), /already used/);
  });

  test("refresh rotates tokens and keeps the connection", async () => {
    const { store } = memoryStore();
    const users = directory();
    const { clientId, tokens } = await connect(store, users);
    const next = await exchangeToken(store, users, new URLSearchParams({ grant_type: "refresh_token", client_id: clientId, refresh_token: tokens.refresh_token }));
    assert.notEqual(next.refresh_token, tokens.refresh_token);
    const p = await authenticateBearer(store, users, `Bearer ${next.access_token}`);
    assert.equal(p.portalUserId, "u-koku");
  });

  test("re-using an old refresh token revokes the whole connection", async () => {
    const { store } = memoryStore();
    const users = directory();
    const { clientId, tokens } = await connect(store, users);
    const next = await exchangeToken(store, users, new URLSearchParams({ grant_type: "refresh_token", client_id: clientId, refresh_token: tokens.refresh_token }));
    await assert.rejects(exchangeToken(store, users, new URLSearchParams({ grant_type: "refresh_token", client_id: clientId, refresh_token: tokens.refresh_token })), /reuse/);
    await assert.rejects(authenticateBearer(store, users, `Bearer ${next.access_token}`), /revoked/);
  });

  test("expired access token is rejected", async () => {
    const { store } = memoryStore();
    const users = directory();
    const { tokens } = await connect(store, users);
    await assert.rejects(authenticateBearer(store, users, `Bearer ${tokens.access_token}`, new Date(Date.now() + 2 * 3600_000)), /expired/);
  });

  test("deactivating or demoting the approving admin cuts access", async () => {
    const { store } = memoryStore();
    const users = directory([{ ...admin }]);
    const { tokens } = await connect(store, users);
    users.users[0].isAdmin = false;
    await assert.rejects(authenticateBearer(store, users, `Bearer ${tokens.access_token}`), /admin/);
  });

  test("revoke ends the connection", async () => {
    const { store } = memoryStore();
    const users = directory();
    const { tokens } = await connect(store, users);
    await revokeToken(store, tokens.refresh_token);
    await assert.rejects(authenticateBearer(store, users, `Bearer ${tokens.access_token}`), /revoked/);
  });

  test("missing or garbage bearer is 401", async () => {
    const { store } = memoryStore();
    await assert.rejects(authenticateBearer(store, directory(), null), (e: unknown) => e instanceof OAuthError && e.status === 401);
    await assert.rejects(authenticateBearer(store, directory(), "Bearer nope"), (e: unknown) => e instanceof OAuthError && e.status === 401);
  });
});

// ─── MCP protocol ───────────────────────────────────────────────────────────

describe("MCP JSON-RPC", () => {
  const tools = [{ name: "ping_tool", title: "t", description: "d", inputSchema: { type: "object" }, annotations: { readOnlyHint: true } }];
  const deps = { tools, call: async (_n: string, args: Record<string, unknown>) => ({ structured: { echo: args } }) };

  test("initialize negotiates a supported version and advertises tools", async () => {
    const r = await handleJsonRpc({ jsonrpc: "2.0", id: 1, method: "initialize", params: { protocolVersion: "2025-06-18" } }, deps);
    const result = r!.result as Record<string, any>;
    assert.equal(result.protocolVersion, "2025-06-18");
    assert.ok(result.capabilities.tools);
    assert.match(result.instructions, /confirm_change only after Koku/);
  });

  test("unknown protocol version falls back to the newest", async () => {
    const r = await handleJsonRpc({ jsonrpc: "2.0", id: 1, method: "initialize", params: { protocolVersion: "1999-01-01" } }, deps);
    assert.equal((r!.result as any).protocolVersion, "2025-06-18");
  });

  test("notifications get no response", async () => {
    assert.equal(await handleJsonRpc({ jsonrpc: "2.0", method: "notifications/initialized" }, deps), null);
  });

  test("tools/call returns text + structured content", async () => {
    const r = await handleJsonRpc({ jsonrpc: "2.0", id: 2, method: "tools/call", params: { name: "ping_tool", arguments: { a: 1 } } }, deps);
    const result = r!.result as any;
    assert.deepEqual(result.structuredContent, { echo: { a: 1 } });
    assert.equal(result.content[0].type, "text");
    assert.equal(result.isError, false);
  });

  test("unknown tool and unknown method are errors", async () => {
    assert.equal((await handleJsonRpc({ jsonrpc: "2.0", id: 3, method: "tools/call", params: { name: "nope" } }, deps))!.error!.code, -32602);
    assert.equal((await handleJsonRpc({ jsonrpc: "2.0", id: 4, method: "resources/list" }, deps))!.error!.code, -32601);
  });
});

// ─── Write-safety logic ─────────────────────────────────────────────────────

describe("write safety", () => {
  const sizes = ["XS", "S", "M", "L", "XL", "2XL", "3XL"];

  test("size matching folds common spellings", () => {
    assert.equal(matchSize("3xl", sizes), "3XL");
    assert.equal(matchSize("XXXL", sizes), "3XL");
    assert.equal(matchSize(" xxl ", sizes), "2XL");
    assert.equal(matchSize("4XL", sizes), null);
    assert.equal(matchSize("m-l", ["S/M", "M/L"]), "M/L");
  });

  test("first test case: +5 on 3XL", () => {
    const r = resolveQuantityChanges({ "3XL": 4, L: 10 }, sizes, [{ size: "3XL", add: 5 }]);
    assert.deepEqual(r, { ok: true, changes: [{ size: "3XL", before: 4, after: 9 }] });
  });

  test("a size not on the order starts from zero", () => {
    const r = resolveQuantityChanges({ L: 10 }, sizes, [{ size: "3XL", add: 5 }]);
    assert.deepEqual(r, { ok: true, changes: [{ size: "3XL", before: 0, after: 5 }] });
  });

  test("rejects unknown sizes, negatives, both/neither of add and set_to, decimals, duplicates", () => {
    assert.equal(resolveQuantityChanges({}, sizes, [{ size: "5XL", add: 1 }]).ok, false);
    assert.equal(resolveQuantityChanges({ M: 2 }, sizes, [{ size: "M", add: -3 }]).ok, false);
    assert.equal(resolveQuantityChanges({}, sizes, [{ size: "M", add: 1, setTo: 2 }]).ok, false);
    assert.equal(resolveQuantityChanges({}, sizes, [{ size: "M" }]).ok, false);
    assert.equal(resolveQuantityChanges({}, sizes, [{ size: "M", add: 1.5 }]).ok, false);
    assert.equal(resolveQuantityChanges({}, sizes, [{ size: "M", add: 1 }, { size: "m", add: 1 }]).ok, false);
    assert.equal(resolveQuantityChanges({}, sizes, []).ok, false);
  });

  test("new order quantities are validated", () => {
    assert.deepEqual(resolveNewOrderQuantities(sizes, { "3xl": 5 }), { ok: true, qtys: { "3XL": 5 }, total: 5 });
    assert.equal(resolveNewOrderQuantities(sizes, { "3XL": 0 }).ok, false);
    assert.equal(resolveNewOrderQuantities(sizes, { "9XL": 5 }).ok, false);
    assert.equal(resolveNewOrderQuantities(sizes, { M: -1 }).ok, false);
  });

  test("duplicate detection only counts open orders on order / in production", () => {
    const d = findPotentialDuplicates([
      { id: 1, status: "open", supplierStatus: "on_order", totalQty: 10, createdAt: new Date() },
      { id: 2, status: "open", supplierStatus: "on_production", totalQty: 10, createdAt: new Date() },
      { id: 3, status: "open", supplierStatus: "in_shipment", totalQty: 10, createdAt: new Date() },
      { id: 4, status: "closed", supplierStatus: "on_order", totalQty: 10, createdAt: new Date() },
      { id: 5, status: "open", supplierStatus: "cancelled", totalQty: 10, createdAt: new Date() },
    ]);
    assert.deepEqual(d.map((o) => o.id), [1, 2]);
  });

  test("notes are appended, never overwritten", () => {
    const date = new Date("2026-11-25T01:00:00Z");
    assert.equal(appendNote("Old note", "Add 5 × 3XL", date), "Old note\n[ChatGPT 25/11/2026] Add 5 × 3XL");
    assert.equal(appendNote(null, "First", date), "[ChatGPT 25/11/2026] First");
  });

  test("stable snapshot comparison ignores key order", () => {
    assert.equal(stableStringify({ a: 1, b: { c: 2, d: 3 } }), stableStringify({ b: { d: 3, c: 2 }, a: 1 }));
    assert.notEqual(stableStringify({ M: 4 }), stableStringify({ M: 5 }));
  });

  test("inline images are stripped before data goes to ChatGPT", () => {
    const out = compactForModel({ img: "data:image/png;base64,AAAA", name: "x".repeat(3000) }) as Record<string, string>;
    assert.equal(out.img, "[image omitted]");
    assert.ok(out.name.length < 2100);
  });

  test("staff recipients resolve by id, full name, unique first name; never guess", () => {
    const users = [
      { id: "u1", name: "Girish Kumar" },
      { id: "u2", name: "Girish Patel" },
      { id: "u3", name: "Kat" },
      { id: "u4", name: "Leela", active: false },
    ];
    // exact id and unique first name
    assert.deepEqual(matchStaffRecipients(users, ["u1", "Kat"]), { ok: true, users: [{ id: "u1", name: "Girish Kumar" }, { id: "u3", name: "Kat" }] });
    // full name disambiguates two Girishes
    assert.deepEqual(matchStaffRecipients(users, ["Girish Patel"]), { ok: true, users: [{ id: "u2", name: "Girish Patel" }] });
    // ambiguous first name is refused, not guessed
    const amb = matchStaffRecipients(users, ["Girish"]);
    assert.equal(amb.ok, false);
    // unknown name is refused
    assert.equal(matchStaffRecipients(users, ["Nobody"]).ok, false);
    // inactive staff are not matchable
    assert.equal(matchStaffRecipients(users, ["Leela"]).ok, false);
    // duplicates collapse to one recipient
    assert.deepEqual(matchStaffRecipients(users, ["Kat", "kat"]), { ok: true, users: [{ id: "u3", name: "Kat" }] });
  });
});

// ─── Background-job guard ───────────────────────────────────────────────────
import { backgroundJobsEnabled } from "../background-jobs.server.ts";
describe("background jobs guard", () => {
  test("runs in production and off Railway, never in a fork", () => {
    assert.equal(backgroundJobsEnabled({ RAILWAY_ENVIRONMENT_NAME: "production" }), true);
    assert.equal(backgroundJobsEnabled({}), true);
    assert.equal(backgroundJobsEnabled({ RAILWAY_ENVIRONMENT_NAME: "production-fork-ab12" }), false);
    assert.equal(backgroundJobsEnabled({ RAILWAY_ENVIRONMENT_NAME: "production", PORTAL_BACKGROUND_JOBS: "off" }), false);
  });
});
