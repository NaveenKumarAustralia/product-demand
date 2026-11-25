// node --test app/api-auth.server.test.ts
import { test, beforeEach } from "node:test";
import assert from "node:assert/strict";
import { createHmac } from "node:crypto";
import { authorizeApiRequest, describeKeyMismatch } from "./api-auth.server.ts";

const CORS = { "Access-Control-Allow-Origin": "*" };
const SECRET = "dashboard-shared-secret";
const APP_SECRET = "shopify-app-secret";

const part = (obj: unknown) => Buffer.from(JSON.stringify(obj)).toString("base64url");
const now = () => Math.floor(Date.now() / 1000);

// A real Shopify session token: HS256, signed with the app's client secret.
function sessionToken(aud: string, opts: { secret?: string; exp?: number; alg?: string; nbf?: number } = {}) {
  const header = part({ alg: opts.alg ?? "HS256", typ: "JWT" });
  const payload = part({ aud, exp: opts.exp ?? now() + 60, nbf: opts.nbf ?? now() - 5 });
  const sig = createHmac("sha256", opts.secret ?? APP_SECRET).update(`${header}.${payload}`).digest("base64url");
  return `${header}.${payload}.${sig}`;
}

const request = (headers: Record<string, string>) =>
  new Request("https://portal.example/api/order-status", { headers });

beforeEach(() => {
  process.env.DASHBOARD_API_KEY = SECRET;
  process.env.SHOPIFY_API_KEY = "shopify-client-id";
  process.env.SHOPIFY_API_SECRET = APP_SECRET;
});

test("accepts the dashboard's shared secret", () => {
  assert.equal(authorizeApiRequest(request({ "x-api-key": SECRET }), CORS), null);
});

test("rejects a wrong shared secret", async () => {
  const res = authorizeApiRequest(request({ "x-api-key": "nope" }), CORS);
  assert.equal(res?.status, 401);
  assert.match((await res!.json()).error, /^Invalid API key/);
});

test("whitespace pasted around the stored secret still authorises", () => {
  // The failure this prevents is nasty because it isn't symmetric: HTTP trims
  // trailing whitespace from header values in transit, so padding on the
  // caller's side vanishes while padding on this side survives — and the
  // resulting mismatch is invisible in any UI that displays the value.
  process.env.DASHBOARD_API_KEY = `  ${SECRET}\n`;
  assert.equal(authorizeApiRequest(request({ "x-api-key": SECRET }), CORS), null);
});

test("a mismatch reports lengths and fingerprints, never the secrets", async () => {
  const res = authorizeApiRequest(request({ "x-api-key": `${SECRET}extra` }), CORS);
  const { error } = await res!.json();

  assert.match(error, new RegExp(`sent ${SECRET.length + 5} chars`));
  assert.match(error, new RegExp(`expected ${SECRET.length} chars`));
  // Neither value may appear in the response.
  assert.equal(error.includes(SECRET), false);
});

test("same length but different value is reported as such", async () => {
  const wrong = "x".repeat(SECRET.length);
  const res = authorizeApiRequest(request({ "x-api-key": wrong }), CORS);
  const { error } = await res!.json();
  assert.match(error, new RegExp(`both ${SECRET.length} chars, but different values`));
});

test("identical secrets fingerprint identically", () => {
  // What makes the fingerprints comparable across two separate services.
  const a = describeKeyMismatch("same-value", "different");
  const b = describeKeyMismatch("same-value", "different");
  assert.equal(a, b);
});

test("an unset DASHBOARD_API_KEY authorises nobody", async () => {
  // The dangerous failure mode: if a missing env var compared as equal to an
  // absent header, every anonymous caller would be let straight in.
  delete process.env.DASHBOARD_API_KEY;
  const res = authorizeApiRequest(request({ "x-api-key": "" }), CORS);
  assert.equal(res?.status, 401);
});

test("still accepts a Shopify session token from the embedded app", () => {
  const token = sessionToken("shopify-client-id");
  assert.equal(authorizeApiRequest(request({ Authorization: `Bearer ${token}` }), CORS), null);
});

test("rejects a session token issued for a different app", async () => {
  const token = sessionToken("someone-elses-app");
  const res = authorizeApiRequest(request({ Authorization: `Bearer ${token}` }), CORS);
  assert.equal(res?.status, 401);
  assert.equal((await res!.json()).error, "Token audience mismatch");
});

test("rejects a request with no credentials at all", async () => {
  const res = authorizeApiRequest(request({}), CORS);
  assert.equal(res?.status, 401);
  assert.equal((await res!.json()).error, "Missing token");
});

test("rejects a forged token with the right audience but no valid signature", async () => {
  // The hole this closes: the client id is public, so audience alone proves nothing.
  const forged = `${part({ alg: "HS256", typ: "JWT" })}.${part({ aud: "shopify-client-id", exp: now() + 60 })}.signature`;
  const res = authorizeApiRequest(request({ Authorization: `Bearer ${forged}` }), CORS);
  assert.equal(res?.status, 401);
  assert.equal((await res!.json()).error, "Invalid token signature");
});

test("rejects a token signed with a different secret", async () => {
  const token = sessionToken("shopify-client-id", { secret: "attacker-secret" });
  const res = authorizeApiRequest(request({ Authorization: `Bearer ${token}` }), CORS);
  assert.equal(res?.status, 401);
});

test("rejects alg none", async () => {
  const token = `${part({ alg: "none", typ: "JWT" })}.${part({ aud: "shopify-client-id", exp: now() + 60 })}.`;
  const res = authorizeApiRequest(request({ Authorization: `Bearer ${token}` }), CORS);
  assert.equal(res?.status, 401);
});

test("rejects an expired token", async () => {
  const token = sessionToken("shopify-client-id", { exp: now() - 600 });
  const res = authorizeApiRequest(request({ Authorization: `Bearer ${token}` }), CORS);
  assert.equal(res?.status, 401);
  assert.equal((await res!.json()).error, "Token expired");
});

test("an unset SHOPIFY_API_SECRET accepts no session token", async () => {
  delete process.env.SHOPIFY_API_SECRET;
  const res = authorizeApiRequest(request({ Authorization: `Bearer ${sessionToken("shopify-client-id")}` }), CORS);
  assert.equal(res?.status, 401);
});

test("a malformed bearer token is rejected, not crashed on", async () => {
  const res = authorizeApiRequest(request({ Authorization: "Bearer not-a-jwt" }), CORS);
  assert.equal(res?.status, 401);
});
