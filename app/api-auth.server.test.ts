// node --test app/api-auth.server.test.ts
import { test, beforeEach } from "node:test";
import assert from "node:assert/strict";
import { authorizeApiRequest } from "./api-auth.server.ts";

const CORS = { "Access-Control-Allow-Origin": "*" };
const SECRET = "dashboard-shared-secret";

// A Shopify session token is only ever read for its audience claim, so an
// unsigned JWT with the right payload is enough to exercise that path.
function sessionToken(aud: string) {
  const part = (obj: unknown) => Buffer.from(JSON.stringify(obj)).toString("base64url");
  return `${part({ alg: "HS256", typ: "JWT" })}.${part({ aud })}.signature`;
}

const request = (headers: Record<string, string>) =>
  new Request("https://portal.example/api/order-status", { headers });

beforeEach(() => {
  process.env.DASHBOARD_API_KEY = SECRET;
  process.env.SHOPIFY_API_KEY = "shopify-client-id";
});

test("accepts the dashboard's shared secret", () => {
  assert.equal(authorizeApiRequest(request({ "x-api-key": SECRET }), CORS), null);
});

test("rejects a wrong shared secret", async () => {
  const res = authorizeApiRequest(request({ "x-api-key": "nope" }), CORS);
  assert.equal(res?.status, 401);
  assert.equal((await res!.json()).error, "Invalid API key");
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

test("a malformed bearer token is rejected, not crashed on", async () => {
  const res = authorizeApiRequest(request({ Authorization: "Bearer not-a-jwt" }), CORS);
  assert.equal(res?.status, 401);
});
