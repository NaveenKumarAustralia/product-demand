import { createHash, createHmac, timingSafeEqual } from "crypto";

/**
 * Authorisation for the portal's public API routes.
 *
 * Two callers need in, and they can't authenticate the same way:
 *
 *  1. The embedded app running inside Shopify admin, which sends a session
 *     token minted by App Bridge. That's the original path — a JWT whose
 *     audience is this app's client id.
 *
 *  2. The Karma East analytics dashboard, a standalone Railway service. It has
 *     no Shopify session and cannot obtain one, so it presents a shared secret
 *     instead. The secret only ever travels server-to-server; the dashboard's
 *     browser never sees it.
 *
 * Returns null when the request is allowed, or the Response to send back when
 * it isn't.
 */

/** Allowed clock drift between Shopify and this server, in seconds. */
const SESSION_TOKEN_LEEWAY_SECONDS = 60;

/**
 * Verify a Shopify App Bridge session token and return its payload.
 *
 * Session tokens are HS256 JWTs signed with the app's client secret
 * (SHOPIFY_API_SECRET). The audience (client id) is public, so checking it
 * alone proves nothing: anyone could mint an unsigned token with the right
 * `aud`. The signature is what proves Shopify issued the token, and `exp` /
 * `nbf` stop an old token being replayed.
 */
export type SessionTokenResult =
  | { ok: true; payload: Record<string, unknown> }
  | { ok: false; error: string };

export function verifyShopifySessionToken(token: string, secret: string, nowSeconds = Math.floor(Date.now() / 1000)): SessionTokenResult {
  const parts = token.split(".");
  if (parts.length !== 3) return { ok: false, error: "Invalid token" };
  const [headerB64, payloadB64, signatureB64] = parts;

  let header: Record<string, unknown>;
  let payload: Record<string, unknown>;
  try {
    header = JSON.parse(Buffer.from(headerB64, "base64url").toString("utf-8"));
    payload = JSON.parse(Buffer.from(payloadB64, "base64url").toString("utf-8"));
  } catch {
    return { ok: false, error: "Invalid token" };
  }
  if (!payload || typeof payload !== "object") return { ok: false, error: "Invalid token" };
  // Only HS256 is ever issued; refusing anything else blocks "alg: none".
  if (header?.alg !== "HS256") return { ok: false, error: "Invalid token" };

  const expected = createHmac("sha256", secret).update(`${headerB64}.${payloadB64}`).digest();
  let provided: Buffer;
  try {
    provided = Buffer.from(signatureB64, "base64url");
  } catch {
    return { ok: false, error: "Invalid token signature" };
  }
  if (provided.length !== expected.length || !timingSafeEqual(provided, expected)) {
    return { ok: false, error: "Invalid token signature" };
  }

  const exp = Number(payload.exp);
  if (!Number.isFinite(exp) || nowSeconds > exp + SESSION_TOKEN_LEEWAY_SECONDS) {
    return { ok: false, error: "Token expired" };
  }
  const nbf = Number(payload.nbf);
  if (Number.isFinite(nbf) && nowSeconds + SESSION_TOKEN_LEEWAY_SECONDS < nbf) {
    return { ok: false, error: "Token not yet valid" };
  }
  return { ok: true, payload };
}

/**
 * Length-safe, timing-safe secret comparison.
 *
 * Both sides are trimmed first. Secrets get pasted into hosting dashboards by
 * hand, and a stray trailing space or newline is invisible in every UI that
 * shows the value. Worse, the two ends don't fail symmetrically: HTTP strips
 * trailing whitespace from header values in transit, so padding on the caller's
 * side quietly disappears while padding on this side does not — producing a
 * mismatch that looks like a wrong key and reads as a wrong key, but isn't one.
 *
 * Trimming costs nothing in strength: whitespace at either end carries no
 * entropy, and the comparison stays constant-time over the trimmed values.
 */
function secretsMatch(provided: string, expected: string): boolean {
  const a = Buffer.from(provided.trim());
  const b = Buffer.from(expected.trim());
  if (a.length !== b.length) return false;
  return timingSafeEqual(a, b);
}

/**
 * A short, irreversible fingerprint of a secret, so a mismatch can be diagnosed
 * without either side's value ever being printed. Two identical secrets always
 * produce the same fingerprint; different ones effectively never do.
 *
 * Safe to return on a 401: reversing eight hex characters of SHA-256 back to a
 * 256-bit random key isn't feasible, and it reveals nothing about the value's
 * content — only whether the two ends agree.
 */
function fingerprint(value: string): string {
  return createHash("sha256").update(value).digest("hex").slice(0, 8);
}

/**
 * Describes *how* two secrets differ. Length is the giveaway for the usual
 * culprit — a trailing space or newline picked up when pasting into a
 * dashboard — which is invisible in every UI that shows the value.
 */
export function describeKeyMismatch(provided: string, expected: string): string {
  // Reported on the trimmed values, since those are what actually get compared.
  const a = provided.trim();
  const b = expected.trim();
  const detail = a.length === b.length
    ? `both ${a.length} chars, but different values (sent ${fingerprint(a)}, expected ${fingerprint(b)})`
    : `sent ${a.length} chars (${fingerprint(a)}), expected ${b.length} chars (${fingerprint(b)})`;
  return `Invalid API key — ${detail}`;
}

/**
 * Authorisation for the read-only feed the Aisling Equation factory app pulls
 * (packing lists and the like). Its own secret, AISLING_API_KEY, and only these
 * routes check it — so that key can read the feed but can never place or edit
 * orders through the routes authorizeApiRequest guards.
 */
export function authorizeAislingRequest(request: Request): Response | null {
  const expected = process.env.AISLING_API_KEY;
  // Same rule as above: an unset secret must never let everyone in.
  if (!expected) {
    return Response.json({ error: "Aisling access not configured (set AISLING_API_KEY)" }, { status: 401 });
  }
  const provided = request.headers.get("x-api-key") ?? "";
  if (!provided) return Response.json({ error: "Missing API key" }, { status: 401 });
  if (!secretsMatch(provided, expected)) {
    return Response.json({ error: describeKeyMismatch(provided, expected) }, { status: 401 });
  }
  return null;
}

export function authorizeApiRequest(request: Request, cors: Record<string, string>): Response | null {
  // Service-to-service: a shared secret in its own header, checked first so a
  // machine caller never has to fake a Shopify token.
  const apiKey = request.headers.get("x-api-key");
  if (apiKey) {
    const expected = process.env.DASHBOARD_API_KEY;
    // An unset env var must never mean "everything is allowed" — without this
    // guard a missing secret on the server would authorise every caller that
    // sent no key at all.
    if (!expected) {
      return Response.json({ error: "Service auth not configured" }, { status: 401, headers: cors });
    }
    if (!secretsMatch(apiKey, expected)) {
      return Response.json({ error: describeKeyMismatch(apiKey, expected) }, { status: 401, headers: cors });
    }
    return null;
  }

  // Embedded app: the original Shopify session-token path, unchanged.
  const authHeader = request.headers.get("Authorization") ?? "";
  const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : null;

  if (!token) {
    return Response.json({ error: "Missing token" }, { status: 401, headers: cors });
  }

  const secret = process.env.SHOPIFY_API_SECRET;
  // Same rule as the shared secret: no configured secret must never mean
  // "accept any token".
  if (!secret) {
    return Response.json({ error: "Session token auth not configured" }, { status: 401, headers: cors });
  }
  const verified = verifyShopifySessionToken(token, secret);
  if (!verified.ok) {
    return Response.json({ error: verified.error }, { status: 401, headers: cors });
  }
  const payload = verified.payload;

  const clientId = process.env.SHOPIFY_API_KEY;
  const aud = payload.aud;
  const audValid = aud === clientId || (Array.isArray(aud) && aud.includes(clientId));

  if (!audValid) {
    return Response.json({ error: "Token audience mismatch" }, { status: 401, headers: cors });
  }

  return null;
}
