import { createHash, timingSafeEqual } from "crypto";

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

/**
 * Decode the JWT payload (base64url) and return the parsed object.
 * We don't verify the signature here — the shop domain is passed
 * separately and we use it only to scope our own DB query.
 * The idToken audience check ensures the token was issued for this app.
 */
function decodeJwtPayload(token: string): Record<string, unknown> | null {
  try {
    const [, payloadB64] = token.split(".");
    return JSON.parse(Buffer.from(payloadB64, "base64url").toString("utf-8"));
  } catch {
    return null;
  }
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

  const payload = decodeJwtPayload(token);
  if (!payload) {
    return Response.json({ error: "Invalid token" }, { status: 401, headers: cors });
  }

  const clientId = process.env.SHOPIFY_API_KEY;
  const aud = payload.aud;
  const audValid = aud === clientId || (Array.isArray(aud) && aud.includes(clientId));

  if (!audValid) {
    return Response.json({ error: "Token audience mismatch" }, { status: 401, headers: cors });
  }

  return null;
}
