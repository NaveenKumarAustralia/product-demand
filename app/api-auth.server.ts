import { timingSafeEqual } from "crypto";

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

/** Length-safe, timing-safe string comparison. */
function secretsMatch(provided: string, expected: string): boolean {
  const a = Buffer.from(provided);
  const b = Buffer.from(expected);
  if (a.length !== b.length) return false;
  return timingSafeEqual(a, b);
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
      return Response.json({ error: "Invalid API key" }, { status: 401, headers: cors });
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
