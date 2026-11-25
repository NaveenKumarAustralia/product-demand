import { OAuthError } from "./oauth-core";

/**
 * Public base URL of this app, as ChatGPT sees it. MCP_PUBLIC_URL wins (set it
 * in Railway); otherwise it is derived from the proxy headers.
 */
export function publicBase(request: Request): string {
  const fixed = process.env.MCP_PUBLIC_URL?.trim().replace(/\/+$/, "");
  if (fixed) return fixed;
  const url = new URL(request.url);
  const host = request.headers.get("x-forwarded-host")?.split(",")[0]?.trim() || request.headers.get("host") || url.host;
  const proto = request.headers.get("x-forwarded-proto")?.split(",")[0]?.trim() || url.protocol.replace(":", "");
  return `${proto}://${host}`;
}

export const NO_STORE = { "Cache-Control": "no-store", Pragma: "no-cache" };
const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type, Mcp-Protocol-Version, Mcp-Session-Id",
};

export function json(body: unknown, init: { status?: number; headers?: Record<string, string> } = {}) {
  return new Response(JSON.stringify(body), {
    status: init.status ?? 200,
    headers: { "Content-Type": "application/json; charset=utf-8", ...CORS, ...init.headers },
  });
}

export function preflight() {
  return new Response(null, { status: 204, headers: CORS });
}

export function oauthErrorResponse(err: unknown) {
  if (err instanceof OAuthError) return json(err.toJSON(), { status: err.status, headers: NO_STORE });
  console.error("OAuth error", err);
  return json({ error: "server_error", error_description: "Unexpected error" }, { status: 500, headers: NO_STORE });
}

/** Simple in-memory limiter: N attempts per key per window. Good enough for one instance. */
const buckets = new Map<string, { count: number; resetAt: number }>();
export function rateLimited(key: string, limit: number, windowMs: number): boolean {
  const now = Date.now();
  const b = buckets.get(key);
  if (!b || b.resetAt < now) {
    buckets.set(key, { count: 1, resetAt: now + windowMs });
    if (buckets.size > 10000) for (const [k, v] of buckets) if (v.resetAt < now) buckets.delete(k);
    return false;
  }
  b.count += 1;
  return b.count > limit;
}

export function clientIp(request: Request): string {
  return request.headers.get("x-forwarded-for")?.split(",")[0]?.trim() || request.headers.get("x-real-ip") || "unknown";
}
