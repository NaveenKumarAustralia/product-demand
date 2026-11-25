import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import { clientIp, json, NO_STORE, oauthErrorResponse, preflight, rateLimited } from "../mcp/http.server";
import { exchangeToken, OAuthError } from "../mcp/oauth-core";
import { portalUserDirectory, prismaOAuthStore } from "../mcp/oauth-store.server";

export const loader = async ({ request }: LoaderFunctionArgs) => (request.method === "OPTIONS" ? preflight() : new Response(null, { status: 405 }));

export const action = async ({ request }: ActionFunctionArgs) => {
  if (request.method === "OPTIONS") return preflight();
  try {
    if (rateLimited(`token:${clientIp(request)}`, 120, 60 * 1000)) throw new OAuthError("temporarily_unavailable", "Too many requests", 429);
    const type = request.headers.get("content-type") ?? "";
    const form = type.includes("application/json")
      ? new URLSearchParams(Object.entries((await request.json()) as Record<string, string>))
      : new URLSearchParams(await request.text());
    const tokens = await exchangeToken(prismaOAuthStore, portalUserDirectory, form);
    return json(tokens, { headers: NO_STORE });
  } catch (err) {
    return oauthErrorResponse(err);
  }
};
