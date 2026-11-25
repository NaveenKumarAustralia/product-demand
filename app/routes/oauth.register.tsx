import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import { clientIp, json, NO_STORE, oauthErrorResponse, preflight, rateLimited } from "../mcp/http.server";
import { allowedRedirectHosts, OAuthError, registerClient } from "../mcp/oauth-core";
import { prismaOAuthStore } from "../mcp/oauth-store.server";

// Dynamic client registration (RFC 7591). Only redirect URIs on allowed hosts
// (ChatGPT by default) are accepted.
export const loader = async ({ request }: LoaderFunctionArgs) => (request.method === "OPTIONS" ? preflight() : new Response(null, { status: 405 }));

export const action = async ({ request }: ActionFunctionArgs) => {
  if (request.method === "OPTIONS") return preflight();
  try {
    if (rateLimited(`register:${clientIp(request)}`, 20, 60 * 60 * 1000)) throw new OAuthError("temporarily_unavailable", "Too many registrations", 429);
    let body: unknown;
    try {
      body = await request.json();
    } catch {
      throw new OAuthError("invalid_client_metadata", "Body must be JSON");
    }
    const client = await registerClient(prismaOAuthStore, body, allowedRedirectHosts());
    return json(client, { status: 201, headers: NO_STORE });
  } catch (err) {
    return oauthErrorResponse(err);
  }
};
