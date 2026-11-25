import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import { NO_STORE, preflight } from "../mcp/http.server";
import { revokeToken } from "../mcp/oauth-core";
import { prismaOAuthStore } from "../mcp/oauth-store.server";

// Token revocation (RFC 7009). Revoking either token ends the whole connection.
export const loader = async ({ request }: LoaderFunctionArgs) => (request.method === "OPTIONS" ? preflight() : new Response(null, { status: 405 }));

export const action = async ({ request }: ActionFunctionArgs) => {
  if (request.method === "OPTIONS") return preflight();
  const form = new URLSearchParams(await request.text());
  await revokeToken(prismaOAuthStore, form.get("token") ?? "").catch((e) => console.error("revoke failed", e));
  return new Response(null, { status: 200, headers: NO_STORE });
};
