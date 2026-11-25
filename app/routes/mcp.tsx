import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import { json, preflight, publicBase } from "../mcp/http.server";
import { authenticateBearer, OAuthError } from "../mcp/oauth-core";
import { portalUserDirectory, prismaOAuthStore } from "../mcp/oauth-store.server";
import { handleJsonRpc, type JsonRpcRequest } from "../mcp/protocol";
import { runTool, TOOLS } from "../mcp/tools.server";

/**
 * Remote MCP endpoint for ChatGPT (Streamable HTTP, stateless JSON responses).
 * Every request needs an OAuth access token issued by /oauth/token.
 */

const toolSpecs = TOOLS.map(({ name, title, description, inputSchema, annotations }) => ({ name, title, description, inputSchema, annotations }));

function unauthorized(request: Request, err: OAuthError) {
  const base = publicBase(request);
  return json(
    { jsonrpc: "2.0", id: null, error: { code: -32001, message: err.description } },
    {
      status: 401,
      headers: {
        "WWW-Authenticate": `Bearer resource_metadata="${base}/.well-known/oauth-protected-resource", error="invalid_token", error_description="${err.description.replace(/"/g, "'")}"`,
      },
    },
  );
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  if (request.method === "OPTIONS") return preflight();
  // No server-initiated stream: this server only answers POSTs.
  return new Response(null, { status: 405, headers: { Allow: "POST, OPTIONS" } });
};

export const action = async ({ request }: ActionFunctionArgs) => {
  if (request.method === "OPTIONS") return preflight();
  if (request.method === "DELETE") return new Response(null, { status: 204 });
  if (request.method !== "POST") return new Response(null, { status: 405, headers: { Allow: "POST, OPTIONS" } });

  let principal;
  try {
    principal = await authenticateBearer(prismaOAuthStore, portalUserDirectory, request.headers.get("authorization"));
  } catch (err) {
    if (err instanceof OAuthError) return unauthorized(request, err);
    console.error("MCP auth error", err);
    return json({ jsonrpc: "2.0", id: null, error: { code: -32603, message: "Internal error" } }, { status: 500 });
  }

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return json({ jsonrpc: "2.0", id: null, error: { code: -32700, message: "Parse error" } }, { status: 400 });
  }

  const deps = { tools: toolSpecs, call: (name: string, args: Record<string, unknown>) => runTool(name, args, { principal }) };
  if (Array.isArray(body)) {
    const responses = (await Promise.all(body.map((m) => handleJsonRpc(m as JsonRpcRequest, deps)))).filter(Boolean);
    return responses.length ? json(responses) : new Response(null, { status: 202 });
  }
  const response = await handleJsonRpc(body as JsonRpcRequest, deps);
  return response ? json(response) : new Response(null, { status: 202 });
};
