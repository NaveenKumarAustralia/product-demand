import type { LoaderFunctionArgs } from "react-router";
import { json, preflight, publicBase } from "../mcp/http.server";
import { authorizationServerMetadata } from "../mcp/oauth-core";

// OAuth authorization-server metadata (RFC 8414) for the ChatGPT connector.
export const loader = async ({ request }: LoaderFunctionArgs) => {
  if (request.method === "OPTIONS") return preflight();
  return json(authorizationServerMetadata(publicBase(request)));
};
