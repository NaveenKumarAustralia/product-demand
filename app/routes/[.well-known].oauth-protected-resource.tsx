import type { LoaderFunctionArgs } from "react-router";
import { json, preflight, publicBase } from "../mcp/http.server";
import { protectedResourceMetadata } from "../mcp/oauth-core";

// OAuth protected-resource metadata (RFC 9728) for the /mcp endpoint.
export const loader = async ({ request }: LoaderFunctionArgs) => {
  if (request.method === "OPTIONS") return preflight();
  return json(protectedResourceMetadata(publicBase(request)));
};
