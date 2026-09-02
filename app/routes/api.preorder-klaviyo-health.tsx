import type { LoaderFunctionArgs } from "react-router";
import { getKlaviyoConnectionStatus } from "../preorder/preorder-klaviyo.server";

export const loader = async (_args: LoaderFunctionArgs) => {
  const status = getKlaviyoConnectionStatus();
  return Response.json(
    { ok: true, configured: status.configured, revision: status.revision },
    { headers: { "Cache-Control": "no-store, max-age=0" } },
  );
};
