import type { LoaderFunctionArgs } from "react-router";
import { getKlaviyoConnectionStatus } from "../preorder/preorder-klaviyo.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

// Klaviyo connection health for the Back-in-Stock panel. Gated behind portal
// auth so the app's config state ("is a Klaviyo key present") is never readable
// by the public — the panel's same-origin fetch carries the session cookie, so
// it keeps working. Only ever returns a boolean + the API revision, never a key.
export const loader = async ({ request }: LoaderFunctionArgs) => {
  try {
    await requirePreorderPortalUser(request);
    const status = getKlaviyoConnectionStatus();
    return Response.json(
      { ok: true, configured: status.configured, revision: status.revision },
      { headers: { "Cache-Control": "no-store, max-age=0" } },
    );
  } catch (error) {
    if (error instanceof Response) return error;
    return Response.json({ ok: false, error: "Could not check Klaviyo health." }, { status: 500 });
  }
};
