import type { LoaderFunctionArgs } from "react-router";
import { getKlaviyoConnectionStatus } from "../preorder/preorder-klaviyo.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  try {
    await requirePreorderPortalUser(request);
    return Response.json({ ok: true, ...getKlaviyoConnectionStatus() });
  } catch (error) {
    if (error instanceof Response) return error;
    return Response.json({ ok: false, configured: false, error: "Could not check Klaviyo connection." }, { status: 500 });
  }
};
