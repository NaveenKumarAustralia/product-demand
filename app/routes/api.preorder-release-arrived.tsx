import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { releaseArrivedPreorders } from "../preorder/preorder-release.server";

// Admin trigger to run the pre-order auto-release immediately (instead of waiting
// for the 10-min timer) — release Shopify holds + tag orders for Pick Pack for any
// batch whose stock has landed. Read-only-safe to call repeatedly (idempotent via
// readyAt). GET so it can be run from the browser while logged into the portal.
async function run(request: Request) {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });
  const result = await releaseArrivedPreorders();
  return Response.json({ ok: true, ...result }, { headers: { "Cache-Control": "no-store" } });
}

export const loader = async ({ request }: LoaderFunctionArgs) => run(request);
export const action = async ({ request }: ActionFunctionArgs) => run(request);
