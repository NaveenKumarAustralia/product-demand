import type { LoaderFunctionArgs } from "react-router";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { reconcileAllPreorderSellingPlanDates } from "../preorder/preorder-selling-plan.service.server";

// Admin: re-stamp the pre-order variant metafields (karmaeast.preorder / dispatch)
// on EVERY live batch's variants, and re-sync each selling plan's date. Use this
// to backfill after the metafield DEFINITION was created — values written before
// the definition existed were only "adopted" and can be invisible to the
// notification email; re-writing them under the definition makes the email banner
// render reliably. Idempotent (daily scheduler runs the same thing).
//   GET /api/preorder-reconcile-plans
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });
  const result = await reconcileAllPreorderSellingPlanDates();
  return Response.json({ ok: true, ...result }, { headers: { "Cache-Control": "no-store" } });
};
