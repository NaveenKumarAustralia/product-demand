import type { LoaderFunctionArgs } from "react-router";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { reportLivePreorderInventoryPolicies, reconcileAllLivePreorderInventoryPolicies } from "../preorder/preorder-inventory-policy.server";

// Make sure no pre-order variant can oversell past its batch: once a batch runs
// out of capacity (reserved ≥ incoming − buffer) its Shopify inventory policy
// must be DENY, so Shop Pay / PayPal / quick-add can't keep selling it. This
// normally happens automatically (on every order + a 3-hourly sweep); this
// endpoint lets an admin verify and force it on demand.
//
//   GET /api/preorder-reconcile-policies           → DRY RUN: report every live
//       pre-order variant, its remaining capacity, Shopify's current policy, the
//       policy it SHOULD have, and whether it's still oversellable. Changes nothing.
//   GET /api/preorder-reconcile-policies?apply=1    → apply the fix (flip full
//       batches to DENY, reopen ones with capacity), then report the new state.
//
// Admin only. Idempotent.
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const apply = new URL(request.url).searchParams.get("apply") === "1";
  if (apply) {
    try { await reconcileAllLivePreorderInventoryPolicies(); }
    catch (e) { return Response.json({ ok: false, applied: true, error: e instanceof Error ? e.message : String(e) }, { status: 500 }); }
  }

  const report = await reportLivePreorderInventoryPolicies();
  return Response.json({
    ok: report.ok,
    applied: apply,
    mode: apply ? "applied" : "dry-run",
    shop: report.shop,
    error: report.error,
    summary: report.summary,
    // After apply, needChange/stillOversellable should be 0.
    stillOversellable: report.rows.filter((r) => r.availableToPreorder <= 0 && (r.currentPolicy ?? "").toUpperCase() === "CONTINUE"),
    rows: report.rows,
  });
};
