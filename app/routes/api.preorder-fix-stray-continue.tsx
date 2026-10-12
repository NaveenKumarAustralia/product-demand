import type { LoaderFunctionArgs } from "react-router";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { sweepStrayContinueInventoryPolicies } from "../preorder/preorder-inventory-policy.server";

// Fix products wrongly left on "continue selling when out of stock" (inventory
// policy CONTINUE) WITHOUT pre-order enabled. Scans every Shopify product and,
// for any that is NOT a currently-enabled pre-order batch, flips its CONTINUE
// variants back to DENY.
//
//   GET /api/preorder-fix-stray-continue            → DRY RUN: list every product
//       that would be fixed (and how many variants). Changes NOTHING.
//   GET /api/preorder-fix-stray-continue?apply=1     → apply: set those variants
//       back to DENY in Shopify.
//
// Admin only. Enabled pre-order products are left untouched.
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const apply = new URL(request.url).searchParams.get("apply") === "1";
  const result = await sweepStrayContinueInventoryPolicies({ apply });

  return Response.json(
    {
      ok: result.ok,
      mode: apply ? "applied" : "dry-run",
      shop: result.shop,
      error: result.error,
      scannedProducts: result.scannedProducts,
      excludedPreorderProducts: result.excludedPreorderProducts,
      productsToFix: result.affected.length,
      variantsToFix: result.totalVariants,
      applied: result.applied,
      changedProducts: result.changedProducts,
      changedVariants: result.changedVariants,
      truncated: result.truncated,
      errors: result.errors,
      affected: result.affected,
      hint: apply
        ? "Done. Re-run the dry run (no ?apply) to confirm 0 remain."
        : "DRY RUN — nothing changed. Review 'affected', then add ?apply=1 to set them all back to DENY.",
    },
    { headers: { "Cache-Control": "no-store" } },
  );
};
