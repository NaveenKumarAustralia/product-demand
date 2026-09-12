import type { LoaderFunctionArgs } from "react-router";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { resplitHeldPreorderOrders } from "../preorder/preorder-release.server";

// Admin: fix EXISTING held pre-order orders so only the pre-order line stays
// held and the in-stock lines ship now (older orders were held whole). Dry-run
// by default; add &apply=1 to actually re-split.
//   GET /api/preorder-resplit-held[&apply=1]
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });
  const apply = new URL(request.url).searchParams.get("apply") === "1";
  const result = await resplitHeldPreorderOrders({ apply });
  return Response.json({ ok: true, ...result }, { headers: { "Cache-Control": "no-store" } });
};
