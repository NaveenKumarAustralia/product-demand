import type { LoaderFunctionArgs } from "react-router";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { captureMissedPreorders, listAffectedForFollowup } from "../preorder/preorder-missed-capture.server";

// Admin: find (and optionally convert) paid orders that sold a pre-order-enabled
// variant WITHOUT the selling plan (bought via quick-add / theme button rather
// than the pre-order button). GET so it runs from the browser while logged in.
//   ?days=21            how far back to scan (default 21)
//   &apply=1            actually reserve + tag + hold them (omit for a dry run)
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });
  const url = new URL(request.url);
  const days = Number(url.searchParams.get("days") ?? "21") || 21;
  // Read-only follow-up list (every affected order + customer email), no gating,
  // no changes — for emailing customers who got a normal confirmation.
  if (url.searchParams.get("followup") === "1") {
    const result = await listAffectedForFollowup({ days });
    return Response.json({ ok: true, ...result }, { headers: { "Cache-Control": "no-store" } });
  }
  const apply = url.searchParams.get("apply") === "1";
  const result = await captureMissedPreorders({ days, apply });
  return Response.json({ ok: true, ...result }, { headers: { "Cache-Control": "no-store" } });
};
