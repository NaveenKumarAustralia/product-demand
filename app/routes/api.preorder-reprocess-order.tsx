import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { processShopifyOrderCreated } from "../preorder/preorder-shopify-order.server";
import { reholdUnheldPreorderOrders } from "../preorder/preorder-fulfillment.server";

// Admin: fix a pre-order that didn't get held/captured. Re-runs the FULL order
// pipeline (capture any live pre-order line → reserve → hold → tag → email) for
// one order — idempotent, so it only fills in what's missing and won't
// double-email. Also: ?action=rehold-all re-applies holds to every reserved-but-
// unheld order at once.
//   /api/preorder-reprocess-order?orderId=1234567890      (numeric Shopify order id)
//   /api/preorder-reprocess-order?name=397534K            (order name, with or without #)
//   /api/preorder-reprocess-order?action=rehold-all
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const url = new URL(request.url);
  const session = await prisma.session.findFirst({
    where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true, accessToken: true },
  });
  if (!session?.shop || !session.accessToken) return Response.json({ ok: false, error: "No Shopify session." }, { status: 500 });
  const shop = session.shop;

  if ((url.searchParams.get("action") ?? "") === "rehold-all") {
    const r = await reholdUnheldPreorderOrders();
    return Response.json({ ok: r.errors === 0, action: "rehold-all", ...r }, { headers: { "Cache-Control": "no-store" } });
  }

  let orderId = (url.searchParams.get("orderId") ?? "").replace(/\D/g, "");
  const name = (url.searchParams.get("name") ?? "").trim().replace(/^#/, "");
  if (!orderId && !name) return Response.json({ ok: false, error: "Pass ?orderId=<numeric>, ?name=<order name>, or ?action=rehold-all." }, { status: 400 });

  // Resolve an order name (e.g. 397534K) to its numeric id.
  if (!orderId && name) {
    const res = await fetch(`https://${shop}/admin/api/2025-10/graphql.json`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": session.accessToken },
      body: JSON.stringify({ query: `#graphql
        query FindOrder($q: String!) { orders(first: 1, query: $q) { nodes { id name } } }`, variables: { q: `name:${name}` } }),
    });
    const json = await res.json().catch(() => ({})) as { data?: { orders?: { nodes?: Array<{ id?: string; name?: string }> } } };
    const node = json.data?.orders?.nodes?.[0];
    if (!node?.id) return Response.json({ ok: false, error: `No order matching "${name}".` }, { status: 404 });
    orderId = String(node.id).replace(/\D/g, "");
  }

  try {
    const result = await processShopifyOrderCreated(shop, { id: orderId });
    // ALWAYS re-hold too: processShopifyOrderCreated only places the hold when it
    // captures something NEW, so an order that was already reserved but never held
    // (the real failure mode) wouldn't get held by the capture path alone. The
    // re-hold works off the reservation ledger, so it fixes that case.
    const rehold = await reholdUnheldPreorderOrders({ orderId });
    return Response.json({ ok: true, orderId, result, rehold, note: "Reprocessed + re-held off the reservation ledger (idempotent)." }, { headers: { "Cache-Control": "no-store" } });
  } catch (error) {
    return Response.json({ ok: false, orderId, error: error instanceof Error ? error.message : String(error) }, { status: 500 });
  }
};
