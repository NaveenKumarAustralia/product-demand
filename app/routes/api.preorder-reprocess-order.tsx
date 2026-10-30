import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { processShopifyOrderCreated } from "../preorder/preorder-shopify-order.server";
import { reholdUnheldPreorderOrders, getAvailableAtLocation } from "../preorder/preorder-fulfillment.server";
import { getPreorderSellingPlanRegistryEntries } from "../preorder/preorder-selling-plan-registry.server";
import { getPreorderLocationSettings, locationForMarket } from "../preorder/preorder-locations.server";
import { calculatePreorderCapacity, isPreorderEligibleStatus, marketFromDestination } from "../preorder/preorder-rules.server";

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

  // Debug (no changes): dump the order's lines so we can see why capture skips a
  // line — whether variant.id is present, what selling plan (if any) it carries,
  // quantity, and fulfillment-order status.
  if (url.searchParams.get("debug") === "1") {
    const res = await fetch(`https://${shop}/admin/api/2025-10/graphql.json`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": session.accessToken },
      body: JSON.stringify({ query: `#graphql
        query ReprocessDebug($id: ID!) {
          order(id: $id) {
            id name displayFulfillmentStatus
            shippingAddress { countryCodeV2 } billingAddress { countryCodeV2 }
            lineItems(first: 50) { nodes { id title quantity currentQuantity sku variant { id title } sellingPlan { name } } }
            fulfillmentOrders(first: 25) { nodes { id status } }
          }
        }`, variables: { id: `gid://shopify/Order/${orderId}` } }),
    });
    const json = await res.json().catch(() => ({}));
    return Response.json({ ok: true, debug: true, orderId, order: (json as { data?: { order?: unknown } }).data?.order ?? null }, { headers: { "Cache-Control": "no-store" } });
  }

  // Capture trace (read-only): replay the capture decision for the order's lines
  // and report exactly where/why it stops — no reservation is created.
  if (url.searchParams.get("debug") === "capture") {
    const num = (s: unknown) => String(s ?? "").replace(/\D/g, "");
    const res = await fetch(`https://${shop}/admin/api/2025-10/graphql.json`, {
      method: "POST", headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": session.accessToken },
      body: JSON.stringify({ query: `#graphql
        query CapTrace($id: ID!) { order(id: $id) { name shippingAddress { countryCodeV2 } billingAddress { countryCodeV2 }
          lineItems(first: 50) { nodes { id quantity variant { id } sellingPlan { name } } } } }`, variables: { id: `gid://shopify/Order/${orderId}` } }),
    });
    const oj = await res.json().catch(() => ({})) as { data?: { order?: { shippingAddress?: { countryCodeV2?: string }; billingAddress?: { countryCodeV2?: string }; lineItems?: { nodes?: Array<{ id: string; quantity: number; variant?: { id?: string }; sellingPlan?: { name?: string } }> } } } };
    const order = oj.data?.order;
    const country = String(order?.shippingAddress?.countryCodeV2 || order?.billingAddress?.countryCodeV2 || "").toUpperCase();
    const market = country === "US" ? "USA" : "AU";
    const [registry, enabledSettings, locations] = await Promise.all([
      getPreorderSellingPlanRegistryEntries(shop),
      prisma.preorderBatchSetting.findMany({ where: { enabled: true }, select: { supplierOrderId: true, safetyBufferPercent: true, safetyBufferQty: true } }),
      getPreorderLocationSettings(),
    ]);
    const activated = new Set(registry.map((r) => r.supplierOrderId));
    const liveIds = enabledSettings.filter((s) => activated.has(s.supplierOrderId)).map((s) => s.supplierOrderId);
    const settingBy = new Map(enabledSettings.map((s) => [s.supplierOrderId, s]));
    const batches = liveIds.length ? await prisma.supplierOrder.findMany({ where: { id: { in: liveIds }, status: "open" }, select: { id: true, destination: true, supplierStatus: true, lines: { select: { variantId: true, qtyOrdered: true, qtyReceived: true } } } }) : [];
    const reserved = await prisma.preorderReservation.groupBy({ by: ["supplierOrderId", "variantId"], where: { supplierOrderId: { in: liveIds }, status: "reserved" }, _sum: { quantity: true } });
    const resBy = new Map(reserved.map((r) => [`${r.supplierOrderId}:${num(r.variantId)}`, r._sum.quantity ?? 0]));
    const token = session.accessToken;
    const lineTrace = [];
    for (const l of order?.lineItems?.nodes ?? []) {
      const vnum = num(l.variant?.id);
      const hits = batches.filter((b) => b.lines.some((ln) => num(ln.variantId) === vnum)).map((b) => ({ batchId: b.id, destination: b.destination, market: marketFromDestination(b.destination), status: b.supplierStatus, eligible: isPreorderEligibleStatus(b.supplierStatus) }));
      let available: number | null = null; let capacity: unknown = null; let locId: string | null = null;
      const hit = hits.find((h) => h.market === market) ?? hits[0];
      if (hit) {
        locId = locationForMarket(locations, hit.market as "AU" | "USA");
        if (locId) { try { available = await getAvailableAtLocation(shop, token, l.variant?.id ?? "", locId); } catch (e) { available = -999; } }
        const b = batches.find((x) => x.id === hit.batchId)!; const bl = b.lines.find((ln) => num(ln.variantId) === vnum)!; const st = settingBy.get(hit.batchId);
        capacity = calculatePreorderCapacity({ confirmedIncomingQty: Math.max(0, (bl.qtyOrdered ?? 0) - (bl.qtyReceived ?? 0)), reservedQty: resBy.get(`${hit.batchId}:${vnum}`) ?? 0, safetyBufferPercent: st?.safetyBufferPercent ?? 0, safetyBufferQty: st?.safetyBufferQty ?? null });
      }
      lineTrace.push({ variantId: l.variant?.id ?? null, vnum, hasPlan: Boolean(l.sellingPlan?.name), qty: l.quantity, hitCount: hits.length, hits, chosenMarket: market, locationId: locId, available, oosGatePass: available != null && available <= 0, capacity });
    }
    return Response.json({ ok: true, debug: "capture", orderId, market, liveBatchIds: liveIds, lines: lineTrace }, { headers: { "Cache-Control": "no-store" } });
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
