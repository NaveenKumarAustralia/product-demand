import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { getPreorderSellingPlanRegistryEntries } from "../preorder/preorder-selling-plan-registry.server";
import { calculatePreorderCapacity, marketFromDestination } from "../preorder/preorder-rules.server";
import { KARMA_EAST_PREORDER_PLAN_PREFIX } from "../preorder/preorder-shopify-order-normalize";

// Admin diagnostic: for a list of Shopify order numbers, tell the TRUTH about how
// each pre-order line was recorded in the app — so we can prove/disprove "ghost
// allocations" (an order placed as a pre-order that has NO reservation counting
// against the incoming batch, hence risks overselling what's coming in).
//
//   GET /api/preorder-order-audit?orders=385992,386250,386223,386220
//
// For each order it shows every line: does a PreorderReservation exist (the app's
// "-1")? which batch? still unfulfilled? Plus a per-batch/per-variant capacity
// reconciliation (incoming vs reserved vs oversold) and any logged allocation
// failures. Read-only — writes nothing.
const API_VERSION = "2025-10";
const numericId = (v: string) => String(v ?? "").split("/").pop()?.replace(/[^0-9]/g, "") ?? "";
function variantIdCandidates(value: string): string[] {
  const raw = String(value ?? "").trim();
  const n = raw.replace(/[^0-9]/g, "");
  return Array.from(new Set([raw, n, n ? `gid://shopify/ProductVariant/${n}` : ""].filter(Boolean)));
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const url = new URL(request.url);
  const tokens = Array.from(new Set(
    (url.searchParams.get("orders") ?? "").split(/[\s,]+/).map((s) => s.trim().replace(/^#/, "")).filter(Boolean),
  ));
  const batchIdsParam = Array.from(new Set(
    (url.searchParams.get("batch") ?? url.searchParams.get("batches") ?? "").split(/[\s,]+/).map((s) => Number(s.replace(/[^0-9]/g, ""))).filter((n) => Number.isFinite(n) && n > 0),
  ));

  const session = await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true, accessToken: true } });
  if (!session?.shop || !session.accessToken) return Response.json({ ok: false, error: "No offline Shopify session." }, { status: 500 });
  const { shop, accessToken } = session;
  const gql = async (query: string, variables: Record<string, unknown>) => {
    const res = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
      method: "POST", headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": accessToken },
      body: JSON.stringify({ query, variables }),
    });
    return res.json() as Promise<any>;
  };

  // BATCH MODE: ?batch=1281 (or batch=1281,954) → reconcile EVERY size in that
  // production batch at once (incoming vs reserved vs available/oversold), no order
  // numbers needed. Also pulls Shopify's LIVE available per size so you can see how
  // the app's reservations line up with Shopify's (often-negative) inventory.
  if (batchIdsParam.length) {
    const bufferRows = await prisma.preorderBatchSetting.findMany({ where: { supplierOrderId: { in: batchIdsParam } }, select: { supplierOrderId: true, safetyBufferPercent: true, safetyBufferQty: true, enabled: true, shipDate: true } });
    const bufferBy = new Map(bufferRows.map((b) => [b.supplierOrderId, b]));
    const orders = await prisma.supplierOrder.findMany({
      where: { id: { in: batchIdsParam } },
      select: { id: true, productTitle: true, destination: true, status: true, eta: true, lines: { select: { variantId: true, variantTitle: true, qtyOrdered: true, qtyReceived: true } } },
    });
    // Shopify's LIVE available per variant (total across locations). Negative =
    // oversold (bought past on-hand while inventory policy is "continue selling").
    const allVariantGids = Array.from(new Set(
      orders.flatMap((o) => o.lines.map((l) => { const n = numericId(l.variantId); return n ? `gid://shopify/ProductVariant/${n}` : ""; }).filter(Boolean)),
    ));
    // Shopify's live inventory per variant: available, on-hand and committed
    // (summed across locations). on_hand = physically counted; committed = in
    // unfulfilled orders; available = on_hand − committed. Negative available =
    // oversold. This is what lets us see EXACTLY why the number is what it is.
    const shopInv = new Map<string, { available: number; onHand: number; committed: number }>();
    for (let i = 0; i < allVariantGids.length; i += 50) {
      const chunk = allVariantGids.slice(i, i + 50);
      const j = await gql(`query VInv($ids: [ID!]!) {
        nodes(ids: $ids) { ... on ProductVariant { id inventoryQuantity
          inventoryItem { inventoryLevels(first: 20) { nodes { quantities(names: ["on_hand","committed","available"]) { name quantity } } } } } }
      }`, { ids: chunk }).catch(() => null);
      for (const n of j?.data?.nodes ?? []) {
        const num = numericId(String(n?.id ?? "")); if (!num) continue;
        let onHand = 0, committed = 0, avail = 0;
        for (const lvl of n?.inventoryItem?.inventoryLevels?.nodes ?? []) {
          for (const q of lvl?.quantities ?? []) {
            if (q?.name === "on_hand") onHand += Number(q.quantity ?? 0);
            else if (q?.name === "committed") committed += Number(q.quantity ?? 0);
            else if (q?.name === "available") avail += Number(q.quantity ?? 0);
          }
        }
        // Fall back to inventoryQuantity for available if levels didn't return it.
        shopInv.set(num, { available: avail || Number(n?.inventoryQuantity ?? 0), onHand, committed });
      }
    }
    // Pre-order reservations for this batch, split by status per size, so you can
    // see how many are STILL WAITING vs already SHIPPED (fulfilled) vs cancelled.
    const statusGroups = await prisma.preorderReservation.groupBy({
      by: ["variantTitle", "status"],
      where: { supplierOrderId: { in: batchIdsParam } },
      _sum: { quantity: true },
    });
    const statusBySize = new Map<string, Record<string, number>>();
    for (const g of statusGroups) {
      const sz = String(g.variantTitle ?? "").trim(); if (!sz) continue;
      const rec = statusBySize.get(sz) ?? {};
      rec[g.status] = (rec[g.status] ?? 0) + (g._sum.quantity ?? 0);
      statusBySize.set(sz, rec);
    }
    const batches = [];
    for (const o of orders) {
      const buf = bufferBy.get(o.id);
      const sizes = [];
      for (const l of o.lines) {
        const incoming = Math.max(0, l.qtyOrdered - l.qtyReceived);
        const reservedAgg = await prisma.preorderReservation.aggregate({
          where: { supplierOrderId: o.id, variantId: { in: variantIdCandidates(l.variantId) }, status: "reserved" },
          _sum: { quantity: true },
        });
        const reservedQty = reservedAgg._sum.quantity ?? 0;
        const cap = calculatePreorderCapacity({ confirmedIncomingQty: incoming, reservedQty, safetyBufferPercent: buf?.safetyBufferPercent ?? 0, safetyBufferQty: buf?.safetyBufferQty ?? null });
        const inv = shopInv.get(numericId(l.variantId)) ?? null;
        const byStatus = statusBySize.get(String(l.variantTitle ?? "").trim()) ?? {};
        const shippedQty = byStatus["fulfilled"] ?? 0;
        const totalPreordersEver = reservedQty + shippedQty + (byStatus["released"] ?? 0);
        sizes.push({
          size: l.variantTitle, qtyOrdered: l.qtyOrdered, qtyReceived: l.qtyReceived, incomingRemaining: incoming,
          reservedWaiting: reservedQty, shippedAlready: shippedQty, cancelled: byStatus["released"] ?? 0, totalPreordersEver,
          availableToPreorder: cap.availableToPreorder, oversoldBy: cap.overallocatedBy,
          shopifyAvailable: inv?.available ?? null, shopifyOnHand: inv?.onHand ?? null, shopifyCommitted: inv?.committed ?? null,
        });
      }
      batches.push({ batchId: o.id, productTitle: o.productTitle, destination: o.destination, status: o.status, enabled: buf?.enabled ?? null, shipDate: buf?.shipDate ?? o.eta ?? null, sizes });
    }
    const oversold = batches.flatMap((b) => b.sizes.filter((s: any) => s.oversoldBy > 0).map((s: any) => `batch ${b.batchId} ${b.productTitle ?? ""} ${s.size ?? ""} oversold by ${s.oversoldBy}`.trim()));
    return Response.json({
      ok: true,
      summary: { batchesChecked: batchIdsParam.length, oversoldSizes: oversold.length, oversold, note: "reservedQty = pre-orders committed against this size. oversoldBy>0 means more reserved than (incoming − buffer)." },
      batches,
    }, { headers: { "Cache-Control": "no-store" } });
  }

  // HOLD-CHECK MODE: ?holds=1178 (optionally &size=3XL) → for every still-waiting
  // reserved order in that batch, show its Shopify fulfilment-order hold status +
  // tags, so we can see WHICH orders aren't on hold (and therefore why Shopify's
  // "committed" is higher than expected). Read-only.
  const holdsBatch = Number((url.searchParams.get("holds") ?? "").replace(/[^0-9]/g, ""));
  if (Number.isFinite(holdsBatch) && holdsBatch > 0) {
    const sizeFilter = (url.searchParams.get("size") ?? "").trim().toLowerCase();
    const reservations = await prisma.preorderReservation.findMany({
      where: { supplierOrderId: holdsBatch, status: "reserved" },
      select: { shopifyOrderId: true, shopifyOrderName: true, shopifyLineItemId: true, variantTitle: true, quantity: true, reservedAt: true },
      orderBy: { reservedAt: "asc" },
    });
    const filtered = sizeFilter ? reservations.filter((r) => String(r.variantTitle ?? "").trim().toLowerCase() === sizeFilter) : reservations;
    // Group by order (an order may have several pre-order lines).
    const byOrder = new Map<string, typeof filtered>();
    for (const r of filtered) { const a = byOrder.get(r.shopifyOrderId) ?? []; a.push(r); byOrder.set(r.shopifyOrderId, a); }

    const rows = [];
    for (const [orderId, rs] of byOrder) {
      const j = await gql(`query OrderHolds($id: ID!) {
        order(id: $id) { id name tags displayFulfillmentStatus
          fulfillmentOrders(first: 25) { nodes { id status
            lineItems(first: 50) { nodes { remainingQuantity lineItem { id } } } } } }
      }`, { id: `gid://shopify/Order/${orderId}` }).catch(() => null);
      const order = j?.data?.order;
      const preLineIds = new Set(rs.map((r) => r.shopifyLineItemId));
      const fos = (order?.fulfillmentOrders?.nodes ?? []).map((fo: any) => {
        const lineIds = (fo.lineItems?.nodes ?? []).map((n: any) => numericId(String(n?.lineItem?.id ?? ""))).filter(Boolean);
        const hasPreLine = lineIds.some((id: string) => preLineIds.has(id));
        return { status: fo.status, hasPreorderLine: hasPreLine, lineIds };
      });
      const preFos = fos.filter((f: any) => f.hasPreorderLine);
      const heldPre = preFos.filter((f: any) => f.status === "ON_HOLD").length;
      const openPre = preFos.filter((f: any) => f.status === "OPEN").length;
      const tags: string[] = order?.tags ?? [];
      rows.push({
        order: order?.name ?? orderId, orderId,
        sizes: rs.map((r) => r.variantTitle),
        reservedAt: rs[0]?.reservedAt,
        fulfillmentStatus: order?.displayFulfillmentStatus ?? null,
        hasPreorderHoldTag: tags.includes("pre-order-hold"),
        hasPreorderTag: tags.some((t) => t === "pre-order" || t.startsWith("pre-order-")),
        preorderFulfillmentOrders: preFos.map((f: any) => f.status),
        verdict: !order ? "order not found in Shopify"
          : preFos.length === 0 ? "🟠 no fulfilment order holds the pre-order line (likely already picked/split or shipped)"
          : openPre > 0 && heldPre === 0 ? "🔴 NOT held — pre-order line is OPEN (would be picked as normal; hold missing)"
          : openPre > 0 ? "🟠 partially held — some pre-order fulfilment orders OPEN"
          : "🟢 held correctly",
      });
    }
    const notHeld = rows.filter((r) => String(r.verdict).startsWith("🔴") || String(r.verdict).startsWith("🟠"));
    return Response.json({
      ok: true,
      summary: { batch: holdsBatch, size: sizeFilter || "(all sizes)", ordersChecked: rows.length, notFullyHeld: notHeld.length,
        note: "🟢 = pre-order line is on hold (correct). 🔴/🟠 = hold missing or partial → shows as Shopify 'committed' and can be picked early. Likely causes: order placed before hold logic, a scope gap when it landed, or the line was split/edited." },
      orders: rows,
    }, { headers: { "Cache-Control": "no-store" } });
  }

  if (!tokens.length) return Response.json({ ok: false, error: "Pass ?orders=385992,386250,… OR ?batch=1281 OR ?holds=1281&size=3XL (hold status per order)." }, { status: 400 });

  // Live batch map: variant → the enabled+activated batch(es) it belongs to, plus
  // that batch's incoming/received/buffer for the capacity math.
  const [enabledSettings, registry, allBatchSettings] = await Promise.all([
    prisma.preorderBatchSetting.findMany({ where: { enabled: true }, select: { supplierOrderId: true, safetyBufferPercent: true, safetyBufferQty: true } }),
    getPreorderSellingPlanRegistryEntries(shop),
    prisma.preorderBatchSetting.findMany({ select: { supplierOrderId: true, safetyBufferPercent: true, safetyBufferQty: true } }),
  ]);
  const activatedIds = new Set(registry.map((r) => r.supplierOrderId));
  const liveIds = enabledSettings.map((s) => s.supplierOrderId).filter((id) => activatedIds.has(id));
  const bufferByBatch = new Map(allBatchSettings.map((b) => [b.supplierOrderId, { percent: b.safetyBufferPercent, qty: b.safetyBufferQty }]));
  const liveBatches = liveIds.length
    ? await prisma.supplierOrder.findMany({
        where: { id: { in: liveIds }, status: "open", destination: { in: ["send_to_au", "send_to_usa"] } },
        select: { id: true, destination: true, productTitle: true, lines: { select: { variantId: true, variantTitle: true, qtyOrdered: true, qtyReceived: true } } },
      })
    : [];
  const variantToBatch = new Map<string, Array<{ batchId: number; market: string | null }>>();
  for (const b of liveBatches) for (const l of b.lines) {
    const n = numericId(l.variantId);
    if (!n) continue;
    const arr = variantToBatch.get(n) ?? [];
    arr.push({ batchId: b.id, market: marketFromDestination(b.destination) });
    variantToBatch.set(n, arr);
  }

  const capacityCache = new Map<string, unknown>();
  async function capacityFor(batchId: number, variantId: string) {
    const vnum = numericId(variantId);
    const key = `${batchId}:${vnum}`;
    if (capacityCache.has(key)) return capacityCache.get(key);
    const batch = liveBatches.find((b) => b.id === batchId);
    const line = batch?.lines.find((l) => numericId(l.variantId) === vnum);
    const incoming = line ? Math.max(0, line.qtyOrdered - line.qtyReceived) : 0;
    const reservedAgg = await prisma.preorderReservation.aggregate({
      where: { supplierOrderId: batchId, variantId: { in: variantIdCandidates(variantId) }, status: "reserved" },
      _sum: { quantity: true },
    });
    const reservedQty = reservedAgg._sum.quantity ?? 0;
    const buf = bufferByBatch.get(batchId);
    const cap = calculatePreorderCapacity({ confirmedIncomingQty: incoming, reservedQty, safetyBufferPercent: buf?.percent ?? 0, safetyBufferQty: buf?.qty ?? null });
    const out = {
      batchId, variantTitle: line?.variantTitle ?? null, batchLive: !!batch, hasLineInBatch: !!line,
      incomingRemaining: incoming, reservedQty, safetyBufferQty: cap.safetyBufferQty,
      availableToPreorder: cap.availableToPreorder, oversoldBy: cap.overallocatedBy,
    };
    capacityCache.set(key, out);
    return out;
  }

  const report: any[] = [];
  const batchesTouched = new Set<string>(); // `${batchId}:${vnum}`

  for (const token of tokens) {
    const search = await gql(
      `query FindOrder($q: String!) {
        orders(first: 3, query: $q) {
          nodes {
            id name email cancelledAt displayFulfillmentStatus
            lineItems(first: 50) { nodes { id title quantity unfulfilledQuantity variant { id title } sellingPlan { name } } }
          }
        }
      }`,
      { q: `name:${token}` },
    ).catch((e) => ({ __err: String(e) }));
    if ((search as any).__err || search?.errors?.length) {
      report.push({ token, error: (search as any).__err || search.errors.map((e: any) => e.message).join("; ") });
      continue;
    }
    const orderNodes: any[] = search?.data?.orders?.nodes ?? [];
    // Prefer an exact name match (with or without #) if the search returned several.
    const order = orderNodes.find((o) => numericId(o.name) === token || String(o.name ?? "").replace(/^#/, "") === token) ?? orderNodes[0];
    if (!order) { report.push({ token, error: "Order not found in Shopify (check the number)." }); continue; }

    const orderIdNum = numericId(order.id);
    const dbRes = await prisma.preorderReservation.findMany({
      where: { OR: [{ shopifyOrderId: orderIdNum }, { shopifyOrderName: order.name }, { shopifyOrderName: `#${token}` }, { shopifyOrderName: token }] },
      select: { shopifyLineItemId: true, supplierOrderId: true, variantId: true, variantTitle: true, quantity: true, status: true, reservedAt: true, releasedAt: true, fulfilledAt: true },
    });
    const resByLine = new Map<string, typeof dbRes>();
    for (const r of dbRes) { const arr = resByLine.get(r.shopifyLineItemId) ?? []; arr.push(r); resByLine.set(r.shopifyLineItemId, arr); }

    const failures = await prisma.activityLog.findMany({
      where: { action: "preorder_allocation_failed", OR: [{ entityId: orderIdNum }, { entityName: order.name }] },
      select: { createdAt: true, toValue: true }, orderBy: { createdAt: "desc" }, take: 5,
    }).catch(() => []);

    const lines: any[] = [];
    for (const li of order.lineItems?.nodes ?? []) {
      const lineId = numericId(li.id);
      const vnum = numericId(String(li.variant?.id ?? ""));
      const planName: string = li.sellingPlan?.name ?? "";
      const isPlan = planName.startsWith(KARMA_EAST_PREORDER_PLAN_PREFIX);
      const inLiveBatch = variantToBatch.get(vnum) ?? [];
      const looksPreorder = isPlan || inLiveBatch.length > 0;
      const res = resByLine.get(lineId) ?? [];
      const reservedQty = res.filter((r) => r.status === "reserved").reduce((s, r) => s + r.quantity, 0);
      const unfulfilled = Number(li.unfulfilledQuantity ?? li.quantity ?? 0);

      // Capacity snapshot for the batch(es) this line touches (reserved ones first,
      // else the live-batch mapping).
      const batchIds = res.length ? Array.from(new Set(res.map((r) => r.supplierOrderId))) : inLiveBatch.map((b) => b.batchId);
      const caps = [];
      for (const bid of batchIds) { const c = await capacityFor(bid, String(li.variant?.id ?? "")); caps.push(c); batchesTouched.add(`${bid}:${vnum}`); }

      let verdict: string;
      if (!looksPreorder) verdict = "not a pre-order line";
      else if (reservedQty >= li.quantity) verdict = "✅ reserved (counts against the batch)";
      else if (res.some((r) => r.status === "released")) verdict = "released/cancelled";
      else if (res.some((r) => r.status === "fulfilled")) verdict = "fulfilled";
      else if (unfulfilled <= 0) verdict = "shipped already (was in stock at the time)";
      else verdict = "🚨 GHOST — placed as a pre-order but NO reservation (not counted against the batch)";

      lines.push({
        line: li.title, size: li.variant?.title ?? null, qty: li.quantity, unfulfilledQty: unfulfilled,
        soldVia: isPlan ? "pre-order button (selling plan)" : (inLiveBatch.length ? "no-plan (quick-add / Shop Pay / express)" : "n/a"),
        reservedQty, reservationStatuses: res.map((r) => r.status),
        batchCapacity: caps, verdict,
      });
    }

    report.push({
      token, name: order.name, orderId: orderIdNum, cancelled: !!order.cancelledAt,
      fulfillmentStatus: order.displayFulfillmentStatus,
      allocationFailuresLogged: failures.map((f) => ({ at: f.createdAt, reason: f.toValue })),
      lines,
    });
  }

  // Batch-level reconciliation for every (batch, variant) touched.
  const batchReconcile = [];
  for (const key of batchesTouched) {
    const [bidStr, vnum] = key.split(":");
    const c = await capacityFor(Number(bidStr), vnum);
    batchReconcile.push(c);
  }

  const ghosts = report.flatMap((o) => (o.lines ?? []).filter((l: any) => String(l.verdict).startsWith("🚨")).map((l: any) => `${o.name} — ${l.line} ${l.size ?? ""}`.trim()));

  return Response.json({
    ok: true,
    summary: {
      ordersChecked: tokens.length,
      ghostsFound: ghosts.length,
      ghosts,
      oversoldBatches: batchReconcile.filter((c: any) => c.oversoldBy > 0),
      note: "A pre-order is recorded as a PreorderReservation in the app (NOT as a -1 on the restock sheet or on Shopify inventory). 'GHOST' = the order was placed as a pre-order but no reservation exists, so it is NOT counted against the incoming batch — that is the real overselling risk.",
    },
    orders: report,
    batchReconcile,
  }, { headers: { "Cache-Control": "no-store" } });
};
