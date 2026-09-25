import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { getOfflineToken, addOrderTags, applyPreorderHoldPolicy } from "../preorder/preorder-fulfillment.server";

// Admin backfill: re-apply the pre-order fulfillment-hold policy to orders that
// were placed but never got their pre-order line(s) held/split (the plan-path
// hold bug fixed Sep 25 2026 — a mixed order outside the combine window, or a
// fully-pre-order plan order, was reserved + tagged but never actually held).
//
//   GET /api/preorder-fix-holds?orders=388695            → fix these order(s)
//   GET /api/preorder-fix-holds?orders=388695&dryRun=1   → show what WOULD happen
//   GET /api/preorder-fix-holds                          → auto-scan all still-
//                                                          reserved orders (dry run
//                                                          unless &apply=1)
//
// The pre-order line item IDs come from PreorderReservation rows (status
// reserved, not yet readyAt/fulfilled), so it only ever holds the lines the app
// itself allocated as pre-orders. applyPreorderHoldPolicy decides combine (hold
// whole order) vs line-only (split + hold just the pre-order line). Idempotent:
// re-running an already-correct order is a no-op.
const numericId = (v: string) => String(v ?? "").split("/").pop()?.replace(/[^0-9]/g, "") ?? "";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const url = new URL(request.url);
  const explicit = Array.from(new Set(
    (url.searchParams.get("orders") ?? "").split(/[\s,]+/).map((s) => s.trim().replace(/^#/, "").replace(/[^0-9]/g, "")).filter(Boolean),
  ));
  // With explicit orders, default to APPLYING; on a full auto-scan default to a
  // dry run so nobody nukes holds across the whole store by accident.
  const dryRun = url.searchParams.get("dryRun") === "1" || (!explicit.length && url.searchParams.get("apply") !== "1");

  const session = await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true } });
  if (!session?.shop) return Response.json({ ok: false, error: "No offline Shopify session." }, { status: 500 });
  const shop = session.shop;
  const token = await getOfflineToken(shop);
  if (!token) return Response.json({ ok: false, error: "No offline Shopify token." }, { status: 500 });

  // Active pre-order reservations (still awaiting stock: reserved, not released,
  // not readied, not fulfilled) — these are the lines that SHOULD be on hold.
  // IMPORTANT: the user types the order NAME number (e.g. 388695, from "#388695K"),
  // but reservations are keyed by Shopify's INTERNAL order id. So match either the
  // internal id OR the order name (which contains the typed number).
  const orderMatch = explicit.length
    ? { OR: explicit.flatMap((t) => [{ shopifyOrderId: t }, { shopifyOrderName: { contains: t } }]) }
    : {};
  const where = {
    shop,
    status: "reserved",
    readyAt: null,
    releasedAt: null,
    fulfilledAt: null,
    ...orderMatch,
  } as const;
  const rows = await prisma.preorderReservation.findMany({
    where,
    select: { shopifyOrderId: true, shopifyOrderName: true, shopifyLineItemId: true, supplierOrderId: true, expectedShipDate: true },
  });

  // Diagnostic: if explicit orders were asked for but nothing active matched, show
  // ANY reservation rows for those orders (any status) so we can see why.
  let diagnostic: Array<Record<string, unknown>> | undefined;
  if (explicit.length && !rows.length) {
    const anyRows = await prisma.preorderReservation.findMany({
      where: { shop, OR: explicit.flatMap((t) => [{ shopifyOrderId: t }, { shopifyOrderName: { contains: t } }]) },
      select: { shopifyOrderId: true, shopifyOrderName: true, shopifyLineItemId: true, variantTitle: true, supplierOrderId: true, status: true, readyAt: true, releasedAt: true, fulfilledAt: true },
    });
    diagnostic = anyRows.map((r) => ({ ...r }));
  }

  // Group reservations by order.
  const byOrder = new Map<string, { name: string | null; lineIds: Set<string>; batchIds: Set<number>; earliestMs: number | null }>();
  for (const r of rows) {
    const oid = numericId(r.shopifyOrderId);
    if (!oid) continue;
    const g = byOrder.get(oid) ?? { name: r.shopifyOrderName ?? null, lineIds: new Set<string>(), batchIds: new Set<number>(), earliestMs: null };
    if (r.shopifyLineItemId) g.lineIds.add(numericId(r.shopifyLineItemId));
    if (r.supplierOrderId) g.batchIds.add(r.supplierOrderId);
    const t = r.expectedShipDate ? r.expectedShipDate.getTime() : NaN;
    if (Number.isFinite(t) && (g.earliestMs === null || t < g.earliestMs)) g.earliestMs = t;
    byOrder.set(oid, g);
  }

  const results: Array<Record<string, unknown>> = [];
  for (const [oid, g] of byOrder) {
    const lineIds = Array.from(g.lineIds).filter(Boolean);
    if (!lineIds.length) { results.push({ order: oid, name: g.name, skipped: "no pre-order line ids on the reservations" }); continue; }
    if (dryRun) { results.push({ order: oid, name: g.name, wouldHoldLines: lineIds, batches: Array.from(g.batchIds), dryRun: true }); continue; }
    try {
      const { wholeOrderHeld } = await applyPreorderHoldPolicy(shop, token, oid, lineIds, g.earliestMs);
      if (wholeOrderHeld) await addOrderTags(shop, token, oid, ["pre-order-hold"]);
      results.push({ order: oid, name: g.name, held: true, wholeOrderHeld, lines: lineIds.length, batches: Array.from(g.batchIds) });
    } catch (error) {
      results.push({ order: oid, name: g.name, error: error instanceof Error ? error.message : String(error) });
    }
  }

  return Response.json(
    { ok: true, shop, dryRun, scannedOrders: byOrder.size, results, ...(diagnostic ? { diagnostic, diagnosticNote: "No ACTIVE (reserved & awaiting stock) pre-order lines matched. These are all reservation rows found for the order — check status/readyAt/releasedAt/fulfilledAt to see why. Tip: you can also pass the internal order id (from the Shopify admin URL /orders/<id>)." } : {}), note: explicit.length ? "Applied to the given orders (add &dryRun=1 to preview)." : (dryRun ? "Dry run of ALL still-reserved orders — add &apply=1 to actually place the holds." : "Applied to ALL still-reserved orders.") },
    { headers: { "Cache-Control": "no-store" } },
  );
};
