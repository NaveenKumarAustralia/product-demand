import prisma from "../db.server";
import { getPreorderLocationSettings, locationForMarket } from "./preorder-locations.server";
import {
  getOfflineToken,
  getAvailableAtLocation,
  releaseOrderPreorderHolds,
  addOrderTags,
  removeOrderTags,
  holdOrderOpenFulfillmentOrders,
  resplitHeldOrderPreorderLines,
} from "./preorder-fulfillment.server";
import { getPreorderCombineWindowDays } from "./preorder-storefront-settings.server";
import type { PreorderMarket } from "./preorder-rules.server";

// Backfill for the "combine mixed orders" feature: hold the in-stock items of
// EXISTING open pre-order orders (whose pre-order is due within the combine
// window) so they ship together when the batch lands. Idempotent — an order
// whose in-stock lines are already held (no OPEN fulfilment orders) is a no-op,
// and shipped lines (CLOSED) are never touched. The existing stock-aware
// release then releases everything when the inventory is loaded into Shopify.
export async function combineExistingPreorderOrders(): Promise<{ scannedOrders: number; heldOrders: number; errors: number; skippedNoScope: boolean; windowDays: number }> {
  const windowDays = await getPreorderCombineWindowDays();
  if (windowDays <= 0) return { scannedOrders: 0, heldOrders: 0, errors: 0, skippedNoScope: false, windowDays };
  const pending = await prisma.preorderReservation.findMany({
    where: { status: "reserved", readyAt: null },
    select: { shop: true, shopifyOrderId: true, expectedShipDate: true },
  });
  // Earliest promised dispatch per (shop, order).
  const orders = new Map<string, { shop: string; orderId: string; earliest: number | null }>();
  for (const r of pending) {
    if (!r.shopifyOrderId) continue;
    const key = `${r.shop}::${r.shopifyOrderId}`;
    const t = r.expectedShipDate ? new Date(r.expectedShipDate).getTime() : NaN;
    const cur = orders.get(key) ?? { shop: r.shop, orderId: r.shopifyOrderId, earliest: null };
    if (Number.isFinite(t) && (cur.earliest === null || t < cur.earliest)) cur.earliest = t;
    orders.set(key, cur);
  }
  const cutoff = Date.now() + windowDays * 86400000;
  const tokenByShop = new Map<string, string | null>();
  let scannedOrders = 0, heldOrders = 0, errors = 0, skippedNoScope = false;
  for (const { shop, orderId, earliest } of orders.values()) {
    if (earliest === null || earliest > cutoff) continue; // not due within the window
    scannedOrders += 1;
    if (!tokenByShop.has(shop)) tokenByShop.set(shop, await getOfflineToken(shop));
    const token = tokenByShop.get(shop);
    if (!token) { skippedNoScope = true; continue; }
    try {
      const held = await holdOrderOpenFulfillmentOrders(shop, token, orderId, "Held to ship with the pre-order item in this order (combine window)");
      if (held > 0) {
        await addOrderTags(shop, token, orderId, ["pre-order-hold"]);
        heldOrders += 1;
        console.log(`[preorder combine backfill] ${shop} order ${orderId}: held ${held} in-stock fulfilment order(s).`);
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      if (/access denied|not approved|scope/i.test(message)) skippedNoScope = true;
      console.warn(`[preorder combine backfill] ${shop} order ${orderId} failed:`, message);
      errors += 1;
    }
  }
  return { scannedOrders, heldOrders, errors, skippedNoScope, windowDays };
}

// One-off backfill: fix EXISTING held pre-order orders so only the pre-order
// line stays held and the in-stock lines ship now (older orders were held whole).
// Dry-run by default (counts pending pre-order orders); apply does the re-split.
export async function resplitHeldPreorderOrders(opts: { apply?: boolean } = {}): Promise<{ scanned: number; fixed: number; errors: number; skippedNoScope: boolean; applied: boolean; fixedOrders: string[]; errorDetails: Array<{ order: string; error: string }> }> {
  const pending = await prisma.preorderReservation.findMany({
    where: { status: "reserved", readyAt: null },
    select: { shop: true, shopifyOrderId: true, shopifyOrderName: true, shopifyLineItemId: true },
  });
  const byOrder = new Map<string, { shop: string; orderId: string; name: string; lineIds: string[] }>();
  for (const r of pending) {
    if (!r.shopifyOrderId) continue;
    const key = `${r.shop}::${r.shopifyOrderId}`;
    const cur = byOrder.get(key) ?? { shop: r.shop, orderId: r.shopifyOrderId, name: r.shopifyOrderName || r.shopifyOrderId, lineIds: [] };
    if (r.shopifyLineItemId) cur.lineIds.push(r.shopifyLineItemId);
    byOrder.set(key, cur);
  }
  const tokenByShop = new Map<string, string | null>();
  let scanned = 0, fixed = 0, errors = 0, skippedNoScope = false;
  const fixedOrders: string[] = [];
  const errorDetails: Array<{ order: string; error: string }> = [];
  for (const { shop, orderId, name, lineIds } of byOrder.values()) {
    if (!lineIds.length) continue;
    scanned += 1;
    if (!opts.apply) continue;
    if (!tokenByShop.has(shop)) tokenByShop.set(shop, await getOfflineToken(shop));
    const token = tokenByShop.get(shop);
    if (!token) { skippedNoScope = true; continue; }
    try {
      const { changed } = await resplitHeldOrderPreorderLines(shop, token, orderId, lineIds);
      if (changed) { fixed += 1; fixedOrders.push(name); }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      if (/access denied|not approved|scope/i.test(message)) skippedNoScope = true;
      console.warn(`[preorder resplit] ${shop} order ${orderId} failed:`, message);
      errors += 1;
      errorDetails.push({ order: name, error: message });
    }
  }
  return { scanned, fixed, errors, skippedNoScope, applied: opts.apply === true, fixedOrders, errorDetails };
}

// When a batch's stock lands in Shopify (available at the market's location covers
// the reservations), release the Shopify fulfilment hold on those orders and tag
// them `pre-order-ready-batch-N` so Pick Pack picks & dispatches them (which fires
// Shopify's native shipping-confirmation email at dispatch). Idempotent via
// PreorderReservation.readyAt; safe to run on a timer. Degrades gracefully if the
// write_orders / fulfilment scopes aren't granted yet (logs, leaves readyAt null,
// retries next cycle once the app is re-authed).
export async function releaseArrivedPreorders(): Promise<{ releasedOrders: number; errors: number; skippedNoScope: boolean }> {
  const pending = await prisma.preorderReservation.findMany({
    where: { status: "reserved", readyAt: null },
    select: { id: true, shop: true, shopifyOrderId: true, supplierOrderId: true, variantId: true, market: true, quantity: true },
  });
  if (!pending.length) return { releasedOrders: 0, errors: 0, skippedNoScope: false };

  const byShop = new Map<string, typeof pending>();
  for (const row of pending) {
    const list = byShop.get(row.shop) ?? [];
    list.push(row);
    byShop.set(row.shop, list);
  }

  let releasedOrders = 0;
  let errors = 0;
  let skippedNoScope = false;

  for (const [shop, rows] of byShop) {
    const token = await getOfflineToken(shop);
    if (!token) { skippedNoScope = true; continue; }
    const locations = await getPreorderLocationSettings();

    // Which (batch, variant) groups have enough stock at their market's location?
    const groupKey = (r: { supplierOrderId: number; variantId: string; market: string }) => `${r.supplierOrderId}::${r.variantId}::${r.market}`;
    const groups = new Map<string, { supplierOrderId: number; variantId: string; market: string; reserved: number }>();
    for (const r of rows) {
      const key = groupKey(r);
      const g = groups.get(key) ?? { supplierOrderId: r.supplierOrderId, variantId: r.variantId, market: r.market, reserved: 0 };
      g.reserved += r.quantity;
      groups.set(key, g);
    }

    const stockCache = new Map<string, number>();
    const arrivedGroups = new Set<string>();
    for (const [key, g] of groups) {
      const locationId = locationForMarket(locations, g.market as PreorderMarket);
      if (!locationId) continue; // market not live
      const stockKey = `${g.variantId}::${locationId}`;
      let available = stockCache.get(stockKey);
      if (available === undefined) {
        try {
          available = await getAvailableAtLocation(shop, token, g.variantId, locationId);
        } catch (error) {
          console.warn(`[preorder release] stock check failed (${shop} ${g.variantId}):`, error instanceof Error ? error.message : error);
          errors += 1;
          available = 0;
        }
        stockCache.set(stockKey, available);
      }
      if (available >= g.reserved && available > 0) arrivedGroups.add(key);
    }

    if (!arrivedGroups.size) continue;

    // An order is releasable only when EVERY one of its pending pre-order lines
    // is in an arrived group (so we never half-release a multi-batch order).
    const byOrder = new Map<string, typeof rows>();
    for (const r of rows) {
      const list = byOrder.get(r.shopifyOrderId) ?? [];
      list.push(r);
      byOrder.set(r.shopifyOrderId, list);
    }

    for (const [orderId, orderRows] of byOrder) {
      const allArrived = orderRows.every((r) => arrivedGroups.has(groupKey(r)));
      if (!allArrived) continue;

      const readyTags = Array.from(new Set(orderRows.map((r) => `pre-order-ready-batch-${r.supplierOrderId}`)));
      try {
        await releaseOrderPreorderHolds(shop, token, orderId);
        await addOrderTags(shop, token, orderId, readyTags);
        await removeOrderTags(shop, token, orderId, ["pre-order-hold"]);
        await prisma.preorderReservation.updateMany({
          where: { id: { in: orderRows.map((r) => r.id) } },
          data: { readyAt: new Date() },
        });
        releasedOrders += 1;
        console.log(`[preorder release] ${shop} order ${orderId}: released + tagged ${readyTags.join(", ")}`);
      } catch (error) {
        // Most likely a missing scope before re-auth — leave readyAt null to retry.
        const message = error instanceof Error ? error.message : String(error);
        if (/access denied|not approved|scope/i.test(message)) skippedNoScope = true;
        console.warn(`[preorder release] ${shop} order ${orderId} failed:`, message);
        errors += 1;
      }
    }
  }

  return { releasedOrders, errors, skippedNoScope };
}

// Run the release check on a timer (the app server is long-lived on Railway).
// Guarded so it starts once per process. First pass 30s after boot, then every
// 10 minutes — a small delay after stock is loaded is fine for the warehouse.
export function startPreorderReleaseScheduler() {
  const g = globalThis as unknown as { __kePreorderReleaseStarted?: boolean };
  if (g.__kePreorderReleaseStarted) return;
  g.__kePreorderReleaseStarted = true;
  const run = () => {
    releaseArrivedPreorders()
      .then((r) => { if (r.releasedOrders || r.errors) console.log("[preorder release] cycle:", r); })
      .catch((error) => console.warn("[preorder release] cycle failed:", error instanceof Error ? error.message : error));
  };
  setTimeout(run, 30_000);
  setInterval(run, 10 * 60 * 1000);
  console.log("[preorder release] scheduler started (every 10 min)");
}
