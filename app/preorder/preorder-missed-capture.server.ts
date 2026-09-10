import prisma from "./../db.server";
import { KARMA_EAST_PREORDER_PLAN_PREFIX } from "./preorder-shopify-order-normalize";
import { reservePreorderLine, PreorderCapacityError } from "./preorder-allocation.server";
import { getOfflineToken, getAvailableAtLocation, addOrderTags, holdOrderOpenFulfillmentOrders } from "./preorder-fulfillment.server";
import { getPreorderSellingPlanRegistryEntries } from "./preorder-selling-plan-registry.server";
import { getPreorderLocationSettings, locationForMarket } from "./preorder-locations.server";
import { marketFromDestination, type PreorderMarket } from "./preorder-rules.server";

const API_VERSION = "2025-10";
const numericId = (gid: string) => String(gid ?? "").split("/").pop()?.replace(/[^0-9]/g, "") ?? "";

type MissedLine = {
  order: string; orderIdNumeric: string; lineId: string; variantId: string; title: string | null; size: string | null;
  qty: number; market: PreorderMarket; batchId: number; available: number;
};

/**
 * Find (and optionally convert) paid orders that sold a LIVE pre-order variant
 * WITHOUT the Karma East selling plan — i.e. an out-of-stock item bought via
 * quick-add / the theme buy button rather than the pre-order button. Tight gates
 * so we never touch a normal in-stock order:
 *   1. variant belongs to a batch that is ENABLED **and activated** (a live
 *      Shopify selling plan) — not merely enabled;
 *   2. the order is UNFULFILLED;
 *   3. the variant is currently OUT OF STOCK (available <= 0) at the market's
 *      location (an in-stock sale would still have stock).
 * With `apply`, each is reserved against the batch, tagged pre-order +
 * pre-order-hold, and held so it ships when the batch lands. Idempotent.
 */
export async function captureMissedPreorders(opts: { days?: number; apply?: boolean } = {}): Promise<{
  scannedOrders: number; candidates: MissedLine[]; converted: number; errors: Array<{ order: string; error: string }>; applied: boolean; skippedNoScope: boolean;
}> {
  const days = Math.max(1, Math.min(120, Math.floor(opts.days ?? 30)));
  const apply = opts.apply === true;

  const session = await prisma.session.findFirst({
    where: { isOnline: false, accessToken: { not: "" } },
    orderBy: { expires: "desc" },
    select: { shop: true, accessToken: true },
  });
  if (!session?.accessToken) return { scannedOrders: 0, candidates: [], converted: 0, errors: [{ order: "-", error: "No offline Shopify session." }], applied: apply, skippedNoScope: true };
  const { shop, accessToken } = session;

  // LIVE batches only: enabled AND activated (has a selling plan). An enabled-but-
  // not-activated batch was never offered as pre-order, so its sales are normal.
  const [enabledSettings, registry, locations] = await Promise.all([
    prisma.preorderBatchSetting.findMany({ where: { enabled: true }, select: { supplierOrderId: true } }),
    getPreorderSellingPlanRegistryEntries(shop),
    getPreorderLocationSettings(),
  ]);
  const activatedIds = new Set(registry.map((r) => r.supplierOrderId));
  const liveIds = enabledSettings.map((s) => s.supplierOrderId).filter((id) => activatedIds.has(id));
  const batches = liveIds.length
    ? await prisma.supplierOrder.findMany({
        where: { id: { in: liveIds }, status: "open", destination: { in: ["send_to_au", "send_to_usa"] } },
        select: { id: true, destination: true, lines: { select: { variantId: true } } },
      })
    : [];
  const variantMarketToBatch = new Map<string, { batchId: number; market: PreorderMarket }>();
  for (const b of batches) {
    const market = marketFromDestination(b.destination);
    if (!market) continue;
    for (const l of b.lines) {
      const num = numericId(l.variantId);
      if (num) variantMarketToBatch.set(`${num}:${market}`, { batchId: b.id, market });
    }
  }
  if (!variantMarketToBatch.size) return { scannedOrders: 0, candidates: [], converted: 0, errors: [], applied: apply, skippedNoScope: false };

  const reserved = await prisma.preorderReservation.findMany({ where: { status: "reserved" }, select: { shopifyLineItemId: true } });
  const reservedLineIds = new Set(reserved.map((r) => r.shopifyLineItemId));
  const stockCache = new Map<string, number>(); // `${numericVariant}:${market}` -> available

  const sinceIso = new Date(Date.now() - days * 86400000).toISOString();
  const candidates: MissedLine[] = [];
  const errors: Array<{ order: string; error: string }> = [];
  let scannedOrders = 0;
  let cursor: string | null = null;
  for (let page = 0; page < 15; page += 1) {
    const res: Response = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": accessToken },
      body: JSON.stringify({
        query: `#graphql
          query MissedPreorders($q: String!, $cursor: String) {
            orders(first: 60, after: $cursor, query: $q, sortKey: CREATED_AT, reverse: true) {
              pageInfo { hasNextPage endCursor }
              nodes {
                id name cancelledAt
                shippingAddress { countryCodeV2 }
                billingAddress { countryCodeV2 }
                lineItems(first: 50) { nodes { id quantity title variant { id title } sellingPlan { name } } }
              }
            }
          }`,
        // Newest first, paid, still unfulfilled.
        variables: { q: `created_at:>=${sinceIso} financial_status:paid fulfillment_status:unfulfilled`, cursor },
      }),
    });
    const json = await res.json() as { data?: { orders?: { pageInfo?: { hasNextPage?: boolean; endCursor?: string | null }; nodes?: Array<{
      id: string; name: string; cancelledAt: string | null;
      shippingAddress?: { countryCodeV2?: string | null } | null;
      billingAddress?: { countryCodeV2?: string | null } | null;
      lineItems?: { nodes?: Array<{ id: string; quantity: number; title: string | null; variant?: { id?: string | null; title?: string | null } | null; sellingPlan?: { name?: string | null } | null }> };
    }> } }; errors?: Array<{ message?: string }> };
    if (json.errors?.length) return { scannedOrders, candidates, converted: 0, errors: [{ order: "-", error: json.errors.map((e) => e.message).join("; ") }], applied: apply, skippedNoScope: false };
    const nodes = json.data?.orders?.nodes ?? [];
    for (const order of nodes) {
      scannedOrders += 1;
      if (order.cancelledAt) continue;
      const country = String(order.shippingAddress?.countryCodeV2 || order.billingAddress?.countryCodeV2 || "").toUpperCase();
      const market: PreorderMarket = country === "US" ? "USA" : "AU";
      const orderIdNumeric = numericId(order.id);
      for (const line of order.lineItems?.nodes ?? []) {
        if ((line.sellingPlan?.name ?? "").startsWith(KARMA_EAST_PREORDER_PLAN_PREFIX)) continue;
        const vnum = numericId(String(line.variant?.id ?? ""));
        const hit = variantMarketToBatch.get(`${vnum}:${market}`);
        if (!hit) continue;
        const lineId = numericId(line.id);
        if (reservedLineIds.has(lineId)) continue;
        // Confirm the variant is actually out of stock now (an in-stock sale
        // would still have stock). This is what separates a missed pre-order
        // from a normal order of a product that also has an incoming batch.
        const locationId = locationForMarket(locations, market);
        if (!locationId) continue;
        const stockKey = `${vnum}:${market}`;
        let available = stockCache.get(stockKey);
        if (available === undefined) {
          try { available = await getAvailableAtLocation(shop, accessToken, String(line.variant?.id ?? ""), locationId); }
          catch { available = 1; } // on error, assume in-stock (skip) to stay safe
          stockCache.set(stockKey, available);
        }
        if (available > 0) continue; // had stock → normal sale, not a pre-order
        candidates.push({ order: order.name, orderIdNumeric, lineId, variantId: String(line.variant?.id ?? ""), title: line.title, size: line.variant?.title ?? null, qty: Number(line.quantity), market, batchId: hit.batchId, available });
      }
    }
    const pi = json.data?.orders?.pageInfo;
    if (!pi?.hasNextPage || !pi.endCursor) break;
    cursor = pi.endCursor;
  }

  if (!apply) return { scannedOrders, candidates, converted: 0, errors, applied: false, skippedNoScope: false };

  const token = await getOfflineToken(shop);
  if (!token) return { scannedOrders, candidates, converted: 0, errors: [{ order: "-", error: "No offline token / missing scopes." }], applied: true, skippedNoScope: true };
  let converted = 0;
  const taggedOrders = new Set<string>();
  for (const c of candidates) {
    try {
      await reservePreorderLine({
        shop, shopifyOrderId: c.orderIdNumeric, shopifyOrderName: c.order,
        shopifyLineItemId: c.lineId, productId: null, variantId: c.variantId,
        variantTitle: c.size, sku: null, market: c.market, quantity: c.qty,
        customerEmail: null, preferredSupplierOrderId: c.batchId,
      });
      converted += 1;
    } catch (error) {
      errors.push({ order: c.order, error: error instanceof PreorderCapacityError ? error.message : (error instanceof Error ? error.message : String(error)) });
      continue;
    }
    if (!taggedOrders.has(c.orderIdNumeric)) {
      taggedOrders.add(c.orderIdNumeric);
      try {
        await addOrderTags(shop, token, c.orderIdNumeric, ["pre-order", `pre-order-batch-${c.batchId}`, "pre-order-hold"]);
        await holdOrderOpenFulfillmentOrders(shop, token, c.orderIdNumeric, "Captured missed pre-order — held until the batch lands");
      } catch (error) {
        errors.push({ order: c.order, error: `tag/hold: ${error instanceof Error ? error.message : String(error)}` });
      }
    }
  }
  return { scannedOrders, candidates, converted, errors, applied: true, skippedNoScope: false };
}

// Read-only list of EVERY order that bought a live pre-order variant WITHOUT the
// selling plan (Shop Pay / express / any no-plan path) — for customer follow-up.
// Unlike the capture, this does NOT gate on stock/fulfilment (a customer who got
// the wrong email should be followed up even if their item has since restocked
// or shipped). Returns order #, customer email, product, size, and the batch's
// expected dispatch date.
export async function listAffectedForFollowup(opts: { days?: number } = {}): Promise<{
  scannedOrders: number; orders: string[]; affected: Array<{ order: string; email: string | null; product: string | null; size: string | null; qty: number; batchId: number; dispatch: string | null }>;
}> {
  const days = Math.max(1, Math.min(120, Math.floor(opts.days ?? 14)));
  const session = await prisma.session.findFirst({
    where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true, accessToken: true },
  });
  if (!session?.accessToken) return { scannedOrders: 0, affected: [] };
  const { shop, accessToken } = session;

  const [enabledSettings, registry, batchSettings, locations] = await Promise.all([
    prisma.preorderBatchSetting.findMany({ where: { enabled: true }, select: { supplierOrderId: true } }),
    getPreorderSellingPlanRegistryEntries(shop),
    prisma.preorderBatchSetting.findMany({ select: { supplierOrderId: true, shipDate: true } }),
    getPreorderLocationSettings(),
  ]);
  const activatedIds = new Set(registry.map((r) => r.supplierOrderId));
  const liveIds = enabledSettings.map((s) => s.supplierOrderId).filter((id) => activatedIds.has(id));
  const shipByBatch = new Map(batchSettings.map((b) => [b.supplierOrderId, b.shipDate ? b.shipDate.toISOString() : null]));
  const batches = liveIds.length
    ? await prisma.supplierOrder.findMany({ where: { id: { in: liveIds }, status: "open", destination: { in: ["send_to_au", "send_to_usa"] } }, select: { id: true, destination: true, eta: true, lines: { select: { variantId: true } } } })
    : [];
  const variantMarketToBatch = new Map<string, { batchId: number; dispatch: string | null }>();
  const etaByBatch = new Map(batches.map((b) => [b.id, b.eta ? b.eta.toISOString() : null]));
  for (const b of batches) {
    const market = marketFromDestination(b.destination);
    if (!market) continue;
    const dispatch = shipByBatch.get(b.id) ?? etaByBatch.get(b.id) ?? null;
    for (const l of b.lines) { const num = numericId(l.variantId); if (num) variantMarketToBatch.set(`${num}:${market}`, { batchId: b.id, dispatch }); }
  }
  if (!variantMarketToBatch.size) return { scannedOrders: 0, orders: [], affected: [] };

  const sinceIso = new Date(Date.now() - days * 86400000).toISOString();
  const affected: Array<{ order: string; email: string | null; product: string | null; size: string | null; qty: number; batchId: number; dispatch: string | null }> = [];
  const stockCache = new Map<string, number>();
  let scannedOrders = 0;
  let cursor: string | null = null;
  for (let page = 0; page < 20; page += 1) {
    const res: Response = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
      method: "POST", headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": accessToken },
      body: JSON.stringify({
        query: `#graphql
          query Affected($q: String!, $cursor: String) {
            orders(first: 100, after: $cursor, query: $q, sortKey: CREATED_AT, reverse: true) {
              pageInfo { hasNextPage endCursor }
              nodes { name email cancelledAt shippingAddress { countryCodeV2 } billingAddress { countryCodeV2 }
                lineItems(first: 50) { nodes { title quantity variant { id title } sellingPlan { name } } } }
            }
          }`,
        variables: { q: `created_at:>=${sinceIso} financial_status:paid`, cursor },
      }),
    });
    const json = await res.json() as { data?: { orders?: { pageInfo?: { hasNextPage?: boolean; endCursor?: string | null }; nodes?: Array<{ name: string; email: string | null; cancelledAt: string | null; shippingAddress?: { countryCodeV2?: string | null } | null; billingAddress?: { countryCodeV2?: string | null } | null; lineItems?: { nodes?: Array<{ title: string | null; quantity: number; variant?: { id?: string | null; title?: string | null } | null; sellingPlan?: { name?: string | null } | null }> } }> } }; errors?: Array<{ message?: string }> };
    if (json.errors?.length) break;
    const nodes = json.data?.orders?.nodes ?? [];
    for (const order of nodes) {
      scannedOrders += 1;
      if (order.cancelledAt) continue;
      const country = String(order.shippingAddress?.countryCodeV2 || order.billingAddress?.countryCodeV2 || "").toUpperCase();
      const market: PreorderMarket = country === "US" ? "USA" : "AU";
      for (const line of order.lineItems?.nodes ?? []) {
        if ((line.sellingPlan?.name ?? "").startsWith(KARMA_EAST_PREORDER_PLAN_PREFIX)) continue;
        const vnum = numericId(String(line.variant?.id ?? ""));
        const hit = variantMarketToBatch.get(`${vnum}:${market}`);
        if (!hit) continue;
        // Only count it as affected if the variant was actually OUT OF STOCK — an
        // in-stock sale of a product that merely has a live batch got the correct
        // (normal) email and shouldn't be followed up.
        const locationId = locationForMarket(locations, market);
        if (!locationId) continue;
        const stockKey = `${vnum}:${market}`;
        let available = stockCache.get(stockKey);
        if (available === undefined) {
          try { available = await getAvailableAtLocation(shop, accessToken, String(line.variant?.id ?? ""), locationId); }
          catch { available = 1; }
          stockCache.set(stockKey, available);
        }
        if (available > 0) continue;
        affected.push({ order: order.name, email: order.email, product: line.title, size: line.variant?.title ?? null, qty: line.quantity, batchId: hit.batchId, dispatch: hit.dispatch });
      }
    }
    const pi = json.data?.orders?.pageInfo;
    if (!pi?.hasNextPage || !pi.endCursor) break;
    cursor = pi.endCursor;
  }
  const orders = Array.from(new Set(affected.map((a) => a.order)));
  return { scannedOrders, orders, affected };
}

// Run the capture automatically so missed pre-orders (Shop Pay / quick-add /
// any no-plan path) are pulled in without anyone remembering to. Tight gates
// (live batch + out-of-stock + unfulfilled) make auto-apply safe. First pass ~2
// min after boot, then every 3 hours, over a short rolling window.
export function startMissedPreorderCaptureScheduler() {
  const g = globalThis as unknown as { __keMissedCaptureStarted?: boolean };
  if (g.__keMissedCaptureStarted) return;
  g.__keMissedCaptureStarted = true;
  const run = () => {
    captureMissedPreorders({ days: 4, apply: true })
      .then((r) => { if (r.converted || r.errors.length) console.log("[preorder missed capture] cycle:", { converted: r.converted, candidates: r.candidates.length, errors: r.errors.length }); })
      .catch((e) => console.warn("[preorder missed capture] cycle failed:", e instanceof Error ? e.message : e));
  };
  setTimeout(run, 120_000);
  setInterval(run, 3 * 60 * 60 * 1000);
  console.log("[preorder missed capture] scheduler started (every 3h)");
}
