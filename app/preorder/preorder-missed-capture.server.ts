import prisma from "./../db.server";
import { KARMA_EAST_PREORDER_PLAN_PREFIX } from "./preorder-shopify-order-normalize";
import { reservePreorderLine, PreorderCapacityError } from "./preorder-allocation.server";
import { getOfflineToken, addOrderTags, holdOrderOpenFulfillmentOrders } from "./preorder-fulfillment.server";
import { marketFromDestination, type PreorderMarket } from "./preorder-rules.server";

const API_VERSION = "2025-10";
const numericId = (gid: string) => String(gid ?? "").split("/").pop()?.replace(/[^0-9]/g, "") ?? "";

type MissedLine = {
  order: string; orderIdNumeric: string; lineId: string; variantId: string; title: string | null; size: string | null;
  qty: number; market: PreorderMarket; batchId: number;
};

/**
 * Find (and optionally convert) paid orders that sold a PRE-ORDER-ENABLED variant
 * WITHOUT the Karma East selling plan — i.e. an out-of-stock item bought via
 * quick-add / the theme buy button rather than the pre-order button. These never
 * get flagged/reserved/held on their own. `apply` reserves them against the
 * batch, tags the order pre-order + pre-order-hold, and holds it so it ships
 * when the stock lands. Idempotent (reservation keyed by line item).
 */
export async function captureMissedPreorders(opts: { days?: number; apply?: boolean } = {}): Promise<{
  scannedOrders: number; candidates: MissedLine[]; converted: number; errors: Array<{ order: string; error: string }>; applied: boolean; skippedNoScope: boolean;
}> {
  const days = Math.max(1, Math.min(120, Math.floor(opts.days ?? 21)));
  const apply = opts.apply === true;

  const session = await prisma.session.findFirst({
    where: { isOnline: false, accessToken: { not: "" } },
    orderBy: { expires: "desc" },
    select: { shop: true, accessToken: true },
  });
  if (!session?.accessToken) return { scannedOrders: 0, candidates: [], converted: 0, errors: [{ order: "-", error: "No offline Shopify session." }], applied: apply, skippedNoScope: true };
  const { shop, accessToken } = session;

  // Map every ENABLED pre-order batch's variants → { batchId, market }. A variant
  // can be in both an AU and a USA batch, so key by numeric variant + market.
  const enabledSettings = await prisma.preorderBatchSetting.findMany({ where: { enabled: true }, select: { supplierOrderId: true } });
  const enabledIds = enabledSettings.map((s) => s.supplierOrderId);
  const batches = enabledIds.length
    ? await prisma.supplierOrder.findMany({
        where: { id: { in: enabledIds }, status: "open", destination: { in: ["send_to_au", "send_to_usa"] } },
        select: { id: true, destination: true, lines: { select: { variantId: true } } },
      })
    : [];
  const variantMarketToBatch = new Map<string, number>(); // `${numericVariant}:${market}` -> batchId
  for (const b of batches) {
    const market = marketFromDestination(b.destination);
    if (!market) continue;
    for (const l of b.lines) {
      const num = numericId(l.variantId);
      if (num) variantMarketToBatch.set(`${num}:${market}`, b.id);
    }
  }
  if (!variantMarketToBatch.size) return { scannedOrders: 0, candidates: [], converted: 0, errors: [], applied: apply, skippedNoScope: false };

  // Already-reserved line items (so we never double-count / re-convert).
  const reserved = await prisma.preorderReservation.findMany({ where: { status: "reserved" }, select: { shopifyLineItemId: true } });
  const reservedLineIds = new Set(reserved.map((r) => r.shopifyLineItemId));

  const sinceIso = new Date(Date.now() - days * 86400000).toISOString();
  const candidates: MissedLine[] = [];
  let scannedOrders = 0;
  let cursor: string | null = null;
  for (let page = 0; page < 20; page += 1) { // cap pages so a huge history can't run away
    const res: Response = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": accessToken },
      body: JSON.stringify({
        query: `#graphql
          query MissedPreorders($q: String!, $cursor: String) {
            orders(first: 100, after: $cursor, query: $q, sortKey: CREATED_AT) {
              pageInfo { hasNextPage endCursor }
              nodes {
                id name email cancelledAt
                shippingAddress { countryCodeV2 }
                billingAddress { countryCodeV2 }
                lineItems(first: 50) { nodes { id quantity title variant { id title } sellingPlan { name } } }
              }
            }
          }`,
        variables: { q: `created_at:>=${sinceIso} financial_status:paid`, cursor },
      }),
    });
    const json = await res.json() as { data?: { orders?: { pageInfo?: { hasNextPage?: boolean; endCursor?: string | null }; nodes?: Array<{
      id: string; name: string; email: string | null; cancelledAt: string | null;
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
        const hasPlan = (line.sellingPlan?.name ?? "").startsWith(KARMA_EAST_PREORDER_PLAN_PREFIX);
        if (hasPlan) continue;
        const vnum = numericId(String(line.variant?.id ?? ""));
        const batchId = variantMarketToBatch.get(`${vnum}:${market}`);
        if (!batchId) continue; // not a pre-order-enabled variant for this market
        const lineId = numericId(line.id);
        if (reservedLineIds.has(lineId)) continue; // already captured
        candidates.push({ order: order.name, orderIdNumeric, lineId, variantId: String(line.variant?.id ?? ""), title: line.title, size: line.variant?.title ?? null, qty: Number(line.quantity), market, batchId });
      }
    }
    const pi = json.data?.orders?.pageInfo;
    if (!pi?.hasNextPage || !pi.endCursor) break;
    cursor = pi.endCursor;
  }

  if (!apply) return { scannedOrders, candidates, converted: 0, errors: [], applied: false, skippedNoScope: false };

  // Apply: reserve + tag + hold each candidate's order.
  const token = await getOfflineToken(shop);
  if (!token) return { scannedOrders, candidates, converted: 0, errors: [{ order: "-", error: "No offline token / missing scopes." }], applied: true, skippedNoScope: true };
  let converted = 0;
  const errors: Array<{ order: string; error: string }> = [];
  const taggedOrders = new Set<string>();
  for (const c of candidates) {
    try {
      await reservePreorderLine({
        shop, shopifyOrderId: c.orderIdNumeric, shopifyOrderName: c.order,
        shopifyLineItemId: c.lineId, productId: null,
        variantId: c.variantId,
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
