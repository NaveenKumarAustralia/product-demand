import prisma from "../db.server";
import {
  PreorderCapacityError,
  fulfillPreorderOrder,
  releasePreorderOrder,
  reservePreorderLine,
} from "./preorder-allocation.server";
import {
  KARMA_EAST_PREORDER_PLAN_PREFIX,
  preorderBatchIdFromPlanName,
  preorderText,
  type ShopifyOrderPayload,
} from "./preorder-shopify-order-normalize";
import { getOfflineToken, addOrderTags, holdOrderOpenFulfillmentOrders } from "./preorder-fulfillment.server";
import { getPreorderCombineWindowDays } from "./preorder-storefront-settings.server";
import { captureNoPlanLinesForOrder } from "./preorder-missed-capture.server";

const API_VERSION = "2025-10";

function numericId(value: string) {
  return String(value ?? "").split("/").pop()?.replace(/[^0-9]/g, "") ?? "";
}

// The ORDERS_CREATE REST webhook payload does NOT expose the selling plan in the
// field our REST normalizer expected, so preorder lines were silently missed
// (order paid, nothing reserved). Read the order back via GraphQL, where the
// selling plan is reliably available as LineItem.sellingPlan.name, and normalize
// from that. This is the authoritative source and keeps the webhook correct
// regardless of REST payload shape.
async function fetchPreorderOrderViaGraphql(shop: string, orderIdNumeric: string) {
  const session = await prisma.session.findFirst({
    where: { shop, isOnline: false, accessToken: { not: "" } },
    orderBy: { expires: "desc" },
    select: { accessToken: true },
  });
  if (!session?.accessToken) throw new PreorderCapacityError("Offline Shopify session missing; cannot read order for preorder reservation.");

  const response = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
    method: "POST",
    headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": session.accessToken },
    body: JSON.stringify({
      query: `#graphql
        query PreorderOrder($id: ID!) {
          order(id: $id) {
            id name email
            shippingAddress { countryCodeV2 }
            billingAddress { countryCodeV2 }
            lineItems(first: 100) {
              nodes { id quantity sku title variant { id title } sellingPlan { name } }
            }
          }
        }
      `,
      variables: { id: `gid://shopify/Order/${orderIdNumeric}` },
    }),
  });
  if (!response.ok) throw new PreorderCapacityError(`Shopify returned HTTP ${response.status} reading the order.`);
  const json = await response.json() as {
    data?: { order?: {
      id: string; name: string | null; email: string | null;
      shippingAddress?: { countryCodeV2?: string | null } | null;
      billingAddress?: { countryCodeV2?: string | null } | null;
      lineItems?: { nodes?: Array<{
        id: string; quantity: number; sku: string | null; title: string | null;
        variant?: { id?: string | null; title?: string | null } | null;
        sellingPlan?: { name?: string | null } | null;
      }> };
    } };
    errors?: Array<{ message?: string }>;
  };
  if (json.errors?.length) throw new PreorderCapacityError(json.errors.map((error) => error.message || "Shopify GraphQL error").join("; "));
  const order = json.data?.order;
  if (!order) return null;

  const country = String(order.shippingAddress?.countryCodeV2 || order.billingAddress?.countryCodeV2 || "").toUpperCase();
  const market = country === "US" ? "USA" as const : "AU" as const;
  const allNodes = order.lineItems?.nodes ?? [];
  const lines = allNodes
    .filter((line) => (line.sellingPlan?.name ?? "").startsWith(KARMA_EAST_PREORDER_PLAN_PREFIX))
    .map((line) => ({
      shopifyLineItemId: numericId(line.id),
      productId: null as string | null,
      variantId: String(line.variant?.id ?? ""),
      variantTitle: line.variant?.title ?? line.title ?? null,
      sku: line.sku ?? null,
      quantity: Number(line.quantity),
      sellingPlanName: line.sellingPlan?.name ?? "",
      preferredSupplierOrderId: preorderBatchIdFromPlanName(line.sellingPlan?.name),
    }));
  // Lines WITHOUT our selling plan (Shop Pay / express / any no-plan path) — the
  // webhook captures the ones that are live pre-order variants in real time.
  const noPlanLines = allNodes
    .filter((line) => !(line.sellingPlan?.name ?? "").startsWith(KARMA_EAST_PREORDER_PLAN_PREFIX) && line.variant?.id && Number(line.quantity) > 0)
    .map((line) => ({
      lineId: numericId(line.id),
      variantId: String(line.variant?.id ?? ""),
      qty: Number(line.quantity),
      title: line.title ?? null,
      size: line.variant?.title ?? null,
    }));

  return {
    shopifyOrderId: numericId(order.id),
    shopifyOrderName: order.name ?? null,
    customerEmail: order.email ?? null,
    market,
    lines,
    noPlanLines,
    totalLines: allNodes.length,
  };
}

export async function processShopifyOrderCreated(shop: string, payload: unknown) {
  const restOrder = (payload && typeof payload === "object" ? payload : {}) as { id?: unknown };
  const orderIdNumeric = String(restOrder.id ?? "").replace(/[^0-9]/g, "");
  if (!orderIdNumeric) return { preorder: false, reservations: 0 };

  const normalized = await fetchPreorderOrderViaGraphql(shop, orderIdNumeric);
  if (!normalized) return { preorder: false, reservations: 0 };
  if (!normalized.shopifyOrderId) throw new PreorderCapacityError("Shopify preorder order ID is missing.");

  // Real-time missed-capture: pull in any line that bought a LIVE pre-order
  // variant with NO selling plan (Shop Pay / express / any no-plan path) the
  // instant the order lands — reserve + tag + hold. This is what makes every
  // checkout path a proper pre-order without waiting for the scheduler.
  let capturedMissed = 0;
  if (normalized.noPlanLines.length) {
    try {
      const r = await captureNoPlanLinesForOrder(shop, orderIdNumeric, normalized.shopifyOrderName, normalized.market, normalized.noPlanLines);
      capturedMissed = r.captured;
    } catch (error) {
      console.warn("[preorder] real-time missed capture failed:", error instanceof Error ? error.message : error);
    }
  }

  // No lines carry our selling plan → nothing more to reserve via the plan path.
  if (!normalized.lines.length) return { preorder: capturedMissed > 0, reservations: capturedMissed };

  // Tag the order so Pick Pack handles it: a fully pre-order order gets
  // `pre-order-hold` (Pick Pack sets it aside entirely); a mixed order gets
  // `pre-order` (Pick Pack hides just the pre-order line). Needs write_orders
  // (granted on re-auth) — fails soft until then. Tags before allocating so the
  // pick-pack team never treats it as normal even if allocation needs review.
  // A fully pre-order order: every line is a pre-order. A mixed order has some
  // in-stock lines too.
  const fullyPreorder = normalized.totalLines > 0 && normalized.lines.length >= normalized.totalLines;
  try {
    const token = await getOfflineToken(shop);
    if (token) {
      const batchIds = Array.from(new Set(normalized.lines.map((line) => line.preferredSupplierOrderId).filter(Boolean)));
      const tags = ["pre-order", ...batchIds.map((id) => `pre-order-batch-${id}`)];
      if (fullyPreorder) tags.push("pre-order-hold");
      await addOrderTags(shop, token, orderIdNumeric, tags);
    }
  } catch (error) {
    console.warn("[preorder] order tagging failed (needs write_orders?):", error instanceof Error ? error.message : error);
  }

  try {
    let reservations = 0;
    let earliestShipMs: number | null = null;
    for (const line of normalized.lines) {
      if (!line.shopifyLineItemId || !line.variantId || !Number.isInteger(line.quantity) || line.quantity <= 0) {
        throw new PreorderCapacityError("Shopify preorder line is missing a valid line ID, variant ID or quantity.");
      }
      if (!line.preferredSupplierOrderId) {
        throw new PreorderCapacityError(
          `Shopify preorder line ${line.shopifyLineItemId} is missing its production batch reference. The order has not been allocated automatically.`,
        );
      }
      const rows = await reservePreorderLine({
        shop,
        shopifyOrderId: normalized.shopifyOrderId,
        shopifyOrderName: normalized.shopifyOrderName,
        shopifyLineItemId: line.shopifyLineItemId,
        productId: line.productId,
        variantId: line.variantId,
        variantTitle: line.variantTitle,
        sku: line.sku,
        market: normalized.market,
        quantity: line.quantity,
        customerEmail: normalized.customerEmail,
        preferredSupplierOrderId: line.preferredSupplierOrderId,
      });
      reservations += rows.reduce((sum, row) => sum + row.quantity, 0);
      for (const row of rows) {
        const t = row.expectedShipDate ? new Date(row.expectedShipDate).getTime() : NaN;
        if (Number.isFinite(t) && (earliestShipMs === null || t < earliestShipMs)) earliestShipMs = t;
      }
    }

    // "Combine window": for a MIXED order whose pre-order is due within N days,
    // hold the in-stock items too (and set the whole order aside for Pick Pack)
    // so it all ships together when the batch lands. The existing stock-aware
    // auto-release releases every hold on the order at once.
    if (!fullyPreorder && earliestShipMs !== null) {
      try {
        const windowDays = await getPreorderCombineWindowDays();
        const cutoff = Date.now() + windowDays * 86400000;
        if (windowDays > 0 && earliestShipMs <= cutoff) {
          const token = await getOfflineToken(shop);
          if (token) {
            const held = await holdOrderOpenFulfillmentOrders(shop, token, orderIdNumeric, "Held to ship with the pre-order item in this order (combine window)");
            if (held > 0) await addOrderTags(shop, token, orderIdNumeric, ["pre-order-hold"]);
            console.log(`[preorder combine] ${shop} order ${orderIdNumeric}: held ${held} in-stock fulfilment order(s) to ship with the pre-order.`);
          }
        }
      } catch (error) {
        // Non-fatal: reservation still succeeded; the in-stock part just ships
        // separately if the hold couldn't be placed (e.g., missing scope).
        console.warn("[preorder combine] hold failed:", error instanceof Error ? error.message : error);
      }
    }

    return { preorder: true, reservations, market: normalized.market };
  } catch (error) {
    await releasePreorderOrder(shop, normalized.shopifyOrderId).catch(() => undefined);
    await prisma.activityLog.create({
      data: {
        userName: "Shopify webhook",
        action: "preorder_allocation_failed",
        entity: "shopify_order",
        entityId: normalized.shopifyOrderId,
        entityName: normalized.shopifyOrderName,
        field: "reservation",
        toValue: error instanceof Error ? error.message : "Unknown preorder allocation error",
      },
    }).catch(() => undefined);
    throw error;
  }
}

export async function processShopifyOrderCancelled(shop: string, payload: unknown) {
  const order = (payload && typeof payload === "object" ? payload : {}) as ShopifyOrderPayload;
  const orderId = preorderText(order.id);
  if (!orderId) return { released: 0 };
  const result = await releasePreorderOrder(shop, orderId);
  return { released: result.count };
}

export async function processShopifyOrderFulfilled(shop: string, payload: unknown) {
  const order = (payload && typeof payload === "object" ? payload : {}) as ShopifyOrderPayload;
  const orderId = preorderText(order.id);
  if (!orderId) return { fulfilled: 0 };
  const result = await fulfillPreorderOrder(shop, orderId);
  return { fulfilled: result.count };
}
