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
  const lines = (order.lineItems?.nodes ?? [])
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

  return { shopifyOrderId: numericId(order.id), shopifyOrderName: order.name ?? null, customerEmail: order.email ?? null, market, lines };
}

export async function processShopifyOrderCreated(shop: string, payload: unknown) {
  const restOrder = (payload && typeof payload === "object" ? payload : {}) as { id?: unknown };
  const orderIdNumeric = String(restOrder.id ?? "").replace(/[^0-9]/g, "");
  if (!orderIdNumeric) return { preorder: false, reservations: 0 };

  const normalized = await fetchPreorderOrderViaGraphql(shop, orderIdNumeric);
  if (!normalized || !normalized.lines.length) return { preorder: false, reservations: 0 };
  if (!normalized.shopifyOrderId) throw new PreorderCapacityError("Shopify preorder order ID is missing.");

  try {
    let reservations = 0;
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
