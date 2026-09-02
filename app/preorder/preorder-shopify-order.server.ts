import prisma from "../db.server";
import {
  PreorderCapacityError,
  fulfillPreorderOrder,
  releasePreorderOrder,
  reservePreorderLine,
} from "./preorder-allocation.server";
import type { PreorderMarket } from "./preorder-rules.server";

export const KARMA_EAST_PREORDER_PLAN_PREFIX = "Karma East Pre-order";

type ShopifySellingPlanAllocation = {
  selling_plan?: { name?: unknown } | null;
};

type ShopifyOrderLine = {
  id?: unknown;
  product_id?: unknown;
  variant_id?: unknown;
  variant_title?: unknown;
  sku?: unknown;
  quantity?: unknown;
  selling_plan_allocation?: ShopifySellingPlanAllocation | null;
};

type ShopifyOrderPayload = {
  id?: unknown;
  name?: unknown;
  email?: unknown;
  contact_email?: unknown;
  shipping_address?: { country_code?: unknown } | null;
  billing_address?: { country_code?: unknown } | null;
  line_items?: ShopifyOrderLine[] | null;
};

function text(value: unknown) {
  return String(value ?? "").trim();
}

function marketFromOrder(order: ShopifyOrderPayload): PreorderMarket {
  const countryCode = text(order.shipping_address?.country_code || order.billing_address?.country_code).toUpperCase();
  return countryCode === "US" ? "USA" : "AU";
}

export function isKarmaEastPreorderLine(line: ShopifyOrderLine) {
  const planName = text(line.selling_plan_allocation?.selling_plan?.name);
  return planName.startsWith(KARMA_EAST_PREORDER_PLAN_PREFIX);
}

export function normalizeShopifyPreorderLines(payload: unknown) {
  const order = (payload && typeof payload === "object" ? payload : {}) as ShopifyOrderPayload;
  const shopifyOrderId = text(order.id);
  const shopifyOrderName = text(order.name) || null;
  const customerEmail = text(order.email || order.contact_email) || null;
  const market = marketFromOrder(order);

  const lines = Array.isArray(order.line_items)
    ? order.line_items.filter(isKarmaEastPreorderLine).map((line) => ({
        shopifyLineItemId: text(line.id),
        productId: text(line.product_id) || null,
        variantId: text(line.variant_id),
        variantTitle: text(line.variant_title) || null,
        sku: text(line.sku) || null,
        quantity: Number(line.quantity),
      }))
    : [];

  return { shopifyOrderId, shopifyOrderName, customerEmail, market, lines };
}

export async function processShopifyOrderCreated(shop: string, payload: unknown) {
  const normalized = normalizeShopifyPreorderLines(payload);
  if (!normalized.lines.length) return { preorder: false, reservations: 0 };
  if (!normalized.shopifyOrderId) throw new PreorderCapacityError("Shopify preorder order ID is missing.");

  try {
    let reservations = 0;
    for (const line of normalized.lines) {
      if (!line.shopifyLineItemId || !line.variantId || !Number.isInteger(line.quantity) || line.quantity <= 0) {
        throw new PreorderCapacityError("Shopify preorder line is missing a valid line ID, variant ID or quantity.");
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
      });
      reservations += rows.reduce((sum, row) => sum + row.quantity, 0);
    }

    return { preorder: true, reservations, market: normalized.market };
  } catch (error) {
    // Never leave a partially allocated multi-line preorder. If one line can't
    // be allocated, release any reservations created earlier in this order and
    // surface the order for staff attention.
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
  const orderId = text(order.id);
  if (!orderId) return { released: 0 };
  const result = await releasePreorderOrder(shop, orderId);
  return { released: result.count };
}

export async function processShopifyOrderFulfilled(shop: string, payload: unknown) {
  const order = (payload && typeof payload === "object" ? payload : {}) as ShopifyOrderPayload;
  const orderId = text(order.id);
  if (!orderId) return { fulfilled: 0 };
  const result = await fulfillPreorderOrder(shop, orderId);
  return { fulfilled: result.count };
}
