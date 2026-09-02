import type { PreorderMarket } from "./preorder-rules.server";

export const KARMA_EAST_PREORDER_PLAN_PREFIX = "Karma East Pre-order";

export type ShopifySellingPlanAllocation = {
  selling_plan?: { name?: unknown } | null;
};

export type ShopifyOrderLine = {
  id?: unknown;
  product_id?: unknown;
  variant_id?: unknown;
  variant_title?: unknown;
  sku?: unknown;
  quantity?: unknown;
  selling_plan_allocation?: ShopifySellingPlanAllocation | null;
};

export type ShopifyOrderPayload = {
  id?: unknown;
  name?: unknown;
  email?: unknown;
  contact_email?: unknown;
  shipping_address?: { country_code?: unknown } | null;
  billing_address?: { country_code?: unknown } | null;
  line_items?: ShopifyOrderLine[] | null;
};

export function preorderText(value: unknown) {
  return String(value ?? "").trim();
}

function marketFromOrder(order: ShopifyOrderPayload): PreorderMarket {
  const countryCode = preorderText(order.shipping_address?.country_code || order.billing_address?.country_code).toUpperCase();
  return countryCode === "US" ? "USA" : "AU";
}

export function isKarmaEastPreorderLine(line: ShopifyOrderLine) {
  const planName = preorderText(line.selling_plan_allocation?.selling_plan?.name);
  return planName.startsWith(KARMA_EAST_PREORDER_PLAN_PREFIX);
}

export function normalizeShopifyPreorderLines(payload: unknown) {
  const order = (payload && typeof payload === "object" ? payload : {}) as ShopifyOrderPayload;
  const shopifyOrderId = preorderText(order.id);
  const shopifyOrderName = preorderText(order.name) || null;
  const customerEmail = preorderText(order.email || order.contact_email) || null;
  const market = marketFromOrder(order);

  const lines = Array.isArray(order.line_items)
    ? order.line_items.filter(isKarmaEastPreorderLine).map((line) => ({
        shopifyLineItemId: preorderText(line.id),
        productId: preorderText(line.product_id) || null,
        variantId: preorderText(line.variant_id),
        variantTitle: preorderText(line.variant_title) || null,
        sku: preorderText(line.sku) || null,
        quantity: Number(line.quantity),
      }))
    : [];

  return { shopifyOrderId, shopifyOrderName, customerEmail, market, lines };
}
