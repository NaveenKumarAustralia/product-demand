import prisma from "../db.server";
import {
  PreorderCapacityError,
  fulfillPreorderOrder,
  releasePreorderOrder,
  reservePreorderLine,
} from "./preorder-allocation.server";
import {
  normalizeShopifyPreorderLines,
  preorderText,
  type ShopifyOrderPayload,
} from "./preorder-shopify-order-normalize";

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
