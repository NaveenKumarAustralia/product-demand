export const KARMA_EAST_PREORDER_PLAN_PREFIX = "Karma East Pre-order";

export type PreorderSellingPlanInput = {
  batchId: number;
  productTitle: string;
  shipDate: Date | string | null;
  productIds?: string[];
  variantIds: string[];
};

function toProductGid(value: string) {
  const text = String(value ?? "").trim();
  if (!text) return null;
  return text.startsWith("gid://shopify/Product/") ? text : `gid://shopify/Product/${text}`;
}

function toVariantGid(value: string) {
  const text = String(value ?? "").trim();
  if (!text) return null;
  return text.startsWith("gid://shopify/ProductVariant/") ? text : `gid://shopify/ProductVariant/${text}`;
}

function expectedLabel(value: Date | string | null) {
  if (!value) return "date to be confirmed";
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return "date to be confirmed";
  return new Intl.DateTimeFormat("en-AU", { day: "numeric", month: "short", year: "numeric", timeZone: "Australia/Adelaide" }).format(date);
}

export function buildPreorderSellingPlanGroup(input: PreorderSellingPlanInput) {
  if (!Number.isInteger(input.batchId) || input.batchId <= 0) throw new Error("A valid production batch ID is required.");

  const productIds = Array.from(new Set((input.productIds ?? []).map(toProductGid).filter(Boolean))) as string[];
  const productVariantIds = Array.from(new Set(input.variantIds.map(toVariantGid).filter(Boolean))) as string[];
  if (!productIds.length && !productVariantIds.length) throw new Error("At least one Shopify product or variant is required.");

  const dateLabel = expectedLabel(input.shipDate);
  const planName = `${KARMA_EAST_PREORDER_PLAN_PREFIX} · Batch #${input.batchId} · Expected ${dateLabel}`;

  return {
    input: {
      name: KARMA_EAST_PREORDER_PLAN_PREFIX,
      merchantCode: `karma-east-preorder-batch-${input.batchId}`,
      options: ["Pre-order"],
      position: 1,
      sellingPlansToCreate: [
        {
          name: planName,
          options: `Expected dispatch ${dateLabel}`,
          category: "PRE_ORDER",
          billingPolicy: {
            fixed: {
              checkoutCharge: {
                type: "PERCENTAGE",
                value: { percentage: 100 },
              },
              remainingBalanceChargeTrigger: "NO_REMAINING_BALANCE",
            },
          },
          deliveryPolicy: {
            fixed: { fulfillmentTrigger: "UNKNOWN" },
          },
          inventoryPolicy: {
            reserve: "ON_FULFILLMENT",
          },
        },
      ],
    },
    resources: {
      productIds,
      productVariantIds,
    },
  };
}
