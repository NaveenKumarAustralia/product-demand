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

// Public: the customer-facing "Expected <date>" label for a dispatch date. Used
// for the variant metafield so the confirmation email can show it on any path.
export function preorderExpectedLabel(value: Date | string | null) {
  return expectedLabel(value);
}

// The customer-facing plan name/options for a given batch + dispatch date. Used
// both when creating and when refreshing the date on an existing plan.
export function preorderPlanNameFor(batchId: number, shipDate: Date | string | null) {
  return `${KARMA_EAST_PREORDER_PLAN_PREFIX} · Batch #${batchId} · Expected ${expectedLabel(shipDate)}`;
}

// Input for sellingPlanGroupUpdate that renames the existing plan to reflect a
// changed dispatch date (so the order line, cart and email show the new date).
export function buildPreorderSellingPlanUpdateInput(input: { batchId: number; shipDate: Date | string | null; sellingPlanId: string }) {
  const dateLabel = expectedLabel(input.shipDate);
  return {
    sellingPlansToUpdate: [
      {
        id: input.sellingPlanId,
        name: `${KARMA_EAST_PREORDER_PLAN_PREFIX} · Batch #${input.batchId} · Expected ${dateLabel}`,
        options: `Expected dispatch ${dateLabel}`,
      },
    ],
  };
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
