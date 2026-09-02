import test from "node:test";
import assert from "node:assert/strict";
import { buildPreorderSellingPlanGroup, KARMA_EAST_PREORDER_PLAN_PREFIX } from "./preorder-selling-plan.ts";

test("builds a 100 percent upfront Shopify preorder selling plan", () => {
  const result = buildPreorderSellingPlanGroup({
    batchId: 123,
    productTitle: "Test Dress",
    shipDate: "2026-10-10T00:00:00.000Z",
    variantIds: ["456", "gid://shopify/ProductVariant/789", "456"],
  });

  const plan = result.input.sellingPlansToCreate[0];
  assert.equal(result.input.name, KARMA_EAST_PREORDER_PLAN_PREFIX);
  assert.equal(result.input.merchantCode, "karma-east-preorder-batch-123");
  assert.equal(plan.category, "PRE_ORDER");
  assert.equal(plan.billingPolicy.fixed.checkoutCharge.type, "PERCENTAGE");
  assert.equal(plan.billingPolicy.fixed.checkoutCharge.value.percentage, 100);
  assert.equal(plan.billingPolicy.fixed.remainingBalanceChargeTrigger, "NO_REMAINING_BALANCE");
  assert.equal(plan.deliveryPolicy.fixed.fulfillmentTrigger, "UNKNOWN");
  assert.equal(plan.inventoryPolicy.reserve, "ON_FULFILLMENT");
  assert.deepEqual(result.resources.productVariantIds, [
    "gid://shopify/ProductVariant/456",
    "gid://shopify/ProductVariant/789",
  ]);
  assert.match(plan.name, /Batch #123/);
  assert.match(plan.options, /10 Oct 2026/);
});

test("rejects selling plans with no Shopify product resources", () => {
  assert.throws(() => buildPreorderSellingPlanGroup({
    batchId: 1,
    productTitle: "Test",
    shipDate: null,
    variantIds: [],
  }), /At least one Shopify product or variant/);
});
