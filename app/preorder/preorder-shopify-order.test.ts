import assert from "node:assert/strict";
import test from "node:test";
import {
  KARMA_EAST_PREORDER_PLAN_PREFIX,
  isKarmaEastPreorderLine,
  normalizeShopifyPreorderLines,
  preorderBatchIdFromPlanName,
} from "./preorder-shopify-order-normalize.ts";

test("only our Karma East preorder selling plan is treated as preorder", () => {
  assert.equal(isKarmaEastPreorderLine({
    selling_plan_allocation: { selling_plan: { name: `${KARMA_EAST_PREORDER_PLAN_PREFIX} · Batch #123 · Expected 6 Oct` } },
  }), true);
  assert.equal(isKarmaEastPreorderLine({
    selling_plan_allocation: { selling_plan: { name: "Subscribe & save" } },
  }), false);
  assert.equal(isKarmaEastPreorderLine({}), false);
});

test("extracts the production batch ID from our selling plan name", () => {
  assert.equal(preorderBatchIdFromPlanName(`${KARMA_EAST_PREORDER_PLAN_PREFIX} · Batch #123 · Expected 10 Oct 2026`), 123);
  assert.equal(preorderBatchIdFromPlanName(`${KARMA_EAST_PREORDER_PLAN_PREFIX} · Expected 10 Oct 2026`), null);
  assert.equal(preorderBatchIdFromPlanName("Subscribe & save · Batch #123"), null);
});

test("US orders allocate to USA market and retain the promised batch", () => {
  const result = normalizeShopifyPreorderLines({
    id: 101,
    name: "#1001",
    email: "customer@example.com",
    shipping_address: { country_code: "US" },
    line_items: [
      {
        id: 1,
        product_id: 11,
        variant_id: 111,
        variant_title: "M",
        sku: "DRESS-M",
        quantity: 2,
        selling_plan_allocation: { selling_plan: { name: `${KARMA_EAST_PREORDER_PLAN_PREFIX} · Batch #456 · Expected 10 Oct 2026` } },
      },
      {
        id: 2,
        product_id: 22,
        variant_id: 222,
        quantity: 1,
      },
    ],
  });
  assert.equal(result.market, "USA");
  assert.equal(result.shopifyOrderId, "101");
  assert.equal(result.lines.length, 1);
  assert.equal(result.lines[0]?.quantity, 2);
  assert.equal(result.lines[0]?.variantId, "111");
  assert.equal(result.lines[0]?.preferredSupplierOrderId, 456);
  assert.match(result.lines[0]?.sellingPlanName ?? "", /Batch #456/);
});

test("AU, NZ and ROW orders use the AU production pool", () => {
  for (const country of ["AU", "NZ", "GB"]) {
    const result = normalizeShopifyPreorderLines({
      id: 1,
      shipping_address: { country_code: country },
      line_items: [],
    });
    assert.equal(result.market, "AU");
  }
});
