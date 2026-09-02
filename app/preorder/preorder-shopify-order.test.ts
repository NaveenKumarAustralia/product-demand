import assert from "node:assert/strict";
import test from "node:test";
import {
  KARMA_EAST_PREORDER_PLAN_PREFIX,
  isKarmaEastPreorderLine,
  normalizeShopifyPreorderLines,
} from "./preorder-shopify-order.server.ts";

test("only our Karma East preorder selling plan is treated as preorder", () => {
  assert.equal(isKarmaEastPreorderLine({
    selling_plan_allocation: { selling_plan: { name: `${KARMA_EAST_PREORDER_PLAN_PREFIX} · Expected 6–11 Oct` } },
  }), true);
  assert.equal(isKarmaEastPreorderLine({
    selling_plan_allocation: { selling_plan: { name: "Subscribe & save" } },
  }), false);
  assert.equal(isKarmaEastPreorderLine({}), false);
});

test("US orders allocate to USA market and ordinary lines are ignored", () => {
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
        selling_plan_allocation: { selling_plan: { name: `${KARMA_EAST_PREORDER_PLAN_PREFIX} · Oct` } },
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
