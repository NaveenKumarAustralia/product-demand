import assert from "node:assert/strict";
import test from "node:test";
import { resolveStorefrontVariantState } from "./preorder-storefront-state.ts";

const base = {
  market: "AU" as const,
  eligible: true,
  enabled: true,
  shopifySellingPlanActive: true,
  sellingPlanId: "gid://shopify/SellingPlan/1",
  expectedShipDate: "2026-10-10T00:00:00.000Z",
  availableToPreorder: 10,
};

test("physical stock always wins over preorder capacity", () => {
  const state = resolveStorefrontVariantState({
    market: "AU",
    physicalAvailable: 3,
    candidates: [{ ...base, batchId: 1 }],
  });
  assert.deepEqual(state, { state: "in_stock", physicalAvailable: 3 });
});

test("earliest eligible active batch is selected for preorder", () => {
  const state = resolveStorefrontVariantState({
    market: "AU",
    physicalAvailable: 0,
    candidates: [
      { ...base, batchId: 2, expectedShipDate: "2026-11-01T00:00:00.000Z", sellingPlanId: "gid://shopify/SellingPlan/2" },
      { ...base, batchId: 1, expectedShipDate: "2026-10-01T00:00:00.000Z", sellingPlanId: "gid://shopify/SellingPlan/1" },
    ],
  });
  assert.equal(state.state, "preorder");
  if (state.state === "preorder") {
    assert.equal(state.batchId, 1);
    assert.equal(state.sellingPlanId, "gid://shopify/SellingPlan/1");
  }
});

test("AU and USA preorder pools never cross", () => {
  const state = resolveStorefrontVariantState({
    market: "USA",
    physicalAvailable: 0,
    candidates: [{ ...base, batchId: 1 }],
  });
  assert.deepEqual(state, { state: "notify_me", physicalAvailable: 0 });
});

test("when a variant has BOTH an AU and a USA plan, each market gets only its own", () => {
  const auPlan = { ...base, batchId: 1, market: "AU" as const, sellingPlanId: "gid://shopify/SellingPlan/AU" };
  const usaPlan = { ...base, batchId: 2, market: "USA" as const, sellingPlanId: "gid://shopify/SellingPlan/USA" };
  const candidates = [auPlan, usaPlan];

  const au = resolveStorefrontVariantState({ market: "AU", physicalAvailable: 0, candidates });
  const usa = resolveStorefrontVariantState({ market: "USA", physicalAvailable: 0, candidates });

  assert.equal(au.state, "preorder");
  assert.equal(usa.state, "preorder");
  if (au.state === "preorder") assert.equal(au.sellingPlanId, "gid://shopify/SellingPlan/AU");
  if (usa.state === "preorder") assert.equal(usa.sellingPlanId, "gid://shopify/SellingPlan/USA");
});

test("no capacity or no live Shopify selling plan falls back to notify me", () => {
  for (const candidate of [
    { ...base, batchId: 1, availableToPreorder: 0 },
    { ...base, batchId: 1, shopifySellingPlanActive: false },
    { ...base, batchId: 1, enabled: false },
  ]) {
    const state = resolveStorefrontVariantState({ market: "AU", physicalAvailable: 0, candidates: [candidate] });
    assert.deepEqual(state, { state: "notify_me", physicalAvailable: 0 });
  }
});
