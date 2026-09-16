import test from "node:test";
import assert from "node:assert/strict";
import { selectFifoReleasableOrders, type ReleasableOrder } from "./preorder-release-fifo.ts";

const G = "1281::gid://shopify/ProductVariant/1::AU"; // one (batch,variant,market) group

function order(id: string, ms: number, need: Record<string, number>, rowIds: number[] = [Number(id)]): ReleasableOrder {
  return { orderId: id, oldestReservedMs: ms, rowIds, needByGroup: need, readyTags: [`pre-order-ready-batch-1281`] };
}

test("one spare unit releases exactly the OLDEST waiting order (a returned unit → oldest order)", () => {
  const orders = [order("300", 3000, { [G]: 1 }), order("100", 1000, { [G]: 1 }), order("200", 2000, { [G]: 1 })];
  const budget = new Map([[G, 1]]);
  const picked = selectFifoReleasableOrders(orders, budget);
  assert.deepEqual(picked.map((o) => o.orderId), ["100"]);
});

test("no spare units releases nothing (oversold / fully committed)", () => {
  const orders = [order("100", 1000, { [G]: 1 }), order("200", 2000, { [G]: 1 })];
  assert.equal(selectFifoReleasableOrders(orders, new Map([[G, 0]])).length, 0);
});

test("budget covers several → releases that many, oldest first", () => {
  const orders = [order("300", 3000, { [G]: 1 }), order("100", 1000, { [G]: 1 }), order("200", 2000, { [G]: 1 })];
  const picked = selectFifoReleasableOrders(orders, new Map([[G, 2]]));
  assert.deepEqual(picked.map((o) => o.orderId), ["100", "200"]);
});

test("a multi-unit order is only released when the WHOLE order fits the budget", () => {
  // Oldest order needs 2 units but only 1 is free → it's skipped, and the 1 free
  // unit is NOT wasted: it goes to the next order that fits (needs 1).
  const orders = [order("100", 1000, { [G]: 2 }), order("200", 2000, { [G]: 1 })];
  const picked = selectFifoReleasableOrders(orders, new Map([[G, 1]]));
  assert.deepEqual(picked.map((o) => o.orderId), ["200"]);
});

test("multi-group order needs EVERY group covered (never half-ship across batches)", () => {
  const G2 = "1282::gid://shopify/ProductVariant/2::AU";
  const orders = [order("100", 1000, { [G]: 1, [G2]: 1 })];
  // Only the first group has a spare unit → the order waits (would otherwise ship
  // half from a batch that hasn't arrived).
  assert.equal(selectFifoReleasableOrders(orders, new Map([[G, 1], [G2, 0]])).length, 0);
  // Both covered → released.
  assert.equal(selectFifoReleasableOrders(orders, new Map([[G, 1], [G2, 1]])).length, 1);
});

test("does not mutate the caller's budget map", () => {
  const budget = new Map([[G, 2]]);
  selectFifoReleasableOrders([order("100", 1000, { [G]: 1 })], budget);
  assert.equal(budget.get(G), 2);
});
