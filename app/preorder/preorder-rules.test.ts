import assert from "node:assert/strict";
import test from "node:test";
import {
  calculatePreorderCapacity,
  getPreorderEligibility,
  marketFromDestination,
  PREORDER_DESTINATION_AU,
  PREORDER_DESTINATION_USA,
  PREORDER_STATUS_ON_PRODUCTION,
} from "./preorder-rules.server.ts";

test("maps production destinations to separate preorder markets", () => {
  assert.equal(marketFromDestination(PREORDER_DESTINATION_AU), "AU");
  assert.equal(marketFromDestination(PREORDER_DESTINATION_USA), "USA");
  assert.equal(marketFromDestination("keep_at_factory"), null);
});

test("requires on-production status, valid destination and explicit enable", () => {
  assert.deepEqual(
    getPreorderEligibility({
      supplierStatus: PREORDER_STATUS_ON_PRODUCTION,
      destination: PREORDER_DESTINATION_AU,
      preorderEnabled: true,
    }),
    { eligible: true, market: "AU", reason: "eligible" },
  );

  assert.equal(
    getPreorderEligibility({
      supplierStatus: "on_order",
      destination: PREORDER_DESTINATION_AU,
      preorderEnabled: true,
    }).eligible,
    false,
  );

  assert.equal(
    getPreorderEligibility({
      supplierStatus: PREORDER_STATUS_ON_PRODUCTION,
      destination: PREORDER_DESTINATION_AU,
      preorderEnabled: false,
    }).reason,
    "not_enabled",
  );
});

test("later production statuses (ready, in shipment, custom) are preorder-eligible; on_order/cancelled are not", () => {
  for (const supplierStatus of ["ready", "in_shipment", "arrived_in_au"]) {
    assert.equal(
      getPreorderEligibility({ supplierStatus, destination: PREORDER_DESTINATION_AU, preorderEnabled: true }).eligible,
      true,
      `${supplierStatus} should be eligible`,
    );
  }
  for (const supplierStatus of ["on_order", "cancelled", ""]) {
    const result = getPreorderEligibility({ supplierStatus, destination: PREORDER_DESTINATION_AU, preorderEnabled: true });
    assert.equal(result.eligible, false, `${supplierStatus || "(blank)"} should not be eligible`);
    assert.equal(result.reason, "not_on_production");
  }
});

test("100 incoming with 5 percent safety produces 95 preorder capacity", () => {
  assert.deepEqual(
    calculatePreorderCapacity({
      confirmedIncomingQty: 100,
      reservedQty: 0,
      safetyBufferPercent: 5,
    }),
    {
      confirmedIncomingQty: 100,
      reservedQty: 0,
      safetyBufferQty: 5,
      availableToPreorder: 95,
      overallocatedBy: 0,
    },
  );
});

test("reserved quantity reduces remaining capacity", () => {
  const result = calculatePreorderCapacity({
    confirmedIncomingQty: 100,
    reservedQty: 90,
    safetyBufferPercent: 5,
  });
  assert.equal(result.availableToPreorder, 5);
  assert.equal(result.overallocatedBy, 0);
});

test("quantity reduction reports overallocation without negative capacity", () => {
  const result = calculatePreorderCapacity({
    confirmedIncomingQty: 80,
    reservedQty: 90,
    safetyBufferPercent: 5,
  });
  assert.equal(result.safetyBufferQty, 4);
  assert.equal(result.availableToPreorder, 0);
  assert.equal(result.overallocatedBy, 14);
});

test("explicit safety quantity overrides percentage and is clamped safely", () => {
  const result = calculatePreorderCapacity({
    confirmedIncomingQty: 50,
    reservedQty: 12,
    safetyBufferPercent: 50,
    safetyBufferQty: 3,
  });
  assert.equal(result.safetyBufferQty, 3);
  assert.equal(result.availableToPreorder, 35);
});
