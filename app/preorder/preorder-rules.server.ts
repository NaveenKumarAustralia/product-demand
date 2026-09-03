export const PREORDER_STATUS_ON_PRODUCTION = "on_production";
export const PREORDER_DESTINATION_AU = "send_to_au";
export const PREORDER_DESTINATION_USA = "send_to_usa";

// Preorder is for products that are committed/incoming but not yet sellable as
// in-stock. Any production-lifecycle status qualifies EXCEPT "on_order" (not yet
// committed to production) and "cancelled". A denylist (rather than an allowlist)
// keeps custom statuses a merchant adds (e.g. "arrived_in_au") preorder-capable;
// the storefront's own physical-stock check hides preorder the moment real stock
// lands at the fulfilment location, so allowing later statuses stays safe.
const PREORDER_INELIGIBLE_STATUSES = new Set(["", "on_order", "cancelled"]);
export function isPreorderEligibleStatus(status?: string | null): boolean {
  const value = String(status ?? "").trim();
  return value.length > 0 && !PREORDER_INELIGIBLE_STATUSES.has(value);
}

export type PreorderMarket = "AU" | "USA";

export type PreorderEligibilityInput = {
  supplierStatus?: string | null;
  destination?: string | null;
  preorderEnabled?: boolean | null;
};

export type PreorderEligibilityResult = {
  eligible: boolean;
  market: PreorderMarket | null;
  reason: "eligible" | "not_on_production" | "invalid_destination" | "not_enabled";
};

export function marketFromDestination(destination?: string | null): PreorderMarket | null {
  if (destination === PREORDER_DESTINATION_AU) return "AU";
  if (destination === PREORDER_DESTINATION_USA) return "USA";
  return null;
}

export function getPreorderEligibility(input: PreorderEligibilityInput): PreorderEligibilityResult {
  if (!isPreorderEligibleStatus(input.supplierStatus)) {
    return { eligible: false, market: marketFromDestination(input.destination), reason: "not_on_production" };
  }

  const market = marketFromDestination(input.destination);
  if (!market) {
    return { eligible: false, market: null, reason: "invalid_destination" };
  }

  if (input.preorderEnabled !== true) {
    return { eligible: false, market, reason: "not_enabled" };
  }

  return { eligible: true, market, reason: "eligible" };
}

export type PreorderCapacityInput = {
  confirmedIncomingQty: number;
  reservedQty: number;
  safetyBufferPercent?: number | null;
  safetyBufferQty?: number | null;
};

export type PreorderCapacityResult = {
  confirmedIncomingQty: number;
  reservedQty: number;
  safetyBufferQty: number;
  availableToPreorder: number;
  overallocatedBy: number;
};

function nonNegativeInt(value: number) {
  if (!Number.isFinite(value)) return 0;
  return Math.max(0, Math.floor(value));
}

export function calculatePreorderCapacity(input: PreorderCapacityInput): PreorderCapacityResult {
  const confirmedIncomingQty = nonNegativeInt(input.confirmedIncomingQty);
  const reservedQty = nonNegativeInt(input.reservedQty);
  const explicitBuffer = input.safetyBufferQty == null ? null : nonNegativeInt(input.safetyBufferQty);
  const percent = Number.isFinite(Number(input.safetyBufferPercent))
    ? Math.max(0, Number(input.safetyBufferPercent))
    : 0;

  // Round safety buffer UP so the reserve is never accidentally smaller than configured.
  const safetyBufferQty = explicitBuffer ?? Math.ceil(confirmedIncomingQty * (percent / 100));
  const rawRemaining = confirmedIncomingQty - safetyBufferQty - reservedQty;

  return {
    confirmedIncomingQty,
    reservedQty,
    safetyBufferQty,
    availableToPreorder: Math.max(0, rawRemaining),
    overallocatedBy: Math.max(0, -rawRemaining),
  };
}
