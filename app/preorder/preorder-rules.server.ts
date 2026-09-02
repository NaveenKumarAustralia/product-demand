export const PREORDER_STATUS_ON_PRODUCTION = "on_production";
export const PREORDER_DESTINATION_AU = "send_to_au";
export const PREORDER_DESTINATION_USA = "send_to_usa";

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
  if (input.supplierStatus !== PREORDER_STATUS_ON_PRODUCTION) {
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
