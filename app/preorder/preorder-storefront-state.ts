export type StorefrontMarket = "AU" | "USA";

export type StorefrontPreorderCandidate = {
  batchId: number;
  market: StorefrontMarket;
  eligible: boolean;
  enabled: boolean;
  shopifySellingPlanActive: boolean;
  sellingPlanId: string | null;
  expectedShipDate: string | null;
  availableToPreorder: number;
};

export type StorefrontVariantState =
  | { state: "in_stock"; physicalAvailable: number }
  | {
      state: "preorder";
      physicalAvailable: 0;
      batchId: number;
      sellingPlanId: string;
      expectedShipDate: string | null;
      availableToPreorder: number;
    }
  | { state: "notify_me"; physicalAvailable: 0 };

function shipTimestamp(value: string | null) {
  if (!value) return Number.MAX_SAFE_INTEGER;
  const time = new Date(value).getTime();
  return Number.isNaN(time) ? Number.MAX_SAFE_INTEGER : time;
}

export function resolveStorefrontVariantState(input: {
  market: StorefrontMarket;
  physicalAvailable: number;
  candidates: StorefrontPreorderCandidate[];
}): StorefrontVariantState {
  const physicalAvailable = Math.max(0, Math.floor(Number(input.physicalAvailable) || 0));
  if (physicalAvailable > 0) return { state: "in_stock", physicalAvailable };

  const candidate = input.candidates
    .filter((row) =>
      row.market === input.market &&
      row.eligible &&
      row.enabled &&
      row.shopifySellingPlanActive &&
      Boolean(row.sellingPlanId) &&
      row.availableToPreorder > 0,
    )
    .sort((a, b) => shipTimestamp(a.expectedShipDate) - shipTimestamp(b.expectedShipDate) || a.batchId - b.batchId)[0];

  if (!candidate?.sellingPlanId) return { state: "notify_me", physicalAvailable: 0 };

  return {
    state: "preorder",
    physicalAvailable: 0,
    batchId: candidate.batchId,
    sellingPlanId: candidate.sellingPlanId,
    expectedShipDate: candidate.expectedShipDate,
    availableToPreorder: Math.max(0, Math.floor(candidate.availableToPreorder)),
  };
}
