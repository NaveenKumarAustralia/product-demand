const KLAVIYO_API_URL = "https://a.klaviyo.com/api/events";
const KLAVIYO_REVISION = "2026-07-15";

export const KLAVIYO_BACK_IN_STOCK_METRIC = "Karma East Back In Stock Available";
export const KLAVIYO_PREORDER_UPDATE_METRIC = "Karma East Preorder Update";
export const KLAVIYO_PREORDER_PLACED_METRIC = "Karma East Pre-order Placed";

export type KlaviyoEventInput = {
  email: string;
  metric: string;
  uniqueId: string;
  properties: Record<string, unknown>;
};

export type KlaviyoConnectionStatus = {
  configured: boolean;
  revision: string;
};

export function getKlaviyoConnectionStatus(): KlaviyoConnectionStatus {
  return {
    configured: Boolean(process.env.KLAVIYO_PRIVATE_API_KEY?.trim()),
    revision: KLAVIYO_REVISION,
  };
}

function requireKlaviyoKey() {
  const key = process.env.KLAVIYO_PRIVATE_API_KEY?.trim();
  if (!key) throw new Error("Klaviyo is not connected. Add KLAVIYO_PRIVATE_API_KEY in Railway first.");
  return key;
}

function normalizeEmail(value: string) {
  const email = value.trim().toLowerCase();
  if (!email || !email.includes("@") || email.length > 100) {
    throw new Error("A valid customer email is required for Klaviyo.");
  }
  return email;
}

export async function createKlaviyoEvent(input: KlaviyoEventInput) {
  const key = requireKlaviyoKey();
  const email = normalizeEmail(input.email);
  const metric = input.metric.trim();
  const uniqueId = input.uniqueId.trim();
  if (!metric) throw new Error("A Klaviyo metric name is required.");
  if (!uniqueId) throw new Error("A unique Klaviyo event ID is required.");

  const response = await fetch(KLAVIYO_API_URL, {
    method: "POST",
    headers: {
      Authorization: `Klaviyo-API-Key ${key}`,
      Accept: "application/vnd.api+json",
      "Content-Type": "application/vnd.api+json",
      revision: KLAVIYO_REVISION,
    },
    body: JSON.stringify({
      data: {
        type: "event",
        attributes: {
          properties: input.properties,
          unique_id: uniqueId,
          metric: {
            data: {
              type: "metric",
              attributes: { name: metric },
            },
          },
          profile: {
            data: {
              type: "profile",
              attributes: { email },
            },
          },
        },
      },
    }),
    signal: AbortSignal.timeout(10_000),
  });

  if (response.status !== 202) {
    const detail = await response.text().catch(() => "");
    throw new Error(`Klaviyo rejected the event (${response.status})${detail ? `: ${detail.slice(0, 500)}` : ""}`);
  }

  return { accepted: true as const, email, metric, uniqueId };
}

/**
 * Fire a "Pre-order Placed" event the moment the app confirms an order contains
 * a pre-order (i.e. a held/reserved pre-order line). This is how NO-PLAN orders
 * (Shop Pay / PayPal / express — which Shopify's confirmation email cannot flag)
 * tell the customer they've bought a pre-order: a Klaviyo flow on this metric
 * sends a branded "your order includes a pre-order — ships <date>" email. The
 * app is the source of truth (it held the order), so this is 100% reliable on
 * every checkout path. Idempotent per order via unique_id.
 */
export function sendPreorderPlacedEvent(input: {
  shop: string;
  orderId: string;
  orderName: string | null;
  email: string;
  market: "AU" | "USA";
  items: Array<{ title: string | null; size: string | null; dispatch: string | null }>;
}) {
  return createKlaviyoEvent({
    email: input.email,
    metric: KLAVIYO_PREORDER_PLACED_METRIC,
    uniqueId: `preorder-placed:${input.shop}:${input.orderId}`,
    properties: {
      shop: input.shop,
      order_id: input.orderId,
      order_name: input.orderName ?? null,
      market: input.market,
      // First item's date drives the headline; the full list is included so the
      // Klaviyo template can list each pre-order line with its own dispatch date.
      dispatch: input.items.find((i) => i.dispatch)?.dispatch ?? null,
      items: input.items,
      source: "karma-east-production-portal",
    },
  });
}

export function sendBackInStockAvailableEvent(input: {
  waitlistId: number;
  email: string;
  shop: string;
  market: "AU" | "USA";
  productId?: string | null;
  productTitle?: string | null;
  variantId: string;
  variantTitle?: string | null;
  sku?: string | null;
  productUrl?: string | null;
  availableQuantity?: number | null;
}) {
  return createKlaviyoEvent({
    email: input.email,
    metric: KLAVIYO_BACK_IN_STOCK_METRIC,
    uniqueId: `back-in-stock:${input.shop}:${input.waitlistId}`,
    properties: {
      waitlist_id: input.waitlistId,
      shop: input.shop,
      market: input.market,
      product_id: input.productId ?? null,
      product_title: input.productTitle ?? null,
      variant_id: input.variantId,
      variant_title: input.variantTitle ?? null,
      sku: input.sku ?? null,
      product_url: input.productUrl ?? null,
      available_quantity: input.availableQuantity ?? null,
      source: "karma-east-production-portal",
    },
  });
}
