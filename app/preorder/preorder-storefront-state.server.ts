import prisma from "../db.server";
import { getPreorderLocationSettings } from "./preorder-locations.server";
import { calculatePreorderCapacity, getPreorderEligibility } from "./preorder-rules.server";
import { getPreorderSellingPlanRegistryEntries } from "./preorder-selling-plan-registry.server";
import {
  resolveStorefrontVariantState,
  type StorefrontMarket,
  type StorefrontPreorderCandidate,
} from "./preorder-storefront-state";

const API_VERSION = "2025-10";

function variantGid(value: string) {
  const text = String(value ?? "").trim();
  if (!text) throw new Error("Shopify variant ID is required.");
  return text.startsWith("gid://shopify/ProductVariant/") ? text : `gid://shopify/ProductVariant/${text}`;
}

function normalizeLocationId(value: string | null) {
  const text = String(value ?? "").trim();
  if (!text) return null;
  if (text.startsWith("gid://shopify/Location/")) return text;
  const numeric = text.replace(/[^0-9]/g, "");
  return numeric ? `gid://shopify/Location/${numeric}` : null;
}

function destinationForMarket(market: StorefrontMarket) {
  return market === "USA" ? "send_to_usa" : "send_to_au";
}

async function getOfflineSession(shop: string) {
  const session = await prisma.session.findFirst({
    where: { shop, isOnline: false, accessToken: { not: "" } },
    orderBy: { expires: "desc" },
    select: { accessToken: true },
  });
  if (!session?.accessToken) throw new Error("Offline Shopify session is missing.");
  return session.accessToken;
}

async function getPhysicalAvailable(shop: string, accessToken: string, variantId: string, locationId: string | null) {
  const response = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "X-Shopify-Access-Token": accessToken,
    },
    body: JSON.stringify({
      query: `#graphql
        query PreorderStorefrontInventory($id: ID!) {
          productVariant(id: $id) {
            inventoryItem {
              inventoryLevels(first: 50) {
                nodes {
                  location { id }
                  quantities(names: ["available"]) { name quantity }
                }
              }
            }
          }
        }
      `,
      variables: { id: variantGid(variantId) },
    }),
  });
  if (!response.ok) throw new Error(`Shopify inventory returned HTTP ${response.status}.`);
  const json = await response.json() as {
    data?: {
      productVariant?: {
        inventoryItem?: {
          inventoryLevels?: {
            nodes?: Array<{
              location?: { id?: string };
              quantities?: Array<{ name?: string; quantity?: number }>;
            }>;
          };
        };
      };
    };
    errors?: Array<{ message?: string }>;
  };
  if (json.errors?.length) throw new Error(json.errors.map((error) => error.message || "Shopify GraphQL error").join("; "));

  let available = 0;
  for (const level of json.data?.productVariant?.inventoryItem?.inventoryLevels?.nodes ?? []) {
    // locationId null = sum EVERY location (total available). Used for markets
    // with no configured preorder location, to still tell in-stock from
    // out-of-stock for the notify-me fallback.
    if (locationId && level.location?.id !== locationId) continue;
    for (const quantity of level.quantities ?? []) {
      if (quantity.name === "available" && Number.isFinite(Number(quantity.quantity))) available += Number(quantity.quantity);
    }
  }
  return Math.max(0, Math.floor(available));
}

export async function getStorefrontPreorderState(input: {
  shop: string;
  variantId: string;
  market: StorefrontMarket;
}) {
  const shop = String(input.shop ?? "").trim();
  const variantId = String(input.variantId ?? "").trim();
  if (!shop || !variantId) throw new Error("Shop and variant are required.");

  // OrderLine.variantId is stored as a Shopify GID (gid://shopify/ProductVariant/N)
  // because it comes from the Admin GraphQL API, but the storefront sends the bare
  // numeric id. Match on every form so the batch line is actually found — without
  // this, an active batch with capacity still fell back to "notify me".
  const variantIdMatch = {
    in: Array.from(new Set([
      variantId,
      variantGid(variantId),
      variantId.replace(/[^0-9]/g, ""),
    ].filter(Boolean))),
  };

  const locations = await getPreorderLocationSettings();
  const locationId = normalizeLocationId(locations[input.market]);
  if (!locationId) {
    // This market has no preorder fulfilment location (e.g. USA before its 3PL
    // is set up), so PRE-ORDER isn't possible here. But we still want the
    // notify-me fallback for genuinely out-of-stock variants, everywhere. With
    // no market location to scope to, judge stock by TOTAL available across all
    // locations: in stock -> hide (normal buy button); out of stock -> notify_me.
    // (Both states are handled by the already-deployed block JS, so no theme
    // redeploy is needed.) The earlier bug returned notify_me WITHOUT checking
    // stock, so it showed on in-stock US variants too.
    try {
      const accessToken = await getOfflineSession(shop);
      const totalAvailable = await getPhysicalAvailable(shop, accessToken, variantId, null);
      return {
        ok: true as const,
        market: input.market,
        variantId,
        locationId: null,
        state: resolveStorefrontVariantState({ market: input.market, physicalAvailable: totalAvailable, candidates: [] }),
      };
    } catch (error) {
      console.warn("[preorder storefront] inventory check failed for unconfigured market:", error);
      // If we can't read inventory, hide rather than risk showing notify on an
      // in-stock item (which was the reported bug).
      return {
        ok: false as const,
        reason: "location_not_configured" as const,
        market: input.market,
        variantId,
        state: { state: "in_stock" as const, physicalAvailable: 0 },
      };
    }
  }

  const [accessToken, orders, registryEntries] = await Promise.all([
    getOfflineSession(shop),
    prisma.supplierOrder.findMany({
      where: {
        shop,
        status: "open",
        // Any production status qualifies (not just on_production) — the
        // per-candidate getPreorderEligibility() below applies the real rule
        // (denies On Order / Cancelled). Hardcoding on_production here made a
        // batch silently stop showing preorder once it moved to Ready/Shipment.
        destination: destinationForMarket(input.market),
        lines: { some: { variantId: variantIdMatch } },
      },
      select: {
        id: true,
        shop: true,
        supplierStatus: true,
        destination: true,
        eta: true,
        createdAt: true,
        lines: {
          where: { variantId: variantIdMatch },
          select: { variantId: true, qtyOrdered: true, qtyReceived: true },
        },
      },
    }),
    getPreorderSellingPlanRegistryEntries(shop),
  ]);

  const physicalAvailable = await getPhysicalAvailable(shop, accessToken, variantId, locationId);
  if (physicalAvailable > 0) {
    return {
      ok: true as const,
      market: input.market,
      variantId,
      locationId,
      state: resolveStorefrontVariantState({ market: input.market, physicalAvailable, candidates: [] }),
    };
  }

  const orderIds = orders.map((order) => order.id);
  const [settings, reservedGroups] = await Promise.all([
    orderIds.length
      ? prisma.preorderBatchSetting.findMany({ where: { supplierOrderId: { in: orderIds } } })
      : Promise.resolve([]),
    orderIds.length
      ? prisma.preorderReservation.groupBy({
          by: ["supplierOrderId"],
          where: { supplierOrderId: { in: orderIds }, variantId: variantIdMatch, status: "reserved" },
          _sum: { quantity: true },
        })
      : Promise.resolve([]),
  ]);

  const settingByOrder = new Map(settings.map((setting) => [setting.supplierOrderId, setting]));
  const registryByOrder = new Map(registryEntries.map((entry) => [entry.supplierOrderId, entry]));
  const reservedByOrder = new Map(reservedGroups.map((row) => [row.supplierOrderId, row._sum.quantity ?? 0]));

  const candidates: StorefrontPreorderCandidate[] = orders.flatMap((order) => {
    const line = order.lines[0];
    const setting = settingByOrder.get(order.id);
    const registry = registryByOrder.get(order.id);
    if (!line) return [];
    const eligibility = getPreorderEligibility({
      supplierStatus: order.supplierStatus,
      destination: order.destination,
      preorderEnabled: setting?.enabled ?? false,
    });
    const capacity = calculatePreorderCapacity({
      confirmedIncomingQty: Math.max(0, line.qtyOrdered - line.qtyReceived),
      reservedQty: reservedByOrder.get(order.id) ?? 0,
      safetyBufferPercent: setting?.safetyBufferPercent ?? 0,
      safetyBufferQty: setting?.safetyBufferQty ?? null,
    });
    return [{
      batchId: order.id,
      market: input.market,
      eligible: eligibility.eligible,
      enabled: setting?.enabled ?? false,
      shopifySellingPlanActive: Boolean(registry),
      sellingPlanId: registry?.sellingPlanId ?? null,
      expectedShipDate: (setting?.shipDate ?? order.eta)?.toISOString() ?? null,
      availableToPreorder: capacity.availableToPreorder,
    }];
  });

  return {
    ok: true as const,
    market: input.market,
    variantId,
    locationId,
    state: resolveStorefrontVariantState({ market: input.market, physicalAvailable: 0, candidates }),
  };
}
