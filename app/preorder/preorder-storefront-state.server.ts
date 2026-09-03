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

async function getPhysicalAvailable(shop: string, accessToken: string, variantId: string, locationId: string) {
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
    if (level.location?.id !== locationId) continue;
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

  const locations = await getPreorderLocationSettings();
  const locationId = normalizeLocationId(locations[input.market]);
  if (!locationId) {
    return {
      ok: false as const,
      reason: "location_not_configured" as const,
      market: input.market,
      variantId,
      state: { state: "notify_me" as const, physicalAvailable: 0 },
    };
  }

  const [accessToken, orders, registryEntries] = await Promise.all([
    getOfflineSession(shop),
    prisma.supplierOrder.findMany({
      where: {
        shop,
        status: "open",
        supplierStatus: "on_production",
        destination: destinationForMarket(input.market),
        lines: { some: { variantId } },
      },
      select: {
        id: true,
        shop: true,
        supplierStatus: true,
        destination: true,
        eta: true,
        createdAt: true,
        lines: {
          where: { variantId },
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
          where: { supplierOrderId: { in: orderIds }, variantId, status: "reserved" },
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
