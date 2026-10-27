import prisma from "../db.server";
import { getPreorderLocationSettings } from "./preorder-locations.server";
import { getPreorderNotifyEnabled, getBackInStockThreshold } from "./preorder-storefront-settings.server";
import { sendBackInStockAvailableEvent } from "./preorder-klaviyo.server";
import { calculatePreorderCapacity, getPreorderEligibility } from "./preorder-rules.server";
import { getPreorderSellingPlanRegistryEntries } from "./preorder-selling-plan-registry.server";
import {
  resolveStorefrontVariantState,
  type StorefrontMarket,
  type StorefrontPreorderCandidate,
  type StorefrontVariantState,
} from "./preorder-storefront-state";

const API_VERSION = "2025-10";

function variantGid(value: string) {
  const text = String(value ?? "").trim();
  if (!text) throw new Error("Shopify variant ID is required.");
  return text.startsWith("gid://shopify/ProductVariant/") ? text : `gid://shopify/ProductVariant/${text}`;
}

function productGid(value: string) {
  const text = String(value ?? "").trim();
  if (!text) return null;
  return text.startsWith("gid://shopify/Product/") ? text : `gid://shopify/Product/${text.replace(/[^0-9]/g, "")}`;
}

// Fetch a product's public storefront URL so the back-in-stock email can deep
// link straight to the product ("Shop now"). Best-effort — returns null if the
// product is not published to the Online Store or on any error.
async function getProductOnlineUrl(shop: string, accessToken: string, productId: string | null): Promise<string | null> {
  const gid = productGid(productId ?? "");
  if (!gid) return null;
  try {
    const response = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": accessToken },
      body: JSON.stringify({
        query: `#graphql
          query BackInStockProductUrl($id: ID!) { product(id: $id) { onlineStoreUrl } }
        `,
        variables: { id: gid },
      }),
    });
    if (!response.ok) return null;
    const json = await response.json() as { data?: { product?: { onlineStoreUrl?: string | null } } };
    const url = json.data?.product?.onlineStoreUrl;
    return typeof url === "string" && url.trim() ? url.trim() : null;
  } catch {
    return null;
  }
}

// Convenience for callers that only have the shop + productId (e.g. the manual
// "Send via Klaviyo" endpoint): resolves the offline session itself. Null on any
// failure so sending never blocks on a missing URL.
export async function resolveProductOnlineUrl(shop: string, productId: string | null): Promise<string | null> {
  try {
    const accessToken = await getOfflineSession(shop);
    return await getProductOnlineUrl(shop, accessToken, productId);
  } catch {
    return null;
  }
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

// When the merchant turns the storefront notify-me block off (e.g. they run a
// separate back-in-stock app), collapse any notify_me result to in_stock so the
// block hides. Pre-order results pass through untouched.
function suppressNotify(state: StorefrontVariantState, notifyEnabled: boolean): StorefrontVariantState {
  if (!notifyEnabled && state.state === "notify_me") return { state: "in_stock", physicalAvailable: 0 };
  return state;
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

  const notifyEnabled = await getPreorderNotifyEnabled();
  const locations = await getPreorderLocationSettings();
  // A market with no fulfilment location of its own (e.g. USA before its 3PL is
  // set up) is fulfilled from the AU batch — so show those shoppers the AU
  // pre-order (button + selling plan + date) rather than only "notify me". We
  // compute against the AU location/batches but keep responding as the requested
  // market, so the order still carries the plan (which makes the confirmation
  // email flag it). When the market later gets its own location, this fallback
  // stops and it uses its own batches.
  let effectiveMarket: StorefrontMarket = input.market;
  let locationId = normalizeLocationId(locations[input.market]);
  if (!locationId && input.market !== "AU") {
    const auLocation = normalizeLocationId(locations.AU);
    if (auLocation) { effectiveMarket = "AU"; locationId = auLocation; }
  }
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
        state: suppressNotify(resolveStorefrontVariantState({ market: input.market, physicalAvailable: totalAvailable, candidates: [] }), notifyEnabled),
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
        destination: destinationForMarket(effectiveMarket),
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
    state: suppressNotify(resolveStorefrontVariantState({ market: input.market, physicalAvailable: 0, candidates }), notifyEnabled),
  };
}

// AUTO back-in-stock: when a product's inventory changes, send the Klaviyo
// "back in stock" event to everyone still WAITING on one of its variants —
// but only once that variant's stock (at the waiter's market location) has
// reached the configured threshold. Marks each notified so it never re-sends.
// Called from the products/update webhook (background). Never throws.
type BackInStockRow = {
  id: number; email: string; productId: string | null; productTitle: string | null;
  variantId: string; variantTitle: string | null; sku: string | null; market: string; status: string;
};
export async function notifyBackInStockForProduct(shop: string, productId: unknown): Promise<{ sent: number; waiting: number }> {
  try {
    const num = String(productId ?? "").replace(/\D/g, "");
    if (!shop || !num) return { sent: 0, waiting: 0 };
    // Only this product's still-waiting signups (productId is stored as gid OR
    // numeric, so match on the numeric part).
    const rows = await prisma.$queryRawUnsafe<BackInStockRow[]>(
      `SELECT "id","email","productId","productTitle","variantId","variantTitle","sku","market","status"
       FROM "PreorderWaitlist"
       WHERE "shop" = $1 AND "status" = 'waiting'
         AND regexp_replace(COALESCE("productId", ''), '[^0-9]', '', 'g') = $2`,
      shop, num,
    );
    if (!rows.length) return { sent: 0, waiting: 0 };

    const threshold = await getBackInStockThreshold();
    const locations = await getPreorderLocationSettings();
    let accessToken: string;
    try { accessToken = await getOfflineSession(shop); } catch { return { sent: 0, waiting: rows.length }; }

    // All rows are the same product (queried by productId), so resolve its public
    // storefront URL once for the "Shop now" button.
    const productUrl = await getProductOnlineUrl(shop, accessToken, rows[0].productId);

    // Group waiters by variant + market (stock is checked per market location).
    const groups = new Map<string, BackInStockRow[]>();
    for (const r of rows) {
      const key = `${r.variantId}|${r.market}`;
      const list = groups.get(key); if (list) list.push(r); else groups.set(key, [r]);
    }

    let sent = 0;
    for (const grp of groups.values()) {
      const market: StorefrontMarket = grp[0].market === "USA" ? "USA" : "AU";
      // Same market→location rule as the storefront state (USA falls back to AU;
      // no location = judge by total stock across all locations).
      let locationId = normalizeLocationId(locations[market]);
      if (!locationId && market !== "AU") locationId = normalizeLocationId(locations.AU);
      let available = 0;
      try { available = await getPhysicalAvailable(shop, accessToken, grp[0].variantId, locationId); } catch { continue; }
      if (available < threshold) continue; // not enough stock yet — keep them waiting
      for (const r of grp) {
        try {
          await sendBackInStockAvailableEvent({
            waitlistId: r.id, email: r.email, shop, market,
            productId: r.productId, productTitle: r.productTitle,
            variantId: r.variantId, variantTitle: r.variantTitle, sku: r.sku,
            productUrl, availableQuantity: available,
          });
          await prisma.$executeRawUnsafe(
            `UPDATE "PreorderWaitlist" SET "status" = 'notified', "notifiedAt" = CURRENT_TIMESTAMP, "updatedAt" = CURRENT_TIMESTAMP WHERE "id" = $1 AND "status" = 'waiting'`,
            r.id,
          );
          sent += 1;
        } catch (e) { console.warn("[backinstock] send failed for waitlist", r.id, e); }
      }
    }
    if (sent) console.log(`[backinstock] ${shop} product ${num}: notified ${sent}/${rows.length} waiting (threshold ${threshold})`);
    return { sent, waiting: rows.length };
  } catch (e) {
    console.warn("[backinstock] notifyBackInStockForProduct failed:", e);
    return { sent: 0, waiting: 0 };
  }
}
