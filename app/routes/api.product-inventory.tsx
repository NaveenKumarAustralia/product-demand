import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// On-demand Shopify variant inventory for a single product. Called when
// staff click the ▼ on a restock row to expand the "Shopify available"
// summary. Pulled out of the portal page loader so the restock page no
// longer waits for N sequential Shopify GraphQL round-trips before
// first paint.
//
// Optional `locationId` keeps this endpoint ready for separate AU / USA stock
// pools. When omitted, response behaviour remains backwards-compatible and
// sums all Shopify locations exactly as the existing portal expects.
//
// Response shape: { variantsBySize: { [variantTitle]: number }, total }.
// variantTitle is the same shape the restock UI's `sizes` list uses,
// so the caller can look up sizes[i] directly.

type InventoryResult = {
  variantsBySize: Record<string, number>;
  total: number;
  locationId?: string | null;
};

type CacheEntry = { expiresAt: number; value: InventoryResult };
const INVENTORY_CACHE_TTL_MS = 30 * 1000;
const INVENTORY_CACHE_MAX = 300;
const inventoryCache = new Map<string, CacheEntry>();

function cacheGet(key: string): InventoryResult | null {
  const entry = inventoryCache.get(key);
  if (!entry) return null;
  if (entry.expiresAt <= Date.now()) {
    inventoryCache.delete(key);
    return null;
  }
  return entry.value;
}

function cacheSet(key: string, value: InventoryResult) {
  if (inventoryCache.size >= INVENTORY_CACHE_MAX) {
    const oldest = inventoryCache.keys().next().value as string | undefined;
    if (oldest) inventoryCache.delete(oldest);
  }
  inventoryCache.set(key, { expiresAt: Date.now() + INVENTORY_CACHE_TTL_MS, value });
}

function normalizeLocationId(value: string | null) {
  const trimmed = value?.trim() ?? "";
  if (!trimmed) return null;
  if (trimmed.startsWith("gid://shopify/Location/")) return trimmed;
  const numeric = trimmed.replace(/[^0-9]/g, "");
  return numeric ? `gid://shopify/Location/${numeric}` : null;
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const url = new URL(request.url);
  const productId = url.searchParams.get("productId")?.trim();
  if (!productId) return Response.json({ variantsBySize: {}, total: 0 });
  const locationId = normalizeLocationId(url.searchParams.get("locationId"));

  const session = await prisma.session.findFirst({
    where: { accessToken: { not: "" } },
    orderBy: { isOnline: "asc" },
  }).catch(() => null);
  if (!session?.shop || !session.accessToken) {
    return Response.json({ variantsBySize: {}, total: 0, error: "no_session" });
  }

  const cacheKey = `${session.shop}:${productId}:${locationId ?? "all"}`;
  const cached = cacheGet(cacheKey);
  if (cached) {
    return Response.json(
      { ...cached, cached: true },
      { headers: { "Cache-Control": "private, max-age=15" } },
    );
  }

  const gql = `#graphql
    query ProductInventory($id: ID!) {
      product(id: $id) {
        variants(first: 100) {
          nodes {
            id
            title
            selectedOptions { name value }
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
      }
    }
  `;

  let json: { data?: { product?: { variants?: { nodes?: unknown[] } } } } | null = null;
  try {
    const res = await fetch(`https://${session.shop}/admin/api/2025-10/graphql.json`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-Shopify-Access-Token": session.accessToken,
      },
      body: JSON.stringify({ query: gql, variables: { id: productId } }),
    });
    if (res.ok) json = await res.json();
  } catch (e) {
    console.warn("[api.product-inventory] graphql failed:", e);
  }

  const nodes = (json?.data?.product?.variants?.nodes ?? []) as Array<{
    title?: string;
    selectedOptions?: Array<{ name?: string; value?: string }>;
    inventoryItem?: {
      inventoryLevels?: {
        nodes?: Array<{
          location?: { id?: string };
          quantities?: Array<{ name?: string; quantity?: number }>;
        }>;
      };
    };
  }>;

  // For a single-variant product whose only variant has no Size option,
  // relabel it "Free Size" so the restock sizes list lines up.
  const isFreeSize = nodes.length === 1 && (() => {
    const opts = nodes[0]?.selectedOptions ?? [];
    if (!opts.length) return true;
    const hasSize = opts.some((o) => (o?.name ?? "").trim().toLowerCase() === "size");
    if (hasSize) return false;
    return opts.every((o) => (o?.name ?? "") === "Title" && (o?.value ?? "") === "Default Title");
  })();

  const variantsBySize: Record<string, number> = {};
  let total = 0;
  for (const variant of nodes) {
    const title = isFreeSize ? "Free Size" : String(variant.title ?? "").trim();
    if (!title) continue;
    let qty = 0;
    for (const level of (variant.inventoryItem?.inventoryLevels?.nodes ?? [])) {
      if (locationId && level.location?.id !== locationId) continue;
      for (const q of (level.quantities ?? [])) {
        if (q?.name === "available" && Number.isFinite(Number(q.quantity))) {
          qty += Number(q.quantity);
        }
      }
    }
    variantsBySize[title] = qty;
    total += qty;
  }

  const value: InventoryResult = { variantsBySize, total, locationId };
  cacheSet(cacheKey, value);
  return Response.json(
    { ...value, cached: false },
    { headers: { "Cache-Control": "private, max-age=15" } },
  );
};
