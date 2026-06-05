import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// On-demand Shopify variant inventory for a single product. Called when
// staff click the ▼ on a restock row to expand the "Shopify available"
// summary. Pulled out of the portal page loader so the restock page no
// longer waits for N sequential Shopify GraphQL round-trips before
// first paint.
//
// Response shape: { variantsBySize: { [variantTitle]: number }, total }.
// variantTitle is the same shape the restock UI's `sizes` list uses,
// so the caller can look up sizes[i] directly.

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const url = new URL(request.url);
  const productId = url.searchParams.get("productId")?.trim();
  if (!productId) return Response.json({ variantsBySize: {}, total: 0 });

  const session = await prisma.session.findFirst({
    where: { accessToken: { not: "" } },
    orderBy: { isOnline: "asc" },
  }).catch(() => null);
  if (!session?.shop || !session.accessToken) {
    return Response.json({ variantsBySize: {}, total: 0, error: "no_session" });
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
              inventoryLevels(first: 20) {
                nodes {
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
    inventoryItem?: { inventoryLevels?: { nodes?: Array<{ quantities?: Array<{ name?: string; quantity?: number }> }> } };
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
      for (const q of (level.quantities ?? [])) {
        if (q?.name === "available" && Number.isFinite(Number(q.quantity))) {
          qty += Number(q.quantity);
        }
      }
    }
    variantsBySize[title] = qty;
    total += qty;
  }

  return Response.json({ variantsBySize, total });
};
