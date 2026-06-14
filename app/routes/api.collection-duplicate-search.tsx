import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// Returns recently-created Shopify products, optionally filtered by
// product type. Used by the Collections "Duplicate From" picker so
// the user can pick an existing product of the same style and
// auto-populate a new row's Description / Tags / Type / HS Code /
// Country / Compare-at-price.
//
// Query params:
//   q                – free-text search across title
//   productType      – exact product_type filter (preferred when known)
//   limit            – cap on results (default 20)

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const url = new URL(request.url);
  const q = url.searchParams.get("q")?.trim() ?? "";
  const productType = url.searchParams.get("productType")?.trim() ?? "";
  const limit = Math.max(1, Math.min(50, Number(url.searchParams.get("limit") ?? 20)));

  const session = await prisma.session.findFirst({
    where: { accessToken: { not: "" } },
    orderBy: { isOnline: "asc" },
  }).catch(() => null);
  if (!session?.shop || !session.accessToken) return Response.json({ products: [], error: "no_session" });

  // Build Shopify filter clauses. Order by created_at desc so the
  // most recent style match comes first — which is exactly what the
  // user wants when duplicating from a recent product.
  const filterClauses: string[] = [];
  if (productType) filterClauses.push(`product_type:"${productType.replace(/"/g, '\\"')}"`);
  if (q) filterClauses.push(`title:*${q.replace(/"/g, '\\"')}*`);
  const filterStr = filterClauses.join(" ");

  const gql = `#graphql
    query CollectionDuplicateSearch($filter: String, $limit: Int!) {
      products(first: $limit, query: $filter, sortKey: CREATED_AT, reverse: true) {
        nodes {
          id
          handle
          title
          productType
          createdAt
          featuredMedia { preview { image { url(transform: { maxWidth: 80, maxHeight: 80 }) } } }
        }
      }
    }
  `;

  let json: { data?: { products?: { nodes?: unknown[] } } } | null = null;
  try {
    const res = await fetch(`https://${session.shop}/admin/api/2025-10/graphql.json`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-Shopify-Access-Token": session.accessToken,
      },
      body: JSON.stringify({ query: gql, variables: { filter: filterStr || null, limit } }),
    });
    if (res.ok) json = await res.json();
  } catch (e) {
    console.warn("[duplicate-search] fetch failed:", e);
  }

  const nodes = (json?.data?.products?.nodes ?? []) as Array<{
    id?: string; handle?: string; title?: string; productType?: string; createdAt?: string;
    featuredMedia?: { preview?: { image?: { url?: string } } };
  }>;
  return Response.json({
    products: nodes.map((p) => ({
      id: String(p.id ?? ""),
      handle: String(p.handle ?? ""),
      title: String(p.title ?? ""),
      productType: String(p.productType ?? ""),
      createdAt: p.createdAt ?? "",
      thumbnail: p.featuredMedia?.preview?.image?.url ?? "",
    })),
  });
};
