import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { unauthenticated } from "../shopify.server";

const VALID_SIZES = new Set(["XS", "S", "M", "L", "XL", "2XL", "3XL", "S-M", "M-L", "L-XL"]);

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const url = new URL(request.url);
  const query = url.searchParams.get("q")?.trim() ?? "";

  if (query.length < 2) {
    return Response.json({ products: [] });
  }

  const session = await prisma.session.findFirst({
    where: { accessToken: { not: "" } },
    orderBy: { isOnline: "asc" },
  }).catch(() => null);

  if (!session?.shop || !session.accessToken) {
    return Response.json({ products: [], error: "no_session" });
  }

  const escaped = query.replace(/[\\"]/g, "\\$&");
  const gqlQuery = `#graphql
    query PackingSearch($query: String) {
      products(first: 20, query: $query, sortKey: TITLE) {
        edges {
          node {
            id
            title
            featuredImage { url }
            variants(first: 100) {
              edges {
                node {
                  id
                  title
                  sku
                  selectedOptions { name value }
                }
              }
            }
          }
        }
      }
    }
  `;

  const mapJson = (json: any, shop: string) =>
    (json?.data?.products?.edges ?? []).map((edge: any) => {
      const seen = new Set<string>();
      const variants = (edge.node.variants?.edges ?? [])
        .map((e: any) => e.node)
        .map((v: any) => {
          const size = (v.selectedOptions as { name: string; value: string }[] | undefined)
            ?.map((o) => o.value)
            .find((val) => VALID_SIZES.has(val)) ?? null;
          return { raw: v, size };
        })
        .filter(({ size }: { size: string | null }) => {
          if (!size || seen.has(size)) return false;
          seen.add(size);
          return true;
        })
        .map(({ raw: v, size }: { raw: any; size: string }) => ({
          id: String(v.id ?? ""),
          title: size,
          sku: v.sku ? String(v.sku) : null,
          availableInventory: null,
        }))
        .filter((v: any) => v.id && v.title);

      return {
        id: edge.node.id,
        shop,
        title: edge.node.title,
        imageUrl: edge.node.featuredImage?.url ?? null,
        skus: Array.from(new Set(variants.map((v: any) => v.sku).filter(Boolean))),
        sizes: variants.map((v: any) => v.title),
        variants,
      };
    });

  const shopifyQuery = `title:*${escaped}* OR sku:*${escaped}*`;
  const needle = query.toLowerCase();
  const matchesLocally = (p: any) =>
    p.title.toLowerCase().includes(needle) || p.skus.some((s: string) => s.toLowerCase().includes(needle));

  const directFetch = async (q: string | null) => {
    try {
      const resp = await fetch(`https://${session.shop}/admin/api/2025-10/graphql.json`, {
        method: "POST",
        headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": session.accessToken },
        body: JSON.stringify({ query: gqlQuery, variables: { query: q } }),
      });
      if (!resp.ok) return [];
      return mapJson(await resp.json(), session.shop);
    } catch {
      return [];
    }
  };

  try {
    const { admin } = await unauthenticated.admin(session.shop);
    const resp = await admin.graphql(gqlQuery, { variables: { query: shopifyQuery } });
    const products = mapJson(await resp.json(), session.shop);
    if (products.length) return Response.json({ products });

    const fallbackResp = await admin.graphql(gqlQuery, { variables: { query: null } });
    const fallback = mapJson(await fallbackResp.json(), session.shop).filter(matchesLocally).slice(0, 8);
    return Response.json({ products: fallback });
  } catch {
    const products = await directFetch(shopifyQuery);
    if (products.length) return Response.json({ products });
    const fallback = (await directFetch(null)).filter(matchesLocally).slice(0, 8);
    return Response.json({ products: fallback });
  }
};
