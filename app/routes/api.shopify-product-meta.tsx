import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// Lists the distinct product TYPES (or TAGS) that already exist in the store,
// so the UI can offer an autocomplete/picker instead of free-typing.
//   /api/shopify-product-meta            → { types: string[] }
//   /api/shopify-product-meta?kind=tags  → { tags: string[] }
// Both are root StringConnection queries in the Admin API (each node is a
// plain string). Cached briefly so repeated opens don't re-hit Shopify.

let cache: { types?: { at: number; values: string[] }; tags?: { at: number; values: string[] } } = {};
const TTL_MS = 5 * 60 * 1000;

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const kind = new URL(request.url).searchParams.get("kind") === "tags" ? "tags" : "types";
  const cached = cache[kind];
  if (cached && Date.now() - cached.at < TTL_MS) {
    return Response.json({ [kind]: cached.values });
  }

  const session = await prisma.session.findFirst({
    where: { accessToken: { not: "" } },
    orderBy: { isOnline: "asc" },
  }).catch(() => null);
  if (!session?.shop || !session.accessToken) {
    return Response.json({ [kind]: [], error: "no_session" });
  }

  const field = kind === "tags" ? "productTags" : "productTypes";
  const gql = `#graphql
    query ProductMeta {
      ${field}(first: 250) { nodes }
    }
  `;

  let values: string[] = [];
  try {
    const res = await fetch(`https://${session.shop}/admin/api/2025-10/graphql.json`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": session.accessToken },
      body: JSON.stringify({ query: gql }),
    });
    if (res.ok) {
      const json = await res.json() as { data?: Record<string, { nodes?: unknown[] }> };
      const nodes = json?.data?.[field]?.nodes ?? [];
      values = (nodes as unknown[]).map((n) => String(n)).filter((s) => s.trim().length > 0);
      values.sort((a, b) => a.localeCompare(b));
      cache = { ...cache, [kind]: { at: Date.now(), values } };
    }
  } catch (e) {
    console.warn("[api.shopify-product-meta] graphql failed:", e);
  }

  return Response.json({ [kind]: values });
};
