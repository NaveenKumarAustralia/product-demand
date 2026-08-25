import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// On-demand Shopify variant SKU + barcode for a single product. Used by the
// "Print barcodes" popup to fill in any barcodes the portal is missing (they
// all exist in Shopify), one product at a time so nothing hangs.
// Response: { variants: [{ title, sku, barcode }] }.
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const productId = new URL(request.url).searchParams.get("productId")?.trim();
  if (!productId) return Response.json({ variants: [] });

  const session = await prisma.session.findFirst({
    where: { accessToken: { not: "" } },
    orderBy: { isOnline: "asc" },
  }).catch(() => null);
  if (!session?.shop || !session.accessToken) return Response.json({ variants: [], error: "no_session" });

  const numeric = productId.replace(/\D/g, "");
  const gid = productId.startsWith("gid://") ? productId : (numeric ? `gid://shopify/Product/${numeric}` : "");
  if (!gid) return Response.json({ variants: [] });

  const gql = `query VariantCodes($id: ID!) { product(id: $id) { variants(first: 100) { nodes { title sku barcode selectedOptions { name value } } } } }`;
  let json: { data?: { product?: { variants?: { nodes?: unknown[] } } } } | null = null;
  try {
    const res = await fetch(`https://${session.shop}/admin/api/2025-10/graphql.json`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": session.accessToken },
      body: JSON.stringify({ query: gql, variables: { id: gid } }),
    });
    if (res.ok) json = await res.json();
  } catch (e) {
    console.warn("[api.product-variant-codes] graphql failed:", e);
  }

  const nodes = (json?.data?.product?.variants?.nodes ?? []) as Array<{
    title?: string; sku?: string | null; barcode?: string | null;
    selectedOptions?: Array<{ name?: string; value?: string }>;
  }>;
  const isFreeSize = nodes.length === 1 && (() => {
    const opts = nodes[0]?.selectedOptions ?? [];
    if (!opts.length) return true;
    if (opts.some((o) => (o?.name ?? "").trim().toLowerCase() === "size")) return false;
    return opts.every((o) => (o?.name ?? "") === "Title" && (o?.value ?? "") === "Default Title");
  })();
  const variants = nodes.map((v) => ({
    title: isFreeSize ? "Free Size" : String(v.title ?? "").trim(),
    sku: v.sku ? String(v.sku) : "",
    barcode: v.barcode ? String(v.barcode) : "",
  })).filter((v) => v.title);

  return Response.json({ variants });
};
