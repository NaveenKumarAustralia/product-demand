import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// TEMP diagnostic/backfill — compares each open order's stored productType with
// its actual Shopify product type. ?apply=1 backfills order.productType from
// Shopify where the stored value is blank. Remove once grouping is fixed.
function toGid(productId: string): string {
  const p = (productId ?? "").trim();
  if (p.startsWith("gid://")) return p;
  const n = p.replace(/\D/g, "");
  return n ? `gid://shopify/Product/${n}` : "";
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const apply = new URL(request.url).searchParams.get("apply") === "1";
  const session = await prisma.session.findFirst({ where: { accessToken: { not: "" } }, orderBy: { isOnline: "asc" }, select: { shop: true, accessToken: true } }).catch(() => null);
  if (!session?.shop || !session.accessToken) return Response.json({ error: "no session" });

  const orders = await prisma.supplierOrder.findMany({
    where: { status: "open", productId: { not: "" } },
    select: { id: true, productTitle: true, productType: true, productId: true },
  });
  const gidToOrders = new Map<string, typeof orders>();
  for (const o of orders) {
    const gid = toGid(o.productId);
    if (!gid) continue;
    const list = gidToOrders.get(gid) ?? [];
    list.push(o); gidToOrders.set(gid, list);
  }
  const gids = Array.from(gidToOrders.keys());
  const shopifyType = new Map<string, string>();
  for (let i = 0; i < gids.length; i += 100) {
    const batch = gids.slice(i, i + 100);
    const res = await fetch(`https://${session.shop}/admin/api/2025-10/graphql.json`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": session.accessToken },
      body: JSON.stringify({ query: `query($ids:[ID!]!){ nodes(ids:$ids){ ... on Product { id productType } } }`, variables: { ids: batch } }),
    }).then((r) => r.json()).catch(() => null) as { data?: { nodes?: Array<{ id?: string; productType?: string } | null> } } | null;
    for (const n of res?.data?.nodes ?? []) { if (n?.id) shopifyType.set(n.id, (n.productType ?? "").trim()); }
  }

  const rows = orders.map((o) => ({ id: o.id, title: o.productTitle, stored: o.productType, shopify: shopifyType.get(toGid(o.productId)) ?? "" }));
  const missingStored = rows.filter((r) => !(r.stored ?? "").trim());
  const missingButHaveShopify = missingStored.filter((r) => r.shopify);
  const shopifyEmptyToo = missingStored.filter((r) => !r.shopify);

  let updated = 0;
  if (apply) {
    for (const r of missingButHaveShopify) {
      await prisma.supplierOrder.update({ where: { id: r.id }, data: { productType: r.shopify } }).then(() => { updated++; }).catch(() => {});
    }
  }

  return Response.json({
    openWithProductId: orders.length,
    missingStoredCount: missingStored.length,
    missingButHaveShopifyCount: missingButHaveShopify.length,
    shopifyAlsoEmptyCount: shopifyEmptyToo.length,
    applied: apply ? updated : "dry-run (add ?apply=1 to backfill)",
    sampleMissingButHaveShopify: missingButHaveShopify.slice(0, 20),
    sampleShopifyAlsoEmpty: shopifyEmptyToo.slice(0, 20).map((r) => ({ id: r.id, title: r.title })),
  });
};
