import prisma from "./db.server";

// Shopify → Collections one-way sync for LOCKED rows. When a product is saved in
// Shopify (products/update webhook), any collection row that is linked to that
// product AND locked ("Shopify is the source of truth") pulls the latest
// descriptive fields + category metafields. Unlocked rows (mid-edit in the
// portal) are left alone. Name, Price (RRP), Price ₹, size quantities, SKU/
// barcode and other portal-managed fields are never overwritten.

const API_VERSION = "2025-10";
const numeric = (s: unknown) => String(s ?? "").replace(/\D/g, "");

// Reserved row keys (kept in sync with portal._index.tsx COL_ROW_* constants).
const K_PRODUCT_ID = "__shopifyProductId";
const K_LOCKED = "__shopifyLocked";
const K_CATEGORY = "__categoryMetafields";

async function gql(shop: string, token: string, query: string, variables: Record<string, unknown>) {
  const res = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
    method: "POST",
    headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": token },
    body: JSON.stringify({ query, variables }),
    signal: AbortSignal.timeout(10_000),
  });
  if (!res.ok) throw new Error(`Shopify HTTP ${res.status}`);
  return res.json() as Promise<any>;
}

// Same shape as portal's buildCategoryMetafieldBlob: category metafields live in
// the reserved `shopify` namespace (filtering the query by it returns nothing, so
// we fetch unfiltered and filter here).
function buildCategoryBlob(product: any): { categoryId: string; categoryName: string; metafields: any[] } | null {
  const categoryId = String(product?.category?.id ?? "");
  const nodes: any[] = product?.metafields?.nodes ?? [];
  const metafields = nodes
    .filter((m) => String(m?.namespace ?? "") === "shopify" && String(m?.value ?? "").trim())
    .map((m) => {
      const refs = [...(m.references?.nodes ?? []), ...(m.reference ? [m.reference] : [])].filter(Boolean);
      return {
        key: String(m.key ?? ""),
        type: String(m.type ?? ""),
        value: String(m.value ?? ""),
        names: refs.map((r: any) => String(r?.displayName ?? "")).filter(Boolean),
        refType: (refs.find((r: any) => r?.type)?.type as string | undefined) ?? null,
        label: String(m?.definition?.name ?? "") || undefined,
      };
    })
    .filter((m) => m.key && m.type && m.value);
  if (!categoryId && !metafields.length) return null;
  return { categoryId, categoryName: String(product?.category?.fullName ?? ""), metafields };
}

export async function syncLockedCollectionRowsForProduct(shop: string, productId: unknown): Promise<{ updated: number }> {
  const num = numeric(productId);
  if (!num || !shop) return { updated: 0 };

  // 1) Cheap first: is this product linked to any LOCKED row anywhere? If not,
  //    return without touching Shopify (products/update fires often).
  const collections = await prisma.collection.findMany({ select: { id: true, rows: true } });
  const matched = collections.filter((c) =>
    (Array.isArray(c.rows) ? (c.rows as Array<Record<string, unknown>>) : []).some(
      (r) => numeric(r?.[K_PRODUCT_ID]) === num && String(r?.[K_LOCKED] ?? "") === "1",
    ),
  );
  if (!matched.length) return { updated: 0 };

  const session = await prisma.session.findFirst({ where: { shop, accessToken: { not: "" } }, orderBy: { isOnline: "asc" }, select: { accessToken: true } }).catch(() => null);
  if (!session?.accessToken) return { updated: 0 };

  // 2) Pull the current descriptive fields + category metafields for this product.
  const json = await gql(shop, session.accessToken, `
    query CollectionsSyncPull($id: ID!) {
      product(id: $id) {
        id descriptionHtml productType vendor tags
        seo { title description }
        category { id fullName }
        metafields(first: 250) {
          nodes {
            namespace key type value
            definition { name }
            references(first: 40) { nodes { __typename ... on Metaobject { id displayName type } } }
            reference { __typename ... on Metaobject { id displayName type } }
          }
        }
        variants(first: 1) { nodes { compareAtPrice inventoryItem { harmonizedSystemCode countryCodeOfOrigin } } }
      }
    }
  `, { id: `gid://shopify/Product/${num}` }).catch(() => null);
  const product = json?.data?.product;
  if (!product?.id) return { updated: 0 };

  const v0 = product.variants?.nodes?.[0] ?? {};
  const shopTags = Array.isArray(product.tags) ? product.tags.map((t: unknown) => String(t).trim()).filter(Boolean) : [];
  const pulled: Record<string, string> = {
    description: String(product.descriptionHtml ?? ""),
    productType: String(product.productType ?? ""),
    vendor: String(product.vendor ?? ""),
    seoTitle: String(product.seo?.title ?? ""),
    seoDescription: String(product.seo?.description ?? ""),
    compareAtPrice: v0.compareAtPrice ? String(v0.compareAtPrice) : "",
    hsCode: String(v0.inventoryItem?.harmonizedSystemCode ?? ""),
    countryOfOrigin: String(v0.inventoryItem?.countryCodeOfOrigin ?? ""),
  };
  // Tags: keep it simple — mirror Shopify's tags (the push already round-trips
  // pre-order tags), but never remove the pre-order system tags if present.
  {
    const kept = shopTags.filter((t: string) => { const s = t.toLowerCase(); return s !== "pre-order" && !s.startsWith("pre-order: ") && !s.startsWith("pre-order ships "); });
    if (product.tags !== undefined) pulled.tags = kept.join(", ");
  }
  const catBlob = buildCategoryBlob(product);
  const catStr = catBlob ? JSON.stringify(catBlob) : "";

  // 3) Update the locked linked rows in each matched collection.
  let updated = 0;
  for (const c of matched) {
    const rows = Array.isArray(c.rows) ? (c.rows as Array<Record<string, unknown>>) : [];
    let changed = false;
    const next = rows.map((row) => {
      if (numeric(row?.[K_PRODUCT_ID]) !== num || String(row?.[K_LOCKED] ?? "") !== "1") return row;
      changed = true; updated += 1;
      return { ...row, ...pulled, ...(catBlob ? { [K_CATEGORY]: catStr } : {}) };
    });
    if (changed) await prisma.collection.update({ where: { id: c.id }, data: { rows: next as unknown as object, updatedAt: new Date() } }).catch(() => {});
  }
  return { updated };
}
