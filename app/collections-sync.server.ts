import prisma from "./db.server";

// Shopify → Collections one-way sync. When a product is saved in Shopify
// (products/update webhook), its linked collection rows pull the latest from
// Shopify:
//  • LOCKED rows ("Shopify is the source of truth") get the FULL pull (incl.
//    price + images), as before.
//  • UNLOCKED rows now also mirror Shopify's DESCRIPTIVE fields (description,
//    SEO, category, tags, type, vendor, status, HS/country) so a Shopify edit
//    shows up in the portal without re-locking — but NEVER price, images, name,
//    quantities or SKU/barcode (portal-managed), and never a field the user has
//    edited in the portal but not yet pushed (tracked in __shopifyEditedFields).

const API_VERSION = "2025-10";
const numeric = (s: unknown) => String(s ?? "").replace(/\D/g, "");

// Reserved row keys (kept in sync with portal._index.tsx COL_ROW_* constants).
const K_PRODUCT_ID = "__shopifyProductId";
const K_LOCKED = "__shopifyLocked";
const K_CATEGORY = "__categoryMetafields";
const K_EDITED = "__shopifyEditedFields";
const K_STATUS = "__shopifyStatus";

// Descriptive fields safe to mirror into an UNLOCKED row (Shopify → portal).
// Excludes price (compareAtPrice) and images (modelPicture) — those stay
// portal-managed; name / quantities / SKU aren't in the pulled set at all.
const UNLOCKED_SYNC_KEYS = new Set<string>([
  "description", "productType", "vendor", "seoTitle", "seoDescription",
  "hsCode", "countryOfOrigin", "tags", K_STATUS, K_CATEGORY,
]);

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

// Fetch a product's current syncable fields from Shopify as a row-patch map:
// description, productType, vendor, seoTitle/Description, compareAtPrice, hsCode,
// countryOfOrigin, tags, modelPicture (image URLs) and __categoryMetafields.
// Returns null if the product can't be read / doesn't exist. Shared by the
// products/update webhook sync AND the "Match Shopify & lock" reconciliation.
export async function pullShopifyProductFields(shop: string, accessToken: string, productIdNumeric: string): Promise<Record<string, string> | null> {
  const num = numeric(productIdNumeric);
  if (!num) return null;
  const json = await gql(shop, accessToken, `
    query CollectionsSyncPull($id: ID!) {
      product(id: $id) {
        id status descriptionHtml productType vendor tags
        seo { title description }
        category { id fullName }
        media(first: 25) { nodes { ... on MediaImage { image { url } } } }
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
  if (!product?.id) return null;

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
  // Keep the row's status chip in step with Shopify (ACTIVE/DRAFT/ARCHIVED), so a
  // product activated in Shopify shows as linked+active here and a later push
  // never reverts it. Reserved row key, not a user field.
  if (product.status) pulled[K_STATUS] = String(product.status);
  // Tags: mirror Shopify's, but never remove the pre-order system tags.
  {
    const kept = shopTags.filter((t: string) => { const s = t.toLowerCase(); return s !== "pre-order" && !s.startsWith("pre-order: ") && !s.startsWith("pre-order ships "); });
    if (product.tags !== undefined) pulled.tags = kept.join(", ");
  }
  // Images → Model Picture (JSON array of CDN URLs). Only when Shopify has
  // images, so an empty/partial fetch never wipes the row.
  const imageUrls: string[] = (product.media?.nodes ?? [])
    .map((n: { image?: { url?: string } }) => String(n?.image?.url ?? ""))
    .filter((u: string) => !!u);
  if (imageUrls.length) pulled.modelPicture = JSON.stringify(imageUrls);
  const catBlob = buildCategoryBlob(product);
  if (catBlob) pulled[K_CATEGORY] = JSON.stringify(catBlob);
  return pulled;
}

export async function syncCollectionRowsForProduct(shop: string, productId: unknown): Promise<{ updated: number }> {
  const num = numeric(productId);
  if (!num || !shop) return { updated: 0 };

  // 1) Cheap first: is this product linked to ANY collection row (locked or
  //    unlocked)? If not, return without touching Shopify (products/update fires
  //    often).
  const collections = await prisma.collection.findMany({ select: { id: true, rows: true } });
  const matched = collections.filter((c) =>
    (Array.isArray(c.rows) ? (c.rows as Array<Record<string, unknown>>) : []).some(
      (r) => numeric(r?.[K_PRODUCT_ID]) === num,
    ),
  );
  if (!matched.length) return { updated: 0 };

  const session = await prisma.session.findFirst({ where: { shop, accessToken: { not: "" } }, orderBy: { isOnline: "asc" }, select: { accessToken: true } }).catch(() => null);
  if (!session?.accessToken) return { updated: 0 };

  // 2) Pull the current fields for this product.
  const pulled = await pullShopifyProductFields(shop, session.accessToken, num);
  if (!pulled) return { updated: 0 };

  // 3) Update the linked rows in each matched collection.
  let updated = 0;
  for (const c of matched) {
    const rows = Array.isArray(c.rows) ? (c.rows as Array<Record<string, unknown>>) : [];
    let changed = false;
    const next = rows.map((row) => {
      if (numeric(row?.[K_PRODUCT_ID]) !== num) return row;
      const locked = String(row?.[K_LOCKED] ?? "") === "1";
      // Fields the user has edited in the portal but not yet pushed — never
      // clobber these from Shopify.
      const edited = new Set(String(row?.[K_EDITED] ?? "").split(",").map((s) => s.trim()).filter(Boolean));

      if (locked) {
        changed = true; updated += 1;
        // Don't overwrite the row's pictures if the user has a pending picture
        // edit (images are portal-curated and editable even when locked).
        if (edited.has("modelPicture")) {
          const { modelPicture: _skip, ...rest } = pulled;
          void _skip;
          return { ...row, ...rest };
        }
        return { ...row, ...pulled };
      }

      // Unlocked: mirror only descriptive fields, skip anything the user is
      // mid-editing, and only when the value actually differs (so a Shopify save
      // doesn't needlessly bump every linked collection's updatedAt).
      const apply: Record<string, string> = {};
      for (const [k, v] of Object.entries(pulled)) {
        if (!UNLOCKED_SYNC_KEYS.has(k) || edited.has(k)) continue;
        if (String(row?.[k] ?? "") !== v) apply[k] = v;
      }
      if (!Object.keys(apply).length) return row;
      changed = true; updated += 1;
      return { ...row, ...apply };
    });
    if (changed) await prisma.collection.update({ where: { id: c.id }, data: { rows: next as unknown as object, updatedAt: new Date() } }).catch(() => {});
  }
  return { updated };
}
