import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

// Admin inspector: dump exactly how a product stores its Shopify CATEGORY
// (taxonomy) node + category metafields, and the taxonomy attributes + allowed
// values for that category. Used to build the Collections "Category metafields"
// editor against the REAL shapes (they vary by store/API version) instead of
// guessing.
//   GET /api/collection-category-inspect?productId=1234567890
//   GET /api/collection-category-inspect?handle=vivien-dress-peacock
const API_VERSION = "2025-10";

async function gql(shop: string, token: string, query: string, variables: Record<string, unknown>) {
  try {
    const res = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": token },
      body: JSON.stringify({ query, variables }),
    });
    const json = await res.json();
    return { ok: res.ok, json };
  } catch (e) {
    return { ok: false, json: { error: e instanceof Error ? e.message : String(e) } };
  }
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const session = await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true, accessToken: true } });
  if (!session?.shop || !session.accessToken) return Response.json({ ok: false, error: "No offline Shopify session." }, { status: 500 });
  const { shop, accessToken } = session;

  const url = new URL(request.url);
  const rawProductId = (url.searchParams.get("productId") || "").replace(/\D/g, "");
  const handle = (url.searchParams.get("handle") || "").trim();
  if (!rawProductId && !handle) return Response.json({ ok: false, error: "Pass ?productId=<numeric id> or ?handle=<product-handle>." }, { status: 400 });

  // Resolve a product GID.
  let productGid = rawProductId ? `gid://shopify/Product/${rawProductId}` : "";
  if (!productGid && handle) {
    const r = await gql(shop, accessToken, `query($q:String!){ products(first:1, query:$q){ nodes { id title } } }`, { q: `handle:${handle}` });
    productGid = r.json?.data?.products?.nodes?.[0]?.id ?? "";
    if (!productGid) return Response.json({ ok: false, error: `No product found for handle "${handle}".`, raw: r.json }, { status: 404 });
  }

  // 1) The product's category node + ALL its metafields (so we can see which are
  //    the category/taxonomy ones, their type, and how values are stored).
  const prod = await gql(shop, accessToken, `
    query CatInspect($id: ID!) {
      product(id: $id) {
        id title handle
        category { id fullName }
        metafields(first: 200) {
          nodes {
            namespace key type value
            definition { name type { name } }
            reference { __typename ... on Metaobject { id handle type displayName } }
            references(first: 30) { nodes { __typename ... on Metaobject { id handle type displayName } } }
          }
        }
      }
    }
  `, { id: productGid });

  const categoryId: string = prod.json?.data?.product?.category?.id ?? "";

  // 2) The taxonomy category's attributes + allowed values (for the pickers). We
  //    try a couple of shapes and report whatever works, so we learn the real
  //    type names for this API version.
  let taxonomy: unknown = null;
  if (categoryId) {
    const t = await gql(shop, accessToken, `
      query CatAttrs($id: ID!) {
        node(id: $id) {
          ... on TaxonomyCategory {
            id fullName
            attributes(first: 100) { nodes { __typename } }
          }
        }
        tca: __type(name: "TaxonomyCategoryAttribute") { kind possibleTypes { name } fields { name } }
        ta: __type(name: "TaxonomyAttribute") { kind fields { name type { kind name ofType { name kind ofType { name } } } } possibleTypes { name } }
        tcla: __type(name: "TaxonomyChoiceListAttribute") { fields { name type { kind name ofType { name } } } }
        tv: __type(name: "TaxonomyValue") { fields { name } }
      }
    `, { id: categoryId });
    taxonomy = t.json;
  }

  return Response.json(
    {
      ok: true,
      shop,
      productGid,
      product: prod.json?.data?.product ?? null,
      productErrors: prod.json?.errors ?? null,
      categoryId,
      taxonomyAttributesRaw: taxonomy,
      hint: "Paste this whole JSON back. I need: product.category.id/fullName, and each category metafield's namespace/key/type/value + reference names, plus the taxonomy attributes block (or its error) so I can build the editor + write-back correctly.",
    },
    { headers: { "Cache-Control": "no-store" } },
  );
};
