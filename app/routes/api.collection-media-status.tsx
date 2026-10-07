import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

// Admin diagnostic: read a product's media and show Shopify's REAL per-image
// processing status + error CODE. Shopify's "Media upload failed / Media
// processing failed" banner is an ASYNC failure that productCreateMedia never
// reports back, so this pulls the truth directly off the product.
//   GET /api/collection-media-status?q=Frankie dress nila
//   GET /api/collection-media-status?productId=1234567890
//   GET /api/collection-media-status?handle=frankie-dress-nila
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
  const q = (url.searchParams.get("q") || "").trim();
  if (!rawProductId && !handle && !q) return Response.json({ ok: false, error: "Pass ?q=<product title>, ?productId=<numeric id> or ?handle=<handle>." }, { status: 400 });

  // Resolve a product GID.
  let productGid = rawProductId ? `gid://shopify/Product/${rawProductId}` : "";
  if (!productGid) {
    const query = handle ? `handle:${handle}` : `title:*${q}*`;
    const r = await gql(shop, accessToken, `query($q:String!){ products(first:5, query:$q){ nodes { id title handle } } }`, { q: query });
    const nodes = r.json?.data?.products?.nodes ?? [];
    productGid = nodes[0]?.id ?? "";
    if (!productGid) return Response.json({ ok: false, error: `No product found for ${handle ? `handle "${handle}"` : `"${q}"`}.`, raw: r.json }, { status: 404 });
  }

  // Pull every media item with its real status + error code. MediaImage carries
  // BOTH mediaErrors (processing) and fileErrors (file-level); grab both plus the
  // source so we can see exactly what failed and why.
  const r = await gql(shop, accessToken, `
    query MediaStatus($id: ID!) {
      product(id: $id) {
        id title handle
        media(first: 50) {
          nodes {
            __typename
            ... on MediaImage {
              id
              status
              mimeType
              mediaErrors { code details message }
              mediaWarnings { code message }
              fileStatus
              fileErrors { code details message }
              image { url width height }
              originalSource { fileSize }
              preview { status }
            }
            ... on Video { id status mediaErrors { code details message } }
            ... on ExternalVideo { id status }
            ... on Model3d { id status mediaErrors { code details message } }
          }
        }
      }
    }
  `, { id: productGid });

  const product = r.json?.data?.product;
  const nodes: any[] = product?.media?.nodes ?? [];
  const summary = nodes.map((n, i) => ({
    i,
    type: n.__typename,
    status: n.status,
    fileStatus: n.fileStatus,
    mediaErrors: n.mediaErrors ?? [],
    fileErrors: n.fileErrors ?? [],
    mediaWarnings: n.mediaWarnings ?? [],
    mime: n.mimeType,
    dims: n.image ? `${n.image.width}x${n.image.height}` : null,
    bytes: n.originalSource?.fileSize ?? null,
    url: n.image?.url ?? null,
  }));
  const failed = summary.filter((s) => String(s.status) === "FAILED" || (s.mediaErrors?.length ?? 0) > 0 || (s.fileErrors?.length ?? 0) > 0);

  return Response.json(
    {
      ok: true,
      shop,
      productGid,
      title: product?.title ?? null,
      handle: product?.handle ?? null,
      mediaCount: nodes.length,
      failedCount: failed.length,
      failed,          // <- the real reason(s) live here: code + details
      all: summary,
      errors: r.json?.errors ?? null,
      hint: "Paste the whole JSON back — I need failed[].status + mediaErrors[].code/details (and fileErrors) to pinpoint the cause.",
    },
    { headers: { "Cache-Control": "no-store" } },
  );
};
