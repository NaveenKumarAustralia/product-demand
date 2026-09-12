import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { getPreorderSellingPlanRegistryEntries } from "../preorder/preorder-selling-plan-registry.server";
import { getOfflineToken } from "../preorder/preorder-fulfillment.server";

// Admin read-only health check for the pre-order PRODUCT TAGS. Lists every live
// pre-order batch and whether its Shopify product carries the expected tags
// (`Pre-order: <size>` for each outstanding pre-order variant, plus the
// `Pre-order` flag and a `Pre-order ships <date>` tag). Flags any live product
// that is missing tags or isn't linked to a Shopify product — so you can SEE the
// tagging is correct rather than trust it. Changes nothing.
//   GET /api/preorder-tag-health
const API_VERSION = "2025-10";
const numeric = (s: string) => String(s ?? "").replace(/\D/g, "");
const toProductGid = (p: string | null) => { const n = numeric(String(p ?? "")); return n ? `gid://shopify/Product/${n}` : null; };

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const session = await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true } });
  if (!session?.shop) return Response.json({ ok: false, error: "No offline Shopify session." }, { status: 500 });
  const token = await getOfflineToken(session.shop);
  if (!token) return Response.json({ ok: false, error: "No offline token / missing scopes." }, { status: 500 });
  const shop = session.shop;

  const [enabled, registry] = await Promise.all([
    prisma.preorderBatchSetting.findMany({ where: { enabled: true }, select: { supplierOrderId: true } }),
    getPreorderSellingPlanRegistryEntries(shop),
  ]);
  const activated = new Set(registry.map((r) => r.supplierOrderId));
  const liveIds = enabled.map((s) => s.supplierOrderId).filter((id) => activated.has(id));
  const batches = liveIds.length
    ? await prisma.supplierOrder.findMany({ where: { id: { in: liveIds } }, select: { id: true, productTitle: true, productId: true, lines: { select: { variantId: true, qtyOrdered: true, qtyReceived: true } } } })
    : [];

  // Bulk-fetch each product's current tags + variant titles.
  const gids = Array.from(new Set(batches.map((b) => toProductGid(b.productId)).filter(Boolean))) as string[];
  const productMap = new Map<string, { tags: string[]; titleById: Map<string, string> }>();
  for (let i = 0; i < gids.length; i += 200) {
    const chunk = gids.slice(i, i + 200);
    const res = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
      method: "POST", headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": token },
      body: JSON.stringify({ query: `#graphql query TagHealth($ids: [ID!]!) { nodes(ids: $ids) { ... on Product { id tags variants(first: 100) { nodes { id title } } } } }`, variables: { ids: chunk } }),
    });
    const json = await res.json() as { data?: { nodes?: Array<{ id?: string; tags?: string[]; variants?: { nodes?: Array<{ id?: string; title?: string }> } } | null> } };
    for (const n of json.data?.nodes ?? []) {
      if (!n?.id) continue;
      productMap.set(n.id, { tags: (n.tags ?? []).map(String), titleById: new Map((n.variants?.nodes ?? []).map((v) => [numeric(String(v.id ?? "")), String(v.title ?? "")])) });
    }
  }

  const isPreTag = (t: string) => { const s = t.trim().toLowerCase(); return s === "pre-order" || s.startsWith("pre-order:") || s.startsWith("pre-order ships"); };
  const results = batches.map((b) => {
    const gid = toProductGid(b.productId);
    const prod = gid ? productMap.get(gid) : null;
    const tags = prod?.tags ?? [];
    const outstanding = (b.lines ?? []).filter((l) => l.qtyOrdered - l.qtyReceived > 0);
    const expectedSizeTags = outstanding.map((l) => prod?.titleById.get(numeric(String(l.variantId ?? "")))).filter(Boolean).map((t) => `Pre-order: ${t}`);
    const missingSizeTags = expectedSizeTags.filter((e) => !tags.some((t) => t.trim() === e));
    const hasFlag = tags.some((t) => t.trim() === "Pre-order");
    const hasShips = tags.some((t) => t.trim().toLowerCase().startsWith("pre-order ships"));
    const ok = Boolean(gid) && hasFlag && hasShips && expectedSizeTags.length > 0 && missingSizeTags.length === 0;
    return {
      supplierOrderId: b.id, title: b.productTitle ?? null, productId: b.productId ?? null,
      hasProductLink: Boolean(gid), foundInShopify: Boolean(prod),
      preorderTags: tags.filter(isPreTag), expectedSizeTags, missingSizeTags, hasFlag, hasShips, ok,
    };
  });

  const summary = {
    liveBatches: results.length,
    correctlyTagged: results.filter((r) => r.ok).length,
    problems: results.filter((r) => !r.ok).map((r) => ({ supplierOrderId: r.supplierOrderId, title: r.title, reason: !r.hasProductLink ? "no product link" : !r.foundInShopify ? "product not found in Shopify" : r.missingSizeTags.length ? `missing size tags: ${r.missingSizeTags.join(", ")}` : !r.hasFlag ? "missing Pre-order flag" : !r.hasShips ? "missing ships-date tag" : "no outstanding variants" })),
  };
  return Response.json({ ok: true, summary, results }, { headers: { "Cache-Control": "no-store" } });
};
