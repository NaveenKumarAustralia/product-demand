import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { getPreorderSellingPlanRegistryEntries } from "../preorder/preorder-selling-plan-registry.server";
import { getOfflineToken } from "../preorder/preorder-fulfillment.server";

// Admin read-only health check for the pre-order PRODUCT TAGS. Lists every live
// pre-order batch and whether its Shopify product carries pre-order tags
// (`Pre-order` flag + at least one `Pre-order: <size>` + a `Pre-order ships`).
// Cheap query (tags only, chunked) so it can't blow Shopify's cost limit.
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
    ? await prisma.supplierOrder.findMany({ where: { id: { in: liveIds } }, select: { id: true, productTitle: true, productId: true } })
    : [];

  // Cheap: fetch ONLY tags, in small chunks, and surface any GraphQL error.
  const gids = Array.from(new Set(batches.map((b) => toProductGid(b.productId)).filter(Boolean))) as string[];
  const tagsByGid = new Map<string, string[]>();
  const queryErrors: string[] = [];
  for (let i = 0; i < gids.length; i += 50) {
    const chunk = gids.slice(i, i + 50);
    const res = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
      method: "POST", headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": token },
      body: JSON.stringify({ query: `query TagHealth($ids: [ID!]!) { nodes(ids: $ids) { ... on Product { id tags } } }`, variables: { ids: chunk } }),
    });
    const json = await res.json() as { data?: { nodes?: Array<{ id?: string; tags?: string[] } | null> }; errors?: Array<{ message?: string }> };
    if (json.errors?.length) { queryErrors.push(json.errors.map((e) => e.message).join("; ")); continue; }
    for (const n of json.data?.nodes ?? []) { if (n?.id) tagsByGid.set(n.id, (n.tags ?? []).map(String)); }
  }

  const isPreTag = (t: string) => { const s = t.trim().toLowerCase(); return s === "pre-order" || s.startsWith("pre-order:") || s.startsWith("pre-order ships"); };
  const results = batches.map((b) => {
    const gid = toProductGid(b.productId);
    const found = gid ? tagsByGid.has(gid) : false;
    const tags = (gid && tagsByGid.get(gid)) || [];
    const preorderTags = tags.filter(isPreTag);
    const hasShips = tags.some((t) => t.trim().toLowerCase().startsWith("pre-order ships"));
    // With the live-stock design the only tag we manage is the ship date; the
    // pre-order FLAG comes from the item being oversold, not a tag.
    const ok = found && hasShips;
    return { supplierOrderId: b.id, title: b.productTitle ?? null, productId: b.productId ?? null, found, preorderTags, hasShipsDateTag: hasShips, ok };
  });

  const summary = {
    liveBatches: results.length,
    haveShipDateTag: results.filter((r) => r.ok).length,
    queryErrors: queryErrors.length ? queryErrors : undefined,
    problems: results.filter((r) => !r.ok).map((r) => ({ supplierOrderId: r.supplierOrderId, title: r.title, reason: !r.found ? "product not returned by Shopify" : "no Pre-order ships <date> tag" })),
  };
  return Response.json({ ok: true, summary, results }, { headers: { "Cache-Control": "no-store" } });
};
