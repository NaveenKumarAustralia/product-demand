import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { getOfflineToken } from "../preorder/preorder-fulfillment.server";

// Admin, READ-ONLY audit: "Which products were set to DRAFT in a given time
// window, and which of them are NON-PROMO (shouldn't have been touched)?"
//
// Built after a staffer bulk-set the sale products to draft on 2026-11-09 and
// swept up several non-sale products by mistake. This scans the Shopify event
// log for every "changed product status ... to draft" in the window, then
// classifies each product as promo (on a discount / BFCM / % off / sale
// collection or tag, or carrying a compare-at price) vs non-promo.
//
//   ?from=2026-11-09T00:00:00Z   window start (UTC, default = the Nov 9 sweep)
//   ?to=2026-11-09T01:00:00Z     window end   (UTC)
//
// Makes NO changes. Reports currentStatus so you can see which are STILL draft.

const API_VERSION = "2025-10";
const norm = (s: unknown) => String(s ?? "").trim().toLowerCase();
// Promo / discount markers (what "put the sale into draft" really targets).
// Matches Sale, Clearance, BFCM, Cyber, Black Friday, "NN%", "% off" — but not
// unrelated words like "Off-Shoulder".
const PROMO_RE = /(sale|clearance|discount|promo|bfcm|cyber|black\s*friday|\d+\s*%|%\s*off)/i;

// Default window = the last 5 days (newest-first scan below), so the bare URL
// covers recent drafting without needing explicit ?from/?to.
const DEFAULT_WINDOW_DAYS = 5;

const EVENTS_QUERY = `
  query DraftSweepEvents($q: String!, $after: String) {
    events(first: 250, after: $after, query: $q, sortKey: CREATED_AT, reverse: true) {
      pageInfo { hasNextPage endCursor }
      nodes { id createdAt message ... on BasicEvent { action subjectType subjectId } }
    }
  }`;

const PRODUCTS_QUERY = `
  query DraftSweepProducts($ids: [ID!]!) {
    nodes(ids: $ids) {
      ... on Product {
        id title handle status productType tags
        collections(first: 25) { nodes { title handle } }
        variants(first: 5) { nodes { price compareAtPrice } }
      }
    }
  }`;

type Gql = (query: string, variables?: Record<string, unknown>) => Promise<any>;

function makeGql(shop: string, token: string): Gql {
  return async (query, variables) => {
    const res = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": token },
      body: JSON.stringify({ query, variables }),
    });
    return (await res.json()) as any;
  };
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return json({ error: "Admin only." }, 403);

  const url = new URL(request.url);
  const nowMs = Date.now();
  const defFrom = new Date(nowMs - DEFAULT_WINDOW_DAYS * 24 * 60 * 60 * 1000).toISOString();
  const defTo = new Date(nowMs).toISOString();
  const from = (url.searchParams.get("from") || defFrom).trim();
  const to = (url.searchParams.get("to") || defTo).trim();

  const shop = (await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true } }))?.shop ?? "";
  const token = shop ? await getOfflineToken(shop) : null;
  if (!shop || !token) return json({ error: "No Shopify session." }, 500);
  const gql = makeGql(shop, token);

  // 1) Walk the product event log in the window; keep "status → draft" events.
  const q = `subject_type:PRODUCT AND created_at:>=${from} AND created_at:<=${to}`;
  const drafted = new Map<string, { at: string; who: string }>(); // productId -> first draft event
  // Newest-first scan (reverse: true). A wide window can hold many events, so
  // the cap is generous; `truncated` tells us if we stopped before the window's
  // oldest edge (so a nil result there isn't mistaken for "nothing drafted").
  const MAX_PAGES = 160; // 160 * 250 = up to 40k events
  let after: string | null = null;
  let pages = 0;
  let scannedEvents = 0;
  let truncated = false;
  while (pages < MAX_PAGES) {
    const r: any = await gql(EVENTS_QUERY, { q, after });
    if (r?.errors?.length) return json({ error: "Shopify error (events)", detail: r.errors.map((e: any) => e.message).join("; ") }, 502);
    const conn = r?.data?.events;
    const nodes: Array<any> = conn?.nodes ?? [];
    scannedEvents += nodes.length;
    for (const n of nodes) {
      const msg = String(n?.message ?? "");
      // Only the status change to draft (one per product), not the per-channel
      // exclusion rows that accompany it.
      if (!/changed product status from .* to draft/i.test(msg)) continue;
      const pid = String(n?.subjectId ?? "");
      if (!pid) continue;
      // Newest-first, so the first time we see a product is its MOST RECENT
      // draft event — the one we care about for "when was it last drafted".
      if (!drafted.has(pid)) {
        const who = msg.split(/ changed product status/i)[0].trim() || "Unknown";
        drafted.set(pid, { at: String(n?.createdAt ?? ""), who });
      }
    }
    pages++;
    if (!conn?.pageInfo?.hasNextPage) break;
    if (pages >= MAX_PAGES) { truncated = true; break; }
    after = conn.pageInfo.endCursor;
  }

  const productIds = [...drafted.keys()];
  if (!productIds.length) {
    return json({ mode: "draft-sweep-audit", window: { from, to }, scannedEvents, truncated, totalDrafted: 0, note: truncated ? "Hit the event-scan cap before reaching the whole window — narrow ?from/?to (UTC) and re-run." : "No product was set to draft in this window. Adjust ?from / ?to (UTC)." });
  }

  // 2) Fetch each drafted product's type/tags/collections/compare-at, in chunks.
  type Row = { id: string; title: string; handle: string; currentStatus: string; productType: string; who: string; at: string; promo: boolean; promoReason: string[] };
  const rows: Row[] = [];
  for (let i = 0; i < productIds.length; i += 200) {
    const ids = productIds.slice(i, i + 200);
    const r: any = await gql(PRODUCTS_QUERY, { ids });
    if (r?.errors?.length) return json({ error: "Shopify error (products)", detail: r.errors.map((e: any) => e.message).join("; ") }, 502);
    for (const p of (r?.data?.nodes ?? [])) {
      if (!p?.id) continue; // a deleted product id resolves to null
      const meta = drafted.get(String(p.id));
      const collections: Array<any> = p.collections?.nodes ?? [];
      const variants: Array<any> = p.variants?.nodes ?? [];
      const onSale = variants.some((v) => v.compareAtPrice && Number(v.compareAtPrice) > Number(v.price ?? 0));
      const saleCollections = collections.filter((c) => PROMO_RE.test(String(c.title ?? "")) || PROMO_RE.test(String(c.handle ?? ""))).map((c) => c.title);
      const saleTags = (Array.isArray(p.tags) ? p.tags : []).filter((t: string) => PROMO_RE.test(String(t)));
      const promoType = PROMO_RE.test(String(p.productType ?? ""));
      const promo = onSale || promoType || saleCollections.length > 0 || saleTags.length > 0;
      const promoReason: string[] = [];
      if (onSale) promoReason.push("compare-at/discount price");
      if (promoType) promoReason.push(`productType "${p.productType}"`);
      if (saleCollections.length) promoReason.push(`in ${saleCollections.join(", ")}`);
      if (saleTags.length) promoReason.push(`tag ${saleTags.join(", ")}`);
      rows.push({
        id: String(p.id), title: String(p.title ?? ""), handle: String(p.handle ?? ""),
        currentStatus: String(p.status ?? ""), productType: String(p.productType || "(none)"),
        who: meta?.who ?? "Unknown", at: meta?.at ?? "",
        promo, promoReason,
      });
    }
  }

  const nonPromo = rows.filter((r) => !r.promo).sort((a, b) => a.at.localeCompare(b.at));
  const promo = rows.filter((r) => r.promo).sort((a, b) => a.at.localeCompare(b.at));
  const nonPromoStillDraft = nonPromo.filter((r) => r.currentStatus === "DRAFT");

  return json({
    mode: "draft-sweep-audit",
    window: { from, to },
    scannedEvents,
    truncated,
    totalDrafted: rows.length,
    counts: { promo: promo.length, nonPromo: nonPromo.length, nonPromoStillDraft: nonPromoStillDraft.length },
    howToRead:
      "nonPromo = products set to draft in this window that have NO promo/sale signal (not on a Sale/BFCM/% Off collection or tag, not discounted, productType isn't a sale type) — these are the ones that likely shouldn't have been drafted. 'currentStatus: DRAFT' means it's STILL hidden and needs switching back on. 'who' and 'at' (UTC) are from the event log. promoReason shows why a product counts as promo.",
    nonPromo_stillDraft_FIX_THESE: nonPromoStillDraft,
    nonPromo_all: nonPromo,
    promo_all: promo,
  });
};

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body, null, 2), { status, headers: { "Content-Type": "application/json", "Cache-Control": "no-store" } });
}
