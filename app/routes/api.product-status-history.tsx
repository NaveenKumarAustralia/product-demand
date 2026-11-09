import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { getOfflineToken } from "../preorder/preorder-fulfillment.server";

// Admin, READ-ONLY diagnostic: "When & why did a product get set to DRAFT?"
// Shopify keeps a per-product event log (status changes, publish/unpublish,
// edits) that names WHO or WHAT did it (a staff member, or an app). This reads
// that log so we can see exactly when a product flipped to draft and whether a
// person, an app, or the portal caused it. Makes NO changes.
//
//   ?q=Celeste Dress Anjali   (default) — search by title (partial, any status)
//   ?id=123456789             — a specific product id (numeric or gid)
//   ?events=60                — how many recent events to pull (default 40)
//
// No token is printed; the offline session is read server-side.

const API_VERSION = "2025-10";

const norm = (s: unknown) => String(s ?? "").trim().toLowerCase();

// Words in an event message/action that point at a status / visibility change,
// so we can surface the relevant entries first.
const STATUS_HINTS = ["draft", "active", "archiv", "unpublish", "publish", "status", "hidden"];

const QUERY = `
  query ProductStatusHistory($q: String!, $events: Int!) {
    products(first: 10, query: $q) {
      nodes {
        id
        title
        handle
        status
        productType
        tags
        createdAt
        updatedAt
        publishedAt
        collections(first: 25) { nodes { id title handle } }
        variants(first: 5) { nodes { price compareAtPrice } }
        events(first: $events, sortKey: CREATED_AT, reverse: true) {
          nodes {
            id
            createdAt
            message
            ... on BasicEvent { action subjectType }
          }
        }
      }
    }
  }`;

type EventNode = { id?: string; createdAt?: string; message?: string; action?: string; subjectType?: string };

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return json({ error: "Admin only." }, 403);

  const url = new URL(request.url);
  const idParam = (url.searchParams.get("id") || "").trim();
  const qParam = (url.searchParams.get("q") || "").trim() || "Celeste Dress Anjali";
  const eventsN = Math.min(Math.max(Number(url.searchParams.get("events") || "40") || 40, 5), 100);

  // Build the Shopify product search. An id wins; otherwise a title search.
  const searchQuery = idParam
    ? `id:${idParam.replace(/\D/g, "")}`
    : `title:${JSON.stringify(qParam)}`; // quote so multi-word titles match as a phrase

  const shop = (await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true } }))?.shop ?? "";
  const token = shop ? await getOfflineToken(shop) : null;
  if (!shop || !token) return json({ error: "No Shopify session." }, 500);

  const res = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
    method: "POST",
    headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": token },
    body: JSON.stringify({ query: QUERY, variables: { q: searchQuery, events: eventsN } }),
  });
  const body = (await res.json()) as { data?: any; errors?: Array<{ message?: string }> };
  if (body.errors?.length) return json({ error: "Shopify error", detail: body.errors.map((e) => e.message).join("; ") }, 502);

  const products: Array<any> = body.data?.products?.nodes ?? [];
  if (!products.length) return json({ mode: "diagnostic", searchedFor: idParam || qParam, found: 0, note: "No product matched. Try ?q=<partial title> or ?id=<productId>." });

  const report = products.map((p) => {
    const events: EventNode[] = p.events?.nodes ?? [];
    // Flag the events that look like a status / visibility change.
    const statusEvents = events.filter((e) => STATUS_HINTS.some((h) => norm(e.message).includes(h) || norm(e.action).includes(h)));
    // "Why would a Sale bulk-draft catch this?" signals:
    const collections: Array<{ id?: string; title?: string; handle?: string }> = p.collections?.nodes ?? [];
    const variants: Array<{ price?: string; compareAtPrice?: string }> = p.variants?.nodes ?? [];
    const onSale = variants.some((v) => v.compareAtPrice && Number(v.compareAtPrice) > Number(v.price ?? 0)); // has a compare-at price > price
    const saleCollections = collections.filter((c) => norm(c.title).includes("sale") || norm(c.handle).includes("sale")).map((c) => c.title);
    const saleTags = (Array.isArray(p.tags) ? p.tags : []).filter((t: string) => norm(t).includes("sale"));
    return {
      id: p.id,
      title: p.title,
      handle: p.handle,
      currentStatus: p.status,
      productType: p.productType || "(none)",
      // Why a "Sale" bulk action could include this product even if productType isn't "Sale":
      whyCaughtInSale: {
        onSale, // a variant has a compare-at (discount) price — automated "Sale" collections include these
        inSaleCollection: saleCollections.length > 0,
        saleCollections,
        saleTags,
      },
      tags: p.tags,
      collections: collections.map((c) => c.title),
      createdAt: p.createdAt,
      updatedAt: p.updatedAt,
      publishedAt: p.publishedAt, // null once unpublished / draft
      // Most-relevant first: the entries that mention status/publish/draft.
      statusChangeEvents: statusEvents.map((e) => ({ at: e.createdAt, action: e.action, who: e.message })),
      // Full recent timeline for context (who did what, when).
      recentEvents: events.map((e) => ({ at: e.createdAt, action: e.action, who: e.message })),
    };
  });

  return json({
    mode: "diagnostic",
    searchedFor: idParam || qParam,
    found: report.length,
    howToRead:
      "statusChangeEvents = entries mentioning status/draft/publish; 'who' names the staff member (or app) that did it. whyCaughtInSale explains how a 'Sale' bulk-draft could include a product whose productType isn't 'Sale': onSale=true means a variant has a compare-at (discount) price, so an AUTOMATED 'Sale' collection would include it; inSaleCollection/saleCollections = it's in a collection named 'Sale'; saleTags = it carries a 'Sale' tag. If all of those are empty and productType isn't Sale, the product was almost certainly swept up by a selection error (e.g. 'select all' without the filter applied).",
    products: report,
  });
};

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body, null, 2), { status, headers: { "Content-Type": "application/json", "Cache-Control": "no-store" } });
}
