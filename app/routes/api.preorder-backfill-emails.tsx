import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

// Admin: backfill customerEmail on pre-order reservations that were captured
// without it (no-plan captures used to store customerEmail=null → the Customer
// Orders tab showed "—"). Looks up each order's email from Shopify and fills the
// blank rows. Dry run by default; add &apply=1 to write.
//   GET /api/preorder-backfill-emails            (dry run)
//   GET /api/preorder-backfill-emails?apply=1    (write)
const API_VERSION = "2025-10";
const toOrderGid = (id: string) => { const n = String(id ?? "").replace(/\D/g, ""); return n ? `gid://shopify/Order/${n}` : null; };

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const apply = new URL(request.url).searchParams.get("apply") === "1";
  const session = await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true, accessToken: true } });
  if (!session?.shop || !session.accessToken) return Response.json({ ok: false, error: "No offline Shopify session." }, { status: 500 });
  const { shop, accessToken } = session;

  // Reservations still missing an email.
  const missing = await prisma.preorderReservation.findMany({
    where: { OR: [{ customerEmail: null }, { customerEmail: "" }] },
    select: { shopifyOrderId: true },
  });
  const orderIds = Array.from(new Set(missing.map((r) => r.shopifyOrderId).filter(Boolean)));
  if (!orderIds.length) return Response.json({ ok: true, rowsMissing: 0, ordersToLookUp: 0, emailsFound: 0, rowsUpdated: 0, note: "Nothing to backfill — every reservation already has an email." }, { headers: { "Cache-Control": "no-store" } });

  // Look up each order's email from Shopify in chunks.
  const emailByOrderId = new Map<string, string>();
  const queryErrors: string[] = [];
  for (let i = 0; i < orderIds.length; i += 50) {
    const chunk = orderIds.slice(i, i + 50);
    const gids = chunk.map(toOrderGid).filter(Boolean) as string[];
    const res = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": accessToken },
      body: JSON.stringify({ query: `query BackfillEmails($ids: [ID!]!) { nodes(ids: $ids) { ... on Order { id email } } }`, variables: { ids: gids } }),
    });
    const json = await res.json() as { data?: { nodes?: Array<{ id?: string; email?: string | null } | null> }; errors?: Array<{ message?: string }> };
    if (json.errors?.length) { queryErrors.push(...json.errors.map((e) => e.message || "Shopify GraphQL error")); continue; }
    for (const n of json.data?.nodes ?? []) {
      const num = String(n?.id ?? "").replace(/\D/g, "");
      const email = (n?.email ?? "").trim();
      if (num && email) emailByOrderId.set(num, email);
    }
  }

  const emailsFound = emailByOrderId.size;
  let rowsUpdated = 0;
  if (apply) {
    for (const [orderId, email] of emailByOrderId) {
      const r = await prisma.preorderReservation.updateMany({
        where: { shopifyOrderId: orderId, OR: [{ customerEmail: null }, { customerEmail: "" }] },
        data: { customerEmail: email },
      });
      rowsUpdated += r.count;
    }
  }

  return Response.json(
    {
      ok: true,
      applied: apply,
      rowsMissing: missing.length,
      ordersToLookUp: orderIds.length,
      emailsFound,
      rowsUpdated,
      ordersWithNoEmailInShopify: orderIds.length - emailsFound,
      queryErrors: queryErrors.length ? queryErrors : undefined,
      note: apply ? `Filled ${rowsUpdated} reservation row(s).` : `Dry run — ${emailsFound} order email(s) found for ${missing.length} blank row(s). Add &apply=1 to write.`,
    },
    { headers: { "Cache-Control": "no-store" } },
  );
};
