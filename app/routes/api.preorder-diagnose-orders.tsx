import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { KARMA_EAST_PREORDER_PLAN_PREFIX, preorderBatchIdFromPlanName } from "../preorder/preorder-shopify-order-normalize";

const API_VERSION = "2025-10";
const numericId = (gid: string) => String(gid ?? "").split("/").pop() ?? "";

// Admin read-only diagnostic for "orders not being flagged as pre-order".
// GET /api/preorder-diagnose-orders?orders=385740,385388,385349
// For each order it reports: does any line carry the Karma East pre-order
// selling plan, the plan/batch, whether it's reserved in our DB, and the order's
// tags — so we can tell a reservation failure (has plan, not reserved) from a
// buy-outside-the-preorder-button case (no plan at all).
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const url = new URL(request.url);
  const names = String(url.searchParams.get("orders") ?? "")
    .split(",").map((s) => s.trim().replace(/^#/, "")).filter(Boolean).slice(0, 30);
  if (!names.length) return Response.json({ ok: false, error: "Pass ?orders=385740,385388,…" }, { status: 400 });

  const session = await prisma.session.findFirst({
    where: { isOnline: false, accessToken: { not: "" } },
    orderBy: { expires: "desc" },
    select: { shop: true, accessToken: true },
  });
  if (!session?.accessToken) return Response.json({ ok: false, error: "No offline Shopify session." }, { status: 500 });
  const { shop, accessToken } = session;

  const query = async (name: string) => {
    const res = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": accessToken },
      body: JSON.stringify({
        query: `#graphql
          query DiagOrder($q: String!) {
            orders(first: 1, query: $q) {
              nodes {
                id name createdAt displayFulfillmentStatus tags
                lineItems(first: 50) { nodes { title quantity sku
                  variant { id title
                    preorderMeta: metafield(namespace: "karmaeast", key: "preorder") { value }
                    dispatchMeta: metafield(namespace: "karmaeast", key: "dispatch") { value }
                  }
                  sellingPlan { name } } }
              }
            }
          }`,
        variables: { q: `name:${name}` },
      }),
    });
    const json = await res.json() as { data?: { orders?: { nodes?: Array<{
      id: string; name: string; createdAt: string; displayFulfillmentStatus: string; tags: string[];
      lineItems?: { nodes?: Array<{ title: string | null; quantity: number; sku: string | null; variant?: { id?: string | null; title?: string | null; preorderMeta?: { value?: string | null } | null; dispatchMeta?: { value?: string | null } | null } | null; sellingPlan?: { name?: string | null } | null }> };
    }> } }; errors?: Array<{ message?: string }> };
    if (json.errors?.length) return { name, error: json.errors.map((e) => e.message).join("; ") };
    const order = json.data?.orders?.nodes?.[0];
    if (!order) return { name, error: "Order not found." };

    const lines = order.lineItems?.nodes ?? [];
    const preorderLines = lines.filter((l) => (l.sellingPlan?.name ?? "").startsWith(KARMA_EAST_PREORDER_PLAN_PREFIX));
    const metaFlag = (l: typeof lines[number]) => String(l.variant?.preorderMeta?.value ?? "").toLowerCase() === "true";
    // The email banner shows a pre-order if ANY line has the selling plan OR the
    // variant's karmaeast.preorder metafield is true. This mirrors the template's
    // condition, so we can tell "email didn't flag it" from a store-config issue.
    const emailBannerWouldShow = preorderLines.length > 0 || lines.some((l) => metaFlag(l));
    const orderIdNumeric = numericId(order.id);
    const reservations = await prisma.preorderReservation.findMany({
      where: { shopifyOrderId: orderIdNumeric },
      select: { status: true, quantity: true, supplierOrderId: true, variantTitle: true },
    });

    return {
      name: order.name,
      createdAt: order.createdAt,
      fulfillment: order.displayFulfillmentStatus,
      hasPreorderPlan: preorderLines.length > 0,
      // Would the order-confirmation email show the pre-order banner? If false but
      // the order is on hold, the variant metafield isn't set (or unreadable) →
      // customer got a plain email. This is the email-vs-hold mismatch check.
      emailBannerWouldShow,
      preorderLines: preorderLines.map((l) => ({
        title: l.title, size: l.variant?.title ?? null, qty: l.quantity,
        plan: l.sellingPlan?.name ?? null, batchId: preorderBatchIdFromPlanName(l.sellingPlan?.name),
        preorderMetafield: l.variant?.preorderMeta?.value ?? null, dispatchMetafield: l.variant?.dispatchMeta?.value ?? null,
      })),
      // Lines with NO plan (e.g. bought via Shop Pay / PayPal express / quick-add).
      // preorderMetafield shows whether the email would still flag it as a pre-order.
      nonPlanLines: lines.filter((l) => !(l.sellingPlan?.name ?? "").startsWith(KARMA_EAST_PREORDER_PLAN_PREFIX)).map((l) => ({ title: l.title, size: l.variant?.title ?? null, qty: l.quantity, preorderMetafield: l.variant?.preorderMeta?.value ?? null, dispatchMetafield: l.variant?.dispatchMeta?.value ?? null })),
      reservedInPortal: reservations.length > 0,
      reservations,
      preorderTags: (order.tags ?? []).filter((t) => t.toLowerCase().startsWith("pre-order")),
    };
  };

  const results = [];
  for (const name of names) results.push(await query(name).catch((e) => ({ name, error: e instanceof Error ? e.message : String(e) })));

  const summary = {
    checked: results.length,
    withPlanNotReserved: results.filter((r) => "hasPreorderPlan" in r && r.hasPreorderPlan && !r.reservedInPortal).map((r) => r.name),
    noPlanAtAll: results.filter((r) => "hasPreorderPlan" in r && !r.hasPreorderPlan).map((r) => r.name),
    reserved: results.filter((r) => "reservedInPortal" in r && r.reservedInPortal).map((r) => r.name),
    // Reserved/held as a pre-order but the email banner would NOT show → customer
    // got a plain confirmation. These need the metafield set (or the template updated).
    heldButEmailWouldNotFlag: results.filter((r) => "reservedInPortal" in r && r.reservedInPortal && "emailBannerWouldShow" in r && !r.emailBannerWouldShow).map((r) => r.name),
  };
  return Response.json({ ok: true, summary, results }, { headers: { "Cache-Control": "no-store" } });
};
