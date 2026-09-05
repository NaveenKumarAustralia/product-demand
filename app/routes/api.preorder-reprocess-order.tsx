import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { KARMA_EAST_PREORDER_PLAN_PREFIX, preorderBatchIdFromPlanName } from "../preorder/preorder-shopify-order-normalize";
import { reservePreorderLine, PreorderCapacityError } from "../preorder/preorder-allocation.server";

const API_VERSION = "2025-10";

// Admin diagnostic + recovery tool. Fetches a real order from Shopify via
// GraphQL (which always includes the selling-plan allocation) and re-runs the
// preorder reservation for it. Lets us (a) reserve an order that was paid before
// the reservation path worked, and (b) SEE the exact line/selling-plan data when
// the webhook path produced nothing. GET so it can be triggered from a browser
// while logged into the portal. Reserving is idempotent (keyed by line item).
function numericId(gid: string) {
  return String(gid ?? "").split("/").pop() ?? "";
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const url = new URL(request.url);
  const orderParam = String(url.searchParams.get("order") ?? "").trim();
  const dryRun = url.searchParams.get("reserve") !== "1";
  if (!orderParam) return Response.json({ ok: false, error: "Pass ?order=<order number or id>, and add &reserve=1 to actually reserve." }, { status: 400 });

  const session = await prisma.session.findFirst({
    where: { isOnline: false, accessToken: { not: "" } },
    orderBy: { expires: "desc" },
    select: { shop: true, accessToken: true },
  });
  if (!session?.accessToken) return Response.json({ ok: false, error: "No offline Shopify session." }, { status: 500 });
  const shop = session.shop;

  const nameQuery = orderParam.replace(/^#/, "");
  const response = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
    method: "POST",
    headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": session.accessToken },
    body: JSON.stringify({
      query: `#graphql
        query ReprocessPreorder($q: String!) {
          orders(first: 1, query: $q) {
            nodes {
              id name email
              shippingAddress { countryCodeV2 }
              billingAddress { countryCodeV2 }
              lineItems(first: 50) {
                nodes {
                  id quantity sku title
                  variant { id title }
                  sellingPlan { name }
                }
              }
            }
          }
        }
      `,
      variables: { q: `name:${nameQuery}` },
    }),
  });
  const json = await response.json() as {
    data?: { orders?: { nodes?: Array<{
      id: string; name: string; email: string | null;
      shippingAddress?: { countryCodeV2?: string | null } | null;
      billingAddress?: { countryCodeV2?: string | null } | null;
      lineItems?: { nodes?: Array<{
        id: string; quantity: number; sku: string | null; title: string | null;
        variant?: { id?: string | null; title?: string | null } | null;
        sellingPlan?: { name?: string | null } | null;
      }> };
    }> } };
    errors?: Array<{ message?: string }>;
  };
  if (json.errors?.length) return Response.json({ ok: false, error: json.errors.map((e) => e.message).join("; ") }, { status: 502 });

  const order = json.data?.orders?.nodes?.[0];
  if (!order) return Response.json({ ok: false, error: `No order found matching "${orderParam}".` }, { status: 404 });

  const country = String(order.shippingAddress?.countryCodeV2 || order.billingAddress?.countryCodeV2 || "").toUpperCase();
  const market = country === "US" ? "USA" as const : "AU" as const;
  const allLines = order.lineItems?.nodes ?? [];

  // What we see, before any filtering — so we can diagnose the exact data.
  const diagnosis = allLines.map((line) => {
    const planName = line.sellingPlan?.name ?? null;
    return {
      title: line.title,
      variantId: line.variant?.id ?? null,
      quantity: line.quantity,
      sellingPlanName: planName,
      isPreorder: Boolean(planName && planName.startsWith(KARMA_EAST_PREORDER_PLAN_PREFIX)),
      batchId: preorderBatchIdFromPlanName(planName),
    };
  });

  const preorderLines = allLines.filter((line) => {
    const planName = line.sellingPlan?.name ?? "";
    return planName.startsWith(KARMA_EAST_PREORDER_PLAN_PREFIX);
  });

  const results: Array<Record<string, unknown>> = [];
  if (!dryRun) {
    for (const line of preorderLines) {
      const planName = line.sellingPlan?.name ?? "";
      const batchId = preorderBatchIdFromPlanName(planName);
      try {
        const rows = await reservePreorderLine({
          shop,
          shopifyOrderId: numericId(order.id),
          shopifyOrderName: order.name,
          shopifyLineItemId: numericId(line.id),
          productId: null,
          variantId: String(line.variant?.id ?? ""),
          variantTitle: line.variant?.title ?? line.title ?? null,
          sku: line.sku ?? null,
          market,
          quantity: Number(line.quantity),
          customerEmail: order.email ?? null,
          preferredSupplierOrderId: batchId ?? undefined,
        });
        results.push({ line: line.title, reserved: rows.reduce((s, r) => s + r.quantity, 0), batchId });
      } catch (error) {
        results.push({ line: line.title, error: error instanceof PreorderCapacityError ? error.message : String(error), batchId });
      }
    }
  }

  return Response.json({
    ok: true,
    order: { id: order.id, name: order.name, market, email: order.email },
    lineCount: allLines.length,
    diagnosis,
    reserved: dryRun ? "(dry run — add &reserve=1 to reserve)" : results,
  }, { headers: { "Cache-Control": "no-store" } });
};
