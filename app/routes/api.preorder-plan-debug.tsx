import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { getOfflineToken } from "../preorder/preorder-fulfillment.server";
import { getPreorderSellingPlanRegistryEntry } from "../preorder/preorder-selling-plan-registry.server";
import { refreshPreorderSellingPlanDate } from "../preorder/preorder-selling-plan.service.server";

const API_VERSION = "2025-10";

// Admin diagnostic for the "stuck dispatch date on the cart/order" issue.
// GET /api/preorder-plan-debug?batch=<supplierOrderId>[&fix=1]
//   - shows the live Shopify selling-plan name/options for the batch
//   - with &fix=1, re-renames it to the current dispatch date and reports errors
async function planName(shop: string, token: string, groupId: string) {
  const res = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
    method: "POST",
    headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": token },
    body: JSON.stringify({
      query: `#graphql
        query PlanDebug($id: ID!) {
          sellingPlanGroup(id: $id) {
            id name merchantCode
            sellingPlans(first: 5) { nodes { id name options } }
          }
        }`,
      variables: { id: groupId },
    }),
  });
  const json = await res.json() as { data?: { sellingPlanGroup?: unknown }; errors?: unknown };
  return json.data?.sellingPlanGroup ?? { errors: json.errors };
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const url = new URL(request.url);
  const batch = Number(url.searchParams.get("batch"));
  const doFix = url.searchParams.get("fix") === "1";
  if (!Number.isInteger(batch) || batch <= 0) return Response.json({ ok: false, error: "Pass ?batch=<supplierOrderId>" }, { status: 400 });

  const session = await prisma.session.findFirst({ where: { accessToken: { not: "" } }, orderBy: { isOnline: "asc" }, select: { shop: true } });
  if (!session?.shop) return Response.json({ ok: false, error: "No Shopify session." }, { status: 500 });
  const shop = session.shop;
  const token = await getOfflineToken(shop);
  if (!token) return Response.json({ ok: false, error: "No offline token (re-auth the app)." }, { status: 500 });

  const registry = await getPreorderSellingPlanRegistryEntry(shop, batch);
  const setting = await prisma.preorderBatchSetting.findUnique({ where: { supplierOrderId: batch }, select: { shipDate: true, enabled: true } });
  const order = await prisma.supplierOrder.findUnique({ where: { id: batch }, select: { eta: true, productTitle: true } });

  if (!registry) {
    return Response.json({ ok: true, batch, live: false, note: "No live Shopify selling plan for this batch (never activated, or turned off).", setting, order }, { headers: { "Cache-Control": "no-store" } });
  }

  const before = await planName(shop, token, registry.sellingPlanGroupId);
  let fixError: string | null = null;
  let after: unknown = null;
  if (doFix) {
    try { await refreshPreorderSellingPlanDate(batch); }
    catch (e) { fixError = e instanceof Error ? e.message : String(e); }
    after = await planName(shop, token, registry.sellingPlanGroupId);
  }

  return Response.json({
    ok: true,
    batch,
    productTitle: order?.productTitle ?? null,
    portalShipDate: setting?.shipDate ?? order?.eta ?? null,
    registry: { sellingPlanGroupId: registry.sellingPlanGroupId, sellingPlanId: registry.sellingPlanId },
    planBefore: before,
    ...(doFix ? { fixError, planAfter: after } : {}),
  }, { headers: { "Cache-Control": "no-store" } });
};
