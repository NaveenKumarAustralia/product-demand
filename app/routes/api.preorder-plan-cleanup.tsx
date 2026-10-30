import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { getPreorderPermissionContext } from "../preorder/preorder-permissions.server";
import { getPreorderSellingPlanRegistryEntries, removePreorderSellingPlanRegistryEntry } from "../preorder/preorder-selling-plan-registry.server";
import { deactivatePreorderSellingPlan } from "../preorder/preorder-selling-plan.service.server";

const API_VERSION = "2025-10";
const numeric = (s: unknown) => String(s ?? "").replace(/\D/g, "");

// Delete a Shopify selling plan group by id WITHOUT touching variant inventory
// policy / metafields / tags. Used when the orphan's product is ALSO a live
// (keep) batch — we must remove the stale plan but leave the live pre-order's
// variant state intact — and for orphans whose batch no longer exists.
async function deleteSellingPlanGroupOnly(shop: string, token: string, sellingPlanGroupId: string) {
  const res = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
    method: "POST",
    headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": token },
    body: JSON.stringify({
      query: `#graphql
        mutation DeleteKarmaEastPreorderGroup($id: ID!) {
          sellingPlanGroupDelete(id: $id) { deletedSellingPlanGroupId userErrors { message } }
        }`,
      variables: { id: sellingPlanGroupId },
    }),
  });
  const json = await res.json() as { data?: { sellingPlanGroupDelete?: { deletedSellingPlanGroupId?: string; userErrors?: Array<{ message?: string }> } }; errors?: Array<{ message?: string }> };
  const errs = [...(json.errors ?? []), ...(json.data?.sellingPlanGroupDelete?.userErrors ?? [])].map((e) => e.message).filter(Boolean);
  // "does not exist" is fine — the group was already gone; we just need the
  // registry entry cleared. Any other error should surface.
  if (errs.length && !errs.some((m) => /not.*exist|not found|invalid id/i.test(String(m)))) {
    throw new Error(errs.join("; "));
  }
}

// Admin: remove every ORPHAN pre-order selling plan (closed/disabled/deleted
// batches). Dry-run by default; add ?apply=1 to execute. Time-bounded so it
// never times out — re-open the SAME ?apply=1 URL until remaining=0. Idempotent:
// a re-run only touches what's left. KEEP batches (open + enabled) are never
// touched.
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });
  const { permissions } = await getPreorderPermissionContext();

  const url = new URL(request.url);
  const apply = url.searchParams.get("apply") === "1" || url.searchParams.get("apply") === "true";

  const session = await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true, accessToken: true } });
  const shop = session?.shop; const token = session?.accessToken;
  if (!shop || !token) return Response.json({ ok: false, error: "No Shopify session." }, { status: 500 });

  const entries = await getPreorderSellingPlanRegistryEntries(shop);
  const orderIds = entries.map((e) => e.supplierOrderId);
  const [orders, settings] = await Promise.all([
    prisma.supplierOrder.findMany({ where: { id: { in: orderIds } }, select: { id: true, status: true, productTitle: true, productId: true } }),
    prisma.preorderBatchSetting.findMany({ where: { supplierOrderId: { in: orderIds } }, select: { supplierOrderId: true, enabled: true } }),
  ]);
  const orderById = new Map(orders.map((o) => [o.id, o]));
  const enabledByOrder = new Map(settings.map((s) => [s.supplierOrderId, s.enabled]));

  const keepProductIds = new Set<string>();
  const orphanEntries: typeof entries = [];
  for (const e of entries) {
    const o = orderById.get(e.supplierOrderId);
    const isKeep = o && o.status === "open" && enabledByOrder.get(e.supplierOrderId) === true;
    if (isKeep) { if (o.productId) keepProductIds.add(numeric(o.productId)); continue; }
    orphanEntries.push(e);
  }
  const keepCount = entries.length - orphanEntries.length;

  if (!apply) {
    return Response.json({
      ok: true, dryRun: true, shop,
      totalRegistered: entries.length, keepCount, orphanCount: orphanEntries.length,
      note: "DRY RUN — nothing changed. Open the same URL with ?apply=1 to remove the orphans. It runs in time-bounded chunks; re-open ?apply=1 until remaining=0.",
    }, { headers: { "Cache-Control": "no-store" } });
  }

  const start = Date.now();
  const BUDGET_MS = 22000;
  let removed = 0; let groupOnly = 0;
  const failed: Array<{ batchId: number; product: string; error: string }> = [];
  for (const e of orphanEntries) {
    if (Date.now() - start > BUDGET_MS) break;
    const o = orderById.get(e.supplierOrderId);
    const overlapsKeep = Boolean(o?.productId && keepProductIds.has(numeric(o.productId)));
    try {
      if (o && !overlapsKeep) {
        // Full deactivate: delete plan + restore DENY + clear metafields/tag.
        await deactivatePreorderSellingPlan({ supplierOrderId: e.supplierOrderId, actor, permissions });
      } else {
        // Group-only (deleted batch, or product shares variants with a live keep).
        await deleteSellingPlanGroupOnly(shop, token, e.sellingPlanGroupId);
        await removePreorderSellingPlanRegistryEntry(shop, e.supplierOrderId);
        groupOnly += 1;
      }
      removed += 1;
    } catch (err) {
      failed.push({ batchId: e.supplierOrderId, product: o?.productTitle ?? "(deleted)", error: err instanceof Error ? err.message : String(err) });
    }
  }

  const after = await getPreorderSellingPlanRegistryEntries(shop);
  const remaining = Math.max(0, after.length - keepCount);
  return Response.json({
    ok: true, applied: true, shop,
    removedThisRun: removed, groupOnlyDeletes: groupOnly, failedThisRun: failed.length, failed,
    keepCount, remainingOrphans: remaining, stillRegistered: after.length,
    note: remaining > 0
      ? `Removed ${removed} this run. ${remaining} orphan(s) left — re-open this same ?apply=1 URL to continue.`
      : `Done. ${after.length} selling plan(s) remain — your ${keepCount} live pre-order(s). Re-run the audit to confirm.`,
  }, { headers: { "Cache-Control": "no-store" } });
};
