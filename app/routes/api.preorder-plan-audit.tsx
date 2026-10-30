import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { getPreorderSellingPlanRegistryEntries } from "../preorder/preorder-selling-plan-registry.server";

// Admin diagnostic (READ-ONLY): every Shopify pre-order selling plan the app has
// registered, classified KEEP vs ORPHAN. A product stays pre-orderable on the
// storefront (and its variants stay CONTINUE / oversellable) as long as its
// selling plan is attached — even after its batch is closed or pre-order is
// turned off in the portal. KEEP = the batch is still open AND pre-order enabled;
// everything else is an ORPHAN whose selling plan should be removed (which also
// restores DENY inventory policy). No changes are made here.
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const session = await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true } });
  const shop = session?.shop;
  if (!shop) return Response.json({ ok: false, error: "No Shopify session." }, { status: 500 });

  const entries = await getPreorderSellingPlanRegistryEntries(shop);
  const orderIds = entries.map((e) => e.supplierOrderId);

  const [orders, settings] = await Promise.all([
    prisma.supplierOrder.findMany({ where: { id: { in: orderIds } }, select: { id: true, status: true, productTitle: true, destination: true } }),
    prisma.preorderBatchSetting.findMany({ where: { supplierOrderId: { in: orderIds } }, select: { supplierOrderId: true, enabled: true } }),
  ]);
  const orderById = new Map(orders.map((o) => [o.id, o]));
  const enabledByOrder = new Map(settings.map((s) => [s.supplierOrderId, s.enabled]));

  type Row = { batchId: number; product: string; status: string; enabled: boolean; destination: string | null; reason: string };
  const keep: Row[] = [];
  const orphans: Row[] = [];
  for (const e of entries) {
    const order = orderById.get(e.supplierOrderId);
    const enabled = enabledByOrder.get(e.supplierOrderId) === true;
    if (!order) {
      orphans.push({ batchId: e.supplierOrderId, product: "(batch deleted)", status: "missing", enabled: false, destination: null, reason: "batch no longer exists" });
      continue;
    }
    const row: Row = { batchId: order.id, product: order.productTitle, status: order.status, enabled, destination: order.destination, reason: "" };
    if (order.status === "open" && enabled) {
      row.reason = "open + pre-order enabled";
      keep.push(row);
    } else {
      row.reason = order.status !== "open" ? `batch ${order.status} (not open)` : "pre-order turned off";
      orphans.push(row);
    }
  }

  const byProduct = (rows: Row[]) => [...rows].sort((a, b) => a.product.localeCompare(b.product));

  return Response.json({
    ok: true,
    shop,
    totalRegistered: entries.length,
    keepCount: keep.length,
    orphanCount: orphans.length,
    keep: byProduct(keep),
    orphans: byProduct(orphans),
    note: "Run the cleanup to remove the ORPHANS (deletes each one's Shopify selling plan + restores DENY inventory policy). KEEP entries are left live.",
  }, { headers: { "Cache-Control": "no-store" } });
};
