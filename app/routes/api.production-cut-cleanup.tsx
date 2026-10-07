import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

// Admin cleanup for the Production Cuts log. The OLD rule logged a cut whenever
// an order entered ANY consumed status (on_production / ready / in_shipment), so
// some records were logged without a real On Order → On Production cut. This
// removes the wrongly-logged ones.
//
//   GET /api/production-cut-cleanup                 → DRY RUN: report every cut
//       record + its order's current status, and how many each mode would delete.
//   GET /api/production-cut-cleanup?apply=1         → delete records whose order
//       is NOT currently in a production status (on_order/cancelled/missing) —
//       these definitely shouldn't have a cut. SAFE: keeps every order that's
//       genuinely in production.
//   GET /api/production-cut-cleanup?apply=1&all=1   → delete ALL cut records
//       (start the log clean; re-flip genuinely-cut orders to repopulate).
//
// Admin only. Dry-run changes nothing.
const PRODUCTION_STATUSES = new Set(["on_production", "ready", "in_shipment"]);

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const url = new URL(request.url);
  const apply = url.searchParams.get("apply") === "1";
  const all = url.searchParams.get("all") === "1";

  const records = await prisma.productionCutRecord.findMany({
    select: { id: true, supplierOrderId: true, productTitle: true, qty: true, productionDate: true },
    orderBy: { productionDate: "desc" },
  });
  const orderIds = records.map((r) => r.supplierOrderId).filter((n): n is number => typeof n === "number");
  const orders = orderIds.length
    ? await prisma.supplierOrder.findMany({ where: { id: { in: orderIds } }, select: { id: true, supplierStatus: true } })
    : [];
  const statusByOrder = new Map(orders.map((o) => [o.id, o.supplierStatus]));

  const rows = records.map((r) => {
    const status = r.supplierOrderId != null ? (statusByOrder.get(r.supplierOrderId) ?? "(order deleted)") : "(no order)";
    const inProduction = PRODUCTION_STATUSES.has(status);
    return {
      id: r.id,
      supplierOrderId: r.supplierOrderId,
      productTitle: r.productTitle,
      qty: r.qty,
      productionDate: r.productionDate,
      currentOrderStatus: status,
      wouldDelete_safe: !inProduction,   // not currently in production → safe to remove
    };
  });

  const safeIds = rows.filter((r) => r.wouldDelete_safe).map((r) => r.id);
  const deletedMode = all ? "all" : "safe";

  if (apply) {
    const where = all ? {} : { id: { in: safeIds } };
    const res = await prisma.productionCutRecord.deleteMany({ where });
    return Response.json({
      ok: true, applied: true, mode: deletedMode, deleted: res.count,
      remaining: records.length - res.count,
      note: all ? "Deleted ALL cut records — re-flip genuinely-cut orders (On Order → On Production) to repopulate." : "Deleted cut records whose order is no longer in production.",
    }, { headers: { "Cache-Control": "no-store" } });
  }

  return Response.json({
    ok: true, applied: false, mode: "dry-run",
    totalRecords: records.length,
    wouldDelete_safe: safeIds.length,          // ?apply=1
    wouldDelete_all: records.length,           // ?apply=1&all=1
    byStatus: rows.reduce((acc: Record<string, number>, r) => { acc[r.currentOrderStatus] = (acc[r.currentOrderStatus] ?? 0) + 1; return acc; }, {}),
    rows,
    hint: "DRY RUN. ?apply=1 removes records whose order isn't in production (safe). ?apply=1&all=1 wipes the whole log. Review rows[] first.",
  }, { headers: { "Cache-Control": "no-store" } });
};
