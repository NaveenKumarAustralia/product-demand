import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import {
  canViewPreorderReports,
  getPreorderPermissionContext,
} from "../preorder/preorder-permissions.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  try {
    const actor = await requirePreorderPortalUser(request);
    const { permissions } = await getPreorderPermissionContext();
    if (!canViewPreorderReports(actor, permissions)) {
      return Response.json({ ok: false, error: "You do not have permission to view preorder reports." }, { status: 403 });
    }

    const [statusMarket, recentReservations, recentFailures] = await Promise.all([
      prisma.preorderReservation.groupBy({
        by: ["status", "market"],
        _sum: { quantity: true },
        _count: { _all: true },
      }),
      prisma.preorderReservation.findMany({
        select: {
          shopifyOrderId: true,
          supplierOrderId: true,
          quantity: true,
          status: true,
          market: true,
          reservedAt: true,
        },
        orderBy: { reservedAt: "desc" },
        take: 5000,
      }),
      prisma.activityLog.findMany({
        where: { action: "preorder_allocation_failed" },
        select: {
          id: true,
          entityId: true,
          entityName: true,
          toValue: true,
          createdAt: true,
        },
        orderBy: { createdAt: "desc" },
        take: 20,
      }),
    ]);

    const uniqueOrders = new Set(recentReservations.map((row) => row.shopifyOrderId));
    const activeByBatch = new Map<number, number>();
    for (const row of recentReservations) {
      if (row.status !== "reserved") continue;
      activeByBatch.set(row.supplierOrderId, (activeByBatch.get(row.supplierOrderId) ?? 0) + row.quantity);
    }

    const batchIds = Array.from(activeByBatch.keys());
    const batches = batchIds.length
      ? await prisma.supplierOrder.findMany({
          where: { id: { in: batchIds } },
          select: {
            id: true, productTitle: true, supplier: true, destination: true,
            lines: { select: { qtyOrdered: true, qtyReceived: true } },
          },
        })
      : [];
    const incomingByBatch = new Map<number, number>(
      batches.map((batch) => [
        batch.id,
        batch.lines.reduce((sum, line) => sum + Math.max(0, line.qtyOrdered - line.qtyReceived), 0),
      ]),
    );

    return Response.json({
      ok: true,
      summary: {
        customerOrders: uniqueOrders.size,
        reservationRows: recentReservations.length,
        quantities: statusMarket.map((row) => ({
          status: row.status,
          market: row.market,
          quantity: row._sum.quantity ?? 0,
          rows: row._count._all,
        })),
      },
      activeBatches: batches
        .map((batch) => {
          const reservedQty = activeByBatch.get(batch.id) ?? 0;
          const incomingQty = incomingByBatch.get(batch.id) ?? 0;
          return {
            id: batch.id,
            productTitle: batch.productTitle,
            supplier: batch.supplier,
            destination: batch.destination,
            reservedQty,
            incomingQty,
            fillPercent: incomingQty > 0 ? Math.min(100, Math.round((reservedQty / incomingQty) * 100)) : 0,
          };
        })
        .sort((a, b) => b.reservedQty - a.reservedQty),
      recentFailures: recentFailures.map((row) => ({
        id: row.id,
        shopifyOrderId: row.entityId,
        orderName: row.entityName,
        message: row.toValue,
        createdAt: row.createdAt.toISOString(),
      })),
    });
  } catch (error) {
    if (error instanceof Response) return error;
    console.error("[preorder report] failed:", error);
    return Response.json({ ok: false, error: "Could not load preorder report." }, { status: 500 });
  }
};
