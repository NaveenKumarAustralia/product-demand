import prisma from "../db.server";
import { calculatePreorderCapacity, getPreorderEligibility } from "./preorder-rules.server";

export type PreorderDashboardVariant = {
  variantId: string;
  variantTitle: string;
  sku: string | null;
  qtyOrdered: number;
  qtyReceived: number;
  incomingRemaining: number;
  reservedQty: number;
  safetyBufferQty: number;
  availableToPreorder: number;
  overallocatedBy: number;
};

export type PreorderDashboardBatch = {
  id: number;
  productId: string;
  productTitle: string;
  supplier: string;
  supplierStatus: string;
  destination: string | null;
  market: "AU" | "USA" | null;
  eligible: boolean;
  eligibilityReason: string;
  enabled: boolean;
  safetyBufferPercent: number;
  safetyBufferQty: number | null;
  shipDate: string | null;
  productionEta: string | null;
  pausedReason: string | null;
  totalIncoming: number;
  totalReserved: number;
  totalAvailable: number;
  variants: PreorderDashboardVariant[];
};

export type PreorderDashboardCustomerOrderLine = {
  reservationId: number;
  supplierOrderId: number;
  productId: string | null;
  variantId: string;
  variantTitle: string | null;
  sku: string | null;
  quantity: number;
  status: string;
  expectedShipDate: string | null;
};

export type PreorderDashboardCustomerOrder = {
  shopifyOrderId: string;
  shopifyOrderName: string | null;
  customerEmail: string | null;
  market: string;
  reservedAt: string;
  totalQuantity: number;
  lines: PreorderDashboardCustomerOrderLine[];
};

export type PreorderDashboardData = {
  batches: PreorderDashboardBatch[];
  customerOrders: PreorderDashboardCustomerOrder[];
  totals: {
    activeBatches: number;
    eligibleBatches: number;
    incomingUnits: number;
    reservedUnits: number;
    availableCapacity: number;
    overallocatedUnits: number;
  };
};

export async function loadPreorderDashboardData(): Promise<PreorderDashboardData> {
  const orders = await prisma.supplierOrder.findMany({
    where: {
      status: "open",
      destination: { in: ["send_to_au", "send_to_usa"] },
    },
    select: {
      id: true,
      productId: true,
      productTitle: true,
      supplier: true,
      supplierStatus: true,
      destination: true,
      eta: true,
      createdAt: true,
      lines: {
        select: {
          variantId: true,
          variantTitle: true,
          sku: true,
          qtyOrdered: true,
          qtyReceived: true,
        },
        orderBy: { id: "asc" },
      },
    },
    orderBy: [{ eta: "asc" }, { createdAt: "asc" }, { id: "asc" }],
  });

  const orderIds = orders.map((order) => order.id);
  const [settings, reservations] = await Promise.all([
    orderIds.length
      ? prisma.preorderBatchSetting.findMany({ where: { supplierOrderId: { in: orderIds } } })
      : Promise.resolve([]),
    orderIds.length
      ? prisma.preorderReservation.findMany({
          where: { supplierOrderId: { in: orderIds } },
          orderBy: [{ reservedAt: "desc" }, { id: "desc" }],
          take: 2000,
        })
      : Promise.resolve([]),
  ]);
  const byOrder = new Map(settings.map((setting) => [setting.supplierOrderId, setting]));

  const reservedByBatchVariant = new Map<string, number>();
  for (const reservation of reservations) {
    if (reservation.status !== "reserved") continue;
    const key = `${reservation.supplierOrderId}:${reservation.variantId}`;
    reservedByBatchVariant.set(key, (reservedByBatchVariant.get(key) ?? 0) + reservation.quantity);
  }

  const batches: PreorderDashboardBatch[] = orders.map((order) => {
    const setting = byOrder.get(order.id);
    const eligibility = getPreorderEligibility({
      supplierStatus: order.supplierStatus,
      destination: order.destination,
      preorderEnabled: setting?.enabled ?? false,
    });

    const variants = order.lines.map((line) => {
      const reservedQty = reservedByBatchVariant.get(`${order.id}:${line.variantId}`) ?? 0;
      const incomingRemaining = Math.max(0, line.qtyOrdered - line.qtyReceived);
      const capacity = calculatePreorderCapacity({
        confirmedIncomingQty: incomingRemaining,
        reservedQty,
        safetyBufferPercent: setting?.safetyBufferPercent ?? 5,
        safetyBufferQty: setting?.safetyBufferQty ?? null,
      });
      return {
        variantId: line.variantId,
        variantTitle: line.variantTitle,
        sku: line.sku ?? null,
        qtyOrdered: line.qtyOrdered,
        qtyReceived: line.qtyReceived,
        incomingRemaining,
        reservedQty,
        safetyBufferQty: capacity.safetyBufferQty,
        availableToPreorder: capacity.availableToPreorder,
        overallocatedBy: capacity.overallocatedBy,
      };
    });

    return {
      id: order.id,
      productId: order.productId,
      productTitle: order.productTitle,
      supplier: order.supplier,
      supplierStatus: order.supplierStatus,
      destination: order.destination,
      market: eligibility.market,
      eligible: eligibility.eligible,
      eligibilityReason: eligibility.reason,
      enabled: setting?.enabled ?? false,
      safetyBufferPercent: setting?.safetyBufferPercent ?? 5,
      safetyBufferQty: setting?.safetyBufferQty ?? null,
      shipDate: setting?.shipDate?.toISOString() ?? null,
      productionEta: order.eta?.toISOString() ?? null,
      pausedReason: setting?.pausedReason ?? null,
      totalIncoming: variants.reduce((sum, row) => sum + row.incomingRemaining, 0),
      totalReserved: variants.reduce((sum, row) => sum + row.reservedQty, 0),
      totalAvailable: variants.reduce((sum, row) => sum + row.availableToPreorder, 0),
      variants,
    };
  });

  const customerOrderMap = new Map<string, PreorderDashboardCustomerOrder>();
  for (const reservation of reservations) {
    let item = customerOrderMap.get(reservation.shopifyOrderId);
    if (!item) {
      item = {
        shopifyOrderId: reservation.shopifyOrderId,
        shopifyOrderName: reservation.shopifyOrderName,
        customerEmail: reservation.customerEmail,
        market: reservation.market,
        reservedAt: reservation.reservedAt.toISOString(),
        totalQuantity: 0,
        lines: [],
      };
      customerOrderMap.set(reservation.shopifyOrderId, item);
    }
    item.totalQuantity += reservation.quantity;
    item.lines.push({
      reservationId: reservation.id,
      supplierOrderId: reservation.supplierOrderId,
      productId: reservation.productId,
      variantId: reservation.variantId,
      variantTitle: reservation.variantTitle,
      sku: reservation.sku,
      quantity: reservation.quantity,
      status: reservation.status,
      expectedShipDate: reservation.expectedShipDate?.toISOString() ?? null,
    });
  }
  const customerOrders = Array.from(customerOrderMap.values()).sort(
    (a, b) => new Date(b.reservedAt).getTime() - new Date(a.reservedAt).getTime(),
  );

  return {
    batches,
    customerOrders,
    totals: {
      activeBatches: batches.filter((batch) => batch.enabled && batch.eligible).length,
      eligibleBatches: batches.filter((batch) => batch.supplierStatus === "on_production").length,
      incomingUnits: batches.reduce((sum, batch) => sum + batch.totalIncoming, 0),
      reservedUnits: batches.reduce((sum, batch) => sum + batch.totalReserved, 0),
      availableCapacity: batches.reduce((sum, batch) => sum + batch.totalAvailable, 0),
      overallocatedUnits: batches.reduce(
        (sum, batch) => sum + batch.variants.reduce((variantSum, row) => variantSum + row.overallocatedBy, 0),
        0,
      ),
    },
  };
}
