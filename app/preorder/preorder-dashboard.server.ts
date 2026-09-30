import prisma from "../db.server";
import { getPreorderLocationSettings } from "./preorder-locations.server";
import { getPreorderNotifyEnabled, getPreorderCombineWindowDays } from "./preorder-storefront-settings.server";
import { getPreorderPermissionContext } from "./preorder-permissions.server";
import { calculatePreorderCapacity, getPreorderEligibility, isPreorderEligibleStatus } from "./preorder-rules.server";
import { getPreorderSellingPlanRegistryEntries } from "./preorder-selling-plan-registry.server";

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
  imageUrl: string | null;
  supplier: string;
  supplierStatus: string;
  destination: string | null;
  market: "AU" | "USA" | null;
  eligible: boolean;
  eligibilityReason: string;
  enabled: boolean;
  shopifySellingPlanActive: boolean;
  shopifySellingPlanGroupId: string | null;
  shopifySellingPlanId: string | null;
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
  productTitle: string | null;
  imageUrl: string | null;
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
  // Declared order value = the Shopify order's total price (best-effort, fetched
  // from Shopify). null when it couldn't be fetched.
  orderValue: number | null;
  orderCurrency: string | null;
  lines: PreorderDashboardCustomerOrderLine[];
};

export type PreorderDashboardData = {
  batches: PreorderDashboardBatch[];
  customerOrders: PreorderDashboardCustomerOrder[];
  configuration: {
    locations: { AU: string | null; USA: string | null };
    notifyBlockEnabled: boolean;
    combineWindowDays: number;
    users: Array<{ id: string; name: string; admin: boolean }>;
    permissions: {
      managePreorderUserIds: string[];
      manageEtaUserIds: string[];
      manageSafetyBufferUserIds: string[];
      sendNotificationUserIds: string[];
      viewReportsUserIds: string[];
    };
  };
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
      productImageUrl: true,
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
  const [settings, reservations, permissionContext, locations, sellingPlanEntries, notifyBlockEnabled, combineWindowDays] = await Promise.all([
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
    getPreorderPermissionContext(),
    getPreorderLocationSettings(),
    getPreorderSellingPlanRegistryEntries(),
    getPreorderNotifyEnabled(),
    getPreorderCombineWindowDays(),
  ]);
  const byOrder = new Map(settings.map((setting) => [setting.supplierOrderId, setting]));
  const sellingPlanByOrder = new Map(sellingPlanEntries.map((entry) => [entry.supplierOrderId, entry]));

  const reservedByBatchVariant = new Map<string, number>();
  for (const reservation of reservations) {
    if (reservation.status !== "reserved") continue;
    const key = `${reservation.supplierOrderId}:${reservation.variantId}`;
    reservedByBatchVariant.set(key, (reservedByBatchVariant.get(key) ?? 0) + reservation.quantity);
  }

  const batches: PreorderDashboardBatch[] = orders.map((order) => {
    const setting = byOrder.get(order.id);
    const sellingPlan = sellingPlanByOrder.get(order.id);
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
        safetyBufferPercent: setting?.safetyBufferPercent ?? 0,
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
      imageUrl: order.productImageUrl ?? null,
      supplier: order.supplier,
      supplierStatus: order.supplierStatus,
      destination: order.destination,
      market: eligibility.market,
      eligible: eligibility.eligible,
      eligibilityReason: eligibility.reason,
      enabled: setting?.enabled ?? false,
      shopifySellingPlanActive: Boolean(sellingPlan),
      shopifySellingPlanGroupId: sellingPlan?.sellingPlanGroupId ?? null,
      shopifySellingPlanId: sellingPlan?.sellingPlanId ?? null,
      safetyBufferPercent: setting?.safetyBufferPercent ?? 0,
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

  const batchInfoById = new Map(orders.map((o) => [o.id, { productTitle: o.productTitle, imageUrl: o.productImageUrl }]));
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
        orderValue: null,
        orderCurrency: null,
        lines: [],
      };
      customerOrderMap.set(reservation.shopifyOrderId, item);
    }
    item.totalQuantity += reservation.quantity;
    item.lines.push({
      reservationId: reservation.id,
      supplierOrderId: reservation.supplierOrderId,
      productId: reservation.productId,
      productTitle: batchInfoById.get(reservation.supplierOrderId)?.productTitle ?? null,
      imageUrl: batchInfoById.get(reservation.supplierOrderId)?.imageUrl ?? null,
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

  // Declared order value = the Shopify order's total price. Batch-fetched from
  // Shopify (capped to the most recent 300 orders to bound the call). Best-effort:
  // any failure just leaves orderValue null (the column shows "—").
  try {
    const sess = await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true, accessToken: true } });
    if (sess?.shop && sess.accessToken) {
      const numOf = (id: string) => String(id ?? "").replace(/\D/g, "");
      const capped = customerOrders.slice(0, 300);
      const gids = Array.from(new Set(capped.map((o) => numOf(o.shopifyOrderId)).filter(Boolean).map((n) => `gid://shopify/Order/${n}`)));
      const valueByOrder = new Map<string, { amount: number; currency: string }>();
      for (let i = 0; i < gids.length; i += 250) {
        const chunk = gids.slice(i, i + 250);
        const res = await fetch(`https://${sess.shop}/admin/api/2025-10/graphql.json`, {
          method: "POST",
          headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": sess.accessToken },
          body: JSON.stringify({ query: `query OrderValues($ids: [ID!]!) { nodes(ids: $ids) { ... on Order { id totalPriceSet { shopMoney { amount currencyCode } } } } }`, variables: { ids: chunk } }),
        }).catch(() => null);
        const j = res ? await res.json().catch(() => null) as { data?: { nodes?: Array<{ id?: string; totalPriceSet?: { shopMoney?: { amount?: string; currencyCode?: string } } } | null> } } : null;
        for (const n of j?.data?.nodes ?? []) {
          const num = numOf(String(n?.id ?? ""));
          const m = n?.totalPriceSet?.shopMoney;
          if (num && m?.amount != null) valueByOrder.set(num, { amount: Number(m.amount) || 0, currency: String(m.currencyCode ?? "") });
        }
      }
      for (const o of customerOrders) {
        const v = valueByOrder.get(numOf(o.shopifyOrderId));
        if (v) { o.orderValue = v.amount; o.orderCurrency = v.currency; }
      }
    }
  } catch { /* best-effort; leave orderValue null */ }

  return {
    batches,
    customerOrders,
    configuration: {
      locations,
      notifyBlockEnabled,
      combineWindowDays,
      users: permissionContext.users.map((user) => ({
        id: user.id,
        name: user.name,
        admin: user.admin === true,
      })),
      permissions: permissionContext.permissions,
    },
    totals: {
      activeBatches: batches.filter((batch) => batch.enabled && batch.eligible).length,
      eligibleBatches: batches.filter((batch) => isPreorderEligibleStatus(batch.supplierStatus)).length,
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
