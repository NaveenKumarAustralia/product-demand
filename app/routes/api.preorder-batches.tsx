import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { authorizeApiRequest } from "../api-auth.server";
import {
  calculatePreorderCapacity,
  getPreorderEligibility,
} from "../preorder/preorder-rules.server";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type, X-Api-Key",
};

export const loader = async ({ request }: LoaderFunctionArgs) => {
  if (request.method === "OPTIONS") {
    return new Response(null, { status: 204, headers: CORS });
  }

  const unauthorized = authorizeApiRequest(request, CORS);
  if (unauthorized) return unauthorized;

  const url = new URL(request.url);
  const shop = (url.searchParams.get("shop") ?? "").trim();
  const market = (url.searchParams.get("market") ?? "").trim().toUpperCase();
  const enabledOnly = url.searchParams.get("enabled") === "1";
  if (!shop) {
    return Response.json({ error: "shop is required" }, { status: 400, headers: CORS });
  }
  if (market && market !== "AU" && market !== "USA") {
    return Response.json({ error: "market must be AU or USA" }, { status: 400, headers: CORS });
  }

  const orders = await prisma.supplierOrder.findMany({
    where: {
      shop,
      status: "open",
      ...(market === "AU" ? { destination: "send_to_au" } : {}),
      ...(market === "USA" ? { destination: "send_to_usa" } : {}),
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

  const settings = orders.length
    ? await prisma.preorderBatchSetting.findMany({
        where: { supplierOrderId: { in: orders.map((order) => order.id) } },
      })
    : [];
  const settingByOrder = new Map(settings.map((setting) => [setting.supplierOrderId, setting]));

  const batches = orders
    .map((order) => {
      const setting = settingByOrder.get(order.id);
      const eligibility = getPreorderEligibility({
        supplierStatus: order.supplierStatus,
        destination: order.destination,
        preorderEnabled: setting?.enabled ?? false,
      });

      const variants = order.lines.map((line) => {
        // Reservation ledger is introduced in the allocation phase. Until then,
        // reservedQty is deliberately zero rather than inferred from negative
        // Shopify inventory or order quantities.
        const reservedQty = 0;
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
          sku: line.sku,
          qtyOrdered: line.qtyOrdered,
          qtyReceived: line.qtyReceived,
          incomingRemaining,
          ...capacity,
        };
      });

      return {
        supplierOrderId: order.id,
        productId: order.productId,
        productTitle: order.productTitle,
        supplier: order.supplier,
        supplierStatus: order.supplierStatus,
        destination: order.destination,
        market: eligibility.market,
        eligibility,
        enabled: setting?.enabled ?? false,
        safetyBufferPercent: setting?.safetyBufferPercent ?? 5,
        safetyBufferQty: setting?.safetyBufferQty ?? null,
        shipDate: setting?.shipDate?.toISOString() ?? null,
        productionEta: order.eta?.toISOString() ?? null,
        pausedReason: setting?.pausedReason ?? null,
        enabledAt: setting?.enabledAt?.toISOString() ?? null,
        variants,
        totalIncomingRemaining: variants.reduce((sum, variant) => sum + variant.incomingRemaining, 0),
        totalAvailableToPreorder: variants.reduce((sum, variant) => sum + variant.availableToPreorder, 0),
      };
    })
    .filter((batch) => !enabledOnly || batch.enabled);

  return Response.json({ ok: true, batches }, { headers: CORS });
};
