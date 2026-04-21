import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type",
};

function decodeJwtPayload(token: string): Record<string, unknown> | null {
  try {
    const [, payloadB64] = token.split(".");
    return JSON.parse(Buffer.from(payloadB64, "base64url").toString("utf-8"));
  } catch {
    return null;
  }
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  if (request.method === "OPTIONS") {
    return new Response(null, { status: 204, headers: CORS });
  }
  return new Response(null, { status: 405, headers: CORS });
};

export const action = async ({ request }: ActionFunctionArgs) => {
  if (request.method === "OPTIONS") {
    return new Response(null, { status: 204, headers: CORS });
  }

  if (request.method !== "POST") {
    return Response.json({ error: "Method not allowed" }, { status: 405, headers: CORS });
  }

  const authHeader = request.headers.get("Authorization") ?? "";
  const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : null;

  if (!token) {
    return Response.json({ error: "Missing token" }, { status: 401, headers: CORS });
  }

  const payload = decodeJwtPayload(token);
  if (!payload) {
    return Response.json({ error: "Invalid token" }, { status: 401, headers: CORS });
  }

  const clientId = process.env.SHOPIFY_API_KEY;
  const aud = payload.aud;
  const audValid = aud === clientId || (Array.isArray(aud) && aud.includes(clientId));

  if (!audValid) {
    return Response.json({ error: "Token audience mismatch" }, { status: 401, headers: CORS });
  }

  let body: {
    shop: string;
    productId: string;
    productTitle: string;
    productType?: string;
    productImageUrl?: string;
    supplier: string;
    poNumber?: string;
    eta?: string;
    notes?: string;
    priority?: string;
    existingOrderId?: number;
    lines?: Array<{
      variantId: string;
      variantTitle: string;
      sku?: string;
      qtyOrdered: number;
      costPrice?: number;
    }>;
  };

  try {
    body = await request.json();
  } catch {
    return Response.json({ error: "Invalid JSON body" }, { status: 400, headers: CORS });
  }

  const { shop, productId, productTitle, productType, productImageUrl, supplier, poNumber, eta, notes, priority, existingOrderId, lines } = body;

  if (!shop || !productId || !supplier) {
    return Response.json(
      { error: "shop, productId, and supplier are required" },
      { status: 400, headers: CORS },
    );
  }

  const orderLines = (lines ?? []).filter((line) => Number(line.qtyOrdered || 0) > 0);

  if (!orderLines.length && !notes?.trim()) {
    return Response.json(
      { error: "Add at least one quantity or an order note" },
      { status: 400, headers: CORS },
    );
  }

  const totalQty = orderLines.reduce((sum, l) => sum + (l.qtyOrdered || 0), 0);

  try {
    if (existingOrderId != null) {
      const orderId = Number(existingOrderId);
      if (!Number.isInteger(orderId)) {
        return Response.json(
          { error: "existingOrderId must be a valid order id" },
          { status: 400, headers: CORS },
        );
      }

      const existingOrder = await prisma.supplierOrder.findFirst({
        where: { id: orderId, shop, productId, status: "open" },
        select: {
          id: true,
          notes: true,
          totalQty: true,
          lines: {
            select: {
              id: true,
              variantId: true,
              qtyOrdered: true,
            },
          },
        },
      });

      if (!existingOrder) {
        return Response.json(
          { error: "Existing open order was not found for this product" },
          { status: 404, headers: CORS },
        );
      }

      const trimmedNotes = notes?.trim();
      const nextNotes = trimmedNotes
        ? [existingOrder.notes, trimmedNotes].filter(Boolean).join("\n")
        : existingOrder.notes;

      const order = await prisma.$transaction(async (tx) => {
        const updatedOrder = await tx.supplierOrder.update({
          where: { id: existingOrder.id },
          data: {
            notes: nextNotes || null,
            totalQty: existingOrder.totalQty + totalQty,
            productType: productType?.trim() || undefined,
            priority: priority || undefined,
          },
          select: { id: true, poNumber: true, totalQty: true },
        });

        for (const line of orderLines) {
          const existingLine = existingOrder.lines.find((item) => item.variantId === line.variantId);
          if (existingLine) {
            await tx.orderLine.update({
              where: { id: existingLine.id },
              data: { qtyOrdered: existingLine.qtyOrdered + line.qtyOrdered },
            });
          } else {
            await tx.orderLine.create({
              data: {
                orderId: existingOrder.id,
                variantId: line.variantId,
                variantTitle: line.variantTitle || "",
                sku: line.sku || null,
                qtyOrdered: line.qtyOrdered,
                costPrice: line.costPrice ?? null,
              },
            });
          }
        }

        return updatedOrder;
      });

      return Response.json({ success: true, order }, { headers: CORS });
    }

    const openOrderCount = await prisma.supplierOrder.count({
      where: { shop, productId, status: "open" },
    });

    if (openOrderCount >= 2) {
      return Response.json(
        { error: "This product already has 2 open orders. Add to an existing order instead." },
        { status: 400, headers: CORS },
      );
    }

    const order = await prisma.supplierOrder.create({
      data: {
        shop,
        productId,
        productTitle: productTitle || "",
        productType: productType?.trim() || null,
        productImageUrl: productImageUrl || null,
        supplier,
        supplierStatus: "on_order",
        priority: priority || null,
        poNumber: poNumber || null,
        eta: eta ? new Date(eta) : null,
        notes: notes || null,
        totalQty,
        status: "open",
        lines: orderLines.length
          ? {
              create: orderLines.map((l) => ({
                variantId: l.variantId,
                variantTitle: l.variantTitle || "",
                sku: l.sku || null,
                qtyOrdered: l.qtyOrdered,
                costPrice: l.costPrice ?? null,
              })),
            }
          : undefined,
      },
      select: { id: true, poNumber: true, totalQty: true },
    });

    return Response.json({ success: true, order }, { headers: CORS });
  } catch (err) {
    console.error("place-order DB error:", err);
    return Response.json({ error: "Database error" }, { status: 500, headers: CORS });
  }
};
