import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { syncOrderNoteMessages } from "../portal-messages.server";
import { authorizeApiRequest } from "../api-auth.server";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type, X-Api-Key",
};
const PRODUCT_GROUP_RENAMES: Record<string, string> = {
  "Short Sleeve Dresses": "Dresses",
};

function normalizeProductGroup(value?: string | null) {
  const trimmed = value?.trim() ?? "";
  return PRODUCT_GROUP_RENAMES[trimmed] ?? trimmed;
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

  // Accepts either a Shopify session token from the embedded app or the
  // dashboard's shared secret. See api-auth.server.ts.
  const unauthorized = authorizeApiRequest(request, CORS);
  if (unauthorized) return unauthorized;

  let body: {
    shop: string;
    orderId: number;
    supplierStatus?: string;
    priority?: string;
    productType?: string;
    destination?: string;
    eta?: string | null;
    notes?: string;
  };

  try {
    body = await request.json();
  } catch {
    return Response.json({ error: "Invalid JSON body" }, { status: 400, headers: CORS });
  }

  const { shop, orderId, supplierStatus, priority, productType, destination, eta, notes } = body;
  const id = Number(orderId);

  if (!shop || !Number.isInteger(id)) {
    return Response.json({ error: "shop and orderId are required" }, { status: 400, headers: CORS });
  }

  const data: Record<string, unknown> = {};
  if (supplierStatus !== undefined) data.supplierStatus = supplierStatus;
  if (priority !== undefined) data.priority = priority || null;
  if (productType !== undefined) data.productType = normalizeProductGroup(productType) || null;
  // Trimmed and length-capped to match how the portal's own destination editor
  // stores it, so a value set from either place looks the same in the table.
  if (destination !== undefined) data.destination = String(destination).trim().slice(0, 64) || null;
  if (eta !== undefined) data.eta = eta ? new Date(eta) : null;
  if (notes !== undefined) data.notes = notes || null;

  if (!Object.keys(data).length) {
    return Response.json({ error: "No updates provided" }, { status: 400, headers: CORS });
  }

  try {
    const existingOrder = await prisma.supplierOrder.findFirst({
      where: { id, shop, status: "open" },
      select: { id: true },
    });

    if (!existingOrder) {
      return Response.json({ error: "Open order not found" }, { status: 404, headers: CORS });
    }

    const order = await prisma.supplierOrder.update({
      where: { id },
      data,
      select: {
        id: true,
        supplierStatus: true,
        priority: true,
        productType: true,
        destination: true,
        eta: true,
        notes: true,
      },
    });

    if (notes !== undefined) {
      await syncOrderNoteMessages({
        orderId: id,
        field: "notes",
        text: notes,
        fromName: "Shopify block",
      });
    }

    return Response.json({
      success: true,
      order: {
        ...order,
        productType: normalizeProductGroup(order.productType) || null,
        eta: order.eta?.toISOString() ?? null,
      },
    }, { headers: CORS });
  } catch (err) {
    console.error("update-order DB error:", err);
    return Response.json({ error: "Database error" }, { status: 500, headers: CORS });
  }
};
