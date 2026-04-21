import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type",
};
const PRODUCT_GROUP_RENAMES: Record<string, string> = {
  "Short Sleeve Dresses": "Dresses",
};

function normalizeProductGroup(value?: string | null) {
  const trimmed = value?.trim() ?? "";
  return PRODUCT_GROUP_RENAMES[trimmed] ?? trimmed;
}

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
    orderId: number;
    supplierStatus?: string;
    priority?: string;
    productType?: string;
    eta?: string | null;
    notes?: string;
  };

  try {
    body = await request.json();
  } catch {
    return Response.json({ error: "Invalid JSON body" }, { status: 400, headers: CORS });
  }

  const { shop, orderId, supplierStatus, priority, productType, eta, notes } = body;
  const id = Number(orderId);

  if (!shop || !Number.isInteger(id)) {
    return Response.json({ error: "shop and orderId are required" }, { status: 400, headers: CORS });
  }

  const data: Record<string, unknown> = {};
  if (supplierStatus !== undefined) data.supplierStatus = supplierStatus;
  if (priority !== undefined) data.priority = priority || null;
  if (productType !== undefined) data.productType = normalizeProductGroup(productType) || null;
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
        eta: true,
        notes: true,
      },
    });

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
