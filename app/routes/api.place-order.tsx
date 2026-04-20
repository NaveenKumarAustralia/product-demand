import type { ActionFunctionArgs } from "react-router";
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
    supplier: string;
    poNumber?: string;
    eta?: string;
    notes?: string;
    lines: Array<{
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

  const { shop, productId, productTitle, supplier, poNumber, eta, notes, lines } = body;

  if (!shop || !productId || !supplier || !lines?.length) {
    return Response.json(
      { error: "shop, productId, supplier, and lines are required" },
      { status: 400, headers: CORS },
    );
  }

  const totalQty = lines.reduce((sum, l) => sum + (l.qtyOrdered || 0), 0);

  try {
    const order = await prisma.supplierOrder.create({
      data: {
        shop,
        productId,
        productTitle: productTitle || "",
        supplier,
        poNumber: poNumber || null,
        eta: eta ? new Date(eta) : null,
        notes: notes || null,
        totalQty,
        status: "open",
        lines: {
          create: lines.map((l) => ({
            variantId: l.variantId,
            variantTitle: l.variantTitle || "",
            sku: l.sku || null,
            qtyOrdered: l.qtyOrdered,
            costPrice: l.costPrice ?? null,
          })),
        },
      },
      select: { id: true, poNumber: true, totalQty: true },
    });

    return Response.json({ success: true, order }, { headers: CORS });
  } catch (err) {
    console.error("place-order DB error:", err);
    return Response.json({ error: "Database error" }, { status: 500, headers: CORS });
  }
};
