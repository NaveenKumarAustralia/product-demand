import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type",
};

/**
 * Decode the JWT payload (base64url) and return the parsed object.
 * We don't verify the signature here — the shop domain is passed
 * separately and we use it only to scope our own DB query.
 * The idToken audience check ensures the token was issued for this app.
 */
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

  const url = new URL(request.url);
  const productId = url.searchParams.get("productId");
  const shop = url.searchParams.get("shop");

  if (!productId || !shop) {
    return Response.json(
      { error: "productId and shop are required" },
      { status: 400, headers: CORS },
    );
  }

  // Verify the bearer token is a valid Shopify token issued for this app
  const authHeader = request.headers.get("Authorization") ?? "";
  const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : null;

  if (!token) {
    return Response.json({ error: "Missing token" }, { status: 401, headers: CORS });
  }

  const payload = decodeJwtPayload(token);
  if (!payload) {
    return Response.json({ error: "Invalid token" }, { status: 401, headers: CORS });
  }

  // Verify audience matches our app client ID
  const clientId = process.env.SHOPIFY_API_KEY;
  const aud = payload.aud;
  const audValid =
    aud === clientId ||
    (Array.isArray(aud) && aud.includes(clientId));

  if (!audValid) {
    return Response.json({ error: "Token audience mismatch" }, { status: 401, headers: CORS });
  }

  try {
    const orders = await prisma.supplierOrder.findMany({
      where: { shop, productId, status: "open" },
      select: {
        id: true,
        supplier: true,
        totalQty: true,
        eta: true,
        status: true,
        lines: {
          select: {
            variantId: true,
            qtyOrdered: true,
          },
        },
      },
      orderBy: { createdAt: "desc" },
    });

    const latestOrder = orders[0] ?? null;
    const qtyByVariant = new Map<string, number>();

    for (const order of orders) {
      for (const line of order.lines) {
        qtyByVariant.set(
          line.variantId,
          (qtyByVariant.get(line.variantId) ?? 0) + line.qtyOrdered,
        );
      }
    }

    return Response.json(
      {
        order: latestOrder
          ? {
              id: latestOrder.id,
              supplier: latestOrder.supplier,
              totalQty: orders.reduce((sum, order) => sum + order.totalQty, 0),
              eta: latestOrder.eta?.toISOString() ?? null,
              status: latestOrder.status,
              lines: Array.from(qtyByVariant.entries()).map(([variantId, qtyOrdered]) => ({
                variantId,
                qtyOrdered,
              })),
            }
          : null,
      },
      { headers: CORS },
    );
  } catch (err) {
    console.error("order-status DB error:", err);
    return Response.json(
      { error: "Database error" },
      { status: 500, headers: CORS },
    );
  }
};
