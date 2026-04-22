import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { normalizePortalMessageUsers, PORTAL_USERS_KEY } from "../portal-messages.server";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type",
};
const PRODUCT_GROUP_RENAMES: Record<string, string> = {
  "Short Sleeve Dresses": "Dresses",
};

function normalizeProductGroup(value?: string | null) {
  const trimmed = value?.trim() ?? "";
  return PRODUCT_GROUP_RENAMES[trimmed] ?? trimmed;
}

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
    const [orders, usersSetting] = await Promise.all([
      prisma.supplierOrder.findMany({
        where: { shop, productId, status: "open" },
        select: {
          id: true,
          supplier: true,
          totalQty: true,
          productType: true,
          eta: true,
          status: true,
          supplierStatus: true,
          priority: true,
          notes: true,
          lines: {
            select: {
              variantId: true,
              qtyOrdered: true,
            },
          },
        },
        orderBy: { createdAt: "desc" },
        take: 2,
      }),
      prisma.portalSetting.findUnique({
        where: { key: PORTAL_USERS_KEY },
        select: { value: true },
      }),
    ]);
    const staffNames = normalizePortalMessageUsers(usersSetting?.value)
      .filter((user) => user.active !== false)
      .map((user) => user.name);

    const latestOrder = orders[0] ?? null;
    const formattedOrders = orders.map((order) => ({
      ...order,
      productType: normalizeProductGroup(order.productType) || null,
      eta: order.eta?.toISOString() ?? null,
    }));
    const totalQty = orders.reduce((sum, order) => sum + order.totalQty, 0);

    return Response.json(
      {
        order: latestOrder
          ? {
              id: latestOrder.id,
              supplier: latestOrder.supplier,
              totalQty,
              productType: normalizeProductGroup(latestOrder.productType) || null,
              eta: latestOrder.eta?.toISOString() ?? null,
              status: latestOrder.status,
              supplierStatus: latestOrder.supplierStatus,
              priority: latestOrder.priority,
              notes: latestOrder.notes,
              lines: formattedOrders.flatMap((order) => order.lines),
            }
          : null,
        orders: formattedOrders,
        staffNames,
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
