import { createHmac } from "crypto";
import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// CORS headers required for cross-origin requests from the admin extension
const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type",
};

/**
 * Verify a Shopify session token (HS256 JWT signed with the app secret)
 * and return the shop domain, or null if invalid.
 */
function verifySessionToken(token: string): string | null {
  try {
    const secret = process.env.SHOPIFY_API_SECRET;
    if (!secret) return null;

    const parts = token.split(".");
    if (parts.length !== 3) return null;

    const [headerB64, payloadB64, sigB64] = parts;

    // Verify HMAC-SHA256 signature
    const expected = createHmac("sha256", secret)
      .update(`${headerB64}.${payloadB64}`)
      .digest("base64url");

    if (expected !== sigB64) return null;

    // Decode payload — dest = "https://{shop}"
    const payload = JSON.parse(
      Buffer.from(payloadB64, "base64url").toString("utf-8"),
    );

    if (!payload.dest) return null;
    return new URL(payload.dest as string).hostname;
  } catch {
    return null;
  }
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  // Handle CORS preflight
  if (request.method === "OPTIONS") {
    return new Response(null, { status: 204, headers: CORS });
  }

  // Extract and verify the bearer token sent by the extension
  const authHeader = request.headers.get("Authorization") ?? "";
  const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : null;

  if (!token) {
    return Response.json({ error: "Missing token" }, { status: 401, headers: CORS });
  }

  const shop = verifySessionToken(token);
  if (!shop) {
    return Response.json({ error: "Invalid token" }, { status: 401, headers: CORS });
  }

  const url = new URL(request.url);
  const productId = url.searchParams.get("productId");

  if (!productId) {
    return Response.json({ error: "productId is required" }, { status: 400, headers: CORS });
  }

  try {
    const order = await prisma.supplierOrder.findFirst({
      where: { shop, productId, status: "open" },
      select: {
        id: true,
        supplier: true,
        totalQty: true,
        eta: true,
        status: true,
      },
      orderBy: { createdAt: "desc" },
    });

    return Response.json(
      {
        order: order
          ? {
              id: order.id,
              supplier: order.supplier,
              totalQty: order.totalQty,
              eta: order.eta?.toISOString() ?? null,
              status: order.status,
            }
          : null,
      },
      { headers: CORS },
    );
  } catch (err) {
    console.error("order-status DB error:", err);
    return Response.json({ error: "Database error" }, { status: 500, headers: CORS });
  }
};
