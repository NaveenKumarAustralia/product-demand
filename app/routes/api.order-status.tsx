import type { LoaderFunctionArgs } from "react-router";
import { authenticate } from "../shopify.server";
import prisma from "../db.server";

// Required for cross-origin requests from the admin extension iframe
const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type",
};

export const loader = async ({ request }: LoaderFunctionArgs) => {
  // Handle CORS preflight
  if (request.method === "OPTIONS") {
    return new Response(null, { status: 204, headers: CORS });
  }

  let shop: string;

  try {
    const { session } = await authenticate.admin(request);
    shop = session.shop;
  } catch {
    return Response.json(
      { error: "Unauthorized" },
      { status: 401, headers: CORS },
    );
  }

  const url = new URL(request.url);
  const productId = url.searchParams.get("productId");

  if (!productId) {
    return Response.json(
      { error: "productId is required" },
      { status: 400, headers: CORS },
    );
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
    return Response.json(
      { error: "Database error" },
      { status: 500, headers: CORS },
    );
  }
};
