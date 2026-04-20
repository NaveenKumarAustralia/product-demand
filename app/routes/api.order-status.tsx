import type { LoaderFunctionArgs } from "react-router";
import { authenticate } from "../shopify.server";
import prisma from "../db.server";

/**
 * GET /api/order-status?productId=<shopify-gid>
 *
 * Called by the Admin UI Extension (product-order-block) to check whether
 * a product already has an open supplier order.
 *
 * Authenticated via the session token that the extension sends in the
 * Authorization: Bearer <token> header.
 */
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const { session } = await authenticate.admin(request);

  const url = new URL(request.url);
  const productId = url.searchParams.get("productId");

  if (!productId) {
    return Response.json({ error: "productId is required" }, { status: 400 });
  }

  const order = await prisma.supplierOrder.findFirst({
    where: {
      shop: session.shop,
      productId,
      status: "open",
    },
    select: {
      id: true,
      supplier: true,
      totalQty: true,
      eta: true,
      status: true,
    },
    orderBy: { createdAt: "desc" },
  });

  return Response.json({
    order: order
      ? {
          id: order.id,
          supplier: order.supplier,
          totalQty: order.totalQty,
          eta: order.eta?.toISOString() ?? null,
          status: order.status,
        }
      : null,
  });
};
