import type { LoaderFunctionArgs } from "react-router";
import { createHash, timingSafeEqual } from "node:crypto";
import prisma from "../db.server";

// Independent credential; never accepts dashboard keys or unverified JWTs.
// Disabled unless both environment variables are configured.
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const headers = { "Cache-Control": "no-store" };
  const expected = process.env.AI_PORTAL_READ_TOKEN;
  const shop = process.env.AI_PORTAL_SHOP;
  const supplied = request.headers.get("Authorization");
  if (!expected || expected.length < 32 || !shop || !supplied?.startsWith("Bearer ")) {
    return Response.json({ error: "Unauthorized" }, { status: 401, headers });
  }
  const digest = (s: string) => createHash("sha256").update(s).digest();
  if (!timingSafeEqual(digest(supplied.slice(7)), digest(expected))) {
    return Response.json({ error: "Unauthorized" }, { status: 401, headers });
  }
  const url = new URL(request.url);
  const rawCursor = url.searchParams.get("cursor");
  const cursor = rawCursor === null ? null : Number(rawCursor);
  if (cursor !== null && (!Number.isSafeInteger(cursor) || cursor < 1)) return Response.json({ error: "Invalid cursor" }, { status: 400, headers });
  const productId = url.searchParams.get("productId");
  try {
    const orders = await prisma.supplierOrder.findMany({
      where: { shop, status: "open", ...(productId ? { productId } : {}), ...(cursor !== null ? { id: { gt: cursor } } : {}) },
      orderBy: { id: "asc" },
      take: 51,
      select: { id: true, productId: true, productTitle: true, createdAt: true, updatedAt: true, fabricKey: true, fabricName: true, metersPerPiece: true, fabricConsumed: true, totalQty: true, eta: true, status: true, supplierStatus: true, destination: true, lines: { select: { variantId: true, variantTitle: true, qtyOrdered: true, qtyReceived: true } } },
    });
    const page = orders.slice(0, 50);
    return Response.json({ orders: page, nextCursor: orders.length > 50 ? page[49].id : null }, { headers });
  } catch {
    return Response.json({ error: "Unable to read orders" }, { status: 500, headers });
  }
};
