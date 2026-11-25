import type { ActionFunctionArgs } from "react-router";
import { createHash, timingSafeEqual } from "node:crypto";
import prisma from "../db.server";

const headers = { "Cache-Control": "no-store" };
const reply = (body: object, status = 200) => Response.json(body, { status, headers });
class Conflict extends Error {}
export const action = async ({ request }: ActionFunctionArgs) => {
  if (request.method !== "POST") return reply({ error: "Method not allowed" }, 405);
  const key = process.env.AI_PORTAL_WRITE_TOKEN;
  const shop = process.env.AI_PORTAL_SHOP;
  const auth = request.headers.get("Authorization");
  const hash = (s: string) => createHash("sha256").update(s).digest();
  if (!key || key.length < 32 || key === process.env.AI_PORTAL_READ_TOKEN || !shop || !auth?.startsWith("Bearer ") || !timingSafeEqual(hash(key), hash(auth.slice(7)))) return reply({ error: "Unauthorized" }, 401);
  let b;
  try { b = await request.json(); } catch { return reply({ error: "Invalid JSON" }, 400); }
  if (!b || typeof b !== "object" || !Number.isSafeInteger(b.orderId) || b.orderId < 1 || typeof b.variantId !== "string" || !b.variantId || typeof b.productId !== "string" || !b.productId || !Number.isSafeInteger(b.expectedTotal) || b.expectedTotal < 0 || !Number.isSafeInteger(b.expectedQty) || b.expectedQty < 0 || !Number.isSafeInteger(b.delta) || b.delta < 1 || b.delta > 10000 || !Number.isSafeInteger(b.expectedTotal + b.delta) || typeof b.expectedUpdatedAt !== "string" || !Number.isFinite(Date.parse(b.expectedUpdatedAt))) return reply({ error: "Invalid adjustment" }, 400);
  // Owner approval must be enforced by the eventual authenticated connector;
  // leave this experimental endpoint disabled until approval/audit integration.
  if (process.env.AI_PORTAL_ENABLE_TEST_WRITES !== "true" || process.env.RAILWAY_ENVIRONMENT_NAME !== "mcp-test") return reply({ error: "Test writes disabled" }, 403);
  try {
    const result = await prisma.$transaction(async tx => {
      const order = await tx.supplierOrder.findFirst({ where: { id: b.orderId, shop, productId: b.productId, status: "open", supplierStatus: "on_order", fabricConsumed: false }, include: { lines: true } });
      if (!order || order.totalQty !== b.expectedTotal || order.updatedAt.toISOString() !== new Date(b.expectedUpdatedAt).toISOString()) throw new Conflict("Order changed or is not eligible");
      const lines = order.lines.filter(l => l.variantId === b.variantId);
      if (lines.length !== 1 || lines[0].qtyOrdered !== b.expectedQty || lines[0].qtyReceived !== 0 || order.lines.reduce((n,l)=>n+l.qtyOrdered,0) !== order.totalQty) throw new Conflict("Size or total mismatch");
      const line = lines[0];
      const changed = await tx.supplierOrder.updateMany({ where: { id: order.id, shop, updatedAt: order.updatedAt, totalQty: b.expectedTotal, status: "open", supplierStatus: "on_order", fabricConsumed: false }, data: { totalQty: { increment: b.delta } } });
      if (changed.count !== 1) throw new Conflict("Order changed");
      const changedLine = await tx.orderLine.updateMany({ where: { id: line.id, qtyOrdered: b.expectedQty, qtyReceived: 0 }, data: { qtyOrdered: { increment: b.delta } } });
      if (changedLine.count !== 1) throw new Conflict("Size changed");
      const after = await tx.supplierOrder.findUniqueOrThrow({ where: { id: order.id }, include: { lines: true } });
      return { before: { total: order.totalQty, sizeQty: line.qtyOrdered }, after: { total: after.totalQty, sizeQty: after.lines.find(l=>l.id===line.id)?.qtyOrdered }, orderId: order.id };
    }, { isolationLevel: "Serializable" });
    return reply(result);
  } catch (err) {
    if (err instanceof Conflict) return reply({ error: err.message }, 409);
    return reply({ error: "Adjustment not confirmed; read order before retrying" }, 500);
  }
};
