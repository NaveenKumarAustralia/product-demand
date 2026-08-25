import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// TEMP diagnostic — what destination values exist on open orders, and which
// (if any) look like USA. Remove once the USA page filter is confirmed.
export const loader = async (_args: LoaderFunctionArgs) => {
  const orders = await prisma.supplierOrder.findMany({
    where: { status: "open" },
    select: { id: true, productTitle: true, destination: true, supplier: true, totalQty: true, supplierStatus: true },
  });
  const byDestination: Record<string, number> = {};
  for (const o of orders) {
    const key = JSON.stringify(o.destination);
    byDestination[key] = (byDestination[key] ?? 0) + 1;
  }
  const usaish = orders
    .filter((o) => (o.destination ?? "").toLowerCase().includes("usa") || (o.destination ?? "").toLowerCase().includes("us"))
    .map((o) => ({ id: o.id, title: o.productTitle, destination: o.destination, supplier: o.supplier, totalQty: o.totalQty, supplierStatus: o.supplierStatus }));
  return Response.json({ openOrderCount: orders.length, byDestination, usaish });
};
