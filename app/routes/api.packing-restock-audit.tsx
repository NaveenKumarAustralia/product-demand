import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

// Admin READ-ONLY audit. Between 2026-10-20 and 2026-10-23 the restock page's
// "in shipment" linking rebuilt a product's packing-list line from the restock
// ORDER quantities, which could have OVERWRITTEN quantities staff entered by
// hand. This lists, per non-hidden packing list, every line whose product also
// has a restock order soft-linked to that same list (packingListId) — the lines
// most likely to have been touched — and flags any whose current quantity
// exactly equals the linked orders' total (the signature of an overwrite that
// hasn't been corrected yet). It changes NOTHING; it's for staff to re-verify.
//   GET /api/packing-restock-audit
const sumQtys = (q: unknown): number => {
  if (!q || typeof q !== "object") return 0;
  return Object.values(q as Record<string, unknown>).reduce<number>((s, n) => s + (Number(n) || 0), 0);
};

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const lists = await prisma.packingList.findMany({
    where: { hiddenAt: null },
    select: {
      id: true, invoiceNumber: true, title: true,
      lines: { where: { isCustom: false, productId: { not: null } }, select: { id: true, productId: true, productTitle: true, qtys: true } },
    },
    orderBy: { id: "desc" },
  });

  // Soft-linked restock orders per (packingListId → productId).
  const linkedOrders = await prisma.supplierOrder.findMany({
    where: { packingListId: { not: null }, productId: { not: "" } },
    select: { packingListId: true, productId: true, lines: { select: { variantTitle: true, qtyOrdered: true } } },
  });
  const numOf = (s: unknown) => String(s ?? "").replace(/\D/g, "");
  const orderSumByKey = new Map<string, number>(); // `${listId}::${numericProductId}` → total ordered
  for (const o of linkedOrders) {
    if (o.packingListId == null) continue;
    const key = `${o.packingListId}::${numOf(o.productId)}`;
    const t = o.lines.reduce((s, l) => s + (l.qtyOrdered || 0), 0);
    orderSumByKey.set(key, (orderSumByKey.get(key) ?? 0) + t);
  }

  const flagged: unknown[] = [];
  for (const list of lists) {
    const shipLabel = list.invoiceNumber || list.title || `#${list.id}`;
    for (const line of list.lines) {
      const key = `${list.id}::${numOf(line.productId)}`;
      const orderSum = orderSumByKey.get(key);
      if (orderSum == null) continue; // no restock order soft-linked here → not touched via restock
      const lineTotal = sumQtys(line.qtys);
      flagged.push({
        shipment: shipLabel,
        packingListId: list.id,
        product: line.productTitle,
        productId: line.productId,
        lineTotalQty: lineTotal,
        linkedOrderTotalQty: orderSum,
        matchesOrderSum: lineTotal === orderSum && orderSum > 0,
        sizes: line.qtys,
      });
    }
  }

  // Most-likely-overwritten first.
  flagged.sort((a: any, b: any) => Number(b.matchesOrderSum) - Number(a.matchesOrderSum));
  return Response.json(
    {
      ok: true,
      note: "Read-only. Lines here have a restock order linked to the same packing list, so the old restock logic could have touched them. matchesOrderSum:true = the line quantity equals the linked orders' total, i.e. likely an overwrite staff hasn't re-corrected. Re-verify these against your records; restock no longer writes to packing lists.",
      flaggedCount: flagged.length,
      likelyOverwritten: flagged.filter((f: any) => f.matchesOrderSum).length,
      flagged,
    },
    { headers: { "Cache-Control": "no-store" } },
  );
};
