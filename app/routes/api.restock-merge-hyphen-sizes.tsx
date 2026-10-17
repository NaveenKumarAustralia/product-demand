import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

// Admin one-time migration: on the restock page the hyphen size variants
// (S-M / M-L / L-XL) are the same as the slash variants (S/M / M/L / L/XL).
// This folds every OrderLine stored under a hyphen label into the matching
// slash label — renaming the line when no slash sibling exists, or summing its
// qty into the existing slash line and deleting the hyphen line when one does.
// Order total quantities are unchanged (qty is preserved, only re-labelled).
//
//   GET /api/restock-merge-hyphen-sizes            → dry run (shows what would change)
//   GET /api/restock-merge-hyphen-sizes?apply=1    → applies the change
const HYPHEN_TO_SLASH: Record<string, string> = { "S-M": "S/M", "M-L": "M/L", "L-XL": "L/XL" };
const HYPHEN_LABELS = Object.keys(HYPHEN_TO_SLASH).flatMap((h) => [h, h.toLowerCase()]);
const toSlash = (s: string) => HYPHEN_TO_SLASH[String(s ?? "").trim().toUpperCase()] ?? String(s ?? "").trim();

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });
  const apply = new URL(request.url).searchParams.get("apply") === "1";

  // All lines stored under a hyphen label.
  const hyphenLines = await prisma.orderLine.findMany({
    where: { variantTitle: { in: HYPHEN_LABELS } },
    select: { id: true, orderId: true, variantTitle: true, qtyOrdered: true },
    orderBy: { id: "asc" },
  });
  if (!hyphenLines.length) {
    return Response.json({ ok: true, apply, message: "Nothing to migrate — no hyphen size lines found.", renamed: 0, merged: 0, ordersAffected: 0 }, { headers: { "Cache-Control": "no-store" } });
  }

  // The slash lines that hyphen lines may merge into (same orders, same slash labels).
  const orderIds = [...new Set(hyphenLines.map((l) => l.orderId))];
  const slashWanted = [...new Set(hyphenLines.map((l) => toSlash(l.variantTitle)))];
  const slashLines = await prisma.orderLine.findMany({
    where: { orderId: { in: orderIds }, variantTitle: { in: slashWanted } },
    select: { id: true, orderId: true, variantTitle: true, qtyOrdered: true },
    orderBy: { id: "asc" },
  });
  const slashByKey = new Map<string, { id: number; qtyOrdered: number }>();
  for (const s of slashLines) {
    const key = `${s.orderId}::${s.variantTitle}`;
    if (!slashByKey.has(key)) slashByKey.set(key, { id: s.id, qtyOrdered: s.qtyOrdered });
  }

  const renames: Array<{ lineId: number; orderId: number; from: string; to: string; qty: number }> = [];
  const merges: Array<{ hyphenLineId: number; intoLineId: number; orderId: number; from: string; to: string; qty: number }> = [];
  // Track slash lines we create via rename so later hyphen lines merge into them.
  const createdSlash = new Map<string, number>(); // key -> lineId

  for (const h of hyphenLines) {
    const slash = toSlash(h.variantTitle);
    const key = `${h.orderId}::${slash}`;
    const existing = slashByKey.get(key) ?? (createdSlash.has(key) ? { id: createdSlash.get(key)!, qtyOrdered: 0 } : null);
    if (existing) {
      merges.push({ hyphenLineId: h.id, intoLineId: existing.id, orderId: h.orderId, from: h.variantTitle, to: slash, qty: h.qtyOrdered });
    } else {
      renames.push({ lineId: h.id, orderId: h.orderId, from: h.variantTitle, to: slash, qty: h.qtyOrdered });
      createdSlash.set(key, h.id);
    }
  }

  if (apply) {
    await prisma.$transaction(async (tx) => {
      // Renames first, so the slash target exists for any merges into it.
      for (const r of renames) {
        await tx.orderLine.update({ where: { id: r.lineId }, data: { variantTitle: r.to } });
      }
      // Merges: add qty into the slash line, delete the hyphen line.
      const addByLine = new Map<number, number>();
      for (const m of merges) addByLine.set(m.intoLineId, (addByLine.get(m.intoLineId) ?? 0) + m.qty);
      for (const [intoLineId, add] of addByLine) {
        await tx.orderLine.update({ where: { id: intoLineId }, data: { qtyOrdered: { increment: add } } });
      }
      await tx.orderLine.deleteMany({ where: { id: { in: merges.map((m) => m.hyphenLineId) } } });
    });
  }

  return Response.json(
    {
      ok: true,
      apply,
      ordersAffected: orderIds.length,
      renamed: renames.length,
      merged: merges.length,
      renames,
      merges,
      hint: apply ? "Applied. Hyphen size lines have been folded into the slash labels." : "Dry run only — add ?apply=1 to apply.",
    },
    { headers: { "Cache-Control": "no-store" } },
  );
};
