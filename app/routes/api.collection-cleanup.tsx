import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// TEMP one-off: remove EMPTY rows from a collection (no content, no Shopify
// product). Safe — never touches a row that has any value or a linked product.
//   /api/collection-cleanup?id=56          → dry run (reports counts only)
//   /api/collection-cleanup?id=56&apply=1  → actually delete + save
// Remove this route after use.

type Row = Record<string, unknown>;

function isEmptyRow(row: Row): boolean {
  if (String(row.__shopifyProductId ?? "").trim()) return false;
  return Object.entries(row).every(([k, v]) => k.startsWith("__") || !String(v ?? "").trim());
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const url = new URL(request.url);
  const id = Number(url.searchParams.get("id"));
  const apply = url.searchParams.get("apply") === "1";
  if (!id) return Response.json({ ok: false, error: "pass ?id=<collectionId>" });

  const collection = await prisma.collection.findUnique({ where: { id }, select: { id: true, name: true, rows: true } }).catch(() => null);
  if (!collection) return Response.json({ ok: false, error: "not_found" });

  const rows: Row[] = Array.isArray(collection.rows) ? collection.rows as Row[] : [];
  const kept = rows.filter((r) => !isEmptyRow(r));
  const removed = rows.length - kept.length;

  if (apply && removed > 0) {
    await prisma.collection.update({ where: { id }, data: { rows: kept, updatedAt: new Date() } });
  }

  return Response.json({
    ok: true,
    id: collection.id,
    name: collection.name,
    before: rows.length,
    kept: kept.length,
    removed,
    applied: apply && removed > 0,
  });
};
