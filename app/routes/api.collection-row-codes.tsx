import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

// Admin diagnostic: dump the RAW SKU + barcode cells (newlines made visible) for
// any collection row whose Name matches, plus the row's size quantities and
// Shopify link state. Lets us compare a correctly-formatted product (e.g. the
// Paige dress) against one whose SKU/barcode format changed after a Shopify
// update — without guessing at the data.
//   GET /api/collection-row-codes?q=paige
const SIZE_IDS = ["freeSize", "xs", "s", "m", "l", "xl", "xxl", "xxxl", "sm", "ml", "lxl"];

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const q = (new URL(request.url).searchParams.get("q") || "").trim().toLowerCase();
  if (!q) return Response.json({ ok: false, error: "Pass ?q=<part of the product name>." }, { status: 400 });

  const collections = await prisma.collection.findMany({ select: { id: true, name: true, rows: true } });
  const show = (s: unknown) => String(s ?? "").replace(/\n/g, "⏎");
  const matches: unknown[] = [];
  for (const c of collections) {
    const rows = Array.isArray(c.rows) ? (c.rows as Array<Record<string, string>>) : [];
    rows.forEach((row, i) => {
      const name = String(row.name ?? row.title ?? "");
      if (!name.toLowerCase().includes(q)) return;
      const sizes: Record<string, string> = {};
      for (const id of SIZE_IDS) if ((Number(row[id]) || 0) > 0) sizes[id] = row[id];
      matches.push({
        collection: c.name,
        collectionId: c.id,
        rowIndex: i,
        name,
        sizesOrdered: sizes,
        sku_raw: row.sku ?? "",
        barcode_raw: row.barcode ?? "",
        sku_visible: show(row.sku),         // newlines shown as ⏎
        barcode_visible: show(row.barcode),
        sku_lineCount: String(row.sku ?? "").split("\n").filter(Boolean).length,
        barcode_lineCount: String(row.barcode ?? "").split("\n").filter(Boolean).length,
        shopifyProductId: row.__shopifyProductId ?? "",
        locked: row.__shopifyLocked ?? "",
        editedFields: row.__shopifyEditedFields ?? "",
      });
    });
  }
  return Response.json(
    { ok: true, query: q, count: matches.length, matches, hint: "Paste this back — compare the GOOD row (Paige) vs the bad one: sku_visible / barcode_visible (⏎ = newline) and line counts show the format difference." },
    { headers: { "Cache-Control": "no-store" } },
  );
};
