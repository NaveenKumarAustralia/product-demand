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
      // Parse modelPicture entries so we can see whether each image is a stored
      // key (CollectionImage), a CDN/remote URL, an inline data URL, or empty.
      let modelPicture: unknown = row.modelPicture ?? "";
      try {
        const raw = String(row.modelPicture ?? "").trim();
        if (raw.startsWith("[")) {
          const arr = JSON.parse(raw) as unknown[];
          modelPicture = arr.map((x, idx) => {
            if (typeof x === "string") return { i: idx, type: x.startsWith("data:") ? "dataURL" : /^https?:/i.test(x) ? "url" : x ? "string" : "empty", src: x.slice(0, 90) };
            const o = (x && typeof x === "object") ? x as Record<string, unknown> : {};
            const thumb = String(o.thumb ?? ""); const key = String(o.key ?? "");
            return { i: idx, hasKey: !!key, key: key || undefined, thumbType: thumb.startsWith("data:") ? "dataURL" : /^https?:/i.test(thumb) ? "url" : thumb ? "string" : "empty", thumb: thumb.slice(0, 90) || undefined, alt: String(o.alt ?? "").slice(0, 40) || undefined, kind: o.kind };
          });
        }
      } catch { /* leave raw */ }
      matches.push({
        collection: c.name,
        collectionId: c.id,
        rowIndex: i,
        name,
        sizesOrdered: sizes,
        modelPicture,
        sku_raw: row.sku ?? "",
        barcode_raw: row.barcode ?? "",
        sku_visible: show(row.sku),         // newlines shown as ⏎
        barcode_visible: show(row.barcode),
        sku_lineCount: String(row.sku ?? "").split("\n").filter(Boolean).length,
        barcode_lineCount: String(row.barcode ?? "").split("\n").filter(Boolean).length,
        shopifyProductId: row.__shopifyProductId ?? "",
        shopifyProductIdNumeric: String(row.__shopifyProductId ?? "").replace(/\D/g, ""),
        locked: row.__shopifyLocked ?? "",
        editedFields: row.__shopifyEditedFields ?? "",
      });
    });
  }

  // Cross-reference: for each matched row's product, show the restock order(s) and
  // packing-list line(s) the collection page tries to match — and whether the
  // product ids line up (the cause of the missing status/shipment chip).
  const nums = Array.from(new Set((matches as Array<{ shopifyProductIdNumeric: string }>).map((m) => m.shopifyProductIdNumeric).filter(Boolean)));
  const [orders, lines] = nums.length ? await Promise.all([
    prisma.supplierOrder.findMany({ where: { productId: { not: "" } }, select: { id: true, productId: true, supplierStatus: true, status: true } }),
    prisma.packingListLine.findMany({ where: { isCustom: false, productId: { not: null }, packingList: { hiddenAt: null } }, select: { productId: true, qtys: true, packingList: { select: { id: true, invoiceNumber: true, title: true } } } }),
  ]) : [[], []];
  const numOf = (s: unknown) => String(s ?? "").replace(/\D/g, "");
  const cross = (matches as Array<{ shopifyProductIdNumeric: string; name: string }>).map((m) => ({
    name: m.name,
    rowProductIdNumeric: m.shopifyProductIdNumeric,
    restockOrders: orders.filter((o) => numOf(o.productId) === m.shopifyProductIdNumeric).map((o) => ({ rawProductId: o.productId, supplierStatus: o.supplierStatus, status: o.status })),
    packingLines: lines.filter((l) => numOf(l.productId) === m.shopifyProductIdNumeric).map((l) => ({ rawProductId: l.productId, ship: l.packingList.invoiceNumber || l.packingList.title || `#${l.packingList.id}`, qtys: l.qtys })),
  }));

  return Response.json(
    { ok: true, query: q, count: matches.length, matches, crossReference: cross, hint: "crossReference shows, per product, the restock orders + packing lines whose productId matches NUMERICALLY. If restockOrders/packingLines is empty, the product isn't in restock/packing (or the id differs); if non-empty but the chip is wrong, the match key format differs." },
    { headers: { "Cache-Control": "no-store" } },
  );
};
