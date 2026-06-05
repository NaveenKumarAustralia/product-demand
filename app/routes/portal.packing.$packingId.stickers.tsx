import type { LoaderFunctionArgs } from "react-router";
import PDFDocument from "pdfkit";
import prisma from "../db.server";

// Server-generated A4 PDF, one page per box. Each page shows the box
// number, then for each product in the box: the product title, a
// size × qty breakdown, and the total qty in that box for that product.
// User prints + sticks one page on top of each physical box so the
// contents are visible at a glance.

function normalizeQtys(value: unknown): Record<string, number> {
  if (!value || typeof value !== "object") return {};
  const out: Record<string, number> = {};
  for (const [k, v] of Object.entries(value as Record<string, unknown>)) {
    const num = Number(v);
    if (Number.isFinite(num) && num > 0) out[k] = num;
  }
  return out;
}

export const loader = async ({ params }: LoaderFunctionArgs) => {
  const packingId = Number(params.packingId);
  if (!Number.isFinite(packingId) || packingId <= 0) {
    throw new Response("Invalid packing list id", { status: 400 });
  }

  const list = await prisma.packingList.findUnique({
    where: { id: packingId },
    include: {
      lines: { orderBy: [{ sortOrder: "asc" }, { id: "asc" }] },
    },
  });
  if (!list) throw new Response("Packing list not found", { status: 404 });

  // Group lines by their effective box number. Convention in this codebase:
  // the first line of a box has the box number written in; subsequent
  // lines inherit until a new box number appears.
  type StickerLine = { productTitle: string; sizes: Array<[string, number]>; total: number };
  const boxes = new Map<string, StickerLine[]>();
  let currentBox = "";
  for (const line of list.lines) {
    const explicit = (line.boxNumber ?? "").trim();
    if (explicit) currentBox = explicit;
    const boxKey = currentBox || "Unboxed";
    const qtys = normalizeQtys(line.qtys);
    const sizes = Object.entries(qtys).sort(([a], [b]) => a.localeCompare(b));
    const total = sizes.reduce((sum, [, q]) => sum + q, 0);
    if (total <= 0) continue;
    const entry: StickerLine = {
      productTitle: line.productTitle || "Untitled product",
      sizes,
      total,
    };
    const bucket = boxes.get(boxKey);
    if (bucket) bucket.push(entry); else boxes.set(boxKey, [entry]);
  }

  if (boxes.size === 0) {
    throw new Response("Packing list has no boxed items to print", { status: 400 });
  }

  // Sort box keys numerically when possible so 1, 2, 10 read correctly.
  const boxKeys = Array.from(boxes.keys()).sort((a, b) => {
    const an = Number(a);
    const bn = Number(b);
    if (Number.isFinite(an) && Number.isFinite(bn)) return an - bn;
    return a.localeCompare(b);
  });

  const doc = new PDFDocument({ size: "A4", margin: 40 });
  const chunks: Buffer[] = [];
  doc.on("data", (chunk) => chunks.push(Buffer.from(chunk)));
  const finished = new Promise<Buffer>((resolve, reject) => {
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);
  });

  const listLabel = (list.invoiceNumber || list.title || `Packing list #${list.id}`).trim();

  boxKeys.forEach((boxKey, idx) => {
    if (idx > 0) doc.addPage();

    // Header strip with the shipment / packing list label and box number.
    doc.fontSize(16).fillColor("#374151")
      .text(listLabel, { align: "left" });
    doc.moveDown(0.3);
    doc.fontSize(72).fillColor("#111827")
      .text(`BOX ${boxKey}`, { align: "center" });
    doc.moveDown(0.5);

    // Divider.
    const pageWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
    doc.moveTo(doc.page.margins.left, doc.y).lineTo(doc.page.margins.left + pageWidth, doc.y)
      .lineWidth(2).strokeColor("#111827").stroke();
    doc.moveDown(0.8);

    // For each product in the box: title, sizes, total.
    const items = boxes.get(boxKey) ?? [];
    items.forEach((item, itemIdx) => {
      if (itemIdx > 0) doc.moveDown(0.8);
      doc.fontSize(28).fillColor("#111827").text(item.productTitle, { align: "left" });
      doc.moveDown(0.3);
      if (item.sizes.length) {
        const sizesLine = item.sizes
          .map(([size, qty]) => `${size}×${qty}`)
          .join("    ");
        doc.fontSize(22).fillColor("#374151").text(sizesLine, { align: "left" });
        doc.moveDown(0.2);
      }
      doc.fontSize(26).fillColor("#0f766e").text(`Total: ${item.total} pcs`, { align: "left" });
    });

    // Footer with box index for cross-reference.
    const footerY = doc.page.height - doc.page.margins.bottom - 16;
    doc.fontSize(10).fillColor("#9ca3af")
      .text(`Sticker ${idx + 1} of ${boxKeys.length}`, doc.page.margins.left, footerY, { align: "right" });
  });

  doc.end();
  const pdfBuffer = await finished;
  const safeName = (list.invoiceNumber || list.title || `packing-${list.id}`)
    .replace(/[^a-z0-9-]+/gi, "-")
    .replace(/^-+|-+$/g, "")
    .toLowerCase() || `packing-${list.id}`;
  return new Response(pdfBuffer, {
    headers: {
      "Content-Type": "application/pdf",
      "Content-Disposition": `attachment; filename="stickers-${safeName}.pdf"`,
      "Content-Length": String(pdfBuffer.length),
    },
  });
};
