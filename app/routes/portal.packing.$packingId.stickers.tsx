import type { LoaderFunctionArgs } from "react-router";
import PDFDocument from "pdfkit";
import prisma from "../db.server";

// Landscape A4 PDF, one sub-table per box. Each sub-table mirrors the
// packing list columns: Box | Name | Free Size | XS | S | M | L | XL |
// 2XL | 3XL | (any extras) | Total. The Box cell is drawn as a tall
// "merged" cell down the left of the sub-table. If a box has more
// lines than fit on one page, the sub-table splits across pages — the
// table header and the merged box cell are redrawn at the top of each
// new page so rows remain aligned. Row heights auto-grow to fit
// wrapped names so text never escapes its cell.

const BASELINE_SIZES = ["Free Size", "XS", "S", "M", "L", "XL", "2XL", "3XL", "S/M", "M/L", "L/XL"];

function canonicalSizeKey(value: string): string {
  return value.trim().toLowerCase().replace(/\s+/g, "").replace(/-/g, "/");
}

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
    include: { lines: { orderBy: [{ sortOrder: "asc" }, { id: "asc" }] } },
  });
  if (!list) throw new Response("Packing list not found", { status: 404 });

  // Master size-column list for the whole packing list.
  const baselineCanon = new Set(BASELINE_SIZES.map(canonicalSizeKey));
  const extraLabelByCanon = new Map<string, string>();
  for (const line of list.lines) {
    const qtys = normalizeQtys(line.qtys);
    for (const key of Object.keys(qtys)) {
      const trimmed = key.trim();
      if (!trimmed) continue;
      const canon = canonicalSizeKey(trimmed);
      if (baselineCanon.has(canon)) continue;
      if (!extraLabelByCanon.has(canon)) extraLabelByCanon.set(canon, trimmed);
    }
  }
  const sizeColumns = [...BASELINE_SIZES, ...extraLabelByCanon.values()];

  const qtyForSize = (qtys: Record<string, number>, label: string): number => {
    if (qtys[label]) return qtys[label];
    const canon = canonicalSizeKey(label);
    for (const [key, qty] of Object.entries(qtys)) {
      if (canonicalSizeKey(key) === canon) return qty;
    }
    return 0;
  };

  // Group lines by effective box number.
  type LineRow = { name: string; qtys: Record<string, number>; total: number };
  type Box = { boxNo: string; lines: LineRow[] };
  const boxes: Box[] = [];
  let currentBox = "";
  for (const line of list.lines) {
    const explicit = (line.boxNumber ?? "").trim();
    if (explicit) currentBox = explicit;
    const qtys = normalizeQtys(line.qtys);
    const total = Object.values(qtys).reduce((sum, q) => sum + q, 0);
    if (total <= 0) continue;
    const boxKey = currentBox || "—";
    const last = boxes[boxes.length - 1];
    if (last && last.boxNo === boxKey) {
      last.lines.push({ name: line.productTitle || "Untitled", qtys, total });
    } else {
      boxes.push({ boxNo: boxKey, lines: [{ name: line.productTitle || "Untitled", qtys, total }] });
    }
  }

  if (boxes.length === 0) {
    throw new Response("Packing list has no boxed items to print", { status: 400 });
  }

  // ── PDF setup ────────────────────────────────────────────────────
  const doc = new PDFDocument({ size: "A4", layout: "landscape", margin: 24 });
  const chunks: Buffer[] = [];
  doc.on("data", (chunk) => chunks.push(Buffer.from(chunk)));
  const finished = new Promise<Buffer>((resolve, reject) => {
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);
  });

  const listLabel = (list.invoiceNumber || list.title || `Packing list #${list.id}`).trim();
  const pageWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;

  const boxColWidth = 60;
  const totalColWidth = 60;
  const nameColWidth = 180;
  const sizeColsTotalWidth = pageWidth - boxColWidth - nameColWidth - totalColWidth;
  const sizeColWidth = Math.max(30, Math.floor(sizeColsTotalWidth / Math.max(1, sizeColumns.length)));

  const baseRowHeight = 22;
  const rowPadY = 6;
  const boxGap = 14;

  const HEADER_FILL = "#f0c8b8";
  const BORDER = "#94a3b8";

  // Header height auto-fits the tallest wrapping label (e.g. "Free
  // Size" wraps to two lines in narrow size columns).
  const headerRowHeight = (() => {
    doc.fontSize(10);
    let max = 22;
    const labels: Array<[string, number]> = [
      ["Box", boxColWidth],
      ["Name", nameColWidth],
      ...sizeColumns.map((label) => [label, sizeColWidth] as [string, number]),
      ["Total", totalColWidth],
    ];
    for (const [label, width] of labels) {
      const h = doc.heightOfString(label, { width: width - 4 });
      max = Math.max(max, Math.ceil(h) + rowPadY * 2);
    }
    return max;
  })();

  // Measure how tall a single data row needs to be so wrapped name text
  // fits inside its cell. PDFKit's heightOfString gives the height the
  // text would occupy at the given width.
  const measureRowHeight = (line: LineRow): number => {
    doc.fontSize(11);
    const nameHeight = doc.heightOfString(line.name, { width: nameColWidth - 12 });
    return Math.max(baseRowHeight, Math.ceil(nameHeight) + rowPadY * 2);
  };

  const drawPageHeader = () => {
    doc.y = doc.page.margins.top;
    doc.fontSize(13).fillColor("#111827")
      .text(listLabel, doc.page.margins.left, doc.y, { width: pageWidth, align: "left" });
    doc.moveDown(0.4);
  };

  // Draw one chunk of a box's data rows at the current y. Chunk may be
  // the full box or just the rows that fit on the remainder of the
  // page. The table header is always drawn above the chunk and the box
  // cell is drawn as a merged cell spanning the chunk's data height.
  const drawBoxChunk = (boxNo: string, chunkLines: LineRow[], chunkHeights: number[]) => {
    const x0 = doc.page.margins.left;
    const y = doc.y;
    const dataHeight = chunkHeights.reduce((sum, h) => sum + h, 0);

    // Header row.
    doc.save().rect(x0, y, pageWidth, headerRowHeight).fill(HEADER_FILL).restore();
    doc.lineWidth(0.6).strokeColor(BORDER);
    let hx = x0;
    const cells: Array<{ label: string; width: number; align: "center" }> = [
      { label: "Box", width: boxColWidth, align: "center" },
      { label: "Name", width: nameColWidth, align: "center" },
      ...sizeColumns.map((label) => ({ label, width: sizeColWidth, align: "center" as const })),
      { label: "Total", width: totalColWidth, align: "center" },
    ];
    for (const cell of cells) {
      doc.rect(hx, y, cell.width, headerRowHeight).stroke();
      doc.fillColor("#111827").fontSize(10)
        .text(cell.label, hx + 2, y + 6, { width: cell.width - 4, align: cell.align });
      hx += cell.width;
    }

    // Box cell merged across this chunk's data rows.
    const dataY = y + headerRowHeight;
    doc.save().rect(x0, dataY, boxColWidth, dataHeight).fill(HEADER_FILL).restore();
    doc.rect(x0, dataY, boxColWidth, dataHeight).stroke();
    doc.fillColor("#111827").fontSize(20)
      .text(boxNo, x0, dataY + (dataHeight / 2) - 12, {
        width: boxColWidth,
        align: "center",
      });

    // Data rows with variable heights.
    let ry = dataY;
    chunkLines.forEach((line, idx) => {
      const rowH = chunkHeights[idx];
      let cx = x0 + boxColWidth;
      // Name cell — wraps within its width; row height grew to fit.
      doc.rect(cx, ry, nameColWidth, rowH).stroke();
      doc.fillColor("#111827").fontSize(11)
        .text(line.name, cx + 6, ry + rowPadY, { width: nameColWidth - 12, align: "left" });
      cx += nameColWidth;
      // Size cells.
      for (const size of sizeColumns) {
        doc.rect(cx, ry, sizeColWidth, rowH).stroke();
        const qty = qtyForSize(line.qtys, size);
        if (qty > 0) {
          doc.fillColor("#111827").fontSize(11)
            .text(String(qty), cx, ry + rowPadY, { width: sizeColWidth, align: "center" });
        }
        cx += sizeColWidth;
      }
      // Total cell.
      doc.rect(cx, ry, totalColWidth, rowH).stroke();
      doc.fillColor("#111827").fontSize(11)
        .text(String(line.total), cx, ry + rowPadY, { width: totalColWidth, align: "center" });
      ry += rowH;
    });

    doc.y = dataY + dataHeight;
  };

  drawPageHeader();

  for (const box of boxes) {
    // Pre-measure every row for this box.
    const allHeights = box.lines.map(measureRowHeight);

    let cursor = 0;
    while (cursor < box.lines.length) {
      const bottomLimit = doc.page.height - doc.page.margins.bottom;
      // Need room for header + at least the next row.
      const minNeeded = headerRowHeight + allHeights[cursor];
      if (doc.y + minNeeded > bottomLimit) {
        doc.addPage();
        drawPageHeader();
      }
      // Fit as many rows as possible on this page.
      let used = doc.y + headerRowHeight;
      let endIdx = cursor;
      while (endIdx < box.lines.length) {
        const h = allHeights[endIdx];
        if (used + h > bottomLimit) break;
        used += h;
        endIdx++;
      }
      // Always advance at least one row (single-row taller than a page
      // would otherwise loop forever — clamp to one row).
      if (endIdx === cursor) endIdx = cursor + 1;

      const chunkLines = box.lines.slice(cursor, endIdx);
      const chunkHeights = allHeights.slice(cursor, endIdx);
      drawBoxChunk(box.boxNo, chunkLines, chunkHeights);
      cursor = endIdx;

      // If more rows remain for this box, start a new page so the
      // remaining rows get a fresh header + merged box cell.
      if (cursor < box.lines.length) {
        doc.addPage();
        drawPageHeader();
      }
    }
    doc.y += boxGap;
  }

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
