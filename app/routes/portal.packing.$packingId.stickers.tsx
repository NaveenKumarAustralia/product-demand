import type { LoaderFunctionArgs } from "react-router";
import PDFDocument from "pdfkit";
import prisma from "../db.server";

// Landscape A4 PDF, one sub-table per box. Each sub-table mirrors the
// packing list columns: Box | Name | XS | S | M | L | XL | 2XL | 3XL |
// (any extras) | Total. The Box cell is drawn as a tall "merged" cell
// down the left of the sub-table. User prints + cuts between boxes to
// make a sticker per box.

const BASELINE_SIZES = ["XS", "S", "M", "L", "XL", "2XL", "3XL", "S/M", "M/L", "L/XL"];

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

  // Build the master size-column list for the whole packing list: every
  // baseline column (so users have a familiar layout) plus any extra
  // canonical key that appears in the data with qty > 0.
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

  // Resolve a line's qty for a given column label using canonical
  // matching so "S-M" data fills the "S/M" column.
  const qtyForSize = (qtys: Record<string, number>, label: string): number => {
    if (qtys[label]) return qtys[label];
    const canon = canonicalSizeKey(label);
    for (const [key, qty] of Object.entries(qtys)) {
      if (canonicalSizeKey(key) === canon) return qty;
    }
    return 0;
  };

  // Group lines by effective box number (first line of a box has the
  // number; later lines inherit until a new box appears). Skip lines
  // with zero total qty.
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

  // Column widths. Box and Total are fixed; Name takes a generous chunk;
  // the remaining width is split equally between size columns.
  const boxColWidth = 60;
  const totalColWidth = 60;
  const nameColWidth = 180;
  const sizeColsTotalWidth = pageWidth - boxColWidth - nameColWidth - totalColWidth;
  const sizeColWidth = Math.max(34, Math.floor(sizeColsTotalWidth / Math.max(1, sizeColumns.length)));

  const dataRowHeight = 22;
  const headerRowHeight = 22;
  const boxGap = 14;

  const HEADER_FILL = "#f0c8b8";
  const BORDER = "#94a3b8";

  // Header label for the page (shipment).
  const drawPageHeader = () => {
    doc.y = doc.page.margins.top;
    doc.fontSize(13).fillColor("#111827")
      .text(listLabel, doc.page.margins.left, doc.y, { width: pageWidth, align: "left" });
    doc.moveDown(0.4);
  };

  // Draw one box's sub-table at the current y. Returns the ending y.
  const drawBox = (box: Box) => {
    const x0 = doc.page.margins.left;
    let y = doc.y;
    const dataHeight = box.lines.length * dataRowHeight;
    const tableHeight = headerRowHeight + dataHeight;

    // ── HEADER ROW ───────────────────────────────────────────────
    doc.save().rect(x0, y, pageWidth, headerRowHeight).fill(HEADER_FILL).restore();
    doc.lineWidth(0.6).strokeColor(BORDER);
    let hx = x0;
    const cells: Array<{ label: string; width: number; align: "left" | "center" }> = [
      { label: "Box", width: boxColWidth, align: "center" },
      { label: "Name", width: nameColWidth, align: "center" },
      ...sizeColumns.map((label) => ({ label, width: sizeColWidth, align: "center" as const })),
      { label: "Total", width: totalColWidth, align: "center" },
    ];
    for (const cell of cells) {
      doc.rect(hx, y, cell.width, headerRowHeight).stroke();
      doc.fillColor("#111827").fontSize(11)
        .text(cell.label, hx + 2, y + 6, { width: cell.width - 4, align: cell.align });
      hx += cell.width;
    }

    // ── BOX CELL (merged across the data rows) ───────────────────
    const dataY = y + headerRowHeight;
    doc.save().rect(x0, dataY, boxColWidth, dataHeight).fill(HEADER_FILL).restore();
    doc.rect(x0, dataY, boxColWidth, dataHeight).stroke();
    doc.fillColor("#111827").fontSize(20)
      .text(box.boxNo, x0, dataY + (dataHeight / 2) - 12, {
        width: boxColWidth,
        align: "center",
      });

    // ── DATA ROWS ────────────────────────────────────────────────
    box.lines.forEach((line, lineIdx) => {
      const ry = dataY + (lineIdx * dataRowHeight);
      let cx = x0 + boxColWidth;
      // Name cell
      doc.rect(cx, ry, nameColWidth, dataRowHeight).stroke();
      doc.fillColor("#111827").fontSize(11)
        .text(line.name, cx + 6, ry + 6, { width: nameColWidth - 12, align: "left", lineBreak: false, ellipsis: true });
      cx += nameColWidth;
      // Size cells
      for (const size of sizeColumns) {
        doc.rect(cx, ry, sizeColWidth, dataRowHeight).stroke();
        const qty = qtyForSize(line.qtys, size);
        if (qty > 0) {
          doc.fillColor("#111827").fontSize(11)
            .text(String(qty), cx, ry + 6, { width: sizeColWidth, align: "center" });
        }
        cx += sizeColWidth;
      }
      // Total cell
      doc.rect(cx, ry, totalColWidth, dataRowHeight).stroke();
      doc.fillColor("#111827").fontSize(11)
        .text(String(line.total), cx, ry + 6, { width: totalColWidth, align: "center" });
    });

    doc.y = dataY + dataHeight;
    return tableHeight;
  };

  drawPageHeader();

  for (const box of boxes) {
    const tableHeight = headerRowHeight + (box.lines.length * dataRowHeight);
    const bottomLimit = doc.page.height - doc.page.margins.bottom;
    if (doc.y + tableHeight > bottomLimit) {
      doc.addPage();
      drawPageHeader();
    }
    drawBox(box);
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
