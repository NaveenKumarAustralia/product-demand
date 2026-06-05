import type { LoaderFunctionArgs } from "react-router";
import PDFDocument from "pdfkit";
import prisma from "../db.server";

// Landscape A4 PDF that mirrors the packing list view: a 4-column grid
// (Box / Name / Variants / Total). One row per packing list line.
// Auto-paginates when the rows overflow the page.

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

  // Convention: first line of each box has the box number; later lines
  // inherit it until a new box appears. Skip lines with zero qty.
  type Row = { box: string; name: string; variants: string; total: number };
  const rows: Row[] = [];
  let currentBox = "";
  for (const line of list.lines) {
    const explicit = (line.boxNumber ?? "").trim();
    if (explicit) currentBox = explicit;
    const qtys = normalizeQtys(line.qtys);
    const sizes = Object.entries(qtys).sort(([a], [b]) => a.localeCompare(b));
    const total = sizes.reduce((sum, [, q]) => sum + q, 0);
    if (total <= 0) continue;
    rows.push({
      box: currentBox || "—",
      name: line.productTitle || "Untitled product",
      variants: sizes.map(([size, qty]) => `${size}×${qty}`).join("  "),
      total,
    });
  }

  if (rows.length === 0) {
    throw new Response("Packing list has no boxed items to print", { status: 400 });
  }

  // Landscape A4 dimensions: 842pt × 595pt. Margins leave room for
  // header + footer.
  const doc = new PDFDocument({ size: "A4", layout: "landscape", margin: 30 });
  const chunks: Buffer[] = [];
  doc.on("data", (chunk) => chunks.push(Buffer.from(chunk)));
  const finished = new Promise<Buffer>((resolve, reject) => {
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);
  });

  const listLabel = (list.invoiceNumber || list.title || `Packing list #${list.id}`).trim();
  const pageWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;

  // Column layout — widths sum to pageWidth.
  const cols = [
    { key: "box",      label: "Box",      width: 80,  align: "center" as const },
    { key: "name",     label: "Name",     width: 240, align: "left" as const },
    { key: "variants", label: "Variants", width: pageWidth - 80 - 240 - 90, align: "left" as const },
    { key: "total",    label: "Total",    width: 90,  align: "center" as const },
  ];
  const rowPaddingY = 6;
  const rowPaddingX = 8;
  const baseRowHeight = 28;
  const headerRowHeight = 24;

  const drawHeaderRow = () => {
    let x = doc.page.margins.left;
    const y = doc.y;
    // Background.
    doc.save()
      .rect(x, y, pageWidth, headerRowHeight)
      .fill("#f1f5f9")
      .restore();
    doc.lineWidth(0.6).strokeColor("#94a3b8");
    for (const col of cols) {
      doc.rect(x, y, col.width, headerRowHeight).stroke();
      doc.fillColor("#111827").fontSize(11)
        .text(col.label.toUpperCase(), x + rowPaddingX, y + rowPaddingY + 2, {
          width: col.width - rowPaddingX * 2,
          align: col.align,
        });
      x += col.width;
    }
    doc.y = y + headerRowHeight;
    doc.fillColor("#111827");
  };

  const drawTopHeader = () => {
    // Reset cursor to the top margin, draw the shipment label, then
    // the table header.
    doc.y = doc.page.margins.top;
    doc.fontSize(16).fillColor("#111827").text(listLabel, doc.page.margins.left, doc.y, {
      width: pageWidth,
      align: "left",
    });
    doc.moveDown(0.5);
    drawHeaderRow();
  };

  drawTopHeader();

  // Estimate a single row's height based on the longest text content
  // wrapping. PDFKit can measure heights with .heightOfString.
  const measureRowHeight = (row: Row): number => {
    const heights = cols.map((col) => {
      const text = row[col.key as keyof Row] as string | number;
      const str = String(text);
      doc.fontSize(11);
      const h = doc.heightOfString(str, { width: col.width - rowPaddingX * 2 });
      return h;
    });
    return Math.max(baseRowHeight, Math.max(...heights) + rowPaddingY * 2);
  };

  const drawRow = (row: Row) => {
    const height = measureRowHeight(row);
    const bottomLimit = doc.page.height - doc.page.margins.bottom - 14; // 14pt reserved for footer
    if (doc.y + height > bottomLimit) {
      doc.addPage();
      drawTopHeader();
    }
    let x = doc.page.margins.left;
    const y = doc.y;
    doc.lineWidth(0.6).strokeColor("#cbd5e1");
    for (const col of cols) {
      doc.rect(x, y, col.width, height).stroke();
      const text = row[col.key as keyof Row] as string | number;
      doc.fillColor("#111827").fontSize(11)
        .text(String(text), x + rowPaddingX, y + rowPaddingY, {
          width: col.width - rowPaddingX * 2,
          align: col.align,
        });
      x += col.width;
    }
    doc.y = y + height;
  };

  for (const row of rows) {
    drawRow(row);
  }

  // Footer with totals + page counters. PDFKit's bufferPages would let
  // us paint footers retroactively, but it's not enabled by default —
  // for now we just write the totals once at the end.
  doc.moveDown(0.8);
  const totalQty = rows.reduce((sum, r) => sum + r.total, 0);
  const totalBoxes = new Set(rows.map((r) => r.box).filter((b) => b && b !== "—")).size;
  doc.fontSize(11).fillColor("#374151").text(
    `${totalBoxes} box${totalBoxes === 1 ? "" : "es"} · ${rows.length} line${rows.length === 1 ? "" : "s"} · ${totalQty} total pcs`,
    doc.page.margins.left,
    doc.y,
    { width: pageWidth, align: "right" },
  );

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
