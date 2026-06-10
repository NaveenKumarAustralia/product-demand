import type { LoaderFunctionArgs } from "react-router";
import PDFDocument from "pdfkit";
import prisma from "../db.server";

// One A4-landscape shipping label per box. Top: invoice number + "BOX
// N/total". Body: To address (Karma East) on the left and From address
// (Aisling) on the right, then delivery instructions across the
// bottom. User prints + sticks one on top of each physical box.
//
// Box numbers come from packingList.lines — distinct box numbers are
// collected in numeric order; total is the count of distinct boxes.
// So a list with boxes 1-116 produces 116 labels (1/116 … 116/116).

const TO_ADDRESS_LINES = [
  "KARMA EAST PTY LTD",
  "NAVEEN KUMAR",
  "0430574666",
  "6 WILLIAM STREET",
  "BEVERLEY, SA 5009",
  "AUSTRALIA",
];
const FROM_ADDRESS_LINES = [
  "AISLING ENTERPRISES PVT. LTD.",
  "PUSHKAR — INDIA.",
];
const DELIVERY_INSTRUCTIONS = "Delivery Instructions: To be delivered on plain pallets only";

export const loader = async ({ params }: LoaderFunctionArgs) => {
  const packingId = Number(params.packingId);
  if (!Number.isFinite(packingId) || packingId <= 0) {
    throw new Response("Invalid packing list id", { status: 400 });
  }

  const list = await prisma.packingList.findUnique({
    where: { id: packingId },
    include: {
      lines: { orderBy: [{ sortOrder: "asc" }, { id: "asc" }], select: { boxNumber: true } },
    },
  });
  if (!list) throw new Response("Packing list not found", { status: 404 });

  // Walk the lines in display order, propagating box numbers the way
  // the packing list does (first line of a box has the number, later
  // lines inherit until a new number appears). Collect distinct boxes.
  const boxNumbers: string[] = [];
  const seen = new Set<string>();
  let currentBox = "";
  for (const line of list.lines) {
    const explicit = (line.boxNumber ?? "").trim();
    if (explicit) currentBox = explicit;
    if (!currentBox) continue;
    if (!seen.has(currentBox)) {
      seen.add(currentBox);
      boxNumbers.push(currentBox);
    }
  }

  // Sort numerically when all are numeric, else alphabetically — same
  // convention as the box stickers route.
  boxNumbers.sort((a, b) => {
    const an = Number(a);
    const bn = Number(b);
    if (Number.isFinite(an) && Number.isFinite(bn)) return an - bn;
    return a.localeCompare(b);
  });

  if (boxNumbers.length === 0) {
    throw new Response("Packing list has no boxes to label", { status: 400 });
  }

  const totalBoxes = boxNumbers.length;
  const invoiceNumber = (list.invoiceNumber || list.title || `Packing list #${list.id}`).trim();

  // ── PDF setup ────────────────────────────────────────────────────
  const doc = new PDFDocument({ size: "A4", layout: "landscape", margin: 24 });
  const chunks: Buffer[] = [];
  doc.on("data", (chunk) => chunks.push(Buffer.from(chunk)));
  const finished = new Promise<Buffer>((resolve, reject) => {
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);
  });

  const pageWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const leftX = doc.page.margins.left;
  const rightX = doc.page.margins.left + pageWidth / 2 + 10;

  boxNumbers.forEach((boxNo, idx) => {
    if (idx > 0) doc.addPage();

    // Outer border around the entire label area.
    const labelY0 = doc.page.margins.top;
    const labelHeight = doc.page.height - doc.page.margins.top - doc.page.margins.bottom;
    doc.lineWidth(1).strokeColor("#111827")
      .rect(leftX, labelY0, pageWidth, labelHeight).stroke();

    // Inner padding.
    const padX = 18;
    const padY = 14;
    const innerX = leftX + padX;
    const innerWidth = pageWidth - padX * 2;
    let y = labelY0 + padY;

    // Sizes tuned so the entire label fits on a single A4-landscape
    // page even with the longest address block. lineBreak: false on
    // every text call prevents PDFKit from auto-paginating mid-label.
    const HEADER_FS = 40;
    const SECTION_LABEL_FS = 22;
    const ADDRESS_FS = 26;
    const FOOTER_FS = 11;

    // ── HEADER: INV. NO. + BOX NO. ──────────────────────────────
    doc.font("Helvetica-Bold").fontSize(HEADER_FS).fillColor("#111827");
    doc.text(`INV. NO.   # ${invoiceNumber}`, innerX, y, { width: innerWidth, align: "left", lineBreak: false });
    y = doc.y + 6;
    doc.text(`BOX NO.    # ${boxNo}/${totalBoxes}`, innerX, y, { width: innerWidth, align: "left", lineBreak: false });
    y = doc.y + 16;

    // Horizontal separator under the header.
    doc.lineWidth(0.8).strokeColor("#9ca3af").dash(4, { space: 4 });
    doc.moveTo(innerX, y).lineTo(innerX + innerWidth, y).stroke();
    doc.undash();
    y += 18;

    // ── BODY: To / From side-by-side ────────────────────────────
    const colWidth = (innerWidth - 30) / 2;
    const toX = innerX;
    const fromX = innerX + colWidth + 30;
    const bodyTop = y;

    // To column.
    doc.font("Helvetica-Bold").fontSize(SECTION_LABEL_FS).fillColor("#374151");
    doc.text("To: -", toX, bodyTop, { width: colWidth, align: "left", lineBreak: false });
    let toY = doc.y + 8;
    doc.font("Helvetica-Bold").fontSize(ADDRESS_FS).fillColor("#111827");
    for (const line of TO_ADDRESS_LINES) {
      doc.text(line, toX, toY, { width: colWidth, align: "left", lineBreak: false });
      toY = doc.y + 2;
    }

    // From column — anchored toward the bottom of the body, matching
    // the example label.
    const addressLineH = ADDRESS_FS * 1.15;
    const fromBlockHeight = SECTION_LABEL_FS + 8 + (FROM_ADDRESS_LINES.length * (addressLineH + 2));
    const bodyBottom = labelY0 + labelHeight - padY - 26; // leave room for instructions
    const fromTop = Math.max(bodyTop, bodyBottom - fromBlockHeight);
    doc.font("Helvetica-Bold").fontSize(SECTION_LABEL_FS).fillColor("#374151");
    doc.text("From: -", fromX, fromTop, { width: colWidth, align: "left", lineBreak: false });
    let fromY = doc.y + 8;
    doc.font("Helvetica-Bold").fontSize(ADDRESS_FS).fillColor("#111827");
    for (const line of FROM_ADDRESS_LINES) {
      doc.text(line, fromX, fromY, { width: colWidth, align: "left", lineBreak: false });
      fromY = doc.y + 2;
    }

    // ── FOOTER: Delivery instructions ───────────────────────────
    const footerY = labelY0 + labelHeight - padY - 14;
    doc.font("Helvetica").fontSize(FOOTER_FS).fillColor("#111827");
    doc.text(DELIVERY_INSTRUCTIONS, innerX, footerY, { width: innerWidth, align: "left", lineBreak: false });
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
      "Content-Disposition": `attachment; filename="shipping-labels-${safeName}.pdf"`,
      "Content-Length": String(pdfBuffer.length),
    },
  });
};
