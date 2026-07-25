import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// Streams the JJ Restock orders as an .xlsx with the product image embedded
// in each order's first row. One row per ordered variant (qty > 0), grouped
// by product; product-level columns (image, colour code, name, style code,
// cost) are written only on the group's first row.
//
// Optional ?q=<term> filters to orders whose title / style / colour matches
// (mirrors the on-screen search) so a filtered view exports the same set.

const THB_AUD_CACHE_KEY = "production-portal-thb-aud-rate-v1";
const FX_BAHT_BUFFER = 1;

const SIZE_ORDER = ["FREE SIZE", "XS", "S", "M", "L", "XL", "2XL", "XXL", "3XL", "XXXL", "S-M", "S/M", "M-L", "M/L", "L-XL", "L/XL"];
function sizeRank(label: string): number {
  const i = SIZE_ORDER.indexOf(label.trim().toUpperCase());
  return i === -1 ? 999 : i;
}

async function thbPerAud(): Promise<number | null> {
  try {
    const setting = await prisma.portalSetting.findUnique({ where: { key: THB_AUD_CACHE_KEY }, select: { value: true } });
    const cached = setting?.value as { thbPerAud?: number } | null | undefined;
    if (cached && typeof cached.thbPerAud === "number" && cached.thbPerAud > 0) return cached.thbPerAud;
  } catch { /* ignore */ }
  return null;
}
function bahtToAud(baht: number | null | undefined, rate: number | null): number | null {
  if (!baht || baht <= 0 || !rate || rate <= FX_BAHT_BUFFER) return null;
  return baht / (rate - FX_BAHT_BUFFER);
}

async function fetchImage(url: string): Promise<{ buffer: Buffer; ext: "png" | "jpeg" | "gif" } | null> {
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    const type = (res.headers.get("content-type") ?? "").toLowerCase();
    const ext: "png" | "jpeg" | "gif" = type.includes("png") ? "png" : type.includes("gif") ? "gif" : "jpeg";
    const buffer = Buffer.from(await res.arrayBuffer());
    if (buffer.length < 64 || buffer.length > 8 * 1024 * 1024) return null;
    return { buffer, ext };
  } catch {
    return null;
  }
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const url = new URL(request.url);
  const q = (url.searchParams.get("q") ?? "").trim().toLowerCase();

  const orders = await prisma.supplierOrder.findMany({
    where: { supplier: "JJ" },
    orderBy: [{ productTitle: "asc" }, { id: "asc" }],
    select: {
      id: true, productTitle: true, productImageUrl: true, costBaht: true,
      colourCode: true, styleCode: true,
      lines: { select: { variantTitle: true, sku: true, barcode: true, qtyOrdered: true } },
    },
  }).catch(() => [] as Array<{
    id: number; productTitle: string; productImageUrl: string | null; costBaht: number | null;
    colourCode: string | null; styleCode: string | null;
    lines: Array<{ variantTitle: string; sku: string | null; barcode: string | null; qtyOrdered: number }>;
  }>);

  const filtered = orders.filter((o) => {
    if (!q) return true;
    return (o.productTitle ?? "").toLowerCase().includes(q)
      || (o.styleCode ?? "").toLowerCase().includes(q)
      || (o.colourCode ?? "").toLowerCase().includes(q);
  });

  const rate = await thbPerAud();

  const ExcelJS = (await import("exceljs")).default;
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("JJ Orders");

  sheet.columns = [
    { header: "Image", key: "image", width: 14 },
    { header: "Colour Code", key: "colour", width: 14 },
    { header: "Product", key: "product", width: 28 },
    { header: "Style Code", key: "style", width: 14 },
    { header: "Size", key: "size", width: 10 },
    { header: "SKU", key: "sku", width: 20 },
    { header: "Barcode", key: "barcode", width: 20 },
    { header: "Qty", key: "qty", width: 8 },
    { header: "Cost (฿)", key: "baht", width: 12 },
    { header: "Cost (A$)", key: "aud", width: 12 },
  ];
  const header = sheet.getRow(1);
  header.font = { bold: true };
  header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEEF2F7" } };
  header.alignment = { vertical: "middle" };
  sheet.views = [{ state: "frozen", ySplit: 1 }];

  let excelRow = 2; // 1-based; row 1 is the header
  for (const order of filtered) {
    const orderedLines = order.lines
      .filter((l) => (l.qtyOrdered || 0) > 0)
      .sort((a, b) => sizeRank(a.variantTitle) - sizeRank(b.variantTitle));
    if (orderedLines.length === 0) continue;

    const firstExcelRow = excelRow;
    const aud = bahtToAud(order.costBaht, rate);

    orderedLines.forEach((line, i) => {
      const row = sheet.getRow(excelRow);
      if (i === 0) {
        row.getCell("colour").value = order.colourCode ?? "";
        row.getCell("product").value = order.productTitle ?? "";
        row.getCell("style").value = order.styleCode ?? "";
        row.getCell("baht").value = order.costBaht ?? "";
        row.getCell("aud").value = aud != null ? Number(aud.toFixed(2)) : "";
      }
      row.getCell("size").value = line.variantTitle ?? "";
      row.getCell("sku").value = line.sku ?? "";
      row.getCell("barcode").value = line.barcode ?? "";
      row.getCell("qty").value = line.qtyOrdered ?? 0;
      row.alignment = { vertical: "middle" };
      excelRow++;
    });

    // Embed the product image spanning the first row of the order group.
    if (order.productImageUrl) {
      const img = await fetchImage(order.productImageUrl);
      if (img) {
        const imageId = workbook.addImage({ buffer: img.buffer as unknown as ArrayBuffer, extension: img.ext });
        // Make the group's first row tall enough for the picture.
        sheet.getRow(firstExcelRow).height = 78;
        sheet.addImage(imageId, {
          tl: { col: 0.1, row: firstExcelRow - 1 + 0.05 },
          ext: { width: 84, height: 100 },
          editAs: "oneCell",
        });
      }
    }
  }

  const body = await workbook.xlsx.writeBuffer();
  const stamp = new URL(request.url).searchParams.get("stamp") ?? "";
  const filename = `jj-orders${stamp ? `-${stamp}` : ""}.xlsx`;
  return new Response(body, {
    status: 200,
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="${filename}"`,
      "Cache-Control": "no-store",
    },
  });
};
