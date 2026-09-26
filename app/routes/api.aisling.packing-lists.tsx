import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { authorizeAislingRequest } from "../api-auth.server";

// Read-only packing list feed for the Aisling Equation factory app, which
// matches shipments to its own records by invoice number and counts pieces.
// Server-to-server only (AISLING_API_KEY), so no CORS headers.
//
// GET                 → { packingLists: [summary…] }   newest shipment first
// GET ?id=<id>        → { packingList: summary + lines }
// GET ?since=<ISO>    → only lists changed since then (for incremental sync)
//
// Pieces are the sum of each line's per-size qtys. priceRupees is passed
// through exactly as stored: it is often empty because the portal derives the
// effective price from its style costing, so no rupee total is claimed here.

type Qtys = Record<string, number>;

function normalizeQtys(value: unknown): Qtys {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  const out: Qtys = {};
  for (const [size, qty] of Object.entries(value as Record<string, unknown>)) {
    const n = Number(qty);
    if (Number.isFinite(n) && n !== 0) out[size] = n;
  }
  return out;
}
const sumQtys = (q: Qtys) => Object.values(q).reduce((a, b) => a + b, 0);

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const unauthorized = authorizeAislingRequest(request);
  if (unauthorized) return unauthorized;

  const url = new URL(request.url);
  const id = Number(url.searchParams.get("id"));
  const since = url.searchParams.get("since");
  const sinceDate = since ? new Date(since) : null;
  if (sinceDate && Number.isNaN(sinceDate.getTime())) {
    return Response.json({ error: "since must be an ISO date" }, { status: 400 });
  }

  try {
    const lists = await prisma.packingList.findMany({
      where: {
        hiddenAt: null,
        ...(id ? { id } : {}),
        ...(sinceDate ? { updatedAt: { gt: sinceDate } } : {}),
      },
      orderBy: [{ shipmentDate: { sort: "desc", nulls: "first" } }, { id: "desc" }],
      select: {
        id: true, title: true, invoiceNumber: true, shipmentDate: true,
        expectedLeaveFactoryDate: true, shippingMethod: true, status: true, notes: true,
        createdAt: true, updatedAt: true,
        // fabricImageData is left out on purpose: it holds inline image data.
        lines: {
          orderBy: [{ sortOrder: "asc" }, { id: "asc" }],
          select: {
            id: true, boxNumber: true, productId: true, productTitle: true, productImageUrl: true,
            sku: true, isCustom: true, qtys: true, priceRupees: true, weight: true, notes: true,
          },
        },
      },
    });

    const shaped = lists.map(({ lines, ...list }) => {
      const shapedLines = lines.map((line) => {
        const qtys = normalizeQtys(line.qtys);
        return { ...line, qtys, pieces: sumQtys(qtys) };
      });
      const boxes = new Set(shapedLines.map((l) => l.boxNumber).filter(Boolean));
      const summary = {
        ...list,
        pieces: shapedLines.reduce((a, l) => a + l.pieces, 0),
        lineCount: shapedLines.length,
        boxCount: boxes.size,
        totalWeight: shapedLines.reduce((a, l) => a + (l.weight ?? 0), 0),
      };
      return id ? { ...summary, lines: shapedLines } : summary;
    });

    if (id) {
      if (!shaped.length) return Response.json({ error: "Packing list not found" }, { status: 404 });
      return Response.json({ packingList: shaped[0] });
    }
    return Response.json({ packingLists: shaped });
  } catch (err) {
    console.error("aisling packing-lists error:", err);
    return Response.json({ error: "Database error" }, { status: 500 });
  }
};
