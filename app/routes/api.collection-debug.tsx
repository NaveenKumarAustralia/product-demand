import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// TEMP diagnostic: reports the byte size of each collection's rows and which
// COLUMNS carry the weight — so we can see what makes one collection (e.g.
// Nila-Indian Summer) freeze the browser while others open fine.
//   /api/collection-debug            → every collection, heaviest first
//   /api/collection-debug?name=nila  → just collections whose name matches
// Returns only sizes/counts, never the actual data. Read-only.

type Row = Record<string, unknown>;

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const nameQ = (new URL(request.url).searchParams.get("name") ?? "").trim().toLowerCase();

  const collections = await prisma.$queryRawUnsafe<Array<{ id: number; name: string; rows: unknown }>>(
    `SELECT id, name, rows FROM "Collection" ORDER BY "sortOrder" ASC`,
  ).catch(() => [] as Array<{ id: number; name: string; rows: unknown }>);

  const out = collections
    .filter((c) => !nameQ || String(c.name ?? "").toLowerCase().includes(nameQ))
    .map((c) => {
      const rows: Row[] = Array.isArray(c.rows) ? c.rows as Row[] : [];
      const totalBytes = JSON.stringify(rows).length;
      // Per-column total bytes across all rows.
      const byColumn = new Map<string, number>();
      let biggestSingleCell = { row: -1, col: "", bytes: 0 };
      rows.forEach((row, ri) => {
        for (const [col, val] of Object.entries(row)) {
          const bytes = typeof val === "string" ? val.length : JSON.stringify(val ?? "").length;
          byColumn.set(col, (byColumn.get(col) ?? 0) + bytes);
          if (bytes > biggestSingleCell.bytes) biggestSingleCell = { row: ri, col, bytes };
        }
      });
      const columns = Array.from(byColumn.entries())
        .map(([col, bytes]) => ({ col, kb: Math.round(bytes / 1024) }))
        .sort((a, b) => b.kb - a.kb)
        .slice(0, 8);
      return {
        id: c.id,
        name: c.name,
        rowCount: rows.length,
        totalKB: Math.round(totalBytes / 1024),
        biggestCell: { row: biggestSingleCell.row, col: biggestSingleCell.col, kb: Math.round(biggestSingleCell.bytes / 1024) },
        columns,
      };
    })
    .sort((a, b) => b.totalKB - a.totalKB);

  return Response.json({ collections: out });
};
