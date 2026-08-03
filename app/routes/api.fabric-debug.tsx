import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// TEMP diagnostic: dumps the raw saved fabric-in-stock rows whose Name (or any
// cell) matches ?q=..., so we can see exactly which column holds the meters for
// a fabric (e.g. Candy / Nila) and whether the headers line up with the cells.
//   /api/fabric-debug?q=candy
// Read-only; safe to leave in briefly and remove once the meters bug is fixed.
const FABRIC_MANUAL_SHEETS_KEY = "production-portal-fabric-manual-sheets-v1";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const q = (new URL(request.url).searchParams.get("q") ?? "").trim().toLowerCase();
  if (!q) return Response.json({ error: "pass ?q=<fabric name>" });

  const setting = await prisma.portalSetting.findUnique({
    where: { key: FABRIC_MANUAL_SHEETS_KEY },
    select: { value: true },
  }).catch(() => null);

  const sheets = Array.isArray(setting?.value) ? setting!.value as Array<Record<string, unknown>> : [];
  const out: Array<Record<string, unknown>> = [];

  for (const sheet of sheets) {
    const name = String(sheet?.name ?? "");
    const gid = String(sheet?.gid ?? "");
    const kind = String(sheet?.kind ?? "");
    const headers = Array.isArray(sheet?.headers) ? (sheet.headers as unknown[]).map((h) => String(h ?? "")) : [];
    const rows = Array.isArray(sheet?.rows) ? sheet.rows as unknown[][] : [];
    // Which header column looks like the fabric Name?
    const nameIdx = headers.findIndex((h) => /^name$/i.test(h.trim()));
    for (const row of rows) {
      const cells = (row as unknown[]).map((c) => String(c ?? ""));
      const nameCell = nameIdx >= 0 ? (cells[nameIdx] ?? "") : "";
      const hay = (nameCell || cells.join(" ")).toLowerCase();
      if (!hay.includes(q)) continue;
      // Pair each header with its cell so misalignment is obvious, and also list
      // any extra cells that hang off the end past the headers (column shift).
      const paired: Record<string, string> = {};
      headers.forEach((h, i) => { paired[`${i}:${h}`] = cells[i] ?? ""; });
      const extraCells = cells.length > headers.length ? cells.slice(headers.length) : [];
      out.push({ sheet: name, gid, kind, nameCell, headerCount: headers.length, cellCount: cells.length, paired, extraCells });
    }
  }

  return Response.json({ q, matches: out.length, sheetCount: sheets.length, sheetNames: sheets.map((s) => String(s?.name ?? "")), rows: out });
};
