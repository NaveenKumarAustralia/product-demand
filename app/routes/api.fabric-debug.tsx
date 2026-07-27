import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// TEMPORARY diagnostic: dump the stored fabric sheets so we can see why a
// specific fabric (e.g. "candy") isn't being indexed / isn't appearing in the
// Collections fabric picker. Read-only.
//   /api/fabric-debug            → every sheet's name/kind/headers + row counts
//   /api/fabric-debug?name=candy → also lists rows whose Name contains "candy"
const FABRIC_MANUAL_SHEETS_KEY = "production-portal-fabric-manual-sheets-v1";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const url = new URL(request.url);
  const needle = (url.searchParams.get("name") ?? "").trim().toLowerCase();

  const setting = await prisma.portalSetting.findUnique({
    where: { key: FABRIC_MANUAL_SHEETS_KEY },
    select: { value: true },
  });
  const raw = setting?.value;
  if (!Array.isArray(raw)) {
    return Response.json({ ok: false, note: "no manual fabric sheets stored", type: typeof raw });
  }

  const sheets = raw as Array<Record<string, unknown>>;
  const nameHeaderRe = /^name$/i;
  const costHeaderRe = /cost\s*per\s*meter/i;

  const summary = sheets.map((sheet) => {
    const headers = Array.isArray(sheet.headers) ? (sheet.headers as unknown[]).map((h) => String(h)) : [];
    const rows = Array.isArray(sheet.rows) ? (sheet.rows as unknown[][]) : [];
    const nameIdx = headers.findIndex((h) => nameHeaderRe.test(h.trim()));
    const costIdx = headers.findIndex((h) => costHeaderRe.test(h.trim()));
    const out: Record<string, unknown> = {
      name: sheet.name,
      gid: sheet.gid,
      kind: sheet.kind,
      headerCount: headers.length,
      headers,
      rowCount: rows.length,
      nameIdx,
      costIdx,
    };
    if (needle) {
      const matches = rows
        .map((row, i) => ({ i, row }))
        .filter(({ row }) => nameIdx >= 0 && String(row[nameIdx] ?? "").toLowerCase().includes(needle))
        .map(({ i, row }) => ({
          rowIndex: i,
          name: nameIdx >= 0 ? row[nameIdx] : null,
          cost: costIdx >= 0 ? row[costIdx] : null,
          cellCount: row.length,
          // Full row so any misalignment is visible.
          cells: row.map((c) => (typeof c === "string" && c.startsWith("data:") ? "<image>" : c)),
        }));
      out.matches = matches;
    }
    return out;
  });

  return Response.json({
    ok: true,
    sheetCount: sheets.length,
    needle: needle || null,
    sheets: summary,
  });
};
