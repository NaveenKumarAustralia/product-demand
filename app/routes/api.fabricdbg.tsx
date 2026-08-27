import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// TEMP diagnostic (no auth) — inspect the raw fabric-manual-sheets blob so we
// can see what sheets/rows exist and how they're classified, to debug why the
// Reorder Planner fabric picker finds nothing. Returns counts + a small sample
// of fabric NAMES only (no sensitive data). REMOVE after debugging.
const FABRIC_MANUAL_SHEETS_KEY = "production-portal-fabric-manual-sheets-v1";
const COMBINED_FABRIC_ON_ORDER_GID = "759049382";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  void request;
  const row = await prisma.portalSetting.findUnique({ where: { key: FABRIC_MANUAL_SHEETS_KEY }, select: { value: true } }).catch(() => null);
  const v = (row?.value ?? null) as unknown;
  if (!v || !Array.isArray(v)) {
    return Response.json({ ok: false, reason: "no-fabric-setting", type: typeof v, isArray: Array.isArray(v) });
  }
  const sheets = v as Array<{ gid?: string; name?: string; kind?: string; headers?: unknown[]; rows?: unknown[][] }>;
  const summary = sheets.map((s) => {
    const headers = Array.isArray(s.headers) ? s.headers.map((h) => String(h)) : [];
    const nameIdx = headers.findIndex((h) => /^name$/i.test(h));
    const rows = Array.isArray(s.rows) ? s.rows : [];
    const sampleNames = nameIdx >= 0
      ? rows.slice(0, 6).map((r) => String((r as unknown[])[nameIdx] ?? "")).filter(Boolean)
      : [];
    return {
      gid: s.gid,
      name: s.name,
      kind: s.kind,
      isCombinedOnOrder: s.gid === COMBINED_FABRIC_ON_ORDER_GID,
      headerCount: headers.length,
      headers,
      nameIdx,
      rowCount: rows.length,
      sampleNames,
    };
  });
  return Response.json({
    ok: true,
    sheetCount: sheets.length,
    kinds: Array.from(new Set(sheets.map((s) => s.kind))),
    summary,
  });
};
