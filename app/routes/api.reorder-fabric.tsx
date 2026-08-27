import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// Reorder Planner fabric picker — a self-contained RESOURCE route returning raw
// JSON. (A POST to the /portal action via fetch() re-renders HTML instead of
// returning JSON, and importing the /portal route module here crashes at load,
// so this route reads the settings directly.) The matching logic MIRRORS
// buildFabricStockIndex + the fabric resolution in portal._index.tsx.
//
// GET ?title=<product title> → { ok, candidates, all, chosenKey, pinned }
// The remembered pick is WRITTEN by the /portal set_title_fabric_override action
// (that write persists even though its HTML response is ignored).

const FABRIC_MANUAL_SHEETS_KEY = "production-portal-fabric-manual-sheets-v1";
const PRODUCT_INFO_KEY = "production-portal-product-info-v2";
const COMBINED_FABRIC_ON_ORDER_GID = "759049382";
const HIDDEN_FABRIC_SHEET_NAMES = new Set(["new fabric on order", "fabric on order"]);
const normName = (name: string) => name.trim().toLowerCase().replace(/\s+/g, " ");

type RawSheet = { gid?: string; name?: string; kind?: string; headers?: unknown[]; rows?: unknown[][] };
type FabricEntry = { name: string; meters: number; sheetName: string; fabricType?: string; kind: "stock" | "order"; styleMeters?: Record<string, number> };

// Verbatim port of buildFabricStockIndex(portal._index) — one entry per fabric
// row in any stock / on-order sheet.
function buildFabricStockIndex(sheets: Array<{ gid: string; kind: string; name: string; headers: string[]; rows: string[][] }>): FabricEntry[] {
  const out: FabricEntry[] = [];
  for (const sheet of sheets) {
    const isStock = sheet.kind === "stock" || sheet.kind === "simple-stock" || sheet.gid === COMBINED_FABRIC_ON_ORDER_GID;
    const isOrder = !isStock && (sheet.kind === "order" || sheet.kind === "wide-order");
    if (!isStock && !isOrder) continue;
    const nameIdx = sheet.headers.findIndex((h) => /^name$/i.test(h));
    const metersIdx = isStock
      ? sheet.headers.findIndex((h) => /meters?\s*in\s*stock|in\s*stock|meters?\s*available|^meters?$/i.test(h))
      : sheet.headers.findIndex((h) => /quantity\s*ordered|meters?\s*ordered/i.test(h));
    const productsIdx = isStock ? sheet.headers.findIndex((h) => /^products?$/i.test(h)) : -1;
    const fabricTypeIdx = sheet.headers.findIndex((h) => /^fabric\s*type$/i.test(h) || /^type$/i.test(h));
    if (nameIdx < 0) continue;
    for (const row of sheet.rows) {
      const name = (row[nameIdx] ?? "").trim();
      if (!name || name.length < 2) continue;
      const cleaned = metersIdx >= 0 ? (row[metersIdx] ?? "").toString().split(/[^0-9.]/)[0] : "";
      const m = Number(cleaned) || 0;
      if (!Number.isFinite(m)) continue;
      if (isOrder && m === 0) continue;
      let styleMeters: Record<string, number> | undefined;
      if (productsIdx >= 0) {
        try {
          const parsed = JSON.parse((row[productsIdx] ?? "").toString() || "{}");
          if (parsed && Array.isArray(parsed.styles)) {
            const map: Record<string, number> = {};
            for (const item of parsed.styles) {
              const styleId = String(item?.styleId ?? "").trim();
              const meters = Number(String(item?.meters ?? "").trim());
              if (styleId && Number.isFinite(meters) && meters > 0) map[styleId] = meters;
            }
            if (Object.keys(map).length) styleMeters = map;
          }
        } catch { /* legacy plain text — ignore */ }
      }
      const fabricType = fabricTypeIdx >= 0 ? (row[fabricTypeIdx] ?? "").toString().trim() : "";
      out.push({ name, meters: m, sheetName: sheet.name ?? "", fabricType: fabricType || undefined, kind: isStock ? "stock" : "order", styleMeters });
    }
  }
  return out;
}

function normalizeSheets(raw: unknown): Array<{ gid: string; kind: string; name: string; headers: string[]; rows: string[][] }> {
  if (!Array.isArray(raw)) return [];
  return (raw as RawSheet[]).map((s) => ({
    gid: String(s.gid ?? ""),
    kind: String(s.kind ?? "stock"),
    name: String(s.name ?? ""),
    headers: Array.isArray(s.headers) ? s.headers.map((h) => String(h ?? "")) : [],
    rows: Array.isArray(s.rows) ? s.rows.filter((r): r is unknown[] => Array.isArray(r)).map((r) => r.map((c) => String(c ?? ""))) : [],
  }));
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const title = (new URL(request.url).searchParams.get("title") ?? "").trim();
  if (!title) return Response.json({ ok: false, candidates: [], all: [] });
  const titleLower = title.toLowerCase();
  const keyFor = (sheetName: string, name: string) => `${sheetName.trim().toLowerCase()}::${name.trim().toLowerCase()}`;

  try {
    // Fabric sheets (small blob). Product-info: pull ONLY style id/name + the two
    // override maps via jsonb, so the 33MB image blob never crosses into JS.
    const [fabricRow, piRows] = await Promise.all([
      prisma.portalSetting.findUnique({ where: { key: FABRIC_MANUAL_SHEETS_KEY }, select: { value: true } }),
      prisma.$queryRawUnsafe<Array<{ categories: unknown; tfo: unknown; tso: unknown }>>(
        `SELECT
           (SELECT jsonb_agg(jsonb_build_object('styles',
              COALESCE((SELECT jsonb_agg(jsonb_build_object('id', s->'id', 'name', s->'name'))
                        FROM jsonb_array_elements(cat->'styles') s), '[]'::jsonb)))
            FROM jsonb_array_elements(value->'categories') cat) AS categories,
           value->'titleFabricOverrides' AS tfo,
           value->'titleStyleOverrides'  AS tso
         FROM "PortalSetting" WHERE key = $1`,
        PRODUCT_INFO_KEY,
      ).catch(() => []),
    ]);

    const sheets = normalizeSheets(fabricRow?.value);
    const stockSheets = sheets.filter((s) => !HIDDEN_FABRIC_SHEET_NAMES.has(normName(s.name)));
    const stockIndex = buildFabricStockIndex(stockSheets);

    // On-order meters by fabric name. Read the order / wide-order sheets AND the
    // combined "On Order" sheet directly — buildFabricStockIndex forces the
    // combined gid to "stock" and only matches "…ordered" columns, so it misses
    // both the On Order tab's Quantity Ordered and the wide-order tab's "Meters".
    const onOrderByName = new Map<string, number>();
    for (const sheet of sheets) {
      const isOrderSheet = sheet.kind === "order" || sheet.kind === "wide-order" || sheet.gid === COMBINED_FABRIC_ON_ORDER_GID;
      if (!isOrderSheet) continue;
      const nameIdx = sheet.headers.findIndex((h) => /^name$/i.test(h));
      const qtyIdx = sheet.headers.findIndex((h) => /quantity\s*ordered|meters?\s*ordered|^meters?$/i.test(h));
      if (nameIdx < 0 || qtyIdx < 0) continue;
      for (const row of sheet.rows) {
        const n = (row[nameIdx] ?? "").trim().toLowerCase();
        if (!n || n.length < 2) continue;
        const m = Number((row[qtyIdx] ?? "").toString().split(/[^0-9.]/)[0]) || 0;
        if (m > 0) onOrderByName.set(n, (onOrderByName.get(n) ?? 0) + m);
      }
    }

    type Merged = { key: string; name: string; fabricType: string; inStock: number; onOrder: number; styleIds: Set<string> };
    const merged = new Map<string, Merged>();
    for (const e of stockIndex) {
      if (e.kind !== "stock") continue;
      const nameLower = e.name.trim().toLowerCase();
      if (!nameLower) continue;
      const key = keyFor(e.sheetName, e.name);
      let c = merged.get(key);
      if (!c) { c = { key, name: e.name, fabricType: e.fabricType ?? "", inStock: 0, onOrder: onOrderByName.get(nameLower) ?? 0, styleIds: new Set() }; merged.set(key, c); }
      if (!c.fabricType && e.fabricType) c.fabricType = e.fabricType;
      c.inStock += Number(e.meters) || 0;
      if (e.styleMeters) for (const sid of Object.keys(e.styleMeters)) c.styleIds.add(sid);
    }

    // Resolve the product's style from its title (longest style-name prefix wins).
    const pi = piRows[0] ?? { categories: [], tfo: {}, tso: {} };
    const cats = (Array.isArray(pi.categories) ? pi.categories : []) as Array<{ styles?: Array<{ id?: string; name?: string }> }>;
    const styleList: Array<{ n: string; id: string }> = [];
    for (const cat of cats) for (const st of cat.styles ?? []) { const n = String(st?.name ?? "").trim().toLowerCase(); if (n) styleList.push({ n, id: String(st?.id ?? "") }); }
    styleList.sort((a, b) => b.n.length - a.n.length);
    const tso = (pi.tso ?? {}) as Record<string, string>;
    const tfo = (pi.tfo ?? {}) as Record<string, string>;
    const overrideStyleId = tso[titleLower];
    const style = (overrideStyleId ? styleList.find((s) => s.id === overrideStyleId) : null)
      ?? styleList.find((s) => titleLower === s.n || titleLower.startsWith(s.n + " ")) ?? null;
    const styleId = style?.id ?? null;

    const nameMatched = new Set<string>();
    for (const nm of new Set([...merged.values()].map((c) => c.name.trim().toLowerCase()).filter((n) => n.length >= 3))) {
      const escaped = nm.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      if (new RegExp(`\\b${escaped}\\b`, "i").test(titleLower)) nameMatched.add(nm);
    }

    const candidates = [...merged.values()]
      .map((c) => ({ key: c.key, name: c.name, fabricType: c.fabricType, inStock: c.inStock, onOrder: c.onOrder, linked: !!(styleId && c.styleIds.has(styleId)) || nameMatched.has(c.name.trim().toLowerCase()) }))
      .filter((c) => c.linked)
      .sort((a, b) => (b.inStock + b.onOrder) - (a.inStock + a.onOrder))
      .map(({ linked: _l, ...c }) => c);

    const all = [...merged.values()]
      .map((c) => ({ key: c.key, name: c.name, fabricType: c.fabricType, inStock: c.inStock, onOrder: c.onOrder }))
      .sort((a, b) => a.name.localeCompare(b.name) || a.fabricType.localeCompare(b.fabricType));

    const pinnedKey = (tfo[titleLower] ?? "").replace(/::\d+$/, "");
    let chosenKey = "";
    if (pinnedKey && merged.has(pinnedKey)) chosenKey = pinnedKey;
    else if (candidates.length === 1) chosenKey = candidates[0].key;

    return Response.json({ ok: true, candidates, all, chosenKey, pinned: Boolean(pinnedKey && merged.has(pinnedKey)) });
  } catch (e) {
    console.error("[api.reorder-fabric]", e);
    return Response.json({ ok: false, candidates: [], all: [], error: "failed" });
  }
};
