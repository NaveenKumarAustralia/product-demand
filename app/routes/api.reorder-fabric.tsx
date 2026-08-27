import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import {
  loadProductInfoForAction,
  saveProductInfo,
  loadManualFabricSheetsForAction,
  combinedFabricSheetsForIndex,
  buildFabricStockIndex,
} from "./portal._index";

// Reorder Planner fabric picker — a dedicated RESOURCE route (returns raw JSON).
// The previous version POSTed to the /portal action via fetch(), but /portal is
// a full route, so a document-mode POST re-renders HTML instead of returning the
// action's JSON — the picker always saw an empty list. A resource route (this
// file, loader/action only, no default export) returns the JSON directly.
//
// GET  ?title=<product title>  → { candidates, all, chosenKey, pinned }
// POST title=<title> fabricKey=<key|"">  → set/clear the remembered pick.

function buildFabricPayload(productInfo: Awaited<ReturnType<typeof loadProductInfoForAction>>, manualFabricSheets: Awaited<ReturnType<typeof loadManualFabricSheetsForAction>>, title: string) {
  const titleLower = title.toLowerCase();
  const keyFor = (sheetName: string, name: string) => `${sheetName.trim().toLowerCase()}::${name.trim().toLowerCase()}`;

  const stockIndex = buildFabricStockIndex(combinedFabricSheetsForIndex(manualFabricSheets));
  const orderIndex = buildFabricStockIndex(manualFabricSheets.filter((s) => s.kind === "order" || s.kind === "wide-order"));
  const onOrderByName = new Map<string, number>();
  for (const e of orderIndex) {
    if (e.kind !== "order") continue;
    const n = e.name.trim().toLowerCase();
    onOrderByName.set(n, (onOrderByName.get(n) ?? 0) + (Number(e.meters) || 0));
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

  const styleList: Array<{ n: string; id: string }> = [];
  for (const cat of productInfo.categories) for (const st of cat.styles) { const n = st.name?.trim().toLowerCase(); if (n) styleList.push({ n, id: st.id }); }
  styleList.sort((a, b) => b.n.length - a.n.length);
  const overrideStyleId = (productInfo.titleStyleOverrides ?? {})[titleLower];
  const style = (overrideStyleId ? styleList.find((s) => s.id === overrideStyleId) : null)
    ?? styleList.find((s) => titleLower === s.n || titleLower.startsWith(s.n + " ")) ?? null;
  const styleId = style?.id ?? null;

  const distinctNames = Array.from(new Set([...merged.values()].map((c) => c.name.trim().toLowerCase()).filter((n) => n.length >= 3)));
  const nameMatched = new Set<string>();
  for (const n of distinctNames) {
    const escaped = n.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    if (new RegExp(`\\b${escaped}\\b`, "i").test(titleLower)) nameMatched.add(n);
  }

  const candidates = [...merged.values()]
    .map((c) => ({ key: c.key, name: c.name, fabricType: c.fabricType, inStock: c.inStock, onOrder: c.onOrder, linked: !!(styleId && c.styleIds.has(styleId)) || nameMatched.has(c.name.trim().toLowerCase()) }))
    .filter((c) => c.linked)
    .sort((a, b) => (b.inStock + b.onOrder) - (a.inStock + a.onOrder))
    .map(({ linked: _l, ...c }) => c);

  const all = [...merged.values()]
    .map((c) => ({ key: c.key, name: c.name, fabricType: c.fabricType, inStock: c.inStock, onOrder: c.onOrder }))
    .sort((a, b) => a.name.localeCompare(b.name) || a.fabricType.localeCompare(b.fabricType));

  const pinnedKey = ((productInfo.titleFabricOverrides ?? {})[titleLower] ?? "").replace(/::\d+$/, "");
  let chosenKey = "";
  if (pinnedKey && merged.has(pinnedKey)) chosenKey = pinnedKey;
  else if (candidates.length === 1) chosenKey = candidates[0].key;

  return { ok: true, candidates, all, chosenKey, pinned: Boolean(pinnedKey && merged.has(pinnedKey)) };
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const title = (new URL(request.url).searchParams.get("title") ?? "").trim();
  if (!title) return Response.json({ ok: false, candidates: [], all: [] });
  try {
    const [productInfo, manualFabricSheets] = await Promise.all([
      loadProductInfoForAction(),
      loadManualFabricSheetsForAction(),
    ]);
    return Response.json(buildFabricPayload(productInfo, manualFabricSheets, title));
  } catch (e) {
    console.error("[api.reorder-fabric] loader", e);
    return Response.json({ ok: false, candidates: [], all: [], error: "failed" });
  }
};

export const action = async ({ request }: ActionFunctionArgs) => {
  const form = await request.formData();
  const title = String(form.get("title") ?? "").trim().toLowerCase();
  const fabricKey = String(form.get("fabricKey") ?? "").trim().toLowerCase();
  if (!title) return Response.json({ ok: false });
  try {
    const productInfo = await loadProductInfoForAction();
    if (!productInfo.titleFabricOverrides) productInfo.titleFabricOverrides = {};
    if (fabricKey) productInfo.titleFabricOverrides[title] = fabricKey;
    else delete productInfo.titleFabricOverrides[title];
    await saveProductInfo(productInfo);
    return Response.json({ ok: true });
  } catch (e) {
    console.error("[api.reorder-fabric] action", e);
    return Response.json({ ok: false, error: "failed" });
  }
};
