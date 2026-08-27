import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// Product Information style images for the Collections "Pick from Product
// Information" picker. The Collections page ships an image-STRIPPED product-info
// (for speed), so the picker fetches the images here on demand — searched
// server-side so the payload stays small. Cached briefly so repeated searches
// don't re-read the whole (large) blob.
const KEY = "production-portal-product-info-v2";
type StyleItem = { id: string; name: string; categoryName: string; imageUrl: string };
let cache: { at: number; styles: StyleItem[] } | null = null;

async function getStyles(): Promise<StyleItem[]> {
  if (cache && Date.now() - cache.at < 60_000) return cache.styles;
  const row = await prisma.portalSetting.findUnique({ where: { key: KEY }, select: { value: true } }).catch(() => null);
  const v = (row?.value ?? {}) as { categories?: Array<{ name?: string; styles?: Array<{ id?: string; name?: string; imageUrl?: string; hidden?: boolean }> }> };
  const styles: StyleItem[] = [];
  for (const cat of v.categories ?? []) {
    for (const s of cat.styles ?? []) {
      if (s?.hidden) continue;
      if (s?.imageUrl && s?.name) styles.push({ id: String(s.id ?? ""), name: String(s.name), categoryName: String(cat?.name ?? ""), imageUrl: String(s.imageUrl) });
    }
  }
  cache = { at: Date.now(), styles };
  return styles;
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const url = new URL(request.url);
  const q = (url.searchParams.get("q") ?? "").trim().toLowerCase();
  const limit = Math.min(300, Math.max(1, Number(url.searchParams.get("limit")) || 120));
  const all = await getStyles();
  const filtered = q ? all.filter((s) => s.name.toLowerCase().includes(q) || s.categoryName.toLowerCase().includes(q)) : all;
  return Response.json({ styles: filtered.slice(0, limit), total: filtered.length });
};
