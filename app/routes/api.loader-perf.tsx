import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// TEMP diagnostic — measures where the collections/photoshoot loader spends its
// time on production. Returns query timings + per-setting blob sizes (no
// sensitive data, just KB + ms). Remove once the slow page is diagnosed.
export const loader = async (_args: LoaderFunctionArgs) => {
  const out: Record<string, unknown> = {};

  const t0 = Date.now();
  const settings = await prisma.portalSetting.findMany({ select: { key: true, value: true } });
  out.settingsQueryMs = Date.now() - t0;

  const sizes = settings
    .map((s) => ({ key: s.key, kb: Math.round(JSON.stringify(s.value ?? "").length / 1024) }))
    .sort((a, b) => b.kb - a.kb);
  out.totalSettingsKb = sizes.reduce((sum, x) => sum + x.kb, 0);
  out.topBlobs = sizes.slice(0, 12);

  const t1 = Date.now();
  try {
    const shoots = await prisma.$queryRawUnsafe(`SELECT id, name, "sortOrder", (thumbnail IS NOT NULL) AS h FROM "PhotoShoot"`);
    out.photoShootQueryMs = Date.now() - t1;
    out.shootCount = Array.isArray(shoots) ? shoots.length : 0;
  } catch (e) { out.photoShootErr = String(e).slice(0, 200); }

  const t2 = Date.now();
  try {
    const cols = await prisma.$queryRawUnsafe(`SELECT count(*)::int AS n FROM "Collection" WHERE "kind" = 'collection'`);
    out.collectionCountQueryMs = Date.now() - t2;
    out.collectionCount = Array.isArray(cols) ? (cols[0] as { n: number })?.n : cols;
  } catch (e) { out.collectionErr = String(e).slice(0, 200); }

  const t3 = Date.now();
  try {
    const s = await prisma.session.findFirst({ where: { accessToken: { not: "" } }, orderBy: { isOnline: "asc" }, select: { shop: true } });
    out.sessionQueryMs = Date.now() - t3;
    out.hasSession = Boolean(s?.shop);
  } catch (e) { out.sessionErr = String(e).slice(0, 200); }

  // Slim product-info (image-stripped in DB) — should be tiny + fast now.
  const t4 = Date.now();
  try {
    const rows = await prisma.$queryRawUnsafe<Array<{ value: unknown }>>(
      `SELECT CASE WHEN jsonb_typeof(value->'categories') = 'array' THEN
         jsonb_set(value, '{categories}', (
           SELECT COALESCE(jsonb_agg(
             CASE WHEN jsonb_typeof(cat->'styles') = 'array'
               THEN jsonb_set(cat, '{styles}', (
                 SELECT COALESCE(jsonb_agg(style - 'imageUrl'), '[]'::jsonb)
                 FROM jsonb_array_elements(cat->'styles') style))
               ELSE cat END), '[]'::jsonb)
           FROM jsonb_array_elements(value->'categories') cat))
         ELSE value END AS value
       FROM "PortalSetting" WHERE key = 'production-portal-product-info-v2'`);
    out.slimProductInfoMs = Date.now() - t4;
    out.slimProductInfoKb = Math.round(JSON.stringify(rows[0]?.value ?? "").length / 1024);
  } catch (e) { out.slimProductInfoErr = String(e).slice(0, 300); }

  const t5 = Date.now();
  try {
    const rows = await prisma.$queryRawUnsafe<Array<{ value: unknown }>>(
      `SELECT CASE WHEN jsonb_typeof(value) = 'array' THEN (
         SELECT COALESCE(jsonb_agg(
           CASE WHEN jsonb_typeof(sheet->'rows') = 'array'
             THEN jsonb_set(sheet, '{rows}', (
               SELECT COALESCE(jsonb_agg(
                 CASE WHEN jsonb_typeof(r) = 'array' AND jsonb_array_length(r) > 2
                   THEN jsonb_set(r, '{2}', '""'::jsonb) ELSE r END), '[]'::jsonb)
               FROM jsonb_array_elements(sheet->'rows') AS r))
             ELSE sheet END), '[]'::jsonb)
         FROM jsonb_array_elements(value) AS sheet) ELSE value END AS value
       FROM "PortalSetting" WHERE key = 'production-portal-fabric-manual-sheets-v1'`);
    out.slimFabricMs = Date.now() - t5;
    out.slimFabricKb = Math.round(JSON.stringify(rows[0]?.value ?? "").length / 1024);
  } catch (e) { out.slimFabricErr = String(e).slice(0, 300); }

  return Response.json(out);
};
