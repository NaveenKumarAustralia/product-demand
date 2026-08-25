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

  return Response.json(out);
};
