import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

// Admin: view or set the auto SKU/barcode counter (the NEXT number to be used).
// Auto-generation (and the "Generate" button) hand out this number then advance it.
//   View:  /api/collection-sku-counter
//   Set:   /api/collection-sku-counter?set=1500   (next product becomes K1500 / 1500)
const KEY = "collection-auto-sku-next";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const url = new URL(request.url);
  const setRaw = url.searchParams.get("set");
  if (setRaw != null) {
    const n = Math.floor(Number(setRaw.replace(/[^0-9]/g, "")));
    if (!Number.isFinite(n) || n < 1) return Response.json({ ok: false, error: "Pass ?set=<a positive number>, e.g. ?set=1500" }, { status: 400 });
    await prisma.portalSetting.upsert({ where: { key: KEY }, create: { key: KEY, value: { next: n } }, update: { value: { next: n } } });
    return Response.json({ ok: true, set: true, next: n, note: `The next product created with a blank SKU will be K${n} / ${n}.` }, { headers: { "Cache-Control": "no-store" } });
  }

  const s = await prisma.portalSetting.findUnique({ where: { key: KEY }, select: { value: true } }).catch(() => null);
  const stored = (s?.value && typeof s.value === "object" && !Array.isArray(s.value)) ? Number((s.value as { next?: unknown }).next) : null;
  return Response.json({ ok: true, storedNext: stored, note: "storedNext is the raw saved counter. Values ≥ 100000 are ignored by the app (it restarts at 5000). Use ?set=<n> to set it." }, { headers: { "Cache-Control": "no-store" } });
};
