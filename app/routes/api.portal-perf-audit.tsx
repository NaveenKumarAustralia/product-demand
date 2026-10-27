import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

// Admin diagnostic (read-only): measures what the portal loader has to haul.
// 1) Every PortalSetting value's size in KB (the loader fetches many of these
//    on every page, so a bloated blob slows the whole app).
// 2) Row counts + rough size of the big order/packing tables the loader joins.
// Open in the browser while logged into the portal as an admin. No content is
// returned — only sizes — so it's safe to view.
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  // Per-key value size. octet_length(value::text) is the serialized JSON size —
  // a close proxy for what crosses the wire into the loader.
  let settingSizes: Array<{ key: string; kb: number }> = [];
  try {
    const rows = await prisma.$queryRawUnsafe<Array<{ key: string; bytes: number | bigint }>>(
      `SELECT key, octet_length(value::text) AS bytes FROM "PortalSetting" ORDER BY bytes DESC`,
    );
    settingSizes = rows.map((r) => ({ key: r.key, kb: Math.round(Number(r.bytes) / 1024) }));
  } catch (e) {
    settingSizes = [{ key: `ERROR: ${e instanceof Error ? e.message : String(e)}`, kb: 0 }];
  }

  const totalSettingsKb = settingSizes.reduce((a, r) => a + r.kb, 0);

  // Row counts for the tables the loader joins (lines included inflate payloads).
  const counts: Record<string, number> = {};
  try {
    const waitlistCount = prisma.$queryRawUnsafe<Array<{ n: number | bigint }>>(`SELECT COUNT(*)::int AS n FROM "PreorderWaitlist"`).then((r) => Number(r[0]?.n ?? -1)).catch(() => -1);
    const [openOrders, orderLines, packingLists, packingLines, waitlist, activity90d] = await Promise.all([
      prisma.supplierOrder.count({ where: { status: "open" } }),
      prisma.orderLine.count(),
      prisma.packingList.count(),
      prisma.packingListLine.count(),
      waitlistCount,
      prisma.activityLog.count({ where: { createdAt: { gte: new Date(Date.now() - 90 * 24 * 60 * 60 * 1000) } } }),
    ]);
    counts.openSupplierOrders = openOrders;
    counts.orderLines_total = orderLines;
    counts.packingLists = packingLists;
    counts.packingListLines_total = packingLines;
    counts.preorderWaitlist = waitlist;
    counts.activityLogs_last90d = activity90d;
  } catch (e) {
    counts.error = -1;
    counts[`msg_${e instanceof Error ? e.message.slice(0, 40) : "unknown"}`] = -1;
  }

  return Response.json(
    {
      ok: true,
      note: "Sizes in KB. Anything in the 1,000s of KB that the loader fetches on every page is a slowdown suspect.",
      totalSettingsKb,
      top20Settings: settingSizes.slice(0, 20),
      allSettings: settingSizes,
      tableCounts: counts,
    },
    { headers: { "Cache-Control": "no-store" } },
  );
};
