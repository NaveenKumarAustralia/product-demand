import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

// Back-in-Stock reports (read-only). Returns three datasets for a date range:
//  - emailSends      : notify-me emails fired (by notifiedAt)
//  - converted       : waitlisted customers marked converted (by updatedAt)
//  - productsInDemand : current waiting signups grouped by product (by createdAt)
// Each is also broken down by product so the report can rank them.
function parseDay(value: string | null, fallbackDaysAgo: number): Date {
  if (value) {
    const d = new Date(`${value}T00:00:00.000Z`);
    if (!Number.isNaN(d.getTime())) return d;
  }
  const d = new Date();
  d.setUTCHours(0, 0, 0, 0);
  d.setUTCDate(d.getUTCDate() - fallbackDaysAgo);
  return d;
}

type ProductCount = { product: string; count: number };
type DemandRow = { product: string; count: number; au: number; usa: number };

export const loader = async ({ request }: LoaderFunctionArgs) => {
  try {
    await requirePreorderPortalUser(request);
    const url = new URL(request.url);
    const from = parseDay(url.searchParams.get("from"), 30);
    const toRaw = parseDay(url.searchParams.get("to"), 0);
    // Make the upper bound inclusive of the whole "to" day.
    const to = new Date(toRaw.getTime());
    to.setUTCDate(to.getUTCDate() + 1);

    const [emailByProduct, convertedByProduct, demandByProductMarket] = await Promise.all([
      prisma.$queryRawUnsafe<Array<{ product: string; count: bigint }>>(
        `SELECT COALESCE(NULLIF("productTitle",''),'Unknown product') AS product, COUNT(*)::bigint AS count
         FROM "PreorderWaitlist"
         WHERE "notifiedAt" IS NOT NULL AND "notifiedAt" >= $1 AND "notifiedAt" < $2
         GROUP BY 1 ORDER BY count DESC`,
        from, to,
      ),
      prisma.$queryRawUnsafe<Array<{ product: string; count: bigint }>>(
        `SELECT COALESCE(NULLIF("productTitle",''),'Unknown product') AS product, COUNT(*)::bigint AS count
         FROM "PreorderWaitlist"
         WHERE "status" = 'converted' AND "updatedAt" >= $1 AND "updatedAt" < $2
         GROUP BY 1 ORDER BY count DESC`,
        from, to,
      ),
      prisma.$queryRawUnsafe<Array<{ product: string; market: string; count: bigint }>>(
        `SELECT COALESCE(NULLIF("productTitle",''),'Unknown product') AS product, "market", COUNT(*)::bigint AS count
         FROM "PreorderWaitlist"
         WHERE "status" = 'waiting' AND "createdAt" >= $1 AND "createdAt" < $2
         GROUP BY product, "market"`,
        from, to,
      ),
    ]);

    const toPc = (rows: Array<{ product: string; count: bigint }>): ProductCount[] =>
      rows.map((r) => ({ product: r.product, count: Number(r.count) }));
    const emailSends = toPc(emailByProduct);
    const converted = toPc(convertedByProduct);

    const demandMap = new Map<string, DemandRow>();
    for (const r of demandByProductMarket) {
      const row = demandMap.get(r.product) ?? { product: r.product, count: 0, au: 0, usa: 0 };
      const n = Number(r.count);
      row.count += n;
      if (r.market === "AU") row.au += n; else if (r.market === "USA") row.usa += n;
      demandMap.set(r.product, row);
    }
    const productsInDemand = Array.from(demandMap.values()).sort((a, b) => b.count - a.count);

    return Response.json({
      ok: true,
      range: { from: from.toISOString().slice(0, 10), to: toRaw.toISOString().slice(0, 10) },
      emailSends: { total: emailSends.reduce((a, r) => a + r.count, 0), byProduct: emailSends },
      converted: { total: converted.reduce((a, r) => a + r.count, 0), byProduct: converted },
      productsInDemand,
    }, { headers: { "Cache-Control": "no-store" } });
  } catch (error) {
    if (error instanceof Response) return error;
    console.error("[preorder waitlist report] failed", error);
    return Response.json({ ok: false, error: "Could not load reports." }, { status: 500 });
  }
};
