import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// TEMP one-off audit: list activity-log entries in a UTC time window (and/or for
// an entityName substring), plus the distinct users active in that window.
//   /api/activity-audit?from=2026-08-13T05:00:00Z&to=2026-08-13T07:30:00Z
//   &q=AEPL-13
// Remove this route after use.
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const url = new URL(request.url);
  const from = url.searchParams.get("from");
  const to = url.searchParams.get("to");
  const q = (url.searchParams.get("q") ?? "").trim();
  if (!from || !to) return Response.json({ ok: false, error: "pass ?from=<ISO>&to=<ISO>" });

  const where: Record<string, unknown> = { createdAt: { gte: new Date(from), lte: new Date(to) } };
  const rows = await prisma.activityLog.findMany({
    where,
    orderBy: { createdAt: "asc" },
    select: { userName: true, action: true, entity: true, entityName: true, field: true, toValue: true, createdAt: true },
  }).catch((e) => { console.warn(e); return []; });

  const inWindow = rows.map((r) => ({ ...r, createdAt: r.createdAt.toISOString() }));
  const activeUsers = Array.from(new Set(inWindow.map((r) => r.userName))).sort();

  // Anything referencing the invoice substring, over a wider net (any time), to
  // catch edits to that specific packing list.
  const byInvoice = q
    ? (await prisma.activityLog.findMany({
        where: { entityName: { contains: q } },
        orderBy: { createdAt: "desc" },
        take: 40,
        select: { userName: true, action: true, entity: true, entityName: true, field: true, toValue: true, createdAt: true },
      }).catch(() => [])).map((r) => ({ ...r, createdAt: r.createdAt.toISOString() }))
    : [];

  return Response.json({ ok: true, window: { from, to }, activeUsers, count: inWindow.length, inWindow, byInvoice });
};
