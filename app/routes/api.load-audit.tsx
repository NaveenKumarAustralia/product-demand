import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// TEMP one-off: for a packing list (invoiceNumber substring), return its load
// timestamp + who (if recorded), and the activity log within ±windowMin minutes
// of that load time so we can see who was active around the real load moment.
//   /api/load-audit?q=AEPL-13&windowMin=45
// Remove after use.
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const url = new URL(request.url);
  const q = (url.searchParams.get("q") ?? "AEPL-13").trim();
  const windowMin = Number(url.searchParams.get("windowMin") ?? "45") || 45;

  const list = await prisma.packingList.findFirst({
    where: { invoiceNumber: { contains: q } },
    select: { id: true, invoiceNumber: true, title: true, status: true, masterInventoryLoadedAt: true, loadedBy: true, createdAt: true, updatedAt: true },
  }).catch(() => null);
  if (!list) return Response.json({ ok: false, error: `no packing list matching "${q}"` });

  const loadedAt = list.masterInventoryLoadedAt;
  let around: unknown[] = [];
  if (loadedAt) {
    const from = new Date(loadedAt.getTime() - windowMin * 60_000);
    const to = new Date(loadedAt.getTime() + windowMin * 60_000);
    const rows = await prisma.activityLog.findMany({
      where: { createdAt: { gte: from, lte: to } },
      orderBy: { createdAt: "asc" },
      select: { userName: true, action: true, entity: true, entityName: true, field: true, toValue: true, createdAt: true },
    }).catch(() => []);
    around = rows.map((r) => ({ ...r, createdAt: r.createdAt.toISOString() }));
  }

  return Response.json({
    ok: true,
    packingList: {
      id: list.id,
      invoiceNumber: list.invoiceNumber,
      status: list.status,
      masterInventoryLoadedAt: loadedAt ? loadedAt.toISOString() : null,
      loadedBy: list.loadedBy ?? null,
      createdAt: list.createdAt.toISOString(),
      updatedAt: list.updatedAt.toISOString(),
    },
    activityAroundLoad: around,
  });
};
