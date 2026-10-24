import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

// Admin READ-ONLY. Lists EVERY collection (tile) — visible and hidden — with its
// row count and which group (folder) it's mapped to, so we can confirm no
// collection was lost and locate ones that seem "missing" (usually hidden).
//   GET /api/collections-audit            → all collections
//   GET /api/collections-audit?group=2027 → only ones whose group name matches
const COLLECTION_GROUPS_KEY = "collections-groups-v1";
const COLLECTION_HIDDEN_KEY = "collections-hidden-v1";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });
  const q = (new URL(request.url).searchParams.get("group") || "").trim().toLowerCase();

  const [cols, groupSetting, hiddenSetting] = await Promise.all([
    prisma.collection.findMany({ select: { id: true, name: true, kind: true, rows: true, updatedAt: true }, orderBy: { id: "asc" } }),
    prisma.portalSetting.findUnique({ where: { key: COLLECTION_GROUPS_KEY }, select: { value: true } }),
    prisma.portalSetting.findUnique({ where: { key: COLLECTION_HIDDEN_KEY }, select: { value: true } }),
  ]);
  // Hidden is a SEPARATE setting (a list of ids) — collections are never deleted.
  const hiddenIds = new Set<number>((Array.isArray(hiddenSetting?.value) ? hiddenSetting!.value as unknown[] : []).map((n) => Number(n)).filter((n) => Number.isInteger(n)));

  // Map collectionId → group name from the saved groups setting.
  const rawGroups = (() => {
    const v = groupSetting?.value as { groups?: unknown } | unknown[] | null | undefined;
    if (v && typeof v === "object" && Array.isArray((v as { groups?: unknown }).groups)) return (v as { groups: unknown[] }).groups;
    return Array.isArray(v) ? v : [];
  })();
  const groupOfCollection = new Map<number, string>();
  const groupById = new Map<string, { name: string; parentId?: string }>();
  for (const g of rawGroups as Array<Record<string, unknown>>) {
    const id = String(g?.id ?? ""); const name = String(g?.name ?? "");
    if (id) groupById.set(id, { name, parentId: g?.parentId ? String(g.parentId) : undefined });
  }
  const pathOf = (gid: string): string => {
    const parts: string[] = []; const seen = new Set<string>(); let cur: string | undefined = gid;
    while (cur && groupById.has(cur) && !seen.has(cur)) { seen.add(cur); const node: { name: string; parentId?: string } = groupById.get(cur)!; parts.unshift(node.name); cur = node.parentId; }
    return parts.join(" › ");
  };
  for (const g of rawGroups as Array<Record<string, unknown>>) {
    const id = String(g?.id ?? "");
    const ids = Array.isArray(g?.collectionIds) ? (g!.collectionIds as unknown[]) : [];
    for (const n of ids) { const num = Number(n); if (Number.isInteger(num)) groupOfCollection.set(num, pathOf(id)); }
  }

  const rowCount = (rows: unknown) => (Array.isArray(rows) ? rows.filter((r) => r && typeof r === "object").length : 0);
  let list = cols.map((c) => {
    const hidden = hiddenIds.has(c.id);
    return {
      id: c.id,
      name: c.name,
      kind: c.kind,
      hidden,
      rows: rowCount(c.rows),
      group: groupOfCollection.get(c.id) ?? (hidden ? "(hidden, ungrouped)" : "(ungrouped)"),
      updatedAt: c.updatedAt,
    };
  });
  if (q) list = list.filter((c) => (c.group || "").toLowerCase().includes(q) || c.name.toLowerCase().includes(q));

  return Response.json(
    {
      ok: true,
      totalCollections: cols.length,
      visible: cols.filter((c) => !hiddenIds.has(c.id)).length,
      hidden: hiddenIds.size,
      note: "Every collection still exists here. 'hidden:true' ones are hidden (not deleted) — open 'Show hidden collections' in the portal to see and unhide them. 'group' shows the folder path (or ungrouped).",
      collections: list,
    },
    { headers: { "Cache-Control": "no-store" } },
  );
};
