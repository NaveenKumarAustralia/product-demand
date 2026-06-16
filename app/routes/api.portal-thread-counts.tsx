import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { PORTAL_USERS_KEY, normalizePortalMessageUsers } from "../portal-messages.server";

const PORTAL_USER_COOKIE = "portal_user_id";
function getCookieValue(request: Request, name: string): string | null {
  const cookie = request.headers.get("cookie") ?? "";
  for (const part of cookie.split(";")) {
    const [k, ...rest] = part.trim().split("=");
    if (k === name) return rest.join("=");
  }
  return null;
}
async function currentUserIdFor(request: Request): Promise<string | null> {
  const userId = decodeURIComponent(getCookieValue(request, PORTAL_USER_COOKIE) ?? "");
  if (!userId) return null;
  const setting = await prisma.portalSetting.findUnique({
    where: { key: PORTAL_USERS_KEY },
    select: { value: true },
  });
  const users = normalizePortalMessageUsers(setting?.value);
  const found = users.find((u) => u.id === userId && u.active !== false);
  return found?.id ?? null;
}

// Returns total + per-user-unread message counts grouped by thread
// key for a given entityType (and optionally entityId). Used by the
// per-cell 💬 badge across the portal pages.
//
// Query params:
//   entityType (required): "supplier_order" | "collection_row" | etc.
//   entityId (optional): restrict to one parent record
//
// Response: { counts: Array<{ key, total, unread }> }
// key = `<entityType>:<orderId>:<entityKey || "">:<field>`
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const url = new URL(request.url);
  const entityType = url.searchParams.get("entityType")?.trim();
  if (!entityType) return Response.json({ counts: [] });
  const entityIdRaw = url.searchParams.get("entityId");
  const entityId = entityIdRaw ? Number(entityIdRaw) : null;

  const where: Record<string, unknown> = { entityType };
  if (Number.isFinite(entityId)) where.orderId = entityId;

  const totals = await prisma.portalMessage.groupBy({
    by: ["orderId", "entityKey", "field"],
    where,
    _count: { _all: true },
  });

  // Pull the current user's unread tally per group so the badge can
  // show "3 / 5" (3 unread of 5 total) inline.
  let unreadMap = new Map<string, number>();
  const currentUserId = await currentUserIdFor(request).catch(() => null);
  if (currentUserId) {
    const unread = await prisma.portalMessage.groupBy({
      by: ["orderId", "entityKey", "field"],
      where: { ...where, userId: currentUserId, readAt: null },
      _count: { _all: true },
    });
    unreadMap = new Map(unread.map((u) => [`${u.orderId}:${u.entityKey ?? ""}:${u.field}`, u._count._all]));
  }

  const counts = totals.map((t) => {
    const localKey = `${t.orderId}:${t.entityKey ?? ""}:${t.field}`;
    return {
      key: `${entityType}:${localKey}`,
      total: t._count._all,
      unread: unreadMap.get(localKey) ?? 0,
    };
  });
  return Response.json({ counts });
};
