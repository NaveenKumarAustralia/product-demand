import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { PORTAL_USERS_KEY, normalizePortalMessageUsers } from "../portal-messages.server";

// Threads a user has marked "handled". Single settings row: { [threadKey]: {by, at} }.
export const THREADS_RESOLVED_KEY = "portal-threads-resolved-v1";

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

// Returns the full thread for a cell — top-level mention + every
// reply — ordered chronologically. The thread side panel calls this
// to populate the message list.
//
// Query params:
//   thread (required): "<entityType>:<entityId>:<entityKey>:<field>"
//   markRead=1: also mark every unread message in this thread for
//     the current user as read.
//
// Response: { messages: PortalMessage[] }
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const url = new URL(request.url);
  const thread = url.searchParams.get("thread")?.trim();
  if (!thread) return Response.json({ messages: [] });
  const parts = thread.split(":");
  if (parts.length < 4) return Response.json({ messages: [] });
  const [entityType, entityIdRaw, entityKey, ...fieldParts] = parts;
  const field = fieldParts.join(":");
  const entityId = Number(entityIdRaw);
  if (!entityType || !field || !Number.isFinite(entityId)) return Response.json({ messages: [] });

  const where = {
    entityType,
    orderId: entityId,
    entityKey: entityKey ? entityKey : null,
    field,
  };
  const currentUserId = await currentUserIdFor(request).catch(() => null);
  const rawMessages = await prisma.portalMessage.findMany({
    where,
    orderBy: { createdAt: "asc" },
  });
  // A single logical note/reply is stored as ONE ROW PER RECIPIENT (so the
  // unread bell stays simple). For the thread view we want each note once, so
  // GROUP those per-recipient copies — same parent, author and body — keeping
  // the earliest as the representative. The copies also tell us, per logical
  // message: who has READ it (seenBy = recipients with readAt set, for the
  // "Seen by" receipt) and whether it is still UNREAD FOR ME (the current
  // user's own copy has readAt === null — drives the "New" divider). This is
  // computed BEFORE the markRead update below so the boundary survives opening.
  const groups = new Map<string, typeof rawMessages>();
  const order: string[] = [];
  for (const m of rawMessages) {
    const key = `${m.parentMessageId ?? "T"}|${(m.fromName ?? "").toLowerCase()}|${m.body}`;
    if (!groups.has(key)) { groups.set(key, []); order.push(key); }
    groups.get(key)!.push(m);
  }
  const messages = order.map((key) => {
    const copies = groups.get(key)!;
    const rep = copies[0];
    const author = (rep.fromName ?? rep.userName ?? "").toLowerCase();
    const seenBy = new Set<string>();
    let unreadForMe = false;
    for (const c of copies) {
      if (currentUserId && c.userId === currentUserId && c.readAt === null) unreadForMe = true;
      if (c.readAt && (c.userName ?? "").trim() && (c.userName ?? "").toLowerCase() !== author) seenBy.add(c.userName);
    }
    return { ...rep, seenBy: Array.from(seenBy), unreadForMe };
  });

  // Thread-level "resolved" flag — stored in a single settings map keyed by the
  // full thread string, so no per-thread rows or schema change.
  let resolved = false;
  let resolvedBy: string | null = null;
  let resolvedAt: string | null = null;
  try {
    const setting = await prisma.portalSetting.findUnique({ where: { key: THREADS_RESOLVED_KEY }, select: { value: true } });
    const map = (setting?.value && typeof setting.value === "object" && !Array.isArray(setting.value))
      ? (setting.value as Record<string, { by?: string; at?: string }>) : {};
    const r = map[thread];
    if (r) { resolved = true; resolvedBy = r.by ?? null; resolvedAt = r.at ?? null; }
  } catch { /* best-effort */ }

  if (url.searchParams.get("markRead") === "1" && currentUserId) {
    await prisma.portalMessage.updateMany({
      where: { ...where, userId: currentUserId, readAt: null },
      data: { readAt: new Date() },
    });
  }

  return Response.json({ messages, resolved, resolvedBy, resolvedAt });
};
