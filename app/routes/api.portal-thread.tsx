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
  const rawMessages = await prisma.portalMessage.findMany({
    where,
    orderBy: { createdAt: "asc" },
  });
  // A single logical note/reply is stored as ONE ROW PER RECIPIENT (so the
  // unread bell stays simple). For the thread view we want each note once, so
  // collapse those per-recipient copies — same parent, author and body — keeping
  // the earliest. Without this, a reply shows once per recipient (the "doubling").
  const seen = new Set<string>();
  const messages = rawMessages.filter((m) => {
    const key = `${m.parentMessageId ?? "T"}|${(m.fromName ?? "").toLowerCase()}|${m.body}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });

  if (url.searchParams.get("markRead") === "1") {
    const currentUserId = await currentUserIdFor(request).catch(() => null);
    if (currentUserId) {
      await prisma.portalMessage.updateMany({
        where: { ...where, userId: currentUserId, readAt: null },
        data: { readAt: new Date() },
      });
    }
  }

  return Response.json({ messages });
};
