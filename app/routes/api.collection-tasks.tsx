import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { PORTAL_USERS_KEY, normalizePortalMessageUsers } from "../portal-messages.server";
import {
  readAllCollectionTasks, tasksForScope,
  COLLECTION_HUB_NOTE_FIELD, GROUP_NOTE_ENTITY, SHEET_NOTE_ENTITY,
  type CollectionScopeType,
} from "../collection-tasks.server";

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
  const setting = await prisma.portalSetting.findUnique({ where: { key: PORTAL_USERS_KEY }, select: { value: true } });
  const users = normalizePortalMessageUsers(setting?.value);
  return users.find((u) => u.id === userId && u.active !== false)?.id ?? null;
}

// Tasks + hub notes for ONE collection group or sheet.
//   ?scopeType=group|collection & scopeId=<id>
//   &markNotesRead=1  → clear the current user's unread hub-note bell for this scope
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const url = new URL(request.url);
  const scopeType = (url.searchParams.get("scopeType") === "group" ? "group" : "collection") as CollectionScopeType;
  const scopeId = (url.searchParams.get("scopeId") ?? "").trim();
  if (!scopeId) return Response.json({ tasks: [], comments: [] });

  const allTasks = await readAllCollectionTasks();
  const tasks = tasksForScope(allTasks, scopeType, scopeId)
    // open first (by due), then done (most-recently-done first)
    .sort((a, b) => {
      if (a.done !== b.done) return a.done ? 1 : -1;
      if (!a.done) return (a.due ?? "9999-99-99").localeCompare(b.due ?? "9999-99-99") || a.createdAt.localeCompare(b.createdAt);
      return (b.doneAt ?? "").localeCompare(a.doneAt ?? "");
    });

  // Hub notes (flat comments) via PortalMessage, one logical note per
  // (author, body, createdAt) — collapse the per-recipient fan-out copies.
  const noteWhere = {
    entityType: scopeType === "group" ? GROUP_NOTE_ENTITY : SHEET_NOTE_ENTITY,
    orderId: scopeType === "collection" ? Number(scopeId) || 0 : 0,
    entityKey: scopeType === "group" ? scopeId : null,
    field: COLLECTION_HUB_NOTE_FIELD,
  };
  const rawNotes = await prisma.portalMessage.findMany({ where: noteWhere, orderBy: { createdAt: "asc" } });
  const seen = new Set<string>();
  const comments = rawNotes.filter((m) => {
    const key = `${(m.fromName ?? "").toLowerCase()}|${m.body}|${new Date(m.createdAt).getTime()}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  }).map((m) => ({ id: m.id, body: m.body, fromName: m.fromName, createdAt: m.createdAt, editedAt: m.editedAt }));

  if (url.searchParams.get("markNotesRead") === "1") {
    const currentUserId = await currentUserIdFor(request).catch(() => null);
    if (currentUserId) {
      await prisma.portalMessage.updateMany({
        where: { ...noteWhere, userId: currentUserId, readAt: null },
        data: { readAt: new Date() },
      });
    }
  }

  return Response.json({ tasks, comments });
};
