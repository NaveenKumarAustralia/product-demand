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

  // Hub notes + per-task comments both live in PortalMessage (so they ride the
  // unread bell). entityType scopes them to this group/sheet; the FIELD splits
  // them: "hub_note" = the Notes tab thread; "task:<taskId>" = a task's comments.
  const scopeEntity = scopeType === "group" ? GROUP_NOTE_ENTITY : SHEET_NOTE_ENTITY;
  const scopeOrderId = scopeType === "collection" ? Number(scopeId) || 0 : 0;
  const scopeEntityKey = scopeType === "group" ? scopeId : null;

  // One logical message per (parent, author, body, createdAt) — collapse the
  // per-recipient fan-out copies, keep the earliest.
  const collapse = (rows: Array<{ id: number; body: string; fromName: string | null; createdAt: Date; editedAt: Date | null; parentMessageId: number | null; field: string }>) => {
    const seen = new Set<string>();
    return rows.filter((m) => {
      const k = `${m.parentMessageId ?? "T"}|${(m.fromName ?? "").toLowerCase()}|${m.body}|${new Date(m.createdAt).getTime()}`;
      if (seen.has(k)) return false;
      seen.add(k);
      return true;
    });
  };

  const rawNotes = await prisma.portalMessage.findMany({
    where: { entityType: scopeEntity, orderId: scopeOrderId, entityKey: scopeEntityKey, field: COLLECTION_HUB_NOTE_FIELD },
    orderBy: { createdAt: "asc" },
  });
  const comments = collapse(rawNotes).map((m) => ({ id: m.id, body: m.body, fromName: m.fromName, createdAt: m.createdAt, editedAt: m.editedAt, parentMessageId: m.parentMessageId }));

  // Task comments: field "task:<taskId>" under the collection_task entity.
  const rawTaskComments = await prisma.portalMessage.findMany({
    where: { entityType: "collection_task", entityKey: `${scopeType}:${scopeId}`, field: { startsWith: "task:" } },
    orderBy: { createdAt: "asc" },
  });
  const taskComments: Record<string, Array<{ id: number; body: string; fromName: string | null; createdAt: Date; editedAt: Date | null }>> = {};
  for (const m of collapse(rawTaskComments)) {
    const taskId = m.field.slice("task:".length);
    (taskComments[taskId] ??= []).push({ id: m.id, body: m.body, fromName: m.fromName, createdAt: m.createdAt, editedAt: m.editedAt });
  }

  if (url.searchParams.get("markNotesRead") === "1") {
    const currentUserId = await currentUserIdFor(request).catch(() => null);
    if (currentUserId) {
      // Clear this user's unread for the whole hub (notes thread + task comments).
      await prisma.portalMessage.updateMany({
        where: { userId: currentUserId, readAt: null, OR: [
          { entityType: scopeEntity, orderId: scopeOrderId, entityKey: scopeEntityKey, field: COLLECTION_HUB_NOTE_FIELD },
          { entityType: "collection_task", entityKey: `${scopeType}:${scopeId}`, field: { startsWith: "task:" } },
        ] },
        data: { readAt: new Date() },
      });
    }
  }

  return Response.json({ tasks, comments, taskComments });
};
