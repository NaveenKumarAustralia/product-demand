import prisma from "./db.server";

// Tasks + hub notes that hang off a WHOLE collection group or a single sheet
// (collection) — not a cell/row. Tasks are a simple checklist: text, an
// assignee, a due date, done/undone. Stored as one settings row (like the
// collection groups themselves) so there's no schema change.
export const COLLECTION_TASKS_KEY = "collections-tasks-v1";
// Group/sheet-level notes reuse PortalMessage (so they ride the same unread
// bell) with these entityTypes + a fixed field.
export const COLLECTION_HUB_NOTE_FIELD = "hub_note";
export const GROUP_NOTE_ENTITY = "collection_group";
export const SHEET_NOTE_ENTITY = "collection_sheet";

export type CollectionScopeType = "group" | "collection";

export type CollectionTask = {
  id: string;
  scopeType: CollectionScopeType;
  scopeId: string; // group id (string) or collection id (numeric-as-string)
  scopeName: string | null; // human label for the group/sheet, for the bell + "my tasks"
  text: string;
  assigneeId: string | null;
  assigneeName: string | null;
  due: string | null; // YYYY-MM-DD
  done: boolean;
  doneAt: string | null;
  createdBy: string | null;
  createdAt: string;
};

export function scopeKey(scopeType: CollectionScopeType, scopeId: string): string {
  return `${scopeType}:${scopeId}`;
}

export function normalizeCollectionTasks(value: unknown): CollectionTask[] {
  const arr = value && typeof value === "object" && Array.isArray((value as { tasks?: unknown }).tasks)
    ? (value as { tasks: unknown[] }).tasks
    : Array.isArray(value) ? value : [];
  const out: CollectionTask[] = [];
  for (const raw of arr) {
    if (!raw || typeof raw !== "object") continue;
    const t = raw as Record<string, unknown>;
    const id = String(t.id ?? "").trim();
    const scopeType = t.scopeType === "group" ? "group" : t.scopeType === "collection" ? "collection" : null;
    const scopeId = String(t.scopeId ?? "").trim();
    const text = String(t.text ?? "").trim();
    if (!id || !scopeType || !scopeId || !text) continue;
    const dueRaw = String(t.due ?? "").trim();
    const due = /^\d{4}-\d{2}-\d{2}$/.test(dueRaw) ? dueRaw : null;
    out.push({
      id,
      scopeType,
      scopeId,
      scopeName: t.scopeName != null ? String(t.scopeName) : null,
      text,
      assigneeId: t.assigneeId ? String(t.assigneeId) : null,
      assigneeName: t.assigneeName ? String(t.assigneeName) : null,
      due,
      done: t.done === true,
      doneAt: t.doneAt ? String(t.doneAt) : null,
      createdBy: t.createdBy ? String(t.createdBy) : null,
      createdAt: t.createdAt ? String(t.createdAt) : new Date().toISOString(),
    });
  }
  return out;
}

export async function readAllCollectionTasks(): Promise<CollectionTask[]> {
  const row = await prisma.portalSetting.findUnique({ where: { key: COLLECTION_TASKS_KEY }, select: { value: true } }).catch(() => null);
  return normalizeCollectionTasks(row?.value);
}

export async function writeAllCollectionTasks(tasks: CollectionTask[]): Promise<void> {
  await prisma.portalSetting.upsert({
    where: { key: COLLECTION_TASKS_KEY },
    create: { key: COLLECTION_TASKS_KEY, value: { tasks } },
    update: { value: { tasks } },
  });
}

export function tasksForScope(tasks: CollectionTask[], scopeType: CollectionScopeType, scopeId: string): CollectionTask[] {
  return tasks.filter((t) => t.scopeType === scopeType && t.scopeId === scopeId);
}

// Map of scopeKey -> number of OPEN (not done) tasks, for tile badges.
export function openTaskCountsByScope(tasks: CollectionTask[]): Record<string, number> {
  const m: Record<string, number> = {};
  for (const t of tasks) {
    if (t.done) continue;
    const k = scopeKey(t.scopeType, t.scopeId);
    m[k] = (m[k] ?? 0) + 1;
  }
  return m;
}
