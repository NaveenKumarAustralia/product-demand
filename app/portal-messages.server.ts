import prisma from "./db.server";

export const PORTAL_USERS_KEY = "supplier-portal-users-v1";

export type PortalMessageUser = {
  id: string;
  name: string;
  admin?: boolean;
  active?: boolean;
};

export function normalizePortalMessageUsers(value: unknown): PortalMessageUser[] {
  if (!Array.isArray(value)) return [];
  return value
    .map((item) => {
      if (!item || typeof item !== "object") return null;
      const user = item as Record<string, unknown>;
      const id = String(user.id ?? "");
      const name = String(user.name ?? "").trim();
      if (!id || !name) return null;
      return {
        id,
        name,
        admin: Boolean(user.admin),
        active: user.active !== false,
      };
    })
    .filter(Boolean) as PortalMessageUser[];
}

function normalizeTag(value: string) {
  return value.toLowerCase().replace(/[^a-z0-9]/g, "");
}

function aliasesForUser(user: PortalMessageUser) {
  const parts = user.name.trim().split(/\s+/).filter(Boolean);
  return new Set([
    normalizeTag(user.name),
    normalizeTag(parts[0] ?? ""),
    normalizeTag(user.name.replace(/\s+/g, "")),
  ].filter(Boolean));
}

export function taggedUsersFromText(text: string | null | undefined, users: PortalMessageUser[]) {
  const tags = new Set(Array.from(text?.matchAll(/@([a-z0-9._-]+)/gi) ?? [], (match) => normalizeTag(match[1])));
  if (!tags.size) return [];

  return users
    .filter((user) => user.active !== false)
    .filter((user) => Array.from(aliasesForUser(user)).some((alias) => tags.has(alias)));
}

export async function syncSampleIterationMessages({
  iterationId,
  sampleName,
  text,
  fromName,
}: {
  iterationId: number;
  sampleName: string;
  text: string | null | undefined;
  fromName?: string | null;
}) {
  const usersSetting = await prisma.portalSetting.findUnique({
    where: { key: PORTAL_USERS_KEY },
    select: { value: true },
  });
  const users = normalizePortalMessageUsers(usersSetting?.value);
  const taggedUsers = taggedUsersFromText(text, users);

  await prisma.portalMessage.deleteMany({
    where: { orderId: iterationId, field: "sample_notes", readAt: null },
  });

  if (!taggedUsers.length) return;

  await prisma.portalMessage.createMany({
    data: taggedUsers.map((user) => ({
      userId: user.id,
      userName: user.name,
      orderId: iterationId,
      field: "sample_notes",
      fromName: fromName || null,
      productTitle: sampleName,
      body: text?.trim() || "",
    })),
  });
}

export async function syncOrderNoteMessages({
  orderId,
  field,
  text,
  fromName,
}: {
  orderId: number;
  field: "factory_notes" | "notes";
  text: string | null | undefined;
  fromName?: string | null;
}) {
  const order = await prisma.supplierOrder.findUnique({
    where: { id: orderId },
    select: { productTitle: true },
  });
  await syncEntityNoteMessages({
    entityType: "supplier_order",
    orderId,
    entityKey: null,
    field,
    text,
    fromName,
    productTitle: order?.productTitle || null,
  });
}

// Generalised sync — used for ANY note cell across the portal.
// Top-level mentions sync the way restock notes always have: delete
// stale unread mentions for this cell, re-create one PortalMessage
// per tagged user. Reply messages go through createReplyMessage()
// below instead so this never wipes them.
export async function syncEntityNoteMessages({
  entityType,
  orderId,
  entityKey,
  field,
  text,
  fromName,
  productTitle,
}: {
  entityType: string;
  orderId: number;
  entityKey: string | null;
  field: string;
  text: string | null | undefined;
  fromName?: string | null;
  productTitle?: string | null;
}) {
  const usersSetting = await prisma.portalSetting.findUnique({
    where: { key: PORTAL_USERS_KEY },
    select: { value: true },
  });
  const users = normalizePortalMessageUsers(usersSetting?.value);
  const taggedUsers = taggedUsersFromText(text, users);

  // Only touch top-level mentions (parentMessageId IS NULL) so
  // existing thread replies aren't wiped when the parent note is
  // edited.
  await prisma.portalMessage.deleteMany({
    where: {
      entityType,
      orderId,
      entityKey: entityKey ?? null,
      field,
      parentMessageId: null,
      readAt: null,
    },
  });

  if (!taggedUsers.length) return;

  await prisma.portalMessage.createMany({
    data: taggedUsers.map((user) => ({
      userId: user.id,
      userName: user.name,
      orderId,
      field,
      entityType,
      entityKey: entityKey ?? null,
      fromName: fromName || null,
      productTitle: productTitle || null,
      body: text?.trim() || "",
    })),
  });
}

// A reply pings the parent message's author + everyone else who's
// posted in the thread + every user @-mentioned in the reply text.
// Each recipient gets their own PortalMessage row so the bell-unread
// query stays simple. Returns the IDs of the created replies.
export async function createReplyMessage({
  parentMessageId,
  fromUserId,
  fromName,
  body,
}: {
  parentMessageId: number;
  fromUserId: string;
  fromName: string;
  body: string;
}): Promise<{ replyIds: number[] }> {
  const parent = await prisma.portalMessage.findUnique({ where: { id: parentMessageId } });
  if (!parent) return { replyIds: [] };

  const usersSetting = await prisma.portalSetting.findUnique({
    where: { key: PORTAL_USERS_KEY },
    select: { value: true },
  });
  const users = normalizePortalMessageUsers(usersSetting?.value);
  const explicitlyTagged = taggedUsersFromText(body, users);

  // Walk the thread to find every participant. Thread = the parent's
  // entityType/orderId/entityKey/field plus every message that ever
  // referenced this parent (via parentMessageId chain).
  const threadKey = {
    entityType: parent.entityType,
    orderId: parent.orderId,
    entityKey: parent.entityKey,
    field: parent.field,
  };
  const allInThread = await prisma.portalMessage.findMany({
    where: threadKey,
    select: { userId: true, fromName: true },
  });
  const participantUserIds = new Set<string>();
  for (const m of allInThread) participantUserIds.add(m.userId);
  // The parent message's fromName needs resolving to a userId via
  // the user list; same for every participant's fromName.
  const allFromNames = new Set<string>();
  for (const m of allInThread) if (m.fromName) allFromNames.add(m.fromName.toLowerCase());
  if (parent.fromName) allFromNames.add(parent.fromName.toLowerCase());
  const lookupByName = new Map<string, PortalMessageUser>();
  for (const u of users) {
    lookupByName.set(u.name.toLowerCase(), u);
    lookupByName.set(u.name.toLowerCase().split(/\s+/)[0] ?? "", u);
  }
  const participantsFromNames = Array.from(allFromNames)
    .map((name) => lookupByName.get(name))
    .filter(Boolean) as PortalMessageUser[];

  // Union: thread participants (by userId), participants resolved
  // from fromName, explicitly @-tagged users. Drop the sender.
  const recipientsByUserId = new Map<string, PortalMessageUser>();
  for (const u of explicitlyTagged) recipientsByUserId.set(u.id, u);
  for (const u of participantsFromNames) recipientsByUserId.set(u.id, u);
  for (const userId of participantUserIds) {
    const found = users.find((u) => u.id === userId);
    if (found) recipientsByUserId.set(found.id, found);
  }
  recipientsByUserId.delete(fromUserId);

  if (!recipientsByUserId.size) return { replyIds: [] };

  const rows = Array.from(recipientsByUserId.values()).map((user) => ({
    userId: user.id,
    userName: user.name,
    orderId: parent.orderId,
    field: parent.field,
    entityType: parent.entityType,
    entityKey: parent.entityKey,
    parentMessageId,
    fromName,
    productTitle: parent.productTitle,
    body: body.trim(),
  }));
  const created = await prisma.$transaction(rows.map((data) => prisma.portalMessage.create({ data, select: { id: true } })));
  return { replyIds: created.map((r) => r.id) };
}

// Edit / delete a message (anyone can — the user said so). Returns
// the updated row count.
export async function editPortalMessage({
  messageId,
  body,
}: {
  messageId: number;
  body: string;
}) {
  return prisma.portalMessage.update({
    where: { id: messageId },
    data: { body: body.trim(), editedAt: new Date() },
  });
}

export async function deletePortalMessage({ messageId }: { messageId: number }) {
  // Deleting a parent cascades to its direct replies for cleanliness.
  await prisma.portalMessage.deleteMany({ where: { parentMessageId: messageId } });
  return prisma.portalMessage.delete({ where: { id: messageId } });
}
