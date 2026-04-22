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
  const usersSetting = await prisma.portalSetting.findUnique({
    where: { key: PORTAL_USERS_KEY },
    select: { value: true },
  });
  const users = normalizePortalMessageUsers(usersSetting?.value);
  const taggedUsers = taggedUsersFromText(text, users);

  await prisma.portalMessage.deleteMany({
    where: { orderId, field, readAt: null },
  });

  if (!taggedUsers.length) return;

  const order = await prisma.supplierOrder.findUnique({
    where: { id: orderId },
    select: { productTitle: true },
  });

  await prisma.portalMessage.createMany({
    data: taggedUsers.map((user) => ({
      userId: user.id,
      userName: user.name,
      orderId,
      field,
      fromName: fromName || null,
      productTitle: order?.productTitle || null,
      body: text?.trim() || "",
    })),
  });
}
