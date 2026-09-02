import prisma from "../db.server";
import { normalizePortalMessageUsers, PORTAL_USERS_KEY, type PortalMessageUser } from "../portal-messages.server";

export const PREORDER_PERMISSIONS_KEY = "preorder-permissions-v1";

// Menu/page visibility deliberately does NOT live here. The Production Portal's
// existing user.pageAccess.preorders permission is the single authority for
// whether a staff member can see/open the Pre-orders area. These permissions
// are for sensitive actions inside that area.
export type PreorderPermissionSettings = {
  managePreorderUserIds: string[];
  manageEtaUserIds: string[];
  manageSafetyBufferUserIds: string[];
  sendNotificationUserIds: string[];
  viewReportsUserIds: string[];
};

const EMPTY_PERMISSIONS: PreorderPermissionSettings = {
  managePreorderUserIds: [],
  manageEtaUserIds: [],
  manageSafetyBufferUserIds: [],
  sendNotificationUserIds: [],
  viewReportsUserIds: [],
};

function normalizeIds(value: unknown) {
  if (!Array.isArray(value)) return [];
  return Array.from(new Set(value.map((id) => String(id ?? "").trim()).filter(Boolean)));
}

export function normalizePreorderPermissionSettings(value: unknown): PreorderPermissionSettings {
  if (!value || typeof value !== "object" || Array.isArray(value)) return { ...EMPTY_PERMISSIONS };
  const settings = value as Record<string, unknown>;
  return {
    managePreorderUserIds: normalizeIds(settings.managePreorderUserIds),
    manageEtaUserIds: normalizeIds(settings.manageEtaUserIds),
    manageSafetyBufferUserIds: normalizeIds(settings.manageSafetyBufferUserIds),
    sendNotificationUserIds: normalizeIds(settings.sendNotificationUserIds),
    viewReportsUserIds: normalizeIds(settings.viewReportsUserIds),
  };
}

export async function getPreorderPermissionContext() {
  const [usersSetting, permissionsSetting] = await Promise.all([
    prisma.portalSetting.findUnique({ where: { key: PORTAL_USERS_KEY }, select: { value: true } }),
    prisma.portalSetting.findUnique({ where: { key: PREORDER_PERMISSIONS_KEY }, select: { value: true } }),
  ]);

  const users = normalizePortalMessageUsers(usersSetting?.value).filter((user) => user.active !== false);
  const permissions = normalizePreorderPermissionSettings(permissionsSetting?.value);
  return { users, permissions };
}

export async function setPreorderPermissionSettings(next: unknown, actorName: string) {
  const normalized = normalizePreorderPermissionSettings(next);
  await prisma.portalSetting.upsert({
    where: { key: PREORDER_PERMISSIONS_KEY },
    create: { key: PREORDER_PERMISSIONS_KEY, value: normalized },
    update: { value: normalized },
  });

  await prisma.activityLog.create({
    data: {
      userName: actorName || "Unknown",
      action: "preorder_permissions_updated",
      entity: "preorder_settings",
      field: "permissions",
      toValue: JSON.stringify(normalized),
    },
  }).catch((error) => console.warn("[preorder permissions audit] failed:", error));

  return normalized;
}

function hasPermission(
  user: PortalMessageUser | null | undefined,
  permittedUserIds: string[],
) {
  if (!user || user.active === false) return false;
  // Existing portal admins retain emergency/admin access for sensitive actions.
  if (user.admin === true) return true;
  return permittedUserIds.includes(user.id);
}

export function canManagePreorders(user: PortalMessageUser | null | undefined, permissions: PreorderPermissionSettings) {
  return hasPermission(user, permissions.managePreorderUserIds);
}

export function canManagePreorderEta(user: PortalMessageUser | null | undefined, permissions: PreorderPermissionSettings) {
  return hasPermission(user, permissions.manageEtaUserIds);
}

export function canManageSafetyBuffer(user: PortalMessageUser | null | undefined, permissions: PreorderPermissionSettings) {
  return hasPermission(user, permissions.manageSafetyBufferUserIds);
}

export function canSendPreorderNotifications(user: PortalMessageUser | null | undefined, permissions: PreorderPermissionSettings) {
  return hasPermission(user, permissions.sendNotificationUserIds);
}

export function canViewPreorderReports(user: PortalMessageUser | null | undefined, permissions: PreorderPermissionSettings) {
  return hasPermission(user, permissions.viewReportsUserIds);
}
