import prisma from "../db.server";
import { normalizePortalMessageUsers, PORTAL_USERS_KEY, type PortalMessageUser } from "../portal-messages.server";

export const PREORDER_PERMISSIONS_KEY = "preorder-permissions-v1";

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

export function canManagePreorders(user: PortalMessageUser | null | undefined, permissions: PreorderPermissionSettings) {
  if (!user || user.active === false) return false;
  // Existing portal admins retain emergency/admin access.
  if (user.admin === true) return true;
  return permissions.managePreorderUserIds.includes(user.id);
}

export function canManagePreorderEta(user: PortalMessageUser | null | undefined, permissions: PreorderPermissionSettings) {
  if (!user || user.active === false) return false;
  if (user.admin === true) return true;
  return permissions.manageEtaUserIds.includes(user.id);
}

export function canManageSafetyBuffer(user: PortalMessageUser | null | undefined, permissions: PreorderPermissionSettings) {
  if (!user || user.active === false) return false;
  if (user.admin === true) return true;
  return permissions.manageSafetyBufferUserIds.includes(user.id);
}

export function canSendPreorderNotifications(user: PortalMessageUser | null | undefined, permissions: PreorderPermissionSettings) {
  if (!user || user.active === false) return false;
  if (user.admin === true) return true;
  return permissions.sendNotificationUserIds.includes(user.id);
}

export function canViewPreorderReports(user: PortalMessageUser | null | undefined, permissions: PreorderPermissionSettings) {
  if (!user || user.active === false) return false;
  if (user.admin === true) return true;
  return permissions.viewReportsUserIds.includes(user.id);
}
