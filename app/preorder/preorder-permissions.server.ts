import prisma from "../db.server";
import { PORTAL_USERS_KEY, type PortalMessageUser } from "../portal-messages.server";

export type PreorderPermissionSettings = {
  managePreorderUserIds: string[];
  manageEtaUserIds: string[];
  manageSafetyBufferUserIds: string[];
  sendNotificationUserIds: string[];
  viewReportsUserIds: string[];
};

export type PreorderPortalUser = PortalMessageUser & {
  preorderAccess?: {
    manage?: boolean;
    eta?: boolean;
    safetyBuffer?: boolean;
    notifications?: boolean;
    reports?: boolean;
  };
};

function activePortalUsers(value: unknown): PreorderPortalUser[] {
  if (!Array.isArray(value)) return [];
  return value.flatMap((item) => {
    if (!item || typeof item !== "object" || Array.isArray(item)) return [];
    const raw = item as Record<string, unknown>;
    const id = String(raw.id ?? "").trim();
    const name = String(raw.name ?? "").trim();
    if (!id || !name || raw.active === false) return [];
    const access = raw.preorderAccess && typeof raw.preorderAccess === "object" && !Array.isArray(raw.preorderAccess)
      ? raw.preorderAccess as Record<string, unknown>
      : {};
    return [{
      id,
      name,
      admin: Boolean(raw.admin) || raw.role === "admin" || raw.role === "superadmin",
      active: true,
      preorderAccess: {
        manage: Boolean(access.manage),
        eta: Boolean(access.eta),
        safetyBuffer: Boolean(access.safetyBuffer),
        notifications: Boolean(access.notifications),
        reports: Boolean(access.reports),
      },
    }];
  });
}

export async function getPreorderPermissionContext() {
  const usersSetting = await prisma.portalSetting.findUnique({
    where: { key: PORTAL_USERS_KEY },
    select: { value: true },
  });
  const users = activePortalUsers(usersSetting?.value);
  const permissions: PreorderPermissionSettings = {
    managePreorderUserIds: users.filter((user) => user.preorderAccess?.manage).map((user) => user.id),
    manageEtaUserIds: users.filter((user) => user.preorderAccess?.eta).map((user) => user.id),
    manageSafetyBufferUserIds: users.filter((user) => user.preorderAccess?.safetyBuffer).map((user) => user.id),
    sendNotificationUserIds: users.filter((user) => user.preorderAccess?.notifications).map((user) => user.id),
    viewReportsUserIds: users.filter((user) => user.preorderAccess?.reports).map((user) => user.id),
  };
  return { users, permissions };
}

function hasPermission(user: PortalMessageUser | null | undefined, permittedUserIds: string[]) {
  if (!user || user.active === false) return false;
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
