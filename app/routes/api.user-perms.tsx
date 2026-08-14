import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// TEMP one-off: list portal users' load-inventory permission (NO password
// hashes). Remove after use.  /api/user-perms
const PORTAL_USERS_KEY = "supplier-portal-users-v1";

export const loader = async (_args: LoaderFunctionArgs) => {
  const setting = await prisma.portalSetting.findUnique({ where: { key: PORTAL_USERS_KEY }, select: { value: true } }).catch(() => null);
  const raw = Array.isArray(setting?.value) ? setting!.value as Array<Record<string, unknown>> : [];
  const users = raw.map((u) => ({
    name: String(u.name ?? ""),
    role: String(u.role ?? ""),
    admin: Boolean(u.admin),
    active: u.active === undefined ? true : Boolean(u.active),
    canLoadInventory: Boolean(u.canLoadInventory),
  }));
  return Response.json({ ok: true, users });
};
