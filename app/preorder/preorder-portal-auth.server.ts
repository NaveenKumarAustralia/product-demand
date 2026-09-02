import prisma from "../db.server";
import {
  normalizePortalMessageUsers,
  PORTAL_USERS_KEY,
  type PortalMessageUser,
} from "../portal-messages.server";

export const PORTAL_USER_COOKIE = "portal_user_id";

function getCookieValue(request: Request, name: string): string | null {
  const cookie = request.headers.get("cookie") ?? "";
  for (const part of cookie.split(";")) {
    const [key, ...rest] = part.trim().split("=");
    if (key === name) return rest.join("=");
  }
  return null;
}

export async function getCurrentPreorderPortalUser(request: Request): Promise<PortalMessageUser | null> {
  const raw = getCookieValue(request, PORTAL_USER_COOKIE) ?? "";
  const userId = decodeURIComponent(raw).trim();
  if (!userId) return null;

  const setting = await prisma.portalSetting.findUnique({
    where: { key: PORTAL_USERS_KEY },
    select: { value: true },
  });
  const users = normalizePortalMessageUsers(setting?.value);
  return users.find((user) => user.id === userId && user.active !== false) ?? null;
}

export async function requirePreorderPortalUser(request: Request): Promise<PortalMessageUser> {
  const user = await getCurrentPreorderPortalUser(request);
  if (!user) throw new Response("Unauthorized", { status: 401 });
  return user;
}

export function requireSameOrigin(request: Request) {
  const origin = request.headers.get("origin");
  if (!origin) return;
  const requestOrigin = new URL(request.url).origin;
  if (origin !== requestOrigin) {
    throw new Response("Invalid origin", { status: 403 });
  }
}
