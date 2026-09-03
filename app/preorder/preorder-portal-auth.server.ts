import prisma from "../db.server";
import {
  normalizePortalMessageUsers,
  PORTAL_USERS_KEY,
  type PortalMessageUser,
} from "../portal-messages.server";

// MUST match the cookie the main Production Portal login sets
// (portal._index.tsx → PORTAL_USER_COOKIE = "supplier_portal_user"). It was
// "portal_user_id", which is set nowhere, so every /api/preorder-* call 401'd
// for the logged-in portal user (readiness panel showed "Network error", and
// all preorder admin actions silently failed auth).
export const PORTAL_USER_COOKIE = "supplier_portal_user";

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
  if (!origin) return; // non-browser / same-origin navigations omit Origin
  // Compare HOST only, honouring the proxy's forwarded host. Behind Railway,
  // TLS terminates at the edge so request.url is often http:// with an internal
  // host, which made a full-origin (protocol+host) comparison wrongly 403 every
  // POST (e.g. "could not save preorder location"). Host-matching still blocks
  // genuine cross-site POSTs.
  let originHost: string;
  try {
    originHost = new URL(origin).host;
  } catch {
    throw new Response("Invalid origin", { status: 403 });
  }
  const host = request.headers.get("x-forwarded-host")?.split(",")[0]?.trim()
    || request.headers.get("host")
    || new URL(request.url).host;
  if (originHost !== host) {
    throw new Response("Invalid origin", { status: 403 });
  }
}
