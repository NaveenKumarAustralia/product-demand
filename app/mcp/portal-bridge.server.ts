/**
 * Runs Production Portal operations through the portal's OWN action handler,
 * in-process, so every write gets exactly the same validation and side effects
 * as a click in the portal: fabric consumption, the production cut log,
 * pre-order alerts, note threads and the portal's activity log.
 *
 * The calls act as a dedicated portal user, "ChatGPT", which cannot log in to
 * the portal (it has no password). That keeps ChatGPT's changes clearly
 * labelled in the portal's own history. Who authorised each change is recorded
 * separately in McpAuditLog.
 */
import prisma from "../db.server";
import { PORTAL_USERS_KEY } from "../portal-messages.server";
import { PORTAL_USER_COOKIE } from "../preorder/preorder-portal-auth.server";
import { action as portalAction } from "../routes/portal._index";
import { loader as reorderFabricLoader } from "../routes/api.reorder-fabric";

export const MCP_PORTAL_USER_ID = "mcp-chatgpt";
export const MCP_PORTAL_USER_NAME = "ChatGPT";

let ensured = false;

/** Make sure the ChatGPT portal user exists (admin role, no password, so no UI login). */
export async function ensureMcpPortalUser(): Promise<void> {
  if (ensured) return;
  await prisma.$transaction(async (tx) => {
    // Lock the users row so a concurrent portal save can't be overwritten.
    await tx.$executeRaw`SELECT 1 FROM "PortalSetting" WHERE "key" = ${PORTAL_USERS_KEY} FOR UPDATE`;
    const row = await tx.portalSetting.findUnique({ where: { key: PORTAL_USERS_KEY }, select: { value: true } });
    const users = Array.isArray(row?.value) ? (row!.value as Array<Record<string, unknown>>) : [];
    if (users.some((u) => u?.id === MCP_PORTAL_USER_ID)) return;
    // Never create the users list from scratch here: the portal seeds its own
    // superadmin on first load, and this must not pre-empt that.
    if (!row) throw new Error("Portal users are not set up yet; open the portal once first");
    const next = [
      ...users,
      {
        id: MCP_PORTAL_USER_ID,
        name: MCP_PORTAL_USER_NAME,
        username: MCP_PORTAL_USER_NAME.toLowerCase(),
        passwordHash: "",
        role: "admin",
        admin: true,
        active: true,
        canLoadInventory: false,
        pageAccess: {},
        preorderAccess: { manage: false, eta: false, safetyBuffer: false, notifications: false, reports: false },
      },
    ];
    await tx.portalSetting.update({ where: { key: PORTAL_USERS_KEY }, data: { value: next as never } });
  });
  ensured = true;
}

/**
 * Invoke one portal action intent as the ChatGPT user. Returns the parsed JSON
 * body when the portal answers with JSON, or null (the portal answers many
 * edits with null; callers verify by re-reading the database).
 */
export async function callPortalIntent(intent: string, fields: Record<string, string | number | null | undefined>): Promise<unknown> {
  await ensureMcpPortalUser();
  const form = new FormData();
  form.set("intent", intent);
  for (const [k, v] of Object.entries(fields)) {
    if (v !== undefined && v !== null) form.set(k, String(v));
  }
  const request = new Request("http://portal.internal/portal", {
    method: "POST",
    body: form,
    headers: { cookie: `${PORTAL_USER_COOKIE}=${encodeURIComponent(MCP_PORTAL_USER_ID)}` },
  });
  const result = await portalAction({ request, params: {}, context: {} } as Parameters<typeof portalAction>[0]);
  if (result instanceof Response) {
    const type = result.headers.get("content-type") ?? "";
    if (type.includes("application/json")) return result.json();
    return null;
  }
  return result ?? null;
}

/** Read the fabric planner (same data the portal's fabric picker uses). */
export async function readFabric(params: { title?: string; pending?: boolean }): Promise<unknown> {
  const url = new URL("http://portal.internal/api/reorder-fabric");
  if (params.title) url.searchParams.set("title", params.title);
  if (params.pending) url.searchParams.set("pending", "1");
  const res = await reorderFabricLoader({ request: new Request(url), params: {}, context: {} } as Parameters<typeof reorderFabricLoader>[0]);
  return res instanceof Response ? res.json() : res;
}
