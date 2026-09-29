import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

// Per-USER column layout (order + hidden) for the Pre-orders → Customer Orders
// table, so a staff member's chosen columns follow their account across
// browsers/devices. Keyed by the portal user id.
const keyFor = (userId: string) => `preorder-order-cols:${userId}`;

function readPrefs(value: unknown): { order: string[]; hidden: string[] } | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const v = value as { order?: unknown; hidden?: unknown };
  return {
    order: Array.isArray(v.order) ? v.order.map(String) : [],
    hidden: Array.isArray(v.hidden) ? v.hidden.map(String) : [],
  };
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  const s = await prisma.portalSetting.findUnique({ where: { key: keyFor(actor.id) }, select: { value: true } }).catch(() => null);
  return Response.json({ ok: true, columns: readPrefs(s?.value) }, { headers: { "Cache-Control": "no-store" } });
};

export const action = async ({ request }: ActionFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  let body: { order?: unknown; hidden?: unknown } = {};
  try { body = await request.json(); } catch { body = {}; }
  const order = Array.isArray(body.order) ? body.order.map(String).slice(0, 40) : [];
  const hidden = Array.isArray(body.hidden) ? body.hidden.map(String).slice(0, 40) : [];
  await prisma.portalSetting.upsert({
    where: { key: keyFor(actor.id) },
    create: { key: keyFor(actor.id), value: { order, hidden } },
    update: { value: { order, hidden } },
  }).catch(() => {});
  return Response.json({ ok: true }, { headers: { "Cache-Control": "no-store" } });
};
