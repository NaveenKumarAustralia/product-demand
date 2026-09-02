import type { ActionFunctionArgs } from "react-router";
import prisma from "../db.server";
import { sendBackInStockAvailableEvent } from "../preorder/preorder-klaviyo.server";
import { canSendPreorderNotifications, getPreorderPermissionContext } from "../preorder/preorder-permissions.server";
import { requirePreorderPortalUser, requireSameOrigin } from "../preorder/preorder-portal-auth.server";

type WaitlistRow = {
  id: number;
  shop: string;
  email: string;
  productId: string | null;
  productTitle: string | null;
  variantId: string;
  variantTitle: string | null;
  sku: string | null;
  market: string;
  status: string;
};

export const action = async ({ request }: ActionFunctionArgs) => {
  try {
    requireSameOrigin(request);
    const actor = await requirePreorderPortalUser(request);
    const { permissions } = await getPreorderPermissionContext();
    if (!canSendPreorderNotifications(actor, permissions)) {
      return Response.json({ ok: false, error: "You do not have permission to send customer notifications." }, { status: 403 });
    }

    const body = await request.json().catch(() => ({})) as { id?: unknown; availableQuantity?: unknown; productUrl?: unknown };
    const id = Number(body.id);
    if (!Number.isInteger(id) || id <= 0) return Response.json({ ok: false, error: "Invalid waitlist entry." }, { status: 400 });

    const rows = await prisma.$queryRawUnsafe<WaitlistRow[]>(
      `SELECT "id", "shop", "email", "productId", "productTitle", "variantId", "variantTitle", "sku", "market", "status"
       FROM "PreorderWaitlist" WHERE "id" = $1 LIMIT 1`, id,
    );
    const row = rows[0];
    if (!row) return Response.json({ ok: false, error: "Waitlist entry not found." }, { status: 404 });
    if (row.status !== "waiting") return Response.json({ ok: false, error: "Only waiting customers can be notified." }, { status: 409 });
    if (row.market !== "AU" && row.market !== "USA") return Response.json({ ok: false, error: "Waitlist market is invalid." }, { status: 400 });

    const availableQuantity = Number(body.availableQuantity);
    const productUrl = typeof body.productUrl === "string" && body.productUrl.trim() ? body.productUrl.trim() : null;
    const event = await sendBackInStockAvailableEvent({
      waitlistId: row.id,
      email: row.email,
      shop: row.shop,
      market: row.market,
      productId: row.productId,
      productTitle: row.productTitle,
      variantId: row.variantId,
      variantTitle: row.variantTitle,
      sku: row.sku,
      productUrl,
      availableQuantity: Number.isFinite(availableQuantity) && availableQuantity >= 0 ? availableQuantity : null,
    });

    await prisma.$executeRawUnsafe(
      `UPDATE "PreorderWaitlist" SET "status" = 'notified', "notifiedAt" = CURRENT_TIMESTAMP, "updatedAt" = CURRENT_TIMESTAMP WHERE "id" = $1 AND "status" = 'waiting'`,
      row.id,
    );
    await prisma.activityLog.create({ data: { userName: actor.name, action: "preorder_waitlist_klaviyo_event_sent", entity: "preorder_waitlist", entityId: String(row.id), field: "status", toValue: "notified" } }).catch(() => undefined);

    return Response.json({ ok: true, accepted: event.accepted, metric: event.metric });
  } catch (error) {
    if (error instanceof Response) return error;
    console.error("[preorder klaviyo] waitlist notification failed", error);
    return Response.json({ ok: false, error: error instanceof Error ? error.message : "Could not send Klaviyo notification event." }, { status: 500 });
  }
};
