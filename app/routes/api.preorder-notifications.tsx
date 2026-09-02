import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import {
  canSendPreorderNotifications,
  getPreorderPermissionContext,
} from "../preorder/preorder-permissions.server";
import {
  requirePreorderPortalUser,
  requireSameOrigin,
} from "../preorder/preorder-portal-auth.server";

type NotificationRow = {
  id: number;
  shop: string;
  type: string;
  status: string;
  shopifyOrderId: string | null;
  waitlistId: number | null;
  supplierOrderId: number | null;
  customerEmail: string | null;
  subject: string;
  body: string;
  createdByUserName: string | null;
  approvedByUserName: string | null;
  approvedAt: Date | null;
  sentAt: Date | null;
  cancelledAt: Date | null;
  createdAt: Date;
};

export const loader = async ({ request }: LoaderFunctionArgs) => {
  try {
    await requirePreorderPortalUser(request);
    const rows = await prisma.$queryRawUnsafe<NotificationRow[]>(
      `SELECT "id", "shop", "type", "status", "shopifyOrderId", "waitlistId", "supplierOrderId", "customerEmail",
              "subject", "body", "createdByUserName", "approvedByUserName", "approvedAt", "sentAt", "cancelledAt", "createdAt"
       FROM "PreorderNotification"
       ORDER BY "createdAt" DESC
       LIMIT 300`,
    );

    const batches = await prisma.supplierOrder.findMany({
      where: { status: "open" },
      select: { id: true, shop: true, productTitle: true, supplier: true, destination: true },
      orderBy: { createdAt: "desc" },
      take: 300,
    });

    return Response.json({
      ok: true,
      rows: rows.map((row) => ({
        ...row,
        approvedAt: row.approvedAt?.toISOString() ?? null,
        sentAt: row.sentAt?.toISOString() ?? null,
        cancelledAt: row.cancelledAt?.toISOString() ?? null,
        createdAt: row.createdAt.toISOString(),
      })),
      batches,
    });
  } catch (error) {
    if (error instanceof Response) return error;
    console.error("[preorder notifications] load failed", error);
    return Response.json({ ok: false, error: "Could not load preorder notifications." }, { status: 500 });
  }
};

export const action = async ({ request }: ActionFunctionArgs) => {
  try {
    requireSameOrigin(request);
    const actor = await requirePreorderPortalUser(request);
    const { permissions } = await getPreorderPermissionContext();
    if (!canSendPreorderNotifications(actor, permissions)) {
      return Response.json({ ok: false, error: "You do not have permission to manage preorder notifications." }, { status: 403 });
    }

    const body = await request.json().catch(() => ({})) as Record<string, unknown>;
    const operation = String(body.operation ?? "");

    if (operation === "create-draft") {
      const supplierOrderId = Number(body.supplierOrderId || 0) || null;
      const batch = supplierOrderId
        ? await prisma.supplierOrder.findUnique({ where: { id: supplierOrderId }, select: { shop: true, productTitle: true } })
        : await prisma.supplierOrder.findFirst({ select: { shop: true, productTitle: true }, orderBy: { createdAt: "desc" } });
      if (!batch?.shop) return Response.json({ ok: false, error: "No Shopify shop could be resolved for this notification." }, { status: 400 });

      const type = String(body.type ?? "eta_update").trim();
      const subject = String(body.subject ?? "").trim();
      const message = String(body.body ?? "").trim();
      const customerEmail = String(body.customerEmail ?? "").trim() || null;
      if (!subject || !message) return Response.json({ ok: false, error: "Subject and message are required." }, { status: 400 });

      const inserted = await prisma.$queryRawUnsafe<Array<{ id: number }>>(
        `INSERT INTO "PreorderNotification"
          ("shop", "type", "status", "supplierOrderId", "customerEmail", "subject", "body", "createdByUserId", "createdByUserName", "createdAt", "updatedAt")
         VALUES ($1, $2, 'draft', $3, $4, $5, $6, $7, $8, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
         RETURNING "id"`,
        batch.shop,
        type,
        supplierOrderId,
        customerEmail,
        subject,
        message,
        actor.id,
        actor.name,
      );
      return Response.json({ ok: true, id: inserted[0]?.id ?? null });
    }

    const id = Number(body.id);
    if (!Number.isInteger(id) || id <= 0) return Response.json({ ok: false, error: "Invalid notification." }, { status: 400 });

    if (operation === "approve") {
      await prisma.$executeRawUnsafe(
        `UPDATE "PreorderNotification"
         SET "status" = 'approved', "approvedByUserId" = $1, "approvedByUserName" = $2,
             "approvedAt" = CURRENT_TIMESTAMP, "updatedAt" = CURRENT_TIMESTAMP
         WHERE "id" = $3 AND "status" = 'draft'`,
        actor.id,
        actor.name,
        id,
      );
    } else if (operation === "cancel") {
      await prisma.$executeRawUnsafe(
        `UPDATE "PreorderNotification"
         SET "status" = 'cancelled', "cancelledAt" = CURRENT_TIMESTAMP, "updatedAt" = CURRENT_TIMESTAMP
         WHERE "id" = $1 AND "status" IN ('draft','approved')`,
        id,
      );
    } else if (operation === "return-to-draft") {
      await prisma.$executeRawUnsafe(
        `UPDATE "PreorderNotification"
         SET "status" = 'draft', "approvedByUserId" = NULL, "approvedByUserName" = NULL,
             "approvedAt" = NULL, "updatedAt" = CURRENT_TIMESTAMP
         WHERE "id" = $1 AND "status" = 'approved'`,
        id,
      );
    } else {
      return Response.json({ ok: false, error: "Unsupported notification operation." }, { status: 400 });
    }

    await prisma.activityLog.create({
      data: {
        userName: actor.name,
        action: `preorder_notification_${operation.replace(/-/g, "_")}`,
        entity: "preorder_notification",
        entityId: String(id),
      },
    }).catch(() => undefined);

    return Response.json({ ok: true });
  } catch (error) {
    if (error instanceof Response) return error;
    console.error("[preorder notifications] action failed", error);
    return Response.json({ ok: false, error: "Could not update preorder notification." }, { status: 500 });
  }
};
