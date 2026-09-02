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
  source: string;
  notifiedAt: Date | null;
  convertedOrderId: string | null;
  createdAt: Date;
};

export const loader = async ({ request }: LoaderFunctionArgs) => {
  try {
    await requirePreorderPortalUser(request);
    const url = new URL(request.url);
    const market = url.searchParams.get("market");
    const status = url.searchParams.get("status");
    const q = (url.searchParams.get("q") || "").trim();

    const clauses: string[] = [];
    const values: unknown[] = [];
    if (market === "AU" || market === "USA") {
      values.push(market);
      clauses.push(`"market" = $${values.length}`);
    }
    if (status && ["waiting", "notified", "converted", "removed"].includes(status)) {
      values.push(status);
      clauses.push(`"status" = $${values.length}`);
    }
    if (q) {
      values.push(`%${q}%`);
      const p = `$${values.length}`;
      clauses.push(`("email" ILIKE ${p} OR COALESCE("productTitle", '') ILIKE ${p} OR COALESCE("variantTitle", '') ILIKE ${p} OR COALESCE("sku", '') ILIKE ${p})`);
    }

    const where = clauses.length ? `WHERE ${clauses.join(" AND ")}` : "";
    const rows = await prisma.$queryRawUnsafe<WaitlistRow[]>(
      `SELECT "id", "shop", "email", "productId", "productTitle", "variantId", "variantTitle", "sku", "market", "status", "source", "notifiedAt", "convertedOrderId", "createdAt"
       FROM "PreorderWaitlist"
       ${where}
       ORDER BY "createdAt" DESC
       LIMIT 500`,
      ...values,
    );

    const summaryRows = await prisma.$queryRawUnsafe<Array<{ market: string; status: string; count: bigint }>>(
      `SELECT "market", "status", COUNT(*)::bigint AS count
       FROM "PreorderWaitlist"
       GROUP BY "market", "status"`,
    );

    return Response.json({
      ok: true,
      rows: rows.map((row) => ({
        ...row,
        notifiedAt: row.notifiedAt?.toISOString() ?? null,
        createdAt: row.createdAt.toISOString(),
      })),
      summary: summaryRows.map((row) => ({ market: row.market, status: row.status, count: Number(row.count) })),
    });
  } catch (error) {
    if (error instanceof Response) return error;
    console.error("[preorder waitlist] load failed", error);
    return Response.json({ ok: false, error: "Could not load preorder waitlist." }, { status: 500 });
  }
};

export const action = async ({ request }: ActionFunctionArgs) => {
  try {
    requireSameOrigin(request);
    const actor = await requirePreorderPortalUser(request);
    const { permissions } = await getPreorderPermissionContext();
    if (!canSendPreorderNotifications(actor, permissions)) {
      return Response.json({ ok: false, error: "You do not have permission to manage customer notifications." }, { status: 403 });
    }

    const body = await request.json().catch(() => ({})) as { id?: unknown; status?: unknown };
    const id = Number(body.id);
    const status = String(body.status ?? "");
    if (!Number.isInteger(id) || id <= 0 || !["waiting", "notified", "converted", "removed"].includes(status)) {
      return Response.json({ ok: false, error: "Invalid waitlist update." }, { status: 400 });
    }

    const notifiedSql = status === "notified" ? `, "notifiedAt" = COALESCE("notifiedAt", CURRENT_TIMESTAMP)` : "";
    await prisma.$executeRawUnsafe(
      `UPDATE "PreorderWaitlist"
       SET "status" = $1, "updatedAt" = CURRENT_TIMESTAMP${notifiedSql}
       WHERE "id" = $2`,
      status,
      id,
    );

    await prisma.activityLog.create({
      data: {
        userName: actor.name,
        action: "preorder_waitlist_status_changed",
        entity: "preorder_waitlist",
        entityId: String(id),
        field: "status",
        toValue: status,
      },
    }).catch(() => undefined);

    return Response.json({ ok: true });
  } catch (error) {
    if (error instanceof Response) return error;
    console.error("[preorder waitlist] update failed", error);
    return Response.json({ ok: false, error: "Could not update preorder waitlist." }, { status: 500 });
  }
};
