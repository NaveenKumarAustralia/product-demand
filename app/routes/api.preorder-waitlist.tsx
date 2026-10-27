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

    const body = await request.json().catch(() => ({})) as Record<string, unknown>;

    // Manually add a customer to the waitlist (e.g. when they email/call to be
    // notified). Re-activates an existing row for the same shop+email+variant+
    // market instead of erroring on the unique index.
    if (String(body.op ?? "") === "add") {
      const email = String(body.email ?? "").trim().toLowerCase();
      const variantId = String(body.variantId ?? "").trim();
      const market = String(body.market ?? "").trim().toUpperCase();
      if (!email || !email.includes("@") || email.length > 100) return Response.json({ ok: false, error: "Enter a valid customer email." }, { status: 400 });
      if (!variantId) return Response.json({ ok: false, error: "Pick a product size/variant." }, { status: 400 });
      if (market !== "AU" && market !== "USA") return Response.json({ ok: false, error: "Pick a market (AU or USA)." }, { status: 400 });
      const str = (v: unknown) => { const t = v == null ? "" : String(v).trim(); return t || null; };
      const productId = str(body.productId), productTitle = str(body.productTitle), variantTitle = str(body.variantTitle), sku = str(body.sku);
      const session = await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true } });
      const shop = session?.shop;
      if (!shop) return Response.json({ ok: false, error: "No Shopify session available." }, { status: 500 });
      await prisma.$executeRawUnsafe(
        `INSERT INTO "PreorderWaitlist" ("shop","email","productId","productTitle","variantId","variantTitle","sku","market","status","source","createdAt","updatedAt")
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,'waiting','manual',CURRENT_TIMESTAMP,CURRENT_TIMESTAMP)
         ON CONFLICT ("shop","email","variantId","market") DO UPDATE SET
           "status" = 'waiting', "updatedAt" = CURRENT_TIMESTAMP,
           "productId" = COALESCE(EXCLUDED."productId", "PreorderWaitlist"."productId"),
           "productTitle" = COALESCE(EXCLUDED."productTitle", "PreorderWaitlist"."productTitle"),
           "variantTitle" = COALESCE(EXCLUDED."variantTitle", "PreorderWaitlist"."variantTitle"),
           "sku" = COALESCE(EXCLUDED."sku", "PreorderWaitlist"."sku")`,
        shop, email, productId, productTitle, variantId, variantTitle, sku, market,
      );
      await prisma.activityLog.create({ data: { userName: actor.name, action: "preorder_waitlist_added", entity: "preorder_waitlist", entityId: variantId, field: "email", toValue: email } }).catch(() => undefined);
      return Response.json({ ok: true });
    }

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
