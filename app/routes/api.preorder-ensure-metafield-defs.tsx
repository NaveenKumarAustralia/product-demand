import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { getOfflineToken, ensurePreorderMetafieldDefinitions } from "../preorder/preorder-fulfillment.server";

// Admin: create the metafield DEFINITIONS for the pre-order variant metafields
// (karmaeast.preorder / karmaeast.dispatch) with storefront read access. Without
// a definition the values are invisible to Liquid, so the order-confirmation
// email can't read them → PayPal/Shop-Pay pre-orders got a plain email. Creating
// the definition adopts the existing values and makes them Liquid-visible
// retroactively. Idempotent — safe to run repeatedly.
//   GET /api/preorder-ensure-metafield-defs
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const session = await prisma.session.findFirst({
    where: { isOnline: false, accessToken: { not: "" } },
    orderBy: { expires: "desc" },
    select: { shop: true },
  });
  if (!session?.shop) return Response.json({ ok: false, error: "No offline Shopify session." }, { status: 500 });
  const token = await getOfflineToken(session.shop);
  if (!token) return Response.json({ ok: false, error: "No offline token / missing scopes." }, { status: 500 });

  const result = await ensurePreorderMetafieldDefinitions(session.shop, token);
  return Response.json({ ok: result.errors.length === 0, ...result }, { headers: { "Cache-Control": "no-store" } });
};
