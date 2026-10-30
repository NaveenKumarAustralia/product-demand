import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { reconcilePreorderForProduct } from "../preorder/preorder-reconcile.server";

// Admin: manually run the pre-order reconciler for ONE product. Dry-run by
// default (shows what it WOULD do); add &apply=1 to execute. This is the manual
// "Reconcile now" backstop while we wire the reconciler into the event paths.
//   /api/preorder-reconcile?productId=1234567890            (dry run)
//   /api/preorder-reconcile?productId=1234567890&apply=1    (execute)
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const url = new URL(request.url);
  const productId = (url.searchParams.get("productId") ?? "").trim();
  if (!productId) return Response.json({ ok: false, error: "Pass ?productId=<numeric or gid>." }, { status: 400 });
  const apply = url.searchParams.get("apply") === "1" || url.searchParams.get("apply") === "true";

  const session = await prisma.session.findFirst({
    where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true },
  });
  if (!session?.shop) return Response.json({ ok: false, error: "No Shopify session." }, { status: 500 });

  const result = await reconcilePreorderForProduct(session.shop, productId, { dryRun: !apply });
  return Response.json(
    { shop: session.shop, ...result, note: apply ? "Applied." : "DRY RUN — add &apply=1 to execute." },
    { headers: { "Cache-Control": "no-store" } },
  );
};
