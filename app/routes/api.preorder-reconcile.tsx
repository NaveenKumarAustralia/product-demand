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
  let productId = (url.searchParams.get("productId") ?? "").trim();
  const title = (url.searchParams.get("title") ?? "").trim();
  const handle = (url.searchParams.get("handle") ?? "").trim();
  if (!productId && !title && !handle) return Response.json({ ok: false, error: "Pass ?productId=<numeric or gid>, ?handle=<handle>, or ?title=<product title>." }, { status: 400 });
  const apply = url.searchParams.get("apply") === "1" || url.searchParams.get("apply") === "true";

  const session = await prisma.session.findFirst({
    where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true, accessToken: true },
  });
  if (!session?.shop || !session.accessToken) return Response.json({ ok: false, error: "No Shopify session." }, { status: 500 });

  // Resolve a title/handle to the numeric product id so the URL is pasteable.
  if (!productId && (title || handle)) {
    const q = handle ? `handle:${handle}` : `title:${title}`;
    const res = await fetch(`https://${session.shop}/admin/api/2025-10/graphql.json`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": session.accessToken },
      body: JSON.stringify({ query: `#graphql
        query ReconcileResolve($q: String!) { products(first: 1, query: $q) { nodes { id title } } }`, variables: { q } }),
    });
    const json = await res.json().catch(() => ({})) as { data?: { products?: { nodes?: Array<{ id?: string; title?: string }> } } };
    const node = json.data?.products?.nodes?.[0];
    if (!node?.id) return Response.json({ ok: false, error: `No product matching "${title || handle}".` }, { status: 404 });
    productId = node.id;
  }

  const result = await reconcilePreorderForProduct(session.shop, productId, { dryRun: !apply });
  return Response.json(
    { shop: session.shop, ...result, note: apply ? "Applied." : "DRY RUN — add &apply=1 to execute." },
    { headers: { "Cache-Control": "no-store" } },
  );
};
