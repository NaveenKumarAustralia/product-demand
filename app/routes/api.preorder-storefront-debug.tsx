import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { getStorefrontPreorderState } from "../preorder/preorder-storefront-state.server";
import { getPreorderNotifyEnabled } from "../preorder/preorder-storefront-settings.server";
import { getPreorderSellingPlanRegistryEntries } from "../preorder/preorder-selling-plan-registry.server";

const API_VERSION = "2025-10";

// Admin diagnostic: for a product, show WHY each size shows what it shows on the
// storefront — the real state plus the inputs (stock, batch enabled/activated,
// capacity, notify toggle). GET so it can be opened in the browser while logged
// into the portal. Read-only.
function numeric(gid: string) {
  return String(gid ?? "").split("/").pop()?.replace(/[^0-9]/g, "") ?? "";
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const url = new URL(request.url);
  const title = String(url.searchParams.get("title") ?? "").trim();
  const handle = String(url.searchParams.get("handle") ?? "").trim();
  if (!title && !handle) return Response.json({ ok: false, error: "Pass ?title=<product title> or ?handle=<handle>." }, { status: 400 });

  const session = await prisma.session.findFirst({
    where: { isOnline: false, accessToken: { not: "" } },
    orderBy: { expires: "desc" },
    select: { shop: true, accessToken: true },
  });
  if (!session?.accessToken) return Response.json({ ok: false, error: "No offline Shopify session." }, { status: 500 });
  const shop = session.shop;

  const q = handle ? `handle:${handle}` : `title:${title}`;
  const res = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
    method: "POST",
    headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": session.accessToken },
    body: JSON.stringify({
      query: `#graphql
        query DebugProduct($q: String!) {
          products(first: 1, query: $q) {
            nodes {
              id title handle status
              variants(first: 60) { nodes { id title sku inventoryQuantity inventoryPolicy } }
            }
          }
        }
      `,
      variables: { q },
    }),
  });
  const json = await res.json() as {
    data?: { products?: { nodes?: Array<{
      id: string; title: string; handle: string; status: string;
      variants?: { nodes?: Array<{ id: string; title: string; sku: string | null; inventoryQuantity: number | null; inventoryPolicy: string | null }> };
    }> } };
    errors?: Array<{ message?: string }>;
  };
  if (json.errors?.length) return Response.json({ ok: false, error: json.errors.map((e) => e.message).join("; ") }, { status: 502 });
  const product = json.data?.products?.nodes?.[0];
  if (!product) return Response.json({ ok: false, error: `No product matching "${title || handle}".` }, { status: 404 });

  const notifyEnabled = await getPreorderNotifyEnabled();
  const registry = await getPreorderSellingPlanRegistryEntries(shop);
  const variantNodes = product.variants?.nodes ?? [];

  const variants = [];
  for (const v of variantNodes) {
    const variantId = numeric(v.id);
    // Batches that carry this variant (any status/destination) — to explain state.
    const orders = await prisma.supplierOrder.findMany({
      where: { shop, lines: { some: { variantId: { in: [v.id, variantId] } } } },
      select: {
        id: true, status: true, supplierStatus: true, destination: true,
        lines: { where: { variantId: { in: [v.id, variantId] } }, select: { qtyOrdered: true, qtyReceived: true } },
      },
    });
    const batchInfo = [];
    for (const o of orders) {
      const setting = await prisma.preorderBatchSetting.findUnique({ where: { supplierOrderId: o.id }, select: { enabled: true, shipDate: true } });
      const line = o.lines[0];
      batchInfo.push({
        batchId: o.id,
        status: o.status,
        supplierStatus: o.supplierStatus,
        destination: o.destination,
        incoming: line ? Math.max(0, line.qtyOrdered - line.qtyReceived) : 0,
        enabled: setting?.enabled === true,
        activatedOnShopify: registry.some((r) => r.supplierOrderId === o.id),
      });
    }

    let state: unknown = null;
    try {
      state = await getStorefrontPreorderState({ shop, variantId, market: "AU" });
    } catch (error) {
      state = { error: error instanceof Error ? error.message : String(error) };
    }

    variants.push({
      size: v.title,
      variantId,
      sku: v.sku,
      shopifyInventoryQuantity: v.inventoryQuantity,
      shopifyInventoryPolicy: v.inventoryPolicy,
      storefrontState: state,
      batches: batchInfo,
    });
  }

  return Response.json({
    ok: true,
    product: { title: product.title, handle: product.handle, status: product.status },
    shop,
    notifyBlockEnabled: notifyEnabled,
    totalActivatedBatchesInShop: registry.length,
    variants,
  }, { headers: { "Cache-Control": "no-store" } });
};
