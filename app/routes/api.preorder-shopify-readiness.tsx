import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

const API_VERSION = "2026-07";

const REQUIRED_SCOPES = [
  { handle: "write_products", label: "Write products", protected: false },
  { handle: "read_all_orders", label: "Read all orders", protected: true },
  { handle: "read_customer_payment_methods", label: "Read customer payment methods", protected: true },
  { handle: "read_purchase_options", label: "Read purchase options", protected: false },
  { handle: "write_purchase_options", label: "Write purchase options", protected: false },
  { handle: "read_payment_mandate", label: "Read payment mandates", protected: true },
  { handle: "write_payment_mandate", label: "Write payment mandates", protected: true },
  { handle: "write_app_proxy", label: "Storefront app proxy", protected: false },
] as const;

async function queryScopes(shop: string, accessToken: string) {
  const response = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
    method: "POST",
    signal: AbortSignal.timeout(8000),
    headers: {
      "Content-Type": "application/json",
      "X-Shopify-Access-Token": accessToken,
    },
    body: JSON.stringify({
      query: `#graphql
        query PreorderAppReadiness {
          currentAppInstallation {
            accessScopes { handle description }
          }
          shop { name myshopifyDomain }
        }
      `,
    }),
  });
  if (!response.ok) throw new Error(`Shopify returned HTTP ${response.status}`);
  const json = await response.json() as {
    data?: {
      currentAppInstallation?: { accessScopes?: Array<{ handle: string; description?: string }> };
      shop?: { name?: string; myshopifyDomain?: string };
    };
    errors?: Array<{ message?: string }>;
  };
  if (json.errors?.length) throw new Error(json.errors.map((error) => error.message || "Shopify GraphQL error").join("; "));
  return json.data;
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  try {
    const actor = await requirePreorderPortalUser(request);
    if (actor.admin !== true) {
      return Response.json({ ok: false, error: "Only a portal admin can view Shopify preorder readiness." }, { status: 403 });
    }

    const productionShops = await prisma.supplierOrder.findMany({ select: { shop: true }, distinct: ["shop"] });
    const shops = Array.from(new Set(productionShops.map((row) => row.shop).filter(Boolean)));
    const results = [];

    for (const shop of shops) {
      const session = await prisma.session.findFirst({
        where: { shop, isOnline: false },
        orderBy: { expires: "desc" },
        select: { accessToken: true },
      });
      if (!session?.accessToken) {
        results.push({ shop, ok: false, ready: false, error: "Offline Shopify session missing", scopes: [] });
        continue;
      }

      try {
        const data = await queryScopes(shop, session.accessToken);
        const granted = new Set((data?.currentAppInstallation?.accessScopes ?? []).map((scope) => scope.handle));
        const scopes = REQUIRED_SCOPES.map((scope) => ({ ...scope, granted: granted.has(scope.handle) }));
        results.push({
          shop,
          shopName: data?.shop?.name ?? null,
          myshopifyDomain: data?.shop?.myshopifyDomain ?? shop,
          ok: true,
          ready: scopes.every((scope) => scope.granted),
          missingScopes: scopes.filter((scope) => !scope.granted).map((scope) => scope.handle),
          scopes,
          grantedScopes: Array.from(granted).sort(),
          error: null,
        });
      } catch (error) {
        results.push({
          shop,
          ok: false,
          ready: false,
          error: error instanceof Error ? error.message : "Could not query Shopify access scopes",
          scopes: [],
        });
      }
    }

    return Response.json({
      ok: true,
      ready: results.length > 0 && results.every((shop) => shop.ready),
      shops: results,
      requiredScopes: REQUIRED_SCOPES,
      note: "Preorder selling plans and the storefront bridge remain explicitly controlled. Shopify readiness requires both purchase-option access and the storefront app-proxy scope.",
    });
  } catch (error) {
    if (error instanceof Response) return error;
    console.error("[preorder Shopify readiness] failed:", error);
    return Response.json({ ok: false, error: "Could not check Shopify preorder readiness." }, { status: 500 });
  }
};
