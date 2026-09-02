import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

const API_VERSION = "2025-10";
const REQUIRED = [
  { topic: "ORDERS_CREATE", path: "/webhooks/app/orders-create", label: "Order created" },
  { topic: "ORDERS_CANCELLED", path: "/webhooks/app/orders-cancelled", label: "Order cancelled" },
  { topic: "ORDERS_FULFILLED", path: "/webhooks/app/orders-fulfilled", label: "Order fulfilled" },
] as const;

function appOrigin() {
  const raw = String(process.env.SHOPIFY_APP_URL || "").trim();
  if (!raw) return null;
  try {
    return new URL(raw).origin;
  } catch {
    return null;
  }
}

async function fetchSubscriptions(shop: string, accessToken: string) {
  const response = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "X-Shopify-Access-Token": accessToken,
    },
    body: JSON.stringify({
      query: `#graphql
        query PreorderWebhookStatus($topics: [WebhookSubscriptionTopic!]) {
          webhookSubscriptions(first: 100, topics: $topics) {
            nodes { id topic uri }
          }
        }
      `,
      variables: { topics: REQUIRED.map((item) => item.topic) },
    }),
  });
  if (!response.ok) throw new Error(`Shopify returned HTTP ${response.status}`);
  const json = await response.json() as {
    data?: { webhookSubscriptions?: { nodes?: Array<{ id: string; topic: string; uri: string }> } };
    errors?: Array<{ message?: string }>;
  };
  if (json.errors?.length) throw new Error(json.errors.map((error) => error.message || "Shopify GraphQL error").join("; "));
  return json.data?.webhookSubscriptions?.nodes ?? [];
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  try {
    const actor = await requirePreorderPortalUser(request);
    if (actor.admin !== true) {
      return Response.json({ ok: false, error: "Only a portal admin can view Shopify webhook health." }, { status: 403 });
    }

    const origin = appOrigin();
    if (!origin) {
      return Response.json({ ok: false, error: "SHOPIFY_APP_URL is not configured." }, { status: 500 });
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
        results.push({ shop, ok: false, error: "Offline Shopify session missing", subscriptions: [] });
        continue;
      }

      try {
        const nodes = await fetchSubscriptions(shop, session.accessToken);
        const subscriptions = REQUIRED.map((required) => {
          const uri = `${origin}${required.path}`;
          const found = nodes.find((node) => node.topic === required.topic && node.uri === uri);
          return { topic: required.topic, label: required.label, uri, registered: Boolean(found), id: found?.id ?? null };
        });
        results.push({ shop, ok: subscriptions.every((item) => item.registered), error: null, subscriptions });
      } catch (error) {
        results.push({ shop, ok: false, error: error instanceof Error ? error.message : "Could not query Shopify", subscriptions: [] });
      }
    }

    return Response.json({
      ok: true,
      origin,
      shops: results,
      healthy: results.length > 0 && results.every((shop) => shop.ok),
    });
  } catch (error) {
    if (error instanceof Response) return error;
    console.error("[preorder webhook status] failed:", error);
    return Response.json({ ok: false, error: "Could not check Shopify webhook health." }, { status: 500 });
  }
};
