import { PrismaClient } from "@prisma/client";

const prisma = new PrismaClient();
const API_VERSION = "2025-10";
const SHOPIFY_REQUEST_TIMEOUT_MS = 10000;

const REQUIRED = [
  { topic: "ORDERS_CREATE", path: "/webhooks/app/orders-create" },
  { topic: "ORDERS_CANCELLED", path: "/webhooks/app/orders-cancelled" },
  { topic: "ORDERS_FULFILLED", path: "/webhooks/app/orders-fulfilled" },
];

function appOrigin() {
  const raw = String(process.env.SHOPIFY_APP_URL || "").trim();
  if (!raw) return null;
  try {
    return new URL(raw).origin;
  } catch {
    return null;
  }
}

async function graphql(shop, accessToken, query, variables) {
  const response = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "X-Shopify-Access-Token": accessToken,
    },
    body: JSON.stringify({ query, variables }),
    signal: AbortSignal.timeout(SHOPIFY_REQUEST_TIMEOUT_MS),
  });
  if (!response.ok) {
    throw new Error(`Shopify GraphQL HTTP ${response.status} for ${shop}`);
  }
  const json = await response.json();
  if (json.errors?.length) {
    throw new Error(json.errors.map((error) => error.message).join("; "));
  }
  return json.data;
}

async function ensureWebhooksForShop(shop, accessToken, origin) {
  const topics = REQUIRED.map((item) => item.topic);
  const data = await graphql(
    shop,
    accessToken,
    `#graphql
      query PreorderWebhookSubscriptions($topics: [WebhookSubscriptionTopic!]) {
        webhookSubscriptions(first: 100, topics: $topics) {
          nodes { id topic uri }
        }
      }
    `,
    { topics },
  );
  const nodes = data?.webhookSubscriptions?.nodes ?? [];

  for (const required of REQUIRED) {
    const uri = `${origin}${required.path}`;
    const exists = nodes.some((node) => node.topic === required.topic && node.uri === uri);
    if (exists) {
      console.log(`[preorder webhooks] ${shop}: ${required.topic} already registered`);
      continue;
    }

    const createData = await graphql(
      shop,
      accessToken,
      `#graphql
        mutation CreatePreorderWebhook($topic: WebhookSubscriptionTopic!, $subscription: WebhookSubscriptionInput!) {
          webhookSubscriptionCreate(topic: $topic, webhookSubscription: $subscription) {
            webhookSubscription { id topic uri }
            userErrors { field message }
          }
        }
      `,
      { topic: required.topic, subscription: { uri } },
    );
    const result = createData?.webhookSubscriptionCreate;
    if (result?.userErrors?.length) {
      throw new Error(result.userErrors.map((error) => error.message).join("; "));
    }
    console.log(`[preorder webhooks] ${shop}: registered ${required.topic}`);
  }
}

async function main() {
  const origin = appOrigin();
  if (!origin) {
    console.warn("[preorder webhooks] SHOPIFY_APP_URL missing or invalid; registration skipped");
    return;
  }

  const productionShops = await prisma.supplierOrder.findMany({
    select: { shop: true },
    distinct: ["shop"],
  });
  const shops = Array.from(new Set(productionShops.map((row) => row.shop).filter(Boolean)));
  if (!shops.length) {
    console.log("[preorder webhooks] no production shops found; registration skipped");
    return;
  }

  for (const shop of shops) {
    const session = await prisma.session.findFirst({
      where: { shop, isOnline: false },
      orderBy: { expires: "desc" },
      select: { accessToken: true },
    });
    if (!session?.accessToken) {
      console.warn(`[preorder webhooks] ${shop}: offline Shopify session missing; skipped`);
      continue;
    }

    try {
      await ensureWebhooksForShop(shop, session.accessToken, origin);
    } catch (error) {
      // Registration should never take the Production Portal offline. A timeout,
      // Shopify outage, or permission issue is logged and retried on a later deploy.
      console.error(`[preorder webhooks] ${shop}: registration failed`, error);
    }
  }
}

try {
  await main();
} finally {
  await prisma.$disconnect();
}
