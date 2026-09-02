import type { ActionFunctionArgs } from "react-router";
import { authenticate } from "../shopify.server";
import { processShopifyOrderCancelled } from "../preorder/preorder-shopify-order.server";

export const action = async ({ request }: ActionFunctionArgs) => {
  const { shop, topic, payload } = await authenticate.webhook(request);
  const result = await processShopifyOrderCancelled(shop, payload);
  if (result.released > 0) {
    console.log(`[preorder] ${topic} ${shop}: released ${result.released} reservation row(s)`);
  }
  return new Response();
};
