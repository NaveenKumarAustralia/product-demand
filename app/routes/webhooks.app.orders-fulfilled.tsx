import type { ActionFunctionArgs } from "react-router";
import { authenticate } from "../shopify.server";
import { processShopifyOrderFulfilled } from "../preorder/preorder-shopify-order.server";

export const action = async ({ request }: ActionFunctionArgs) => {
  const { shop, topic, payload } = await authenticate.webhook(request);
  const result = await processShopifyOrderFulfilled(shop, payload);
  if (result.fulfilled > 0) {
    console.log(`[preorder] ${topic} ${shop}: marked ${result.fulfilled} reservation row(s) fulfilled`);
  }
  return new Response();
};
