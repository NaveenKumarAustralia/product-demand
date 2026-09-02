import type { ActionFunctionArgs } from "react-router";
import { authenticate } from "../shopify.server";
import { PreorderCapacityError } from "../preorder/preorder-allocation.server";
import { processShopifyOrderCreated } from "../preorder/preorder-shopify-order.server";

export const action = async ({ request }: ActionFunctionArgs) => {
  const { shop, topic, payload } = await authenticate.webhook(request);

  try {
    const result = await processShopifyOrderCreated(shop, payload);
    if (result.preorder) {
      console.log(`[preorder] ${topic} ${shop}: reserved ${result.reservations} unit(s) for ${result.market}`);
    }
  } catch (error) {
    if (error instanceof PreorderCapacityError) {
      // Capacity failures need staff attention, not repeated Shopify delivery.
      // The ingestion service has already released partial allocations and
      // written an ActivityLog exception.
      console.error(`[preorder] ${topic} allocation exception for ${shop}:`, error.message);
      return new Response();
    }
    console.error(`[preorder] ${topic} processing failed for ${shop}:`, error);
    throw error;
  }

  return new Response();
};
