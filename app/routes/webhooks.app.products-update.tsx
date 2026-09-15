import type { ActionFunctionArgs } from "react-router";
import { authenticate } from "../shopify.server";
import { syncLockedCollectionRowsForProduct } from "../collections-sync.server";

// products/update → keep LOCKED collection rows in sync with Shopify. Fires the
// moment a product is saved in Shopify admin; updates the descriptive fields +
// category metafields on the linked, locked collection row(s). Best-effort — a
// failure is logged but never re-queued (Shopify would keep retrying).
export const action = async ({ request }: ActionFunctionArgs) => {
  const { shop, topic, payload } = await authenticate.webhook(request);
  try {
    const productId = (payload && typeof payload === "object" ? (payload as { id?: unknown }).id : "") ?? "";
    const r = await syncLockedCollectionRowsForProduct(shop, productId);
    if (r.updated) console.log(`[collections sync] ${topic} ${shop}: updated ${r.updated} locked row(s) for product ${productId}`);
  } catch (error) {
    console.error(`[collections sync] ${topic} failed for ${shop}:`, error instanceof Error ? error.message : error);
  }
  return new Response();
};
