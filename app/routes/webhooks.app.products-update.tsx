import type { ActionFunctionArgs } from "react-router";
import { authenticate } from "../shopify.server";
import { syncLockedCollectionRowsForProduct } from "../collections-sync.server";

// products/update → keep LOCKED collection rows in sync with Shopify. Fires the
// moment a product is saved in Shopify admin; updates the descriptive fields +
// category metafields on the linked, locked collection row(s). Best-effort — a
// failure is logged but never re-queued (Shopify would keep retrying).
export const action = async ({ request }: ActionFunctionArgs) => {
  const { shop, topic, payload } = await authenticate.webhook(request);
  const productId = (payload && typeof payload === "object" ? (payload as { id?: unknown }).id : "") ?? "";
  // Respond 200 IMMEDIATELY. Shopify requires a webhook response within ~5s and
  // DISABLES a topic whose deliveries keep timing out (which is exactly what
  // happened). The locked-row sync does a DB scan + a Shopify GraphQL call that
  // can exceed 5s, so run it in the BACKGROUND — Railway is a long-lived Node
  // process, so this promise keeps running after the response is sent. Errors
  // are logged, never re-queued.
  void syncLockedCollectionRowsForProduct(shop, productId)
    .then((r) => { if (r.updated) console.log(`[collections sync] ${topic} ${shop}: updated ${r.updated} locked row(s) for product ${productId}`); })
    .catch((error) => { console.error(`[collections sync] ${topic} failed for ${shop}:`, error instanceof Error ? error.message : error); });
  return new Response(null, { status: 200 });
};
