import prisma from "../db.server";
import { getPreorderCombineWindowDays } from "./preorder-storefront-settings.server";

// Shopify write helpers for the pre-order auto-release: read stock at a location,
// release fulfilment holds, and add/remove order tags. All require scopes the app
// gains on re-auth (write_orders + read/write merchant-managed fulfilment orders);
// callers wrap these in try/catch so a missing scope degrades gracefully.
const API_VERSION = "2025-10";

export class PreorderFulfillmentError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "PreorderFulfillmentError";
  }
}

export async function getOfflineToken(shop: string): Promise<string | null> {
  const session = await prisma.session.findFirst({
    where: { shop, isOnline: false, accessToken: { not: "" } },
    orderBy: { expires: "desc" },
    select: { accessToken: true },
  }).catch(() => null);
  return session?.accessToken ?? null;
}

async function graphql<T>(shop: string, token: string, query: string, variables?: Record<string, unknown>): Promise<T> {
  const response = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
    method: "POST",
    headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": token },
    body: JSON.stringify({ query, variables }),
    signal: AbortSignal.timeout(10_000),
  });
  if (!response.ok) throw new PreorderFulfillmentError(`Shopify HTTP ${response.status}`);
  const json = await response.json() as { data?: T; errors?: Array<{ message?: string }> };
  if (json.errors?.length) throw new PreorderFulfillmentError(json.errors.map((e) => e.message || "GraphQL error").join("; "));
  if (!json.data) throw new PreorderFulfillmentError("Shopify returned no data.");
  return json.data;
}

function variantGid(value: string) {
  const text = String(value ?? "").trim();
  return text.startsWith("gid://shopify/ProductVariant/") ? text : `gid://shopify/ProductVariant/${text}`;
}

function locationGid(value: string | null) {
  const text = String(value ?? "").trim();
  if (!text) return null;
  if (text.startsWith("gid://shopify/Location/")) return text;
  const numeric = text.replace(/[^0-9]/g, "");
  return numeric ? `gid://shopify/Location/${numeric}` : null;
}

/** Available stock for a variant at a specific location. Needs read_inventory (we have it). */
export async function getAvailableAtLocation(shop: string, token: string, variantId: string, locationId: string | null): Promise<number> {
  const loc = locationGid(locationId);
  if (!loc) return 0;
  const data = await graphql<{
    productVariant?: { inventoryItem?: { inventoryLevels?: { nodes?: Array<{ location?: { id?: string }; quantities?: Array<{ name?: string; quantity?: number }> }> } } };
  }>(shop, token, `#graphql
    query PreorderStock($id: ID!) {
      productVariant(id: $id) {
        inventoryItem { inventoryLevels(first: 50) { nodes { location { id } quantities(names: ["available"]) { name quantity } } } }
      }
    }
  `, { id: variantGid(variantId) });
  let available = 0;
  for (const node of data.productVariant?.inventoryItem?.inventoryLevels?.nodes ?? []) {
    if (node.location?.id !== loc) continue;
    for (const q of node.quantities ?? []) {
      if (q.name === "available" && Number.isFinite(Number(q.quantity))) available += Number(q.quantity);
    }
  }
  return Math.max(0, Math.floor(available));
}

/** Release every ON_HOLD fulfilment order on an order so it can be picked/shipped. */
export async function releaseOrderPreorderHolds(shop: string, token: string, orderIdNumeric: string): Promise<number> {
  const data = await graphql<{ order?: { fulfillmentOrders?: { nodes?: Array<{ id: string; status?: string }> } } }>(
    shop, token, `#graphql
      query PreorderHolds($id: ID!) {
        order(id: $id) { fulfillmentOrders(first: 25) { nodes { id status } } }
      }
    `, { id: `gid://shopify/Order/${orderIdNumeric}` },
  );
  const held = (data.order?.fulfillmentOrders?.nodes ?? []).filter((fo) => fo.status === "ON_HOLD");
  let released = 0;
  for (const fo of held) {
    const result = await graphql<{ fulfillmentOrderReleaseHold?: { userErrors?: Array<{ message?: string }> } }>(
      shop, token, `#graphql
        mutation PreorderReleaseHold($id: ID!) {
          fulfillmentOrderReleaseHold(id: $id) { userErrors { message } }
        }
      `, { id: fo.id },
    );
    const errs = result.fulfillmentOrderReleaseHold?.userErrors;
    if (errs?.length) throw new PreorderFulfillmentError(errs.map((e) => e.message || "release hold error").join("; "));
    released += 1;
  }
  return released;
}

function toVariantGid(value: string) {
  const text = String(value ?? "").trim();
  if (text.startsWith("gid://shopify/ProductVariant/")) return text;
  const numeric = text.replace(/[^0-9]/g, "");
  return numeric ? `gid://shopify/ProductVariant/${numeric}` : null;
}

function toProductGid(value: string | null) {
  const text = String(value ?? "").trim();
  if (!text) return null;
  if (text.startsWith("gid://shopify/Product/")) return text;
  const numeric = text.replace(/[^0-9]/g, "");
  return numeric ? `gid://shopify/Product/${numeric}` : null;
}

const PREORDER_FLAG_TAG = "Pre-order";
const PREORDER_SIZE_TAG_PREFIX = "Pre-order: ";
const PREORDER_SHIPS_TAG_PREFIX = "Pre-order ships ";
// A tag the pre-order system owns (so we strip only these, never merchant tags).
export const isPreorderSystemTag = (t: string): boolean => {
  const s = String(t).trim().toLowerCase();
  return s === "pre-order" || s.startsWith("pre-order: ") || s.startsWith("pre-order ships ");
};

/**
 * Tag/untag a product as a pre-order so the confirmation email can detect it
 * RELIABLY (notification Liquid reads `line.product.tags`, unlike variant
 * metafields) AND per-SIZE, so an in-stock size of a product that also has
 * pre-order sizes is NOT flagged. For each pre-order variant we add
 * `Pre-order: <size>` (its variant title); plus a product-level `Pre-order`
 * flag (merchant visibility) and one `Pre-order ships <label>` (date). The email
 * flags a line only when `Pre-order: <that line's size>` is present. Strips all
 * pre-order tags when inactive. Read-modify-write; leaves merchant tags untouched.
 * Best-effort; needs write_products.
 */
export async function setProductPreorderTag(shop: string, token: string, productId: string | null, variantIds: string[], opts: { active: boolean; dispatchLabel: string | null }): Promise<void> {
  const gid = toProductGid(productId);
  if (!gid) return;
  const wanted = new Set((variantIds ?? []).map((v) => String(v).replace(/\D/g, "")).filter(Boolean));
  const data = await graphql<{ product?: { id?: string; tags?: string[]; variants?: { nodes?: Array<{ id?: string; title?: string; inventoryQuantity?: number }> } } }>(
    shop, token, `query KePreorderProdTags($id: ID!) { product(id: $id) { id tags variants(first: 100) { nodes { id title inventoryQuantity } } } }`, { id: gid },
  );
  if (!data.product?.id) return;
  const current = (data.product.tags ?? []).map((t) => String(t));
  const base = current.filter((t) => !isPreorderSystemTag(t));
  let next = base;
  if (opts.active) {
    // Tag a size ONLY when it's a batch variant AND currently out of stock — a
    // size is "on pre-order" exactly when it has no stock. In-stock sizes (even
    // in the batch) get no tag, so an in-stock sale is never flagged.
    const sizeTags = (data.product.variants?.nodes ?? [])
      .filter((v) => wanted.has(String(v.id ?? "").replace(/\D/g, "")) && (Number(v.inventoryQuantity) || 0) <= 0)
      .map((v) => `${PREORDER_SIZE_TAG_PREFIX}${String(v.title ?? "").trim()}`)
      .filter((t) => t.length > PREORDER_SIZE_TAG_PREFIX.length);
    // Only mark the product as pre-order if at least one size is actually OOS.
    next = sizeTags.length
      ? [...base, PREORDER_FLAG_TAG, ...sizeTags, ...(opts.dispatchLabel ? [`${PREORDER_SHIPS_TAG_PREFIX}${opts.dispatchLabel}`] : [])]
      : base;
  }
  next = Array.from(new Set(next));
  const same = next.length === current.length && next.every((t) => current.includes(t));
  if (same) return;
  const result = await graphql<{ productUpdate?: { userErrors?: Array<{ message?: string }> } }>(
    shop, token, `#graphql
      mutation KePreorderProdTagUpdate($input: ProductInput!) { productUpdate(input: $input) { userErrors { message } } }
    `, { input: { id: gid, tags: next } },
  );
  const errs = result.productUpdate?.userErrors;
  if (errs?.length) throw new PreorderFulfillmentError(errs.map((e) => e.message || "productUpdate error").join("; "));
}

let _preorderDefsEnsured = false;

/**
 * Create the metafield DEFINITIONS for the pre-order variant metafields. Without
 * a definition a metafield is "unstructured" — visible in the Admin API but NOT
 * to Liquid, so the confirmation-email banner can't read it (the exact bug that
 * left PayPal/Shop-Pay pre-orders getting a plain email). Creating the definition
 * with storefront read access adopts the existing values and makes them Liquid-
 * visible (theme + notifications), retroactively — no re-stamping needed.
 * Idempotent: an already-created definition returns a TAKEN userError we ignore.
 */
export async function ensurePreorderMetafieldDefinitions(shop: string, token: string): Promise<{ created: string[]; existing: string[]; errors: string[] }> {
  const defs = [
    { name: "Pre-order", namespace: "karmaeast", key: "preorder", type: "boolean", description: "Marks a variant that is currently on pre-order (set by the pre-order app)." },
    { name: "Pre-order dispatch", namespace: "karmaeast", key: "dispatch", type: "single_line_text_field", description: "Expected dispatch label for a pre-order variant, e.g. \"12 Oct 2026\"." },
  ];
  const created: string[] = [], existing: string[] = [], errors: string[] = [];
  for (const d of defs) {
    const result = await graphql<{ metafieldDefinitionCreate?: { createdDefinition?: { id?: string } | null; userErrors?: Array<{ code?: string; message?: string }> } }>(
      shop, token, `#graphql
        mutation KeDefCreate($def: MetafieldDefinitionInput!) {
          metafieldDefinitionCreate(definition: $def) {
            createdDefinition { id }
            userErrors { code message }
          }
        }
      `, { def: { name: d.name, namespace: d.namespace, key: d.key, description: d.description, type: d.type, ownerType: "PRODUCTVARIANT" } },
    );
    const errs = result.metafieldDefinitionCreate?.userErrors ?? [];
    if (result.metafieldDefinitionCreate?.createdDefinition?.id) created.push(d.key);
    else if (errs.some((e) => (e.code ?? "").toUpperCase() === "TAKEN")) existing.push(d.key);
    else if (errs.length) errors.push(`${d.key}: ${errs.map((e) => e.message || e.code || "error").join("; ")}`);
  }
  return { created, existing, errors };
}

/**
 * Set the pre-order metafields on a batch's variants so the confirmation email
 * (and storefront) can detect a pre-order from the VARIANT itself — this makes
 * the pre-order banner show on EVERY checkout path (Shop Pay, express, plan or
 * no plan), because the flag is on the product when the email renders.
 * namespace `karmaeast`: `preorder` (boolean), `dispatch` (text label).
 * Best-effort; needs write_products (we have it).
 */
export async function setVariantsPreorderMetafields(shop: string, token: string, variantIds: string[], opts: { preorder: boolean; dispatchLabel: string | null }): Promise<void> {
  const owners = Array.from(new Set(variantIds.map(toVariantGid).filter(Boolean))) as string[];
  if (!owners.length) return;
  // Ensure the definitions exist once per process so the values are Liquid-visible
  // (needed for the notification email to read them). Best-effort.
  if (!_preorderDefsEnsured) {
    _preorderDefsEnsured = true;
    try { await ensurePreorderMetafieldDefinitions(shop, token); }
    catch (e) { _preorderDefsEnsured = false; console.warn("[preorder] ensure metafield definitions failed:", e instanceof Error ? e.message : e); }
  }
  const metafields = owners.flatMap((ownerId) => ([
    { ownerId, namespace: "karmaeast", key: "preorder", type: "boolean", value: opts.preorder ? "true" : "false" },
    { ownerId, namespace: "karmaeast", key: "dispatch", type: "single_line_text_field", value: opts.dispatchLabel ?? "" },
  ]));
  for (let i = 0; i < metafields.length; i += 25) {
    const chunk = metafields.slice(i, i + 25);
    const result = await graphql<{ metafieldsSet?: { userErrors?: Array<{ field?: string[]; message?: string }> } }>(
      shop, token, `#graphql
        mutation KePreorderMeta($metafields: [MetafieldsSetInput!]!) {
          metafieldsSet(metafields: $metafields) { userErrors { field message } }
        }
      `, { metafields: chunk },
    );
    const errs = result.metafieldsSet?.userErrors;
    if (errs?.length) throw new PreorderFulfillmentError(errs.map((e) => e.message || "metafieldsSet error").join("; "));
  }
}

/**
 * Put a hold on every OPEN fulfilment order of an order — used to keep the
 * in-stock items of a MIXED order from shipping before the pre-order item, so
 * the whole order ships together. Pre-order lines are already ON_HOLD (Shopify's
 * deferred selling plan). Returns how many fulfilment orders were newly held.
 */
export async function holdOrderOpenFulfillmentOrders(shop: string, token: string, orderIdNumeric: string, reasonNotes: string): Promise<number> {
  const data = await graphql<{ order?: { fulfillmentOrders?: { nodes?: Array<{ id: string; status?: string }> } } }>(
    shop, token, `#graphql
      query PreorderOpenFOs($id: ID!) {
        order(id: $id) { fulfillmentOrders(first: 25) { nodes { id status } } }
      }
    `, { id: `gid://shopify/Order/${orderIdNumeric}` },
  );
  const open = (data.order?.fulfillmentOrders?.nodes ?? []).filter((fo) => fo.status === "OPEN");
  let held = 0;
  for (const fo of open) {
    const result = await graphql<{ fulfillmentOrderHold?: { userErrors?: Array<{ message?: string }> } }>(
      shop, token, `#graphql
        mutation PreorderHold($id: ID!, $hold: FulfillmentOrderHoldInput!) {
          fulfillmentOrderHold(id: $id, fulfillmentHold: $hold) { userErrors { message } }
        }
      `, { id: fo.id, hold: { reason: "OTHER", reasonNotes } },
    );
    const errs = result.fulfillmentOrderHold?.userErrors;
    if (errs?.length) throw new PreorderFulfillmentError(errs.map((e) => e.message || "hold error").join("; "));
    held += 1;
  }
  return held;
}

type PreorderFoNode = { id: string; status?: string; lineItems?: { nodes?: Array<{ id: string; remainingQuantity?: number; lineItem?: { id?: string } }> } };

async function fetchOrderFulfillmentOrders(shop: string, token: string, orderIdNumeric: string): Promise<PreorderFoNode[]> {
  const data = await graphql<{ order?: { fulfillmentOrders?: { nodes?: PreorderFoNode[] } } }>(
    shop, token, `#graphql
      query KeOrderFOs($id: ID!) {
        order(id: $id) { fulfillmentOrders(first: 25) { nodes { id status lineItems(first: 50) { nodes { id remainingQuantity lineItem { id } } } } } }
      }
    `, { id: `gid://shopify/Order/${orderIdNumeric}` },
  );
  return data.order?.fulfillmentOrders?.nodes ?? [];
}

async function splitFulfillmentOrder(shop: string, token: string, foId: string, lineItems: Array<{ id: string; remainingQuantity?: number }>): Promise<void> {
  const r = await graphql<{ fulfillmentOrderSplit?: { userErrors?: Array<{ message?: string }> } }>(
    shop, token, `#graphql
      mutation KeSplitFO($splits: [FulfillmentOrderSplitInput!]!) {
        fulfillmentOrderSplit(fulfillmentOrderSplits: $splits) { userErrors { message } }
      }
    `, { splits: [{ fulfillmentOrderId: foId, fulfillmentOrderLineItems: lineItems.map((l) => ({ id: l.id, quantity: Math.max(1, Number(l.remainingQuantity) || 1) })) }] },
  );
  const errs = r.fulfillmentOrderSplit?.userErrors;
  if (errs?.length) throw new PreorderFulfillmentError(errs.map((e) => e.message || "split error").join("; "));
}

/**
 * Normalize an order so ONLY its pre-order lines are held and everything else
 * ships now — from ANY current state (whole-order held, half-held, unheld, or a
 * previous split that held the wrong side). `preorderLineItemIds` are numeric
 * Shopify LineItem ids. Steps: (1) no-op if already correct; (2) release every
 * held fulfilment order so all are OPEN + splittable; (3) split each mixed FO to
 * isolate the pre-order lines; (4) re-query and hold every FO that is now
 * ENTIRELY pre-order — so we hold the right side by fact, never by guessing which
 * side of the split it is. Returns whether anything changed and how many FOs held.
 */
export async function normalizeOrderPreorderHolds(shop: string, token: string, orderIdNumeric: string, preorderLineItemIds: string[], reasonNotes: string): Promise<{ changed: boolean; held: number; fullyHeld: boolean }> {
  const wanted = new Set(preorderLineItemIds.map((id) => String(id).replace(/\D/g, "")).filter(Boolean));
  if (!wanted.size) return { changed: false, held: 0, fullyHeld: false };
  const isPre = (l: { lineItem?: { id?: string } }) => wanted.has(String(l.lineItem?.id ?? "").replace(/\D/g, ""));

  let fos = await fetchOrderFulfillmentOrders(shop, token, orderIdNumeric);
  const active = fos.filter((fo) => fo.status === "OPEN" || fo.status === "ON_HOLD");
  // fullyHeld = the order is entirely pre-order (no in-stock line ships now). Used
  // to decide the order-level `pre-order-hold` tag (Pick Pack sets the whole order
  // aside) — only when there's nothing shipping now.
  const fullyHeld = !active.some((fo) => (fo.lineItems?.nodes ?? []).some((l) => !isPre(l)));
  // Already correct? Every held FO is all-pre-order AND no OPEN FO holds a pre-order line.
  let correct = true;
  for (const fo of active) {
    const lines = fo.lineItems?.nodes ?? [];
    if (fo.status === "ON_HOLD" && lines.some((l) => !isPre(l))) correct = false; // in-stock held
    if (fo.status === "OPEN" && lines.some(isPre)) correct = false;                // pre-order not held
  }
  if (correct) return { changed: false, held: 0, fullyHeld };

  // (2) Release every held FO → OPEN + splittable.
  for (const fo of active) if (fo.status === "ON_HOLD") await releaseFulfillmentOrderGraceful(shop, token, fo.id);
  // (3) Split each mixed OPEN FO to isolate the pre-order lines.
  fos = await fetchOrderFulfillmentOrders(shop, token, orderIdNumeric);
  for (const fo of fos) {
    if (fo.status !== "OPEN") continue;
    const lines = fo.lineItems?.nodes ?? [];
    const preLines = lines.filter(isPre);
    if (preLines.length && lines.some((l) => !isPre(l))) await splitFulfillmentOrder(shop, token, fo.id, preLines);
  }
  // (4) Hold every OPEN FO that is now entirely pre-order (the in-stock FOs stay open).
  fos = await fetchOrderFulfillmentOrders(shop, token, orderIdNumeric);
  let held = 0;
  for (const fo of fos) {
    if (fo.status !== "OPEN") continue;
    const lines = fo.lineItems?.nodes ?? [];
    if (lines.length && lines.every(isPre)) { await holdFulfillmentOrderGraceful(shop, token, fo.id, reasonNotes); held += 1; }
  }
  return { changed: true, held, fullyHeld };
}

/** Hold only the pre-order lines of a new order (thin wrapper over normalize). */
export async function holdPreorderLinesOnly(shop: string, token: string, orderIdNumeric: string, preorderLineItemIds: string[], reasonNotes: string): Promise<number> {
  const r = await normalizeOrderPreorderHolds(shop, token, orderIdNumeric, preorderLineItemIds, reasonNotes);
  return r.held;
}

/**
 * The ONE pre-order hold policy, shared by both order paths:
 *  - if the earliest pre-order dispatch is within the combine window → hold the
 *    WHOLE order so it ships together;
 *  - otherwise → hold ONLY the pre-order line(s); in-stock lines ship now.
 * Returns whether the whole order ended up held, so the caller can decide the
 * order-level `pre-order-hold` tag (Pick Pack sets the whole order aside).
 */
export async function applyPreorderHoldPolicy(shop: string, token: string, orderIdNumeric: string, preorderLineItemIds: string[], earliestDispatchMs: number | null): Promise<{ wholeOrderHeld: boolean }> {
  const windowDays = await getPreorderCombineWindowDays();
  const combine = windowDays > 0 && earliestDispatchMs != null && earliestDispatchMs <= Date.now() + windowDays * 86400000;
  if (combine) {
    const held = await holdOrderOpenFulfillmentOrders(shop, token, orderIdNumeric, "Held to ship with the pre-order item in this order (combine window)");
    return { wholeOrderHeld: held > 0 };
  }
  const r = await normalizeOrderPreorderHolds(shop, token, orderIdNumeric, preorderLineItemIds, "Pre-order — held until the batch lands");
  return { wholeOrderHeld: r.fullyHeld };
}

async function holdFulfillmentOrderGraceful(shop: string, token: string, foId: string, reasonNotes: string): Promise<void> {
  const r = await graphql<{ fulfillmentOrderHold?: { userErrors?: Array<{ message?: string }> } }>(
    shop, token, `#graphql
      mutation KeHoldFO($id: ID!, $hold: FulfillmentOrderHoldInput!) { fulfillmentOrderHold(id: $id, fulfillmentHold: $hold) { userErrors { message } } }
    `, { id: foId, hold: { reason: "OTHER", reasonNotes } },
  );
  const bad = (r.fulfillmentOrderHold?.userErrors ?? []).filter((e) => !/already|on hold/i.test(e.message || ""));
  if (bad.length) throw new PreorderFulfillmentError(bad.map((e) => e.message || "hold error").join("; "));
}

async function releaseFulfillmentOrderGraceful(shop: string, token: string, foId: string): Promise<void> {
  const r = await graphql<{ fulfillmentOrderReleaseHold?: { userErrors?: Array<{ message?: string }> } }>(
    shop, token, `#graphql
      mutation KeReleaseFO($id: ID!) { fulfillmentOrderReleaseHold(id: $id) { userErrors { message } } }
    `, { id: foId },
  );
  const bad = (r.fulfillmentOrderReleaseHold?.userErrors ?? []).filter((e) => !/not.*hold|no hold/i.test(e.message || ""));
  if (bad.length) throw new PreorderFulfillmentError(bad.map((e) => e.message || "release error").join("; "));
}

/** Fix an EXISTING held order so only the pre-order line stays held and the
 * in-stock lines ship now (thin wrapper over normalizeOrderPreorderHolds). */
export async function resplitHeldOrderPreorderLines(shop: string, token: string, orderIdNumeric: string, preorderLineItemIds: string[]): Promise<{ changed: boolean }> {
  const r = await normalizeOrderPreorderHolds(shop, token, orderIdNumeric, preorderLineItemIds, "Pre-order — held until the batch lands");
  return { changed: r.changed };
}

export async function addOrderTags(shop: string, token: string, orderIdNumeric: string, tags: string[]): Promise<void> {
  if (!tags.length) return;
  const result = await graphql<{ tagsAdd?: { userErrors?: Array<{ message?: string }> } }>(
    shop, token, `#graphql
      mutation PreorderTagsAdd($id: ID!, $tags: [String!]!) { tagsAdd(id: $id, tags: $tags) { userErrors { message } } }
    `, { id: `gid://shopify/Order/${orderIdNumeric}`, tags },
  );
  const errs = result.tagsAdd?.userErrors;
  if (errs?.length) throw new PreorderFulfillmentError(errs.map((e) => e.message || "tagsAdd error").join("; "));
}

export async function removeOrderTags(shop: string, token: string, orderIdNumeric: string, tags: string[]): Promise<void> {
  if (!tags.length) return;
  const result = await graphql<{ tagsRemove?: { userErrors?: Array<{ message?: string }> } }>(
    shop, token, `#graphql
      mutation PreorderTagsRemove($id: ID!, $tags: [String!]!) { tagsRemove(id: $id, tags: $tags) { userErrors { message } } }
    `, { id: `gid://shopify/Order/${orderIdNumeric}`, tags },
  );
  const errs = result.tagsRemove?.userErrors;
  if (errs?.length) throw new PreorderFulfillmentError(errs.map((e) => e.message || "tagsRemove error").join("; "));
}
