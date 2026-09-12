import prisma from "../db.server";

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
