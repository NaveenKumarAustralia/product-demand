import prisma from "../db.server";
import { canManagePreorders, type PreorderPermissionSettings } from "./preorder-permissions.server";
import { getPreorderEligibility } from "./preorder-rules.server";
import { getPreorderLocationSettings, locationForMarket } from "./preorder-locations.server";
import { buildPreorderSellingPlanGroup } from "./preorder-selling-plan";
import {
  getPreorderSellingPlanRegistryEntry,
  removePreorderSellingPlanRegistryEntry,
  savePreorderSellingPlanRegistryEntry,
} from "./preorder-selling-plan-registry.server";

const API_VERSION = "2025-10";
const REQUIRED_ACTIVATION_SCOPES = ["write_products", "write_purchase_options"] as const;

type Actor = {
  id: string;
  name: string;
  // Optional to match PortalMessageUser (the portal user shape passed in from
  // the route). Absent = not an admin; every check here uses `=== true`.
  admin?: boolean;
  active?: boolean;
};

export class PreorderSellingPlanError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "PreorderSellingPlanError";
  }
}

async function shopifyGraphql<T>(shop: string, accessToken: string, query: string, variables?: Record<string, unknown>) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 10_000);
  try {
    const response = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-Shopify-Access-Token": accessToken,
      },
      body: JSON.stringify({ query, variables }),
      signal: controller.signal,
    });
    if (!response.ok) throw new PreorderSellingPlanError(`Shopify returned HTTP ${response.status}.`);
    const json = await response.json() as { data?: T; errors?: Array<{ message?: string }> };
    if (json.errors?.length) {
      throw new PreorderSellingPlanError(json.errors.map((error) => error.message || "Shopify GraphQL error").join("; "));
    }
    if (!json.data) throw new PreorderSellingPlanError("Shopify returned no data.");
    return json.data;
  } catch (error) {
    if (error instanceof PreorderSellingPlanError) throw error;
    if (error instanceof Error && error.name === "AbortError") {
      throw new PreorderSellingPlanError("Shopify did not respond in time.");
    }
    throw new PreorderSellingPlanError(error instanceof Error ? error.message : "Could not contact Shopify.");
  } finally {
    clearTimeout(timer);
  }
}

async function offlineAccessToken(shop: string) {
  const session = await prisma.session.findFirst({
    where: { shop, isOnline: false },
    orderBy: { expires: "desc" },
    select: { accessToken: true },
  });
  if (!session?.accessToken) throw new PreorderSellingPlanError("Offline Shopify session is missing. Open the app in Shopify to reconnect it.");
  return session.accessToken;
}

async function assertActivationScopes(shop: string, accessToken: string) {
  const data = await shopifyGraphql<{
    currentAppInstallation?: { accessScopes?: Array<{ handle?: string }> };
  }>(shop, accessToken, `#graphql
    query PreorderActivationScopes {
      currentAppInstallation {
        accessScopes { handle }
      }
    }
  `);
  const granted = new Set((data.currentAppInstallation?.accessScopes ?? []).map((scope) => String(scope.handle ?? "")));
  const missing = REQUIRED_ACTIVATION_SCOPES.filter((scope) => !granted.has(scope));
  if (missing.length) {
    throw new PreorderSellingPlanError(`Shopify has not granted: ${missing.join(", ")}. Open the app in Shopify and approve the new permissions first.`);
  }
}

// Pre-order variants are out of stock, so Shopify blocks checkout ("sold out")
// unless the variant is allowed to sell past zero. Activation flips the batch's
// pre-order variants to CONTINUE (continue selling when out of stock); turning a
// batch off flips them back to DENY so they behave as normal sold-out variants.
async function setVariantsInventoryPolicy(
  shop: string,
  accessToken: string,
  productId: string | null,
  variantIds: string[],
  policy: "CONTINUE" | "DENY",
) {
  if (!variantIds.length) return;
  const pid = String(productId ?? "").trim();
  if (!pid) throw new PreorderSellingPlanError("This batch has no linked Shopify product, so pre-order variants can't be set to continue selling.");
  const productGid = pid.startsWith("gid://") ? pid : `gid://shopify/Product/${pid}`;
  const variants = variantIds.map((id) => ({
    id: id.startsWith("gid://") ? id : `gid://shopify/ProductVariant/${id}`,
    inventoryPolicy: policy,
  }));
  const data = await shopifyGraphql<{
    productVariantsBulkUpdate?: { userErrors?: Array<{ field?: string[]; message?: string }> };
  }>(shop, accessToken, `#graphql
    mutation KEPreorderInventoryPolicy($productId: ID!, $variants: [ProductVariantsBulkInput!]!) {
      productVariantsBulkUpdate(productId: $productId, variants: $variants) {
        userErrors { field message }
      }
    }
  `, { productId: productGid, variants });
  const errors = data.productVariantsBulkUpdate?.userErrors;
  if (errors?.length) {
    throw new PreorderSellingPlanError(`Could not update pre-order variant stock policy: ${errors.map((error) => error.message || "Shopify error").join("; ")}`);
  }
}

function assertManagePermission(actor: Actor, permissions: PreorderPermissionSettings) {
  if (!canManagePreorders(actor, permissions)) {
    throw new PreorderSellingPlanError("You do not have permission to manage Shopify preorders.");
  }
}

export async function activatePreorderSellingPlan(input: {
  supplierOrderId: number;
  actor: Actor;
  permissions: PreorderPermissionSettings;
}) {
  assertManagePermission(input.actor, input.permissions);

  const order = await prisma.supplierOrder.findUnique({
    where: { id: input.supplierOrderId },
    select: {
      id: true,
      shop: true,
      status: true,
      supplierStatus: true,
      destination: true,
      productId: true,
      productTitle: true,
      eta: true,
      lines: {
        select: { variantId: true, qtyOrdered: true, qtyReceived: true },
      },
    },
  });
  if (!order) throw new PreorderSellingPlanError("Production batch was not found.");
  if (order.status !== "open") throw new PreorderSellingPlanError("Only an open production batch can be activated for preorder.");

  const setting = await prisma.preorderBatchSetting.findUnique({
    where: { supplierOrderId: order.id },
    select: { enabled: true, shipDate: true },
  });

  const eligibility = getPreorderEligibility({
    supplierStatus: order.supplierStatus,
    destination: order.destination,
    preorderEnabled: setting?.enabled === true,
  });
  if (!eligibility.eligible) {
    throw new PreorderSellingPlanError(`Batch is not eligible for preorder (${eligibility.reason}).`);
  }

  // A market's preorders can only go live once that market has a fulfilment
  // location configured. This keeps USA fully off until a US location is set
  // (no US 3PL yet) and lets it "plug in" automatically later — set the USA
  // location and USA batches become activatable with no code change.
  const locationSettings = await getPreorderLocationSettings();
  if (eligibility.market && !locationForMarket(locationSettings, eligibility.market)) {
    throw new PreorderSellingPlanError(`Set the ${eligibility.market} fulfilment location in Pre-orders → Settings before activating ${eligibility.market} preorders.`);
  }

  const variantIds = order.lines
    .filter((line) => line.qtyOrdered - line.qtyReceived > 0)
    .map((line) => String(line.variantId ?? "").trim())
    .filter(Boolean);
  if (!variantIds.length) throw new PreorderSellingPlanError("This batch has no incoming Shopify variants available to preorder.");

  const accessToken = await offlineAccessToken(order.shop);
  await assertActivationScopes(order.shop, accessToken);

  // Allow these variants to be purchased while out of stock — without this the
  // pre-order adds to cart but Shopify drops it as "sold out" at checkout. Runs
  // on every activate (idempotent) so a re-activation always re-asserts it.
  await setVariantsInventoryPolicy(order.shop, accessToken, order.productId, variantIds, "CONTINUE");

  const existing = await getPreorderSellingPlanRegistryEntry(order.shop, order.id);
  if (existing) return { created: false, registry: existing };

  const variables = buildPreorderSellingPlanGroup({
    batchId: order.id,
    productTitle: order.productTitle,
    shipDate: setting?.shipDate ?? order.eta ?? null,
    // Attach the selling plan to ONLY this batch's incoming variants — never the
    // whole product. Passing productIds here (previously
    // `order.productId ? [order.productId] : []`) associates the plan with the
    // product's *entire* variant set, which would expose a preorder purchase
    // option on variants that have no confirmed incoming capacity in this batch.
    productIds: [],
    variantIds,
  });

  const data = await shopifyGraphql<{
    sellingPlanGroupCreate?: {
      sellingPlanGroup?: {
        id?: string;
        sellingPlans?: { edges?: Array<{ node?: { id?: string } }> };
      };
      userErrors?: Array<{ field?: string[]; message?: string }>;
    };
  }>(order.shop, accessToken, `#graphql
    mutation CreateKarmaEastPreorder($input: SellingPlanGroupInput!, $resources: SellingPlanGroupResourceInput) {
      sellingPlanGroupCreate(input: $input, resources: $resources) {
        sellingPlanGroup {
          id
          sellingPlans(first: 1) { edges { node { id } } }
        }
        userErrors { field message }
      }
    }
  `, variables as unknown as Record<string, unknown>);

  const payload = data.sellingPlanGroupCreate;
  if (payload?.userErrors?.length) {
    throw new PreorderSellingPlanError(payload.userErrors.map((error) => error.message || "Shopify rejected the preorder selling plan.").join("; "));
  }
  const sellingPlanGroupId = String(payload?.sellingPlanGroup?.id ?? "").trim();
  const sellingPlanId = String(payload?.sellingPlanGroup?.sellingPlans?.edges?.[0]?.node?.id ?? "").trim();
  if (!sellingPlanGroupId || !sellingPlanId) throw new PreorderSellingPlanError("Shopify created no usable selling plan IDs.");

  const registry = await savePreorderSellingPlanRegistryEntry({
    supplierOrderId: order.id,
    shop: order.shop,
    sellingPlanGroupId,
    sellingPlanId,
  });

  await prisma.activityLog.create({
    data: {
      userName: input.actor.name,
      action: "preorder_selling_plan_created",
      entity: "supplier_order",
      entityId: String(order.id),
      entityName: order.productTitle,
      field: "sellingPlanGroupId",
      toValue: sellingPlanGroupId,
    },
  }).catch(() => undefined);

  return { created: true, registry };
}

export async function deactivatePreorderSellingPlan(input: {
  supplierOrderId: number;
  actor: Actor;
  permissions: PreorderPermissionSettings;
}) {
  assertManagePermission(input.actor, input.permissions);

  const order = await prisma.supplierOrder.findUnique({
    where: { id: input.supplierOrderId },
    select: {
      id: true, shop: true, productTitle: true, productId: true,
      lines: { select: { variantId: true, qtyOrdered: true, qtyReceived: true } },
    },
  });
  if (!order) throw new PreorderSellingPlanError("Production batch was not found.");

  const existing = await getPreorderSellingPlanRegistryEntry(order.shop, order.id);
  if (!existing) return { removed: false };

  const accessToken = await offlineAccessToken(order.shop);
  await assertActivationScopes(order.shop, accessToken);

  // Restore normal sold-out behaviour on the variants we opened for pre-order.
  const variantIds = order.lines
    .filter((line) => line.qtyOrdered - line.qtyReceived > 0)
    .map((line) => String(line.variantId ?? "").trim())
    .filter(Boolean);
  await setVariantsInventoryPolicy(order.shop, accessToken, order.productId, variantIds, "DENY").catch((error) => {
    console.warn("[preorder deactivate] could not restore DENY inventory policy:", error);
  });

  const data = await shopifyGraphql<{
    sellingPlanGroupDelete?: { deletedSellingPlanGroupId?: string; userErrors?: Array<{ field?: string[]; message?: string }> };
  }>(order.shop, accessToken, `#graphql
    mutation DeleteKarmaEastPreorder($id: ID!) {
      sellingPlanGroupDelete(id: $id) {
        deletedSellingPlanGroupId
        userErrors { field message }
      }
    }
  `, { id: existing.sellingPlanGroupId });

  const payload = data.sellingPlanGroupDelete;
  if (payload?.userErrors?.length) {
    throw new PreorderSellingPlanError(payload.userErrors.map((error) => error.message || "Shopify rejected preorder removal.").join("; "));
  }

  await removePreorderSellingPlanRegistryEntry(order.shop, order.id);
  await prisma.activityLog.create({
    data: {
      userName: input.actor.name,
      action: "preorder_selling_plan_removed",
      entity: "supplier_order",
      entityId: String(order.id),
      entityName: order.productTitle,
      field: "sellingPlanGroupId",
      toValue: existing.sellingPlanGroupId,
    },
  }).catch(() => undefined);

  return { removed: true };
}
