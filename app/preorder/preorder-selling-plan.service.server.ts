import prisma from "../db.server";
import { canManagePreorders, type PreorderPermissionSettings } from "./preorder-permissions.server";
import { getPreorderEligibility } from "./preorder-rules.server";
import { buildPreorderSellingPlanGroup } from "./preorder-selling-plan";
import {
  getPreorderSellingPlanRegistryEntry,
  savePreorderSellingPlanRegistryEntry,
} from "./preorder-selling-plan-registry.server";

const API_VERSION = "2025-10";
const REQUIRED_ACTIVATION_SCOPES = ["write_products", "write_purchase_options"] as const;

type Actor = {
  id: string;
  name: string;
  role: "superadmin" | "admin" | "user";
  admin: boolean;
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

export async function activatePreorderSellingPlan(input: {
  supplierOrderId: number;
  actor: Actor;
  permissions: PreorderPermissionSettings;
}) {
  if (!canManagePreorders(input.actor, input.permissions)) {
    throw new PreorderSellingPlanError("You do not have permission to activate Shopify preorders.");
  }

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
      preorderSetting: {
        select: { enabled: true, shipDate: true },
      },
    },
  });
  if (!order) throw new PreorderSellingPlanError("Production batch was not found.");
  if (order.status !== "open") throw new PreorderSellingPlanError("Only an open production batch can be activated for preorder.");

  const eligibility = getPreorderEligibility({
    productionStatus: order.supplierStatus,
    destination: order.destination,
    preorderEnabled: order.preorderSetting?.enabled === true,
  });
  if (!eligibility.eligible) {
    throw new PreorderSellingPlanError(`Batch is not eligible for preorder (${eligibility.reason}).`);
  }

  const variantIds = order.lines
    .filter((line) => line.qtyOrdered - line.qtyReceived > 0)
    .map((line) => String(line.variantId ?? "").trim())
    .filter(Boolean);
  if (!variantIds.length) throw new PreorderSellingPlanError("This batch has no incoming Shopify variants available to preorder.");

  const existing = await getPreorderSellingPlanRegistryEntry(order.shop, order.id);
  if (existing) return { created: false, registry: existing };

  const accessToken = await offlineAccessToken(order.shop);
  await assertActivationScopes(order.shop, accessToken);

  const variables = buildPreorderSellingPlanGroup({
    batchId: order.id,
    productTitle: order.productTitle,
    shipDate: order.preorderSetting?.shipDate ?? order.eta ?? null,
    productIds: order.productId ? [order.productId] : [],
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
