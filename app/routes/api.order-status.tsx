import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { normalizePortalMessageUsers, PORTAL_USERS_KEY } from "../portal-messages.server";
import { authorizeApiRequest } from "../api-auth.server";
import { getPreorderSellingPlanRegistryEntries } from "../preorder/preorder-selling-plan-registry.server";
import { marketFromDestination } from "../preorder/preorder-rules.server";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type, X-Api-Key",
};
const PRODUCT_GROUP_RENAMES: Record<string, string> = {
  "Short Sleeve Dresses": "Dresses",
};
const RESTOCK_SETTINGS_KEY = "supplier-portal-restock-settings-v1";
// The restock page's chip lists and colours (portal._index.tsx). Each option is
// stored with its own bg/color in the restock settings and staff can edit them,
// so callers get the colours here rather than keeping their own copy. Keep the
// defaults and the obsolete/"ready" rules in step with normalizeRestockSettings
// there.
type RestockOption = { value: string; label: string; bg: string; color: string };

const DEFAULT_STATUS_OPTIONS = [
  { value: "on_order", label: "On Order" },
  { value: "on_production", label: "On Production" },
  { value: "ready", label: "Ready" },
  { value: "in_shipment", label: "In Shipment" },
  { value: "cancelled", label: "Cancelled" },
];
const DEFAULT_STATUS_COLORS: Record<string, { bg: string; color: string }> = {
  on_order: { bg: "#fef9c3", color: "#374151" },
  on_production: { bg: "#dbeafe", color: "#374151" },
  ready: { bg: "#ede9fe", color: "#4c1d95" },
  in_shipment: { bg: "#dcfce7", color: "#14532d" },
  cancelled: { bg: "#fee2e2", color: "#991b1b" },
};
// Merged into "ready" by the portal; never offered as a chip any more.
const OBSOLETE_STATUS_VALUES = new Set(["packed", "ready_to_send", "arrived", "arrived_loaded"]);
const STATUS_LABEL_NOISE = /^\s*invoice\b/i;
const DEFAULT_PRIORITY_OPTIONS = [
  { value: "low", label: "LOW", bg: "#3b82f6", color: "#fff" },
  { value: "high", label: "HIGH", bg: "#7c3aed", color: "#fff" },
  { value: "urgent", label: "URGENT", bg: "#dc2626", color: "#fff" },
  { value: "cancelled", label: "Cancelled", bg: "#d97706", color: "#fff" },
];
// Callers that let a destination be changed need the same list the portal offers.
const DEFAULT_DESTINATION_OPTIONS = [
  { value: "keep_at_factory", label: "Keep at factory", bg: "#fee2e2", color: "#b91c1c" },
  { value: "send_to_au", label: "Send to AUS", bg: "#fef9c3", color: "#a16207" },
  { value: "send_to_usa", label: "Send to USA", bg: "#cffafe", color: "#0e7490" },
];

function normalizeProductGroup(value?: string | null) {
  const trimmed = value?.trim() ?? "";
  return PRODUCT_GROUP_RENAMES[trimmed] ?? trimmed;
}

function slugForOption(label: string) {
  return label.trim().toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
}

function normalizeHexColor(value: unknown, fallback: string) {
  const color = String(value ?? "").trim();
  return /^#[0-9a-f]{6}$/i.test(color) ? color : fallback;
}

function normalizeOptions(
  value: unknown,
  defaults: Array<{ value: string; label: string; bg?: string; color?: string }>,
  defaultColors?: Record<string, { bg: string; color: string }>,
): RestockOption[] {
  const usingDefaults = !Array.isArray(value);
  const items = usingDefaults ? defaults : (value as unknown[]);
  const seen = new Set<string>();
  const normalized = items
    .map((item) => {
      if (!item || typeof item !== "object") return null;
      const option = item as Record<string, unknown>;
      const label = String(option.label ?? "").trim();
      const optionValue = slugForOption(String(option.value ?? label));
      if (!label || !optionValue || seen.has(optionValue)) return null;
      seen.add(optionValue);
      const fallback = (defaultColors?.[optionValue] ?? defaults.find((d) => d.value === optionValue) ?? {}) as { bg?: string; color?: string };
      return {
        value: optionValue,
        label,
        bg: normalizeHexColor(option.bg, fallback.bg ?? "#f3f4f6"),
        color: normalizeHexColor(option.color, fallback.color ?? "#374151"),
      };
    })
    .filter(Boolean) as RestockOption[];
  if (usingDefaults) {
    for (const d of defaults) {
      if (seen.has(d.value)) continue;
      const fallback = defaultColors?.[d.value] ?? d;
      normalized.push({
        value: d.value,
        label: d.label,
        bg: normalizeHexColor(d.bg, fallback.bg ?? "#f3f4f6"),
        color: normalizeHexColor(d.color, fallback.color ?? "#374151"),
      });
    }
  }
  return normalized.length ? normalized : normalizeOptions(defaults, defaults, defaultColors);
}

function normalizeRestockSettings(value: unknown) {
  const settings = value && typeof value === "object" && !Array.isArray(value)
    ? value as Record<string, unknown>
    : {};
  const statusOptions = normalizeOptions(settings.statusOptions, DEFAULT_STATUS_OPTIONS, DEFAULT_STATUS_COLORS)
    .filter((option) => !OBSOLETE_STATUS_VALUES.has(option.value) && !STATUS_LABEL_NOISE.test(option.label));
  if (!statusOptions.some((option) => option.value === "ready")) {
    const insertAt = Math.max(0, statusOptions.findIndex((option) => option.value === "on_production") + 1);
    statusOptions.splice(insertAt, 0, { value: "ready", label: "Ready", ...DEFAULT_STATUS_COLORS.ready });
  }
  return {
    statusOptions,
    priorityOptions: normalizeOptions(settings.priorityOptions, DEFAULT_PRIORITY_OPTIONS),
    destinationOptions: normalizeOptions(settings.destinationOptions, DEFAULT_DESTINATION_OPTIONS),
  };
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  if (request.method === "OPTIONS") {
    return new Response(null, { status: 204, headers: CORS });
  }

  const url = new URL(request.url);
  const productId = url.searchParams.get("productId");
  const shop = url.searchParams.get("shop");

  // The product-page block has room for a couple of orders, so 2 stays the
  // default and its behaviour is unchanged. The analytics dashboard lists them
  // all and asks for more; capped so a caller can't request the whole table.
  const requestedLimit = Number(url.searchParams.get("limit"));
  const take = Number.isInteger(requestedLimit) && requestedLimit > 0
    ? Math.min(requestedLimit, 50)
    : 2;

  if (!productId || !shop) {
    return Response.json(
      { error: "productId and shop are required" },
      { status: 400, headers: CORS },
    );
  }

  // Accepts either a Shopify session token from the embedded app or the
  // dashboard's shared secret. See api-auth.server.ts.
  const unauthorized = authorizeApiRequest(request, CORS);
  if (unauthorized) return unauthorized;

  try {
    const [orders, usersSetting, restockSettingsSetting] = await Promise.all([
      prisma.supplierOrder.findMany({
        where: { shop, productId, status: "open" },
        select: {
          id: true,
          supplier: true,
          totalQty: true,
          productType: true,
          eta: true,
          status: true,
          supplierStatus: true,
          priority: true,
          destination: true,
          notes: true,
          lines: {
            // variantTitle so a caller can label a size without having to hold
            // its own variant lookup; qtyReceived so "on order" can be told
            // apart from what has already landed.
            select: {
              variantId: true,
              variantTitle: true,
              qtyOrdered: true,
              qtyReceived: true,
            },
          },
        },
        orderBy: { createdAt: "desc" },
        take,
      }),
      prisma.portalSetting.findUnique({
        where: { key: PORTAL_USERS_KEY },
        select: { value: true },
      }),
      prisma.portalSetting.findUnique({
        where: { key: RESTOCK_SETTINGS_KEY },
        select: { value: true },
      }),
    ]);
    const restockSettings = normalizeRestockSettings(restockSettingsSetting?.value);

    // Pre-order state per order (batch), read-only — the same three facts the
    // restock page's Status cell shows: enabled in the portal, live on Shopify
    // (has a selling plan in the registry), and the customer ship date.
    const orderIds = orders.map((order) => order.id);
    const [preorderSettings, registryEntries] = orderIds.length
      ? await Promise.all([
          prisma.preorderBatchSetting.findMany({
            where: { supplierOrderId: { in: orderIds } },
            select: { supplierOrderId: true, enabled: true, shipDate: true },
          }),
          getPreorderSellingPlanRegistryEntries(shop),
        ])
      : [[], []];
    const preorderSettingById = new Map(preorderSettings.map((setting) => [setting.supplierOrderId, setting]));
    const liveOrderIds = new Set(registryEntries.map((entry) => entry.supplierOrderId));
    const staffNames = normalizePortalMessageUsers(usersSetting?.value)
      .filter((user) => user.active !== false)
      .map((user) => user.name);

    const latestOrder = orders[0] ?? null;
    const formattedOrders = orders.map((order) => ({
      ...order,
      productType: normalizeProductGroup(order.productType) || null,
      eta: order.eta?.toISOString() ?? null,
      preorder: {
        enabled: preorderSettingById.get(order.id)?.enabled === true,
        live: liveOrderIds.has(order.id),
        shipDate: preorderSettingById.get(order.id)?.shipDate?.toISOString() ?? null,
        market: marketFromDestination(order.destination),
      },
    }));
    const totalQty = orders.reduce((sum, order) => sum + order.totalQty, 0);

    return Response.json(
      {
        order: latestOrder
          ? {
              id: latestOrder.id,
              supplier: latestOrder.supplier,
              totalQty,
              productType: normalizeProductGroup(latestOrder.productType) || null,
              eta: latestOrder.eta?.toISOString() ?? null,
              status: latestOrder.status,
              supplierStatus: latestOrder.supplierStatus,
              priority: latestOrder.priority,
              destination: latestOrder.destination,
              notes: latestOrder.notes,
              lines: formattedOrders.flatMap((order) => order.lines),
            }
          : null,
        orders: formattedOrders,
        staffNames,
        statusOptions: restockSettings.statusOptions,
        priorityOptions: restockSettings.priorityOptions,
        destinationOptions: restockSettings.destinationOptions,
      },
      { headers: CORS },
    );
  } catch (err) {
    console.error("order-status DB error:", err);
    return Response.json(
      { error: "Database error" },
      { status: 500, headers: CORS },
    );
  }
};
