/**
 * MCP tools for the Karma East Production Portal.
 *
 * READ tools answer straight from the portal database and the portal's own
 * read logic (reorder planner, fabric planner, Shopify search).
 *
 * WRITE tools never change anything directly. Each write is two steps:
 *   1. preview_*  → validates, checks for duplicates, snapshots the current
 *                   values and stores a pending change. Nothing is written.
 *   2. confirm_change(change_id) → only after Koku explicitly approves in chat.
 *      Re-checks that nothing changed since the preview, executes through the
 *      portal's own action handler, and records before/after in McpAuditLog.
 * A change id can be confirmed once only, which is also what stops a retried
 * call from creating a second order.
 */
import { randomUUID } from "node:crypto";
import prisma from "../db.server";
import { createReplyMessage, resolveReplyRecipients } from "../portal-messages.server";
import type { McpPrincipal } from "./oauth-core";
import { appendFabricOrderFromMcp, readFabricOrderSheets, readFabricOrderSheetSnapshot, callPortalIntent, ensureMcpPortalUser, MCP_PORTAL_USER_ID, MCP_PORTAL_USER_NAME, readFabric } from "./portal-bridge.server";
import {
  appendNote,
  compactForModel,
  findPotentialDuplicates,
  isValidIsoDate,
  matchSize,
  PENDING_CHANGE_TTL_MS,
  resolveNewOrderQuantities,
  resolveQuantityChanges,
  stableStringify,
  type ResolvedSizeChange,
  type SizeQtyChange,
} from "./mcp-logic";

export type ToolContext = { principal: McpPrincipal };
export type ToolResult = { structured: unknown; text?: string; isError?: boolean };

type JsonSchema = Record<string, unknown>;
export type ToolDef = {
  name: string;
  title: string;
  description: string;
  inputSchema: JsonSchema;
  scope: "portal:read" | "portal:write";
  annotations: { readOnlyHint: boolean; destructiveHint?: boolean; idempotentHint?: boolean; openWorldHint?: boolean };
  handler: (args: Record<string, unknown>, ctx: ToolContext) => Promise<ToolResult>;
};

class ToolInputError extends Error {}

const STATUS_VALUES = ["on_order", "on_production", "ready", "in_shipment", "cancelled"];
const PRIORITY_FALLBACK = ["low", "high", "urgent"];

// ─── Small helpers ──────────────────────────────────────────────────────────

const str = (v: unknown, name: string, opts: { required?: boolean; max?: number } = {}): string | undefined => {
  if (v === undefined || v === null || v === "") {
    if (opts.required) throw new ToolInputError(`${name} is required`);
    return undefined;
  }
  const s = String(v).trim();
  if (opts.max && s.length > opts.max) throw new ToolInputError(`${name} is too long (max ${opts.max})`);
  return s;
};
const int = (v: unknown, name: string, opts: { required?: boolean; min?: number; max?: number } = {}): number | undefined => {
  if (v === undefined || v === null || v === "") {
    if (opts.required) throw new ToolInputError(`${name} is required`);
    return undefined;
  }
  const n = Number(v);
  if (!Number.isInteger(n)) throw new ToolInputError(`${name} must be a whole number`);
  if (opts.min !== undefined && n < opts.min) throw new ToolInputError(`${name} must be ≥ ${opts.min}`);
  if (opts.max !== undefined && n > opts.max) throw new ToolInputError(`${name} must be ≤ ${opts.max}`);
  return n;
};
const isoDaysAgo = (days: number) => new Date(Date.now() - days * 86400000).toISOString().slice(0, 10);
const today = () => new Date().toISOString().slice(0, 10);
function dateArg(v: unknown, name: string, fallback: string): string {
  const s = str(v, name);
  if (!s) return fallback;
  if (!isValidIsoDate(s)) throw new ToolInputError(`${name} must be a date like 2026-11-25`);
  return s;
}

async function restockOptions() {
  const row = await prisma.portalSetting.findUnique({ where: { key: "supplier-portal-restock-settings-v1" }, select: { value: true } });
  const v = (row?.value ?? {}) as Record<string, unknown>;
  const values = (list: unknown) => (Array.isArray(list) ? list.map((o) => String((o as { value?: unknown })?.value ?? "")).filter(Boolean) : []);
  return {
    statuses: values(v.statusOptions).length ? values(v.statusOptions) : STATUS_VALUES,
    priorities: values(v.priorityOptions).length ? values(v.priorityOptions) : PRIORITY_FALLBACK,
    destinations: values(v.destinationOptions),
  };
}

const orderSelect = {
  id: true, shop: true, poNumber: true, supplier: true, productId: true, productTitle: true, productType: true,
  status: true, supplierStatus: true, priority: true, destination: true, eta: true, notes: true, factoryNotes: true,
  supplierNotes: true, totalQty: true, fabricName: true, metersPerPiece: true, fabricConsumed: true, fabricConsumedMeters: true,
  productionDate: true, packingListId: true, createdAt: true, updatedAt: true,
  lines: { select: { variantTitle: true, sku: true, qtyOrdered: true, qtyReceived: true, costPrice: true }, orderBy: { id: "asc" as const } },
} as const;

type OrderRow = Awaited<ReturnType<typeof loadOrder>>;
async function loadOrder(id: number) {
  return prisma.supplierOrder.findUnique({ where: { id }, select: orderSelect });
}

function sizesOf(order: NonNullable<OrderRow>): Record<string, number> {
  const out: Record<string, number> = {};
  for (const l of order.lines) out[l.variantTitle] = (out[l.variantTitle] ?? 0) + l.qtyOrdered;
  return out;
}

function presentOrder(o: NonNullable<OrderRow>) {
  return {
    order_id: o.id,
    product: o.productTitle,
    product_id: o.productId,
    product_type: o.productType,
    supplier: o.supplier,
    po_number: o.poNumber,
    open: o.status === "open",
    production_status: o.supplierStatus,
    priority: o.priority,
    destination: o.destination,
    eta: o.eta?.toISOString().slice(0, 10) ?? null,
    production_date: o.productionDate?.toISOString().slice(0, 10) ?? null,
    total_qty: o.totalQty,
    sizes: o.lines.map((l) => ({ size: l.variantTitle, ordered: l.qtyOrdered, received: l.qtyReceived, cost_price: l.costPrice, sku: l.sku })),
    fabric: o.fabricName ? { name: o.fabricName, meters_per_piece: o.metersPerPiece, consumed: o.fabricConsumed, consumed_meters: o.fabricConsumedMeters } : null,
    notes: o.notes,
    factory_notes: o.factoryNotes,
    supplier_notes: o.supplierNotes,
    packing_list_id: o.packingListId,
    created_at: o.createdAt.toISOString(),
    updated_at: o.updatedAt.toISOString(),
  };
}

type SearchProduct = { id: string; shop?: string; title: string; imageUrl: string | null; sizes: string[]; variants: Array<{ id: string; title: string; sku: string | null }> };

// Shopify product search is expensive (it shares the live portal's API rate
// limit) and the portal's helper returns [] when throttled. So: cache results
// for 5 minutes and retry once after a pause when a search comes back empty.
const searchCache = new Map<string, { at: number; products: SearchProduct[] }>();
const SEARCH_TTL_MS = 5 * 60 * 1000;

async function searchProducts(q: string): Promise<SearchProduct[]> {
  const key = q.trim().toLowerCase();
  const hit = searchCache.get(key);
  if (hit && Date.now() - hit.at < SEARCH_TTL_MS) return hit.products;
  let products: SearchProduct[] = [];
  for (let attempt = 0; attempt < 2 && !products.length; attempt++) {
    if (attempt) await new Promise((r) => setTimeout(r, 2000));
    const res = (await callPortalIntent("reorder_search", { q })) as { products?: SearchProduct[] } | null;
    products = res?.products ?? [];
  }
  if (products.length) {
    searchCache.set(key, { at: Date.now(), products });
    if (searchCache.size > 500) searchCache.delete(searchCache.keys().next().value!);
  }
  return products;
}

/** Cheap lookup: product id + title from the portal's own orders, no Shopify call. */
async function productFromOrders(args: Record<string, unknown>): Promise<{ id: string; title: string } | null> {
  const id = str(args.product_id, "product_id");
  const title = str(args.product_title, "product_title");
  const rows = await prisma.supplierOrder.findMany({
    where: id ? { productId: id } : { productTitle: { equals: title ?? "", mode: "insensitive" } },
    select: { productId: true, productTitle: true },
    distinct: ["productId"],
    take: 3,
  });
  return rows.length === 1 ? { id: rows[0].productId, title: rows[0].productTitle } : null;
}

/** Resolve a product from a Shopify GID or an exact title (needs Shopify for sizes). */
async function resolveProduct(args: Record<string, unknown>): Promise<SearchProduct> {
  const id = str(args.product_id, "product_id");
  const title = str(args.product_title, "product_title");
  if (!id && !title) throw new ToolInputError("Give product_id or product_title");
  if (id && !id.startsWith("gid://shopify/Product/")) throw new ToolInputError("product_id must look like gid://shopify/Product/123");
  const known = await productFromOrders(args);
  const query = title ?? known?.title;
  if (!query) throw new ToolInputError("Product not found in the portal; search it by product_title");
  const wantId = id ?? known?.id;
  const results = await searchProducts(query);
  if (wantId) {
    const exact = results.find((p) => p.id === wantId);
    if (exact) return exact;
  }
  const byTitle = results.filter((p) => p.title.trim().toLowerCase() === query.trim().toLowerCase());
  if (byTitle.length === 1) return byTitle[0];
  if (byTitle.length > 1) throw new ToolInputError(`Several products are called "${query}". Use product_id: ${byTitle.map((p) => p.id).join(", ")}`);
  if (!results.length) throw new ToolInputError(`Shopify search for "${query}" returned nothing (Shopify may be busy). Try again in a minute.`);
  throw new ToolInputError(`No exact product "${query}". Closest: ${results.slice(0, 8).map((p) => `${p.title} (${p.id})`).join("; ")}`);
}

/** Product id + title without Shopify where possible (falls back to search). */
async function resolveProductLite(args: Record<string, unknown>): Promise<{ id: string; title: string; sizes: string[] }> {
  const known = await productFromOrders(args);
  if (known) return { ...known, sizes: [] };
  const p = await resolveProduct(args);
  return { id: p.id, title: p.title, sizes: p.sizes };
}

async function audit(entry: {
  ctx: ToolContext; tool: string; outcome: string; kind?: string; changeId?: string;
  entityType?: string; entityId?: string; input?: unknown; before?: unknown; after?: unknown; error?: string;
}) {
  const p = entry.ctx.principal;
  await prisma.mcpAuditLog.create({
    data: {
      grantId: p.grantId, clientId: p.clientId, portalUserId: p.portalUserId, actorName: `${MCP_PORTAL_USER_NAME} (approved by ${p.portalUserName})`,
      tool: entry.tool, kind: entry.kind ?? null, changeId: entry.changeId ?? null,
      entityType: entry.entityType ?? null, entityId: entry.entityId ?? null, outcome: entry.outcome,
      input: (entry.input ?? undefined) as never, before: (entry.before ?? undefined) as never, after: (entry.after ?? undefined) as never,
      error: entry.error ?? null,
    },
  }).catch((e) => console.error("MCP audit write failed", e));
}

// ─── Pending changes ────────────────────────────────────────────────────────

type ChangeKind = "create_order" | "amend_quantities" | "update_order" | "add_note" | "replace_note" | "delete_order" | "staff_reply" | "create_fabric_order";

async function stagePendingChange(ctx: ToolContext, kind: ChangeKind, payload: unknown, before: unknown, summary: string, entityId: string | null, preview: unknown): Promise<ToolResult> {
  const id = `chg_${randomUUID()}`;
  const expiresAt = new Date(Date.now() + PENDING_CHANGE_TTL_MS);
  await prisma.mcpPendingChange.create({
    data: { id, grantId: ctx.principal.grantId, portalUserId: ctx.principal.portalUserId, kind, payload: payload as never, before: (before ?? undefined) as never, summary, expiresAt },
  });
  await audit({ ctx, tool: `preview:${kind}`, outcome: "previewed", kind, changeId: id, entityType: "SupplierOrder", entityId: entityId ?? undefined, input: payload, before });
  return {
    structured: {
      change_id: id,
      status: "awaiting_approval",
      expires_at: expiresAt.toISOString(),
      summary,
      preview,
      next_step: "Show this preview to Koku and ask for an explicit yes. Only after he approves, call confirm_change with this change_id. Do not confirm on your own.",
    },
    text: `${summary}\n\nNothing has been changed yet. Ask Koku to approve, then call confirm_change with change_id ${id}.`,
  };
}

// ─── Read tools ─────────────────────────────────────────────────────────────

const readTools: ToolDef[] = [
  {
    name: "list_fabric_order_sheets",
    title: "List fabric order sheets",
    description: "List the Production Portal fabric order sheets that accept orders in meters.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
    scope: "portal:read",
    annotations: { readOnlyHint: true, openWorldHint: false },
    handler: async () => ({ structured: { sheets: await readFabricOrderSheets() } }),
  },
  {
    name: "list_samples",
    title: "List product samples",
    description: "Search samples by name and read their latest recorded iteration, including size, fabric, status and factory cost.",
    inputSchema: { type: "object", properties: { search: { type: "string", maxLength: 200 }, limit: { type: "integer", minimum: 1, maximum: 50 } }, additionalProperties: false },
    scope: "portal:read",
    annotations: { readOnlyHint: true, openWorldHint: false },
    handler: async (args) => {
      const search = str(args.search, "search", { max: 200 });
      const limit = int(args.limit, "limit", { min: 1, max: 50 }) ?? 20;
      const samples = await prisma.sample.findMany({
        where: search ? { name: { contains: search, mode: "insensitive" } } : {},
        orderBy: [{ sortOrder: "asc" }, { id: "asc" }],
        take: limit + 1,
        select: {
          id: true, name: true, updatedAt: true,
          iterations: {
            orderBy: [{ version: "desc" }, { id: "desc" }],
            take: 1,
            select: { id: true, version: true, name: true, fabricType: true, sampleSize: true, buttonType: true, factoryCost: true, status: true },
          },
        },
      });
      return { structured: { samples: samples.slice(0, limit), truncated: samples.length > limit } };
    },
  },
  {
    name: "list_vision_boards",
    title: "List vision boards",
    description: "Read the names and item counts of saved vision boards. Choose the original or version 2 view.",
    inputSchema: { type: "object", properties: { version: { type: "integer", enum: [1, 2] }, limit: { type: "integer", minimum: 1, maximum: 50 } }, additionalProperties: false },
    scope: "portal:read",
    annotations: { readOnlyHint: true, openWorldHint: false },
    handler: async (args) => {
      const version = int(args.version, "version", { min: 1, max: 2 }) ?? 2;
      const limit = int(args.limit, "limit", { min: 1, max: 50 }) ?? 20;
      const query = {
        orderBy: [{ sortOrder: "asc" as const }, { id: "asc" as const }],
        take: limit + 1,
        select: { id: true, name: true, updatedAt: true, _count: { select: { items: true } } },
      };
      const boards = version === 1
        ? await prisma.visionBoard.findMany(query)
        : await prisma.visionBoardV2.findMany(query);
      return { structured: { version, boards: boards.slice(0, limit), truncated: boards.length > limit } };
    },
  },
  {
    name: "get_product_information",
    title: "Recorded product style information",
    description: "Read the Product Information style record (excluding image data), title overrides and recorded fabric consumption. A style match is not proof of a complete bill of materials; report missing component data.",
    inputSchema: { type: "object", properties: { product_title: { type: "string", minLength: 2, maxLength: 200 } }, required: ["product_title"], additionalProperties: false },
    scope: "portal:read",
    annotations: { readOnlyHint: true, openWorldHint: false },
    handler: async (args) => {
      const title = str(args.product_title, "product_title", { required: true, max: 200 })!;
      const key = title.toLowerCase();
      const rows = await prisma.$queryRaw<Array<{ style: unknown; category: unknown; style_override: unknown; fabric_override: unknown }>>`
        SELECT (sty.style - 'images' - 'image' - 'thumbnail') AS style, cat.category->'name' AS category,
          ps."value"->'titleStyleOverrides'->${key} AS style_override,
          ps."value"->'titleFabricOverrides'->${key} AS fabric_override
        FROM "PortalSetting" AS ps
        CROSS JOIN LATERAL jsonb_array_elements(ps."value"->'categories') AS cat(category)
        CROSS JOIN LATERAL jsonb_array_elements(cat.category->'styles') AS sty(style)
        WHERE ps.key = 'production-portal-product-info-v2'
          AND (sty.style->>'id' = ps."value"->'titleStyleOverrides'->>${key}
            OR lower(trim(sty.style->>'name')) = ${key}
            OR left(${key}, length(trim(sty.style->>'name')) + 1) = lower(trim(sty.style->>'name')) || ' ')
        ORDER BY (sty.style->>'id' = ps."value"->'titleStyleOverrides'->>${key}) DESC NULLS LAST,
          length(trim(sty.style->>'name')) DESC
        LIMIT 10`;
      const fabric = await readFabric({ title });
      return { structured: compactForModel({ product_title: title, matching_styles: rows,
        fabric_planner: fabric, warning: "Check exact style match and all component fabrics before committing production." }) };
    },
  },
  {
    name: "get_portal_options",
    title: "Portal options",
    description: "Lists the valid production statuses, priorities, destinations and suppliers used in the Production Portal. Call this before suggesting a status, priority or destination.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
    scope: "portal:read",
    annotations: { readOnlyHint: true, openWorldHint: false },
    handler: async () => {
      const [opts, suppliers] = await Promise.all([
        restockOptions(),
        prisma.supplierOrder.findMany({ distinct: ["supplier"], select: { supplier: true }, take: 100 }),
      ]);
      return { structured: { ...opts, suppliers: suppliers.map((s) => s.supplier).filter(Boolean) } };
    },
  },
  {
    name: "search_products",
    title: "Search products",
    description: "Searches Karma East's Shopify products by title or SKU. Returns product ids, sizes and variant SKUs.",
    inputSchema: { type: "object", properties: { query: { type: "string", minLength: 2 } }, required: ["query"], additionalProperties: false },
    scope: "portal:read",
    annotations: { readOnlyHint: true, openWorldHint: false },
    handler: async (args) => {
      const products = await searchProducts(str(args.query, "query", { required: true, max: 200 })!);
      return { structured: { products: products.slice(0, 25).map((p) => ({ product_id: p.id, title: p.title, sizes: p.sizes, skus: p.variants.map((v) => v.sku).filter(Boolean) })) } };
    },
  },
  {
    name: "list_production_orders",
    title: "List production orders",
    description: "Lists production / restock orders in the portal with their production status, sizes and quantities. Filter by product title text, status, supplier or destination. Status 'ready' means ready stock at the factory in India; 'in_shipment' means on the way.",
    inputSchema: {
      type: "object",
      properties: {
        product: { type: "string", description: "Part of the product title" },
        product_id: { type: "string" },
        statuses: { type: "array", items: { type: "string" }, description: "Production statuses, e.g. on_order, on_production, ready, in_shipment (see get_portal_options)" },
        supplier: { type: "string" },
        destination: { type: "string" },
        include_closed: { type: "boolean", default: false },
        limit: { type: "integer", minimum: 1, maximum: 200, default: 50 },
      },
      additionalProperties: false,
    },
    scope: "portal:read",
    annotations: { readOnlyHint: true, openWorldHint: false },
    handler: async (args) => {
      const statuses = Array.isArray(args.statuses) ? args.statuses.map(String) : undefined;
      const product = str(args.product, "product", { max: 200 });
      const where = {
        ...(args.include_closed ? {} : { status: "open" }),
        ...(product ? { productTitle: { contains: product, mode: "insensitive" as const } } : {}),
        ...(args.product_id ? { productId: String(args.product_id) } : {}),
        ...(statuses?.length ? { supplierStatus: { in: statuses } } : {}),
        ...(args.supplier ? { supplier: String(args.supplier) } : {}),
        ...(args.destination ? { destination: String(args.destination) } : {}),
      };
      const limit = int(args.limit, "limit", { min: 1, max: 200 }) ?? 50;
      const [orders, total] = await Promise.all([
        prisma.supplierOrder.findMany({ where, select: orderSelect, orderBy: { createdAt: "desc" }, take: limit }),
        prisma.supplierOrder.count({ where }),
      ]);
      return { structured: { total_matching: total, returned: orders.length, orders: orders.map(presentOrder) } };
    },
  },
  {
    name: "get_production_order",
    title: "Get production order",
    description: "Full detail of one production order: sizes, quantities, costs, fabric, notes, shipment link, cut record and recent history.",
    inputSchema: { type: "object", properties: { order_id: { type: "integer" } }, required: ["order_id"], additionalProperties: false },
    scope: "portal:read",
    annotations: { readOnlyHint: true, openWorldHint: false },
    handler: async (args) => {
      const id = int(args.order_id, "order_id", { required: true, min: 1 })!;
      const order = await loadOrder(id);
      if (!order) throw new ToolInputError(`Order ${id} not found`);
      const [cut, packing, activity, preorder] = await Promise.all([
        prisma.productionCutRecord.findUnique({ where: { supplierOrderId: id } }),
        order.packingListId ? prisma.packingList.findUnique({ where: { id: order.packingListId }, select: { id: true, title: true, status: true, shipmentDate: true, shippingMethod: true } }) : null,
        prisma.activityLog.findMany({ where: { entityId: String(id), entity: "Restock Order" }, orderBy: { createdAt: "desc" }, take: 20 }),
        prisma.preorderBatchSetting.findUnique({ where: { supplierOrderId: id }, select: { enabled: true } }).catch(() => null),
      ]);
      return {
        structured: {
          ...presentOrder(order),
          live_preorder: preorder?.enabled === true,
          cut_record: cut ? { qty: cut.qty, production_date: cut.productionDate.toISOString().slice(0, 10), by: cut.createdByName } : null,
          shipment: packing,
          history: activity.map((a) => ({ at: a.createdAt.toISOString(), by: a.userName, action: a.action, field: a.field, value: a.toValue })),
        },
      };
    },
  },
  {
    name: "get_product_production_view",
    title: "Product stock, sales and orders",
    description: "Everything needed for a reorder decision on ONE product, per size: live Shopify stock, units sold in the window, quantities already on order or in production, ready stock in India, quantities in shipments, plus the open orders themselves. Always call this before proposing a reorder.",
    inputSchema: {
      type: "object",
      properties: {
        product_title: { type: "string" },
        product_id: { type: "string" },
        since: { type: "string", description: "YYYY-MM-DD, default 90 days ago" },
        until: { type: "string", description: "YYYY-MM-DD, default today" },
      },
      additionalProperties: false,
    },
    scope: "portal:read",
    annotations: { readOnlyHint: true, openWorldHint: false },
    handler: async (args) => {
      const product = await resolveProductLite(args);
      const since = dateArg(args.since, "since", isoDaysAgo(90));
      const until = dateArg(args.until, "until", today());
      const [plan, orders, shipmentLines] = await Promise.all([
        callPortalIntent("reorder_plan", { productId: product.id, productTitle: product.title, since, until }) as Promise<{ ok?: boolean; sizes?: Array<{ size: string; stock: number; unitsSold: number }>; salesAvailable?: boolean; error?: string } | null>,
        prisma.supplierOrder.findMany({ where: { productId: product.id, status: "open" }, select: orderSelect, orderBy: { createdAt: "desc" } }),
        prisma.packingListLine.findMany({
          where: { OR: [{ productId: product.id }, { productTitle: { equals: product.title, mode: "insensitive" } }], packingList: { hiddenAt: null, masterInventoryLoadedAt: null } },
          select: { qtys: true, packingList: { select: { id: true, title: true, status: true, shipmentDate: true, shippingMethod: true } } },
        }),
      ]);
      const bySize = new Map<string, { size: string; shopify_stock: number | null; units_sold: number | null; on_order: number; in_production: number; ready_in_india: number; in_shipment_orders: number; other_open_orders: number; in_packing_lists: number }>();
      const row = (size: string) => {
        const key = matchSize(size, Array.from(bySize.keys())) ?? size;
        if (!bySize.has(key)) bySize.set(key, { size: key, shopify_stock: null, units_sold: null, on_order: 0, in_production: 0, ready_in_india: 0, in_shipment_orders: 0, other_open_orders: 0, in_packing_lists: 0 });
        return bySize.get(key)!;
      };
      for (const s of product.sizes) row(s);
      for (const s of plan?.sizes ?? []) { const r = row(s.size); r.shopify_stock = s.stock; r.units_sold = s.unitsSold; }
      for (const o of orders) {
        for (const l of o.lines) {
          const r = row(l.variantTitle);
          const remaining = Math.max(0, l.qtyOrdered - l.qtyReceived);
          if (o.supplierStatus === "on_order") r.on_order += remaining;
          else if (o.supplierStatus === "on_production") r.in_production += remaining;
          else if (o.supplierStatus === "ready") r.ready_in_india += remaining;
          else if (o.supplierStatus === "in_shipment") r.in_shipment_orders += remaining;
          else if (o.supplierStatus !== "cancelled") r.other_open_orders += remaining;
        }
      }
      for (const line of shipmentLines) {
        const q = (line.qtys ?? {}) as Record<string, unknown>;
        for (const [size, qty] of Object.entries(q)) row(size).in_packing_lists += Number(qty) || 0;
      }
      return {
        structured: {
          product: { product_id: product.id, title: product.title },
          window: { since, until, sales_available: plan?.salesAvailable ?? false },
          sizes: Array.from(bySize.values()),
          open_orders: orders.map(presentOrder),
          shipments: Array.from(new Map(shipmentLines.map((l) => [l.packingList.id, l.packingList])).values()),
          notes: "on_order/in_production/ready_in_india/in_shipment_orders/other_open_orders (custom statuses such as packed_ready_to_leave or arrived_in_au) are remaining (ordered − received) on open orders; see each order's production_status. in_packing_lists counts shipment lines not yet loaded to Shopify; it can overlap with ANY open order status, including on_production and ready. Do not sum without verified allocation.",
        },
      };
    },
  },
  {
    name: "reorder_planner_overview",
    title: "Reorder Planner overview",
    description: "The portal's Reorder Planner: every active product ranked by weeks of stock cover (most urgent first), with stock, units sold and what's already on order. 50 products per page.",
    inputSchema: {
      type: "object",
      properties: {
        since: { type: "string", description: "YYYY-MM-DD, default 90 days ago" },
        until: { type: "string", description: "YYYY-MM-DD, default today" },
        search: { type: "string" },
        product_type: { type: "string" },
        tag: { type: "string" },
        hide_sale: { type: "boolean", default: true },
        page: { type: "integer", minimum: 1, default: 1 },
      },
      additionalProperties: false,
    },
    scope: "portal:read",
    annotations: { readOnlyHint: true, openWorldHint: false },
    handler: async (args) => {
      const res = await callPortalIntent("reorder_overview", {
        since: dateArg(args.since, "since", isoDaysAgo(90)),
        until: dateArg(args.until, "until", today()),
        q: str(args.search, "search", { max: 200 }) ?? "",
        type: str(args.product_type, "product_type") ?? "",
        tag: str(args.tag, "tag") ?? "",
        hideSale: args.hide_sale === false ? "" : "1",
        page: int(args.page, "page", { min: 1 }) ?? 1,
      });
      return { structured: compactForModel(res) };
    },
  },
  {
    name: "list_shipments",
    title: "List shipments",
    description: "Packing lists (shipments from the factory): status, dates, shipping method, total pieces and whether they've been loaded into Shopify stock.",
    inputSchema: { type: "object", properties: { include_loaded: { type: "boolean", default: false }, limit: { type: "integer", minimum: 1, maximum: 100, default: 30 } }, additionalProperties: false },
    scope: "portal:read",
    annotations: { readOnlyHint: true, openWorldHint: false },
    handler: async (args) => {
      const lists = await prisma.packingList.findMany({
        where: { hiddenAt: null, ...(args.include_loaded ? {} : { masterInventoryLoadedAt: null }) },
        orderBy: { createdAt: "desc" },
        take: int(args.limit, "limit", { min: 1, max: 100 }) ?? 30,
        select: { id: true, title: true, invoiceNumber: true, status: true, shipmentDate: true, expectedLeaveFactoryDate: true, shippingMethod: true, masterInventoryLoadedAt: true, notes: true, lines: { select: { qtys: true } } },
      });
      return {
        structured: {
          shipments: lists.map((l) => ({
            shipment_id: l.id, title: l.title, invoice: l.invoiceNumber, status: l.status,
            shipment_date: l.shipmentDate?.toISOString().slice(0, 10) ?? null,
            leaves_factory: l.expectedLeaveFactoryDate?.toISOString().slice(0, 10) ?? null,
            shipping_method: l.shippingMethod, loaded_to_shopify: Boolean(l.masterInventoryLoadedAt), notes: l.notes,
            products: l.lines.length,
            pieces: l.lines.reduce((sum, line) => sum + Object.values((line.qtys ?? {}) as Record<string, unknown>).reduce((a: number, q) => a + (Number(q) || 0), 0), 0),
          })),
        },
      };
    },
  },
  {
    name: "get_shipment",
    title: "Get shipment",
    description: "One packing list in detail: every product line with per-size quantities, box, price in rupees and weight.",
    inputSchema: { type: "object", properties: { shipment_id: { type: "integer" } }, required: ["shipment_id"], additionalProperties: false },
    scope: "portal:read",
    annotations: { readOnlyHint: true, openWorldHint: false },
    handler: async (args) => {
      const id = int(args.shipment_id, "shipment_id", { required: true, min: 1 })!;
      const list = await prisma.packingList.findUnique({
        where: { id },
        select: {
          id: true, title: true, invoiceNumber: true, status: true, shipmentDate: true, shippingMethod: true, masterInventoryLoadedAt: true, lockedFxRate: true, notes: true,
          lines: { orderBy: { sortOrder: "asc" }, select: { boxNumber: true, productId: true, productTitle: true, sku: true, qtys: true, shopifyLoadedQtys: true, priceRupees: true, weight: true, notes: true } },
        },
      });
      if (!list) throw new ToolInputError(`Shipment ${id} not found`);
      return { structured: compactForModel(list) };
    },
  },
  {
    name: "get_fabric_availability",
    title: "Fabric availability",
    description: "Fabric stock and consumption from the portal's fabric planner. With product_title: the fabrics that product can use, meters available and meters per piece. Without: meters reserved by on-order products per fabric.",
    inputSchema: { type: "object", properties: { product_title: { type: "string" } }, additionalProperties: false },
    scope: "portal:read",
    annotations: { readOnlyHint: true, openWorldHint: false },
    handler: async (args) => {
      const title = str(args.product_title, "product_title", { max: 200 });
      const res = await readFabric(title ? { title } : { pending: true });
      return { structured: compactForModel(res) };
    },
  },
  {
    name: "list_new_collections",
    title: "New collections",
    description: "New-collection planning sheets in the portal (names, size of each sheet, last update). Use get_new_collection for the rows.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
    scope: "portal:read",
    annotations: { readOnlyHint: true, openWorldHint: false },
    handler: async () => {
      const rows = await prisma.$queryRaw<Array<{ id: number; name: string; kind: string; rows: number; updatedAt: Date }>>`
        SELECT "id", "name", "kind", COALESCE(jsonb_array_length("rows"), 0)::int AS "rows", "updatedAt" FROM "Collection" ORDER BY "sortOrder" ASC, "updatedAt" DESC LIMIT 200`;
      return { structured: { collections: rows.map((r) => ({ collection_id: r.id, name: r.name, kind: r.kind, rows: r.rows, updated_at: r.updatedAt.toISOString() })) } };
    },
  },
  {
    name: "get_new_collection",
    title: "Get new collection",
    description: "Columns and rows of one new-collection sheet (images removed, long cells shortened).",
    inputSchema: { type: "object", properties: { collection_id: { type: "integer" }, max_rows: { type: "integer", minimum: 1, maximum: 500, default: 200 } }, required: ["collection_id"], additionalProperties: false },
    scope: "portal:read",
    annotations: { readOnlyHint: true, openWorldHint: false },
    handler: async (args) => {
      const id = int(args.collection_id, "collection_id", { required: true, min: 1 })!;
      const c = await prisma.collection.findUnique({ where: { id }, select: { id: true, name: true, kind: true, columns: true, rows: true, updatedAt: true } });
      if (!c) throw new ToolInputError(`Collection ${id} not found`);
      const max = int(args.max_rows, "max_rows", { min: 1, max: 500 }) ?? 200;
      const rows = Array.isArray(c.rows) ? c.rows : [];
      return { structured: compactForModel({ ...c, total_rows: rows.length, rows: rows.slice(0, max) }) };
    },
  },
  {
    name: "recent_activity",
    title: "Recent activity",
    description: "Recent changes in the portal (who changed what) and every ChatGPT action with before/after values.",
    inputSchema: { type: "object", properties: { limit: { type: "integer", minimum: 1, maximum: 200, default: 50 }, order_id: { type: "integer" } }, additionalProperties: false },
    scope: "portal:read",
    annotations: { readOnlyHint: true, openWorldHint: false },
    handler: async (args) => {
      const limit = int(args.limit, "limit", { min: 1, max: 200 }) ?? 50;
      const orderId = int(args.order_id, "order_id", { min: 1 });
      const [portal, ai] = await Promise.all([
        prisma.activityLog.findMany({ where: orderId ? { entityId: String(orderId) } : {}, orderBy: { createdAt: "desc" }, take: limit }),
        prisma.mcpAuditLog.findMany({ where: orderId ? { entityId: String(orderId) } : {}, orderBy: { createdAt: "desc" }, take: limit }),
      ]);
      return {
        structured: {
          portal_activity: portal.map((a) => ({ at: a.createdAt.toISOString(), by: a.userName, action: a.action, entity: a.entity, entity_id: a.entityId, name: a.entityName, field: a.field, value: a.toValue })),
          chatgpt_actions: ai.map((a) => ({ at: a.createdAt.toISOString(), tool: a.tool, outcome: a.outcome, change_id: a.changeId, entity_id: a.entityId, before: a.before, after: a.after, error: a.error })),
        },
      };
    },
  },
  {
    name: "list_pending_changes",
    title: "Pending changes",
    description: "Previewed changes still waiting for Koku's approval.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
    scope: "portal:read",
    annotations: { readOnlyHint: true, openWorldHint: false },
    handler: async (_args, ctx) => {
      const rows = await prisma.mcpPendingChange.findMany({ where: { grantId: ctx.principal.grantId, status: "pending", expiresAt: { gt: new Date() } }, orderBy: { createdAt: "desc" }, take: 50 });
      return { structured: { pending: rows.map((r) => ({ change_id: r.id, kind: r.kind, summary: r.summary, expires_at: r.expiresAt.toISOString() })) } };
    },
  },
];

// ─── Preview (write step 1) ─────────────────────────────────────────────────

const PREVIEW_NOTE = "Nothing is changed by this tool. It returns a change_id that needs Koku's explicit approval and a confirm_change call.";

async function deletableOrderSnapshot(id: number) {
  const order = await loadOrder(id);
  if (!order) throw new ToolInputError("Order not found");
  if (order.status !== "open" || order.supplierStatus !== "on_order" || order.fabricConsumed || order.packingListId || order.lines.some(l => l.qtyReceived > 0)) {
    throw new ToolInputError("Only open, uncut, unshipped orders with no received pieces can be deleted through this tool.");
  }
  const [batch, reservations, cut] = await Promise.all([
    prisma.preorderBatchSetting.findUnique({ where: { supplierOrderId: id } }),
    prisma.preorderReservation.count({ where: { supplierOrderId: id } }),
    prisma.productionCutRecord.findUnique({ where: { supplierOrderId: id } }),
  ]);
  if (batch || reservations || cut) throw new ToolInputError("Order has preorder or cutting history; review in portal before deleting.");
  return presentOrder(order);
}

const previewTools: ToolDef[] = [
  {
    name: "preview_create_fabric_order",
    title: "Preview recording a fabric order",
    description: `Record fabric and ordered meters on a selected portal order sheet. This does not send a supplier purchase order. ${PREVIEW_NOTE}`,
    inputSchema: {
      type: "object",
      properties: { sheet_id: { type: "string" }, fabric: { type: "string", minLength: 2, maxLength: 200 }, meters: { type: "number", minimum: 0.001, maximum: 1000000 } },
      required: ["sheet_id", "fabric", "meters"],
      additionalProperties: false,
    },
    scope: "portal:write",
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    handler: async (args, ctx) => {
      const sheetId = str(args.sheet_id, "sheet_id", { required: true, max: 200 })!;
      const fabric = str(args.fabric, "fabric", { required: true, max: 200 })!;
      const meters = args.meters;
      if (fabric.length < 2 || typeof meters !== "number" || !Number.isFinite(meters) || meters < 0.001 || meters > 1000000) {
        throw new ToolInputError("Enter a fabric name and a quantity from 0.001 to 1000000 meters");
      }
      const current = await readFabricOrderSheetSnapshot(sheetId, fabric);
      if (!current) throw new ToolInputError("Pick a fabric order sheet from list_fabric_order_sheets");
      const summary = `Record ${meters} meters of ${fabric} on ${current.name}.`;
      return stagePendingChange(ctx, "create_fabric_order", { sheetId, fabric, meters, expected: JSON.stringify(current) }, current,
        summary, sheetId, { sheet: current.name, fabric, meters, existing_matching_quantities: current.matchingQuantities });
    },
  },
  {
    name: "preview_staff_thread_reply",
    title: "Preview a staff thread reply",
    description: `Preview a reply to an existing portal message. Shows the exact text and everyone the portal would notify. ${PREVIEW_NOTE}`,
    inputSchema: {
      type: "object",
      properties: { parent_message_id: { type: "integer", minimum: 1 }, message: { type: "string", minLength: 1, maxLength: 2000 } },
      required: ["parent_message_id", "message"],
      additionalProperties: false,
    },
    scope: "portal:write",
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    handler: async (args, ctx) => {
      const parentMessageId = int(args.parent_message_id, "parent_message_id", { required: true, min: 1 })!;
      const body = str(args.message, "message", { required: true, max: 2000 })!;
      const resolved = await resolveReplyRecipients({ parentMessageId, body, fromUserId: MCP_PORTAL_USER_ID });
      if (!resolved) throw new ToolInputError("Portal message not found");
      const recipients = [...resolved.recipientsByUserId.values()].map(({ id, name }) => ({ id, name }));
      if (!recipients.length) throw new ToolInputError("This reply would notify no staff");
      const summary = `Reply to portal message #${parentMessageId} and notify ${recipients.map(({ name }) => name).join(", ")}.`;
      const before = { parentBody: resolved.parent.body };
      return stagePendingChange(ctx, "staff_reply", { parentMessageId, body, recipientIds: recipients.map(({ id }) => id).sort() },
        before, summary, String(parentMessageId), { parent_message_id: parentMessageId, recipients, message: body });
    },
  },
  {
    name: "preview_delete_production_order",
    title: "Preview deletion of an uncut production order",
    description: "Permanently deletes an open on-order batch through the portal's existing deletion action, only after explicit approval. Blocks cut, shipped, received or preorder-linked batches. " + PREVIEW_NOTE,
    inputSchema: { type: "object", properties: { order_id: { type: "integer", minimum: 1 } }, required: ["order_id"], additionalProperties: false },
    scope: "portal:write",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false },
    handler: async (args, ctx) => {
      const id = int(args.order_id, "order_id", { required: true, min: 1 })!;
      const before = await deletableOrderSnapshot(id);
      const summary = "Permanently delete order #" + id + " " + before.product + " (" + before.total_qty + " pieces), including its attached messages.";
      return stagePendingChange(ctx, "delete_order", { orderId: id }, before, summary, String(id), { before, after: null, warning: "Deletion removes attached messages; restoration may not preserve every field." });
    },
  },
  {
    name: "preview_create_production_order",
    title: "Preview new production order",
    description: `Prepares a NEW production / reorder order for a product with per-size quantities. Refuses if the product already has an open order that is on order or in production (amend that one instead) unless allow_duplicate is true. ${PREVIEW_NOTE}`,
    inputSchema: {
      type: "object",
      properties: {
        product_title: { type: "string" },
        product_id: { type: "string" },
        quantities: { type: "object", additionalProperties: { type: "integer", minimum: 0 }, description: "Size → pieces, e.g. {\"3XL\": 5}" },
        supplier: { type: "string", description: "Default: Karma East" },
        priority: { type: "string" },
        eta: { type: "string", description: "YYYY-MM-DD" },
        destination: { type: "string" },
        notes: { type: "string", maxLength: 2000 },
        allow_duplicate: { type: "boolean", default: false },
      },
      required: ["quantities"],
      additionalProperties: false,
    },
    scope: "portal:write",
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    handler: async (args, ctx) => {
      const product = await resolveProduct(args);
      const sizes = product.variants.map((v) => v.title);
      const q = resolveNewOrderQuantities(sizes, (args.quantities ?? {}) as Record<string, number>);
      if (!q.ok) throw new ToolInputError(q.error);
      const opts = await restockOptions();
      const priority = str(args.priority, "priority");
      if (priority && !opts.priorities.includes(priority)) throw new ToolInputError(`priority must be one of ${opts.priorities.join(", ")}`);
      const destination = str(args.destination, "destination", { max: 64 });
      if (destination && opts.destinations.length && !opts.destinations.includes(destination)) throw new ToolInputError(`destination must be one of ${opts.destinations.join(", ")}`);
      const eta = str(args.eta, "eta");
      if (eta && !isValidIsoDate(eta)) throw new ToolInputError("eta must be a date like 2026-12-20");
      const supplier = str(args.supplier, "supplier", { max: 100 }) ?? "Karma East";
      const notes = str(args.notes, "notes", { max: 2000 });

      const existing = await prisma.supplierOrder.findMany({ where: { productId: product.id }, select: { id: true, supplierStatus: true, status: true, totalQty: true, createdAt: true } });
      const dups = findPotentialDuplicates(existing);
      if (dups.length && args.allow_duplicate !== true) {
        await audit({ ctx, tool: "preview:create_order", outcome: "blocked_duplicate", kind: "create_order", entityType: "Product", entityId: product.id, input: args, before: dups });
        return {
          isError: true,
          structured: {
            status: "blocked_possible_duplicate",
            product: product.title,
            existing_open_orders: dups.map((d) => ({ order_id: d.id, production_status: d.supplierStatus, total_qty: d.totalQty })),
            advice: "This product already has an open order on order / in production. Prefer preview_amend_order_quantities on that order. Only if Koku explicitly wants a separate order, call again with allow_duplicate: true.",
          },
          text: `Blocked: ${product.title} already has open order(s) ${dups.map((d) => `#${d.id} (${d.supplierStatus}, ${d.totalQty} pcs)`).join(", ")}. Amend instead, or ask Koku whether he really wants a second order.`,
        };
      }
      const payload = { productId: product.id, productTitle: product.title, shop: product.shop ?? null, imageUrl: product.imageUrl, variants: product.variants, qtys: q.qtys, supplier, priority: priority ?? null, eta: eta ?? null, destination: destination ?? null, notes: notes ?? null, allowDuplicate: args.allow_duplicate === true };
      const summary = `Create a new ${supplier} order for ${product.title}: ${Object.entries(q.qtys).map(([s, n]) => `${s} × ${n}`).join(", ")} (total ${q.total} pcs), status on_order${priority ? `, priority ${priority}` : ""}${eta ? `, ETA ${eta}` : ""}${destination ? `, destination ${destination}` : ""}.`;
      return stagePendingChange(ctx, "create_order", payload, { existing_open_orders: dups.map((d) => d.id) }, summary, null, { product: product.title, quantities: q.qtys, total_qty: q.total, supplier, priority, eta, destination, notes, other_open_orders: dups.map((d) => d.id) });
    },
  },
  {
    name: "preview_amend_order_quantities",
    title: "Preview size quantity change",
    description: `Prepares a change to the per-size quantities of an existing order. For each size give either add (e.g. +5) or set_to (absolute). Shows before → after. ${PREVIEW_NOTE}`,
    inputSchema: {
      type: "object",
      properties: {
        order_id: { type: "integer" },
        changes: {
          type: "array",
          minItems: 1,
          items: { type: "object", properties: { size: { type: "string" }, add: { type: "integer" }, set_to: { type: "integer", minimum: 0 } }, required: ["size"], additionalProperties: false },
        },
      },
      required: ["order_id", "changes"],
      additionalProperties: false,
    },
    scope: "portal:write",
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    handler: async (args, ctx) => {
      const id = int(args.order_id, "order_id", { required: true, min: 1 })!;
      const order = await loadOrder(id);
      if (!order) throw new ToolInputError(`Order ${id} not found`);
      if (order.status !== "open") throw new ToolInputError(`Order ${id} is closed and can't be amended`);
      if (["in_shipment", "cancelled"].includes(order.supplierStatus)) throw new ToolInputError(`Order ${id} is ${order.supplierStatus}; amend a different order or create a new one`);
      const current = sizesOf(order);
      const raw = (Array.isArray(args.changes) ? args.changes : []) as Array<Record<string, unknown>>;
      const changes: SizeQtyChange[] = raw.map((c) => ({ size: String(c.size ?? ""), add: c.add as number | undefined, setTo: c.set_to as number | undefined }));
      let resolved = resolveQuantityChanges(current, Object.keys(current), changes);
      // A size not yet on the order: check the product's real sizes in Shopify.
      if (!resolved.ok && /not one of this product's sizes/.test(resolved.error)) {
        try {
          const product = await resolveProduct({ product_id: order.productId });
          resolved = resolveQuantityChanges(current, Array.from(new Set([...Object.keys(current), ...product.variants.map((v) => v.title)])), changes);
        } catch { /* keep the original error */ }
      }
      if (!resolved.ok) throw new ToolInputError(resolved.error);
      const effective = resolved.changes.filter((c) => c.before !== c.after);
      if (!effective.length) throw new ToolInputError("These quantities are already on the order; nothing to change");
      const totalBefore = order.totalQty;
      const totalAfter = totalBefore + effective.reduce((a, c) => a + (c.after - c.before), 0);
      const summary = `Order #${id} ${order.productTitle} (${order.supplierStatus}): ${effective.map((c) => `${c.size} ${c.before} → ${c.after}`).join(", ")}. Total ${totalBefore} → ${totalAfter} pcs.`;
      return stagePendingChange(ctx, "amend_quantities", { orderId: id, changes: effective }, { sizes: current, totalQty: totalBefore, supplierStatus: order.supplierStatus }, summary, String(id), { order_id: id, product: order.productTitle, production_status: order.supplierStatus, changes: effective, total_before: totalBefore, total_after: totalAfter, warning: order.supplierStatus === "on_production" ? "This order is already in production; check with the factory that extra pieces can still be cut." : undefined });
    },
  },
  {
    name: "preview_update_order",
    title: "Preview order status/priority/ETA change",
    description: `Prepares a change to an order's production status, priority, ETA or destination. Status changes run the portal's usual side effects (fabric deduction when going on production, cut log). ${PREVIEW_NOTE}`,
    inputSchema: {
      type: "object",
      properties: {
        order_id: { type: "integer" },
        production_status: { type: "string", description: "One of the statuses from get_portal_options, e.g. on_production" },
        priority: { type: "string", description: "Empty string clears it" },
        eta: { type: "string", description: "YYYY-MM-DD, or empty string to clear" },
        destination: { type: "string" },
      },
      required: ["order_id"],
      additionalProperties: false,
    },
    scope: "portal:write",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true, openWorldHint: false },
    handler: async (args, ctx) => {
      const id = int(args.order_id, "order_id", { required: true, min: 1 })!;
      const order = await loadOrder(id);
      if (!order) throw new ToolInputError(`Order ${id} not found`);
      const opts = await restockOptions();
      const set: Record<string, string> = {};
      const before: Record<string, string | null> = {};
      if (args.production_status !== undefined) {
        const s = String(args.production_status);
        if (!opts.statuses.includes(s)) throw new ToolInputError(`production_status must be one of ${opts.statuses.join(", ")}`);
        if (s !== order.supplierStatus) { set.status = s; before.status = order.supplierStatus; }
      }
      if (args.priority !== undefined) {
        const p = String(args.priority);
        if (p && !opts.priorities.includes(p)) throw new ToolInputError(`priority must be one of ${opts.priorities.join(", ")}`);
        if (p !== (order.priority ?? "")) { set.priority = p; before.priority = order.priority; }
      }
      if (args.eta !== undefined) {
        const e = String(args.eta);
        if (e && !isValidIsoDate(e)) throw new ToolInputError("eta must be a date like 2026-12-20");
        const cur = order.eta?.toISOString().slice(0, 10) ?? "";
        if (e !== cur) { set.eta = e; before.eta = cur || null; }
      }
      if (args.destination !== undefined) {
        const d = String(args.destination).trim().slice(0, 64);
        if (d && opts.destinations.length && !opts.destinations.includes(d)) throw new ToolInputError(`destination must be one of ${opts.destinations.join(", ")}`);
        if (d !== (order.destination ?? "")) { set.destination = d; before.destination = order.destination; }
      }
      if (!Object.keys(set).length) throw new ToolInputError("Nothing to change (values are already the same)");
      const summary = `Order #${id} ${order.productTitle}: ${Object.entries(set).map(([k, v]) => `${k} ${before[k] ?? "(none)"} → ${v || "(cleared)"}`).join(", ")}.`;
      const warnings: string[] = [];
      if (set.status === "on_production") warnings.push("Going On Production deducts this order's fabric from stock and records a production cut.");
      if (set.status === "cancelled") warnings.push("Cancelling stops this order counting as incoming stock.");
      return stagePendingChange(ctx, "update_order", { orderId: id, set }, before, summary, String(id), { order_id: id, product: order.productTitle, before, after: set, warnings });
    },
  },
  {
    name: "preview_replace_order_note",
    title: "Preview replacement or removal of an order note",
    description: `Replace the entire selected note field, or clear it with an empty string. Shows before and after. ${PREVIEW_NOTE}`,
    inputSchema: { type: "object", properties: { order_id: { type: "integer" }, text: { type: "string", maxLength: 2000 }, field: { type: "string", enum: ["notes", "factory_notes"], default: "notes" } }, required: ["order_id", "text"], additionalProperties: false },
    scope: "portal:write",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false },
    handler: async (args, ctx) => {
      const id = int(args.order_id, "order_id", { required: true, min: 1 })!;
      if (typeof args.text !== "string" || args.text.length > 2000) throw new ToolInputError("text must be a string of at most 2000 characters; use an empty string to clear");
      if (args.field !== undefined && args.field !== "notes" && args.field !== "factory_notes") throw new ToolInputError("Invalid note field");
      const field = args.field === "factory_notes" ? "factory_notes" : "notes";
      const order = await loadOrder(id);
      if (!order) throw new ToolInputError(`Order ${id} not found`);
      const current = field === "factory_notes" ? order.factoryNotes : order.notes;
      const next = args.text;
      if ((current ?? "") === next) throw new ToolInputError("Nothing to change");
      const summary = `${next === "" ? "Clear" : "Replace"} ${field} on order #${id} ${order.productTitle}`;
      return stagePendingChange(ctx, "replace_note", { orderId: id, field, text: next }, { [field]: current ?? null }, summary, String(id), { order_id: id, field, before: current ?? null, after: next });
    },
  },
  {
    name: "preview_add_order_note",
    title: "Preview note on an order",
    description: `Prepares a note or recommendation added to an order's notes (or factory notes, which the factory sees). Existing text is kept; the note is appended with a dated [ChatGPT] tag. ${PREVIEW_NOTE}`,
    inputSchema: {
      type: "object",
      properties: { order_id: { type: "integer" }, note: { type: "string", minLength: 1, maxLength: 2000 }, field: { type: "string", enum: ["notes", "factory_notes"], default: "notes" } },
      required: ["order_id", "note"],
      additionalProperties: false,
    },
    scope: "portal:write",
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    handler: async (args, ctx) => {
      const id = int(args.order_id, "order_id", { required: true, min: 1 })!;
      const note = str(args.note, "note", { required: true, max: 2000 })!;
      const field = args.field === "factory_notes" ? "factory_notes" : "notes";
      const order = await loadOrder(id);
      if (!order) throw new ToolInputError(`Order ${id} not found`);
      const current = field === "factory_notes" ? order.factoryNotes : order.notes;
      const next = appendNote(current, note);
      const summary = `Add to ${field === "factory_notes" ? "factory notes (visible to the factory)" : "notes"} on order #${id} ${order.productTitle}: "${note}"`;
      return stagePendingChange(ctx, "add_note", { orderId: id, field, text: next }, { [field]: current ?? null }, summary, String(id), { order_id: id, field, before: current ?? null, after: next });
    },
  },
];

// ─── Confirm (write step 2) ─────────────────────────────────────────────────

async function executeCreate(p: Record<string, unknown>) {
  if (p.allowDuplicate !== true) {
    const existing = await prisma.supplierOrder.findMany({ where: { productId: String(p.productId) }, select: { id: true, supplierStatus: true, status: true, totalQty: true, createdAt: true } });
    const dups = findPotentialDuplicates(existing);
    if (dups.length) throw new ToolInputError(`Another open order (#${dups.map((d) => d.id).join(", #")}) appeared for this product since the preview. Re-check and preview again.`);
  }
  const res = (await callPortalIntent("create_restock_order_from_portal", {
    product: JSON.stringify({ id: p.productId, title: p.productTitle, shop: p.shop ?? undefined, imageUrl: p.imageUrl ?? null, skus: [], sizes: [], variants: p.variants }),
    qtys: JSON.stringify(p.qtys),
    supplier: String(p.supplier ?? "Karma East"),
    supplierStatus: "on_order",
    priority: (p.priority as string | null) ?? "",
    eta: (p.eta as string | null) ?? "",
    notes: (p.notes as string | null) ?? "",
  })) as { success?: boolean; orderId?: number } | null;
  if (!res?.success || !res.orderId) throw new Error("The portal did not create the order (it rejected the input). Nothing was saved.");
  if (p.destination) await callPortalIntent("update_destination", { orderId: res.orderId, value: String(p.destination) });
  const after = await loadOrder(res.orderId);
  return { entityId: String(res.orderId), after: after ? presentOrder(after) : null };
}

async function executeAmend(p: { orderId: number; changes: ResolvedSizeChange[] }, before: { sizes: Record<string, number>; totalQty: number }) {
  const order = await loadOrder(p.orderId);
  if (!order) throw new ToolInputError(`Order ${p.orderId} no longer exists`);
  const now = sizesOf(order);
  for (const c of p.changes) {
    if ((now[c.size] ?? 0) !== c.before) throw new ToolInputError(`Order #${p.orderId} ${c.size} is now ${now[c.size] ?? 0}, not ${c.before} as previewed. Preview again.`);
  }
  if (stableStringify(now) !== stableStringify(before.sizes)) throw new ToolInputError(`Order #${p.orderId} quantities changed since the preview. Preview again.`);
  for (const c of p.changes) await callPortalIntent("update_qty", { orderId: p.orderId, size: c.size, value: c.after });
  const after = await loadOrder(p.orderId);
  const afterSizes = after ? sizesOf(after) : {};
  const mismatch = p.changes.filter((c) => (afterSizes[c.size] ?? 0) !== c.after);
  if (mismatch.length) throw new Error(`Portal saved different values for ${mismatch.map((m) => m.size).join(", ")}; check order #${p.orderId}`);
  return { entityId: String(p.orderId), after: after ? presentOrder(after) : null };
}

async function executeUpdate(p: { orderId: number; set: Record<string, string> }, before: Record<string, string | null>) {
  const order = await loadOrder(p.orderId);
  if (!order) throw new ToolInputError(`Order ${p.orderId} no longer exists`);
  const current: Record<string, string | null> = { status: order.supplierStatus, priority: order.priority, eta: order.eta?.toISOString().slice(0, 10) ?? null, destination: order.destination };
  for (const k of Object.keys(p.set)) {
    if ((current[k] ?? null) !== (before[k] ?? null)) throw new ToolInputError(`Order #${p.orderId} ${k} changed since the preview (now ${current[k] ?? "(none)"}). Preview again.`);
  }
  const intentFor: Record<string, string> = { status: "update_status", priority: "update_priority", eta: "update_eta", destination: "update_destination" };
  for (const [k, v] of Object.entries(p.set)) await callPortalIntent(intentFor[k], { orderId: p.orderId, value: v });
  const after = await loadOrder(p.orderId);
  return { entityId: String(p.orderId), after: after ? presentOrder(after) : null };
}

async function executeNote(p: { orderId: number; field: string; text: string }, before: Record<string, string | null>) {
  const order = await loadOrder(p.orderId);
  if (!order) throw new ToolInputError(`Order ${p.orderId} no longer exists`);
  const current = p.field === "factory_notes" ? order.factoryNotes : order.notes;
  if ((current ?? null) !== (before[p.field] ?? null)) throw new ToolInputError(`The ${p.field} on order #${p.orderId} changed since the preview. Preview again.`);
  await callPortalIntent(p.field === "factory_notes" ? "update_factory_notes" : "update_notes", { orderId: p.orderId, value: p.text });
  const after = await loadOrder(p.orderId);
  if (!after || ((p.field === "factory_notes" ? after.factoryNotes : after.notes) ?? "") !== p.text) throw new Error("Portal did not save the expected note; check the order");
  return { entityId: String(p.orderId), after: presentOrder(after) };
}

const confirmTools: ToolDef[] = [
  {
    name: "confirm_change",
    title: "Apply an approved change",
    description: "Applies a previewed change to the Production Portal. ONLY call this after Koku has seen the preview in this conversation and explicitly said yes to it. Each change_id works once.",
    inputSchema: { type: "object", properties: { change_id: { type: "string" } }, required: ["change_id"], additionalProperties: false },
    scope: "portal:write",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true, openWorldHint: false },
    handler: async (args, ctx) => {
      const id = str(args.change_id, "change_id", { required: true, max: 100 })!;
      const change = await prisma.mcpPendingChange.findUnique({ where: { id } });
      if (!change || change.grantId !== ctx.principal.grantId) throw new ToolInputError("Unknown change_id");
      if (change.status === "done") return { structured: { status: "already_applied", change_id: id, result: change.result }, text: "This change was already applied earlier; nothing was done again." };
      if (change.status !== "pending") throw new ToolInputError(`This change is ${change.status} and can't be applied`);
      if (change.expiresAt.getTime() < Date.now()) {
        await prisma.mcpPendingChange.updateMany({ where: { id, status: "pending" }, data: { status: "expired" } });
        throw new ToolInputError("This preview expired (30 minutes). Preview again.");
      }
      // Claim it: only one confirm can move pending → executing.
      const claimed = await prisma.mcpPendingChange.updateMany({ where: { id, status: "pending" }, data: { status: "executing" } });
      if (claimed.count !== 1) throw new ToolInputError("This change is already being applied");

      const payload = change.payload as Record<string, unknown>;
      const before = (change.before ?? {}) as Record<string, unknown>;
      try {
        let result: { entityId: string; after: unknown };
        if (change.kind === "create_order") result = await executeCreate(payload);
        else if (change.kind === "amend_quantities") result = await executeAmend(payload as never, before as never);
        else if (change.kind === "create_fabric_order") {
          const { sheetId, fabric, meters, expected } = payload as { sheetId: string; fabric: string; meters: number; expected: string };
          const after = await appendFabricOrderFromMcp(sheetId, fabric, meters, expected);
          result = { entityId: sheetId, after };
        }
        else if (change.kind === "staff_reply") {
          const { parentMessageId, body, recipientIds } = payload as { parentMessageId: number; body: string; recipientIds: string[] };
          const parent = await prisma.portalMessage.findUnique({ where: { id: parentMessageId }, select: { body: true } });
          if (!parent || stableStringify({ parentBody: parent.body }) !== stableStringify(before)) {
            throw new ToolInputError("Parent message changed since the preview. Preview the reply again.");
          }
          await ensureMcpPortalUser();
          const { replyIds } = await createReplyMessage({
            parentMessageId, body, fromUserId: MCP_PORTAL_USER_ID, fromName: MCP_PORTAL_USER_NAME,
            expectedRecipientIds: recipientIds,
          });
          if (!replyIds.length) throw new ToolInputError("The portal did not save the reply");
          result = { entityId: String(parentMessageId), after: { body, reply_ids: replyIds } };
        }
        else if (change.kind === "delete_order") {
          const orderId = int(payload.orderId, "orderId", { required: true, min: 1 })!;
          const current = await deletableOrderSnapshot(orderId);
          if (stableStringify(current) !== stableStringify(before)) throw new ToolInputError("Order changed since preview; preview again.");
          await callPortalIntent("delete_order", { orderId });
          if (await loadOrder(orderId)) throw new Error("Order still exists; deletion not verified.");
          result = { entityId: String(orderId), after: { deleted: true, order_id: orderId } };
        }
        else if (change.kind === "update_order") result = await executeUpdate(payload as never, before as never);
        else if ((change.kind === "add_note" || change.kind === "replace_note")) result = await executeNote(payload as never, before as never);
        else throw new Error(`Unknown change kind ${change.kind}`);
        await prisma.mcpPendingChange.update({ where: { id }, data: { status: "done", executedAt: new Date(), result: result as never } });
        await audit({ ctx, tool: "confirm_change", outcome: "applied", kind: change.kind, changeId: id, entityType: "SupplierOrder", entityId: result.entityId, input: payload, before, after: result.after });
        return { structured: { status: "applied", change_id: id, summary: change.summary, order: result.after }, text: `Done: ${change.summary}` };
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        await prisma.mcpPendingChange.update({ where: { id }, data: { status: "failed", error: message, executedAt: new Date() } });
        await audit({ ctx, tool: "confirm_change", outcome: "failed", kind: change.kind, changeId: id, entityType: "SupplierOrder", input: payload, before, error: message });
        throw err instanceof ToolInputError ? err : new Error(`Change failed: ${message}`);
      }
    },
  },
  {
    name: "cancel_change",
    title: "Discard a previewed change",
    description: "Discards a previewed change Koku did not approve.",
    inputSchema: { type: "object", properties: { change_id: { type: "string" } }, required: ["change_id"], additionalProperties: false },
    scope: "portal:write",
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    handler: async (args, ctx) => {
      const id = str(args.change_id, "change_id", { required: true, max: 100 })!;
      const res = await prisma.mcpPendingChange.updateMany({ where: { id, grantId: ctx.principal.grantId, status: "pending" }, data: { status: "cancelled" } });
      if (res.count) await audit({ ctx, tool: "cancel_change", outcome: "cancelled", changeId: id });
      return { structured: { status: res.count ? "cancelled" : "not_pending", change_id: id } };
    },
  },
];

export const TOOLS: ToolDef[] = [...readTools, ...previewTools, ...confirmTools];

export async function runTool(name: string, args: Record<string, unknown>, ctx: ToolContext): Promise<ToolResult> {
  const tool = TOOLS.find((t) => t.name === name);
  if (!tool) return { isError: true, structured: { error: `Unknown tool ${name}` } };
  if (!ctx.principal.scopes.includes(tool.scope)) return { isError: true, structured: { error: `This connection lacks the ${tool.scope} permission` } };
  try {
    return await tool.handler(args ?? {}, ctx);
  } catch (err) {
    if (err instanceof ToolInputError) return { isError: true, structured: { error: err.message }, text: err.message };
    console.error(`MCP tool ${name} failed`, err);
    return { isError: true, structured: { error: err instanceof Error ? err.message : "Tool failed" } };
  }
}
