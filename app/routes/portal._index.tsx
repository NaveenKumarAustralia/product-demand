import { useEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";
import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import { useFetcher, useLoaderData, useSearchParams } from "react-router";
import prisma from "../db.server";
import { unauthenticated } from "../shopify.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const url = new URL(request.url);
  const page = url.searchParams.get("page") ?? "restock";
  const selectedProductGroup = normalizeProductGroup(
    url.searchParams.get("productGroup") ?? url.searchParams.get("productType") ?? "",
  );
  const selectedStatus = url.searchParams.get("status") ?? "";
  const selectedPriority = url.searchParams.get("priority") ?? "";
  const searchTitle = url.searchParams.get("q") ?? "";
  const packingId = Number(url.searchParams.get("packingId") ?? 0) || null;
  const productSearch = url.searchParams.get("productSearch") ?? "";
  const packingSearchLineId = Number(url.searchParams.get("packingSearchLineId") ?? 0) || null;
  const sortBy = url.searchParams.get("sortBy") ?? "orderDateDesc";
  const [allOrders, columnWidthsSetting, packingColumnWidthsSetting, loginRequiredSetting, usersSetting, activeUsersSetting, packingLists] = await Promise.all([
    prisma.supplierOrder.findMany({
      where: { status: "open" },
      include: { lines: { orderBy: { id: "asc" } } },
      orderBy: { createdAt: "desc" },
    }),
    prisma.portalSetting.findUnique({
      where: { key: COLUMN_WIDTHS_KEY },
      select: { value: true },
    }),
    prisma.portalSetting.findUnique({
      where: { key: PACKING_COLUMN_WIDTHS_KEY },
      select: { value: true },
    }),
    prisma.portalSetting.findUnique({
      where: { key: PORTAL_LOGIN_REQUIRED_KEY },
      select: { value: true },
    }),
    prisma.portalSetting.findUnique({
      where: { key: PORTAL_USERS_KEY },
      select: { value: true },
    }),
    prisma.portalSetting.findUnique({
      where: { key: PORTAL_ACTIVE_USERS_KEY },
      select: { value: true },
    }),
    prisma.packingList.findMany({
      orderBy: { createdAt: "desc" },
      include: { lines: { orderBy: [{ boxNumber: "asc" }, { sortOrder: "asc" }, { id: "asc" }] } },
    }),
  ]);
  const users = normalizePortalUsers(usersSetting?.value);
  const loginRequired = normalizeBooleanSetting(loginRequiredSetting?.value);
  const currentUser = getCurrentPortalUser(request, users);
  const activeUsers = await recordAndGetActiveUsers(currentUser, users, activeUsersSetting?.value);
  const normalizedOrders = allOrders.map((order) => ({
    ...order,
    productType: normalizeProductGroup(order.productType) || null,
  }));
  const productGroups = Array.from(new Set(normalizedOrders.map((order) => order.productType).filter(Boolean) as string[]))
    .sort((a, b) => a.localeCompare(b));
  const statusFilters = Array.from(new Set(normalizedOrders.map((order) => order.supplierStatus).filter(Boolean)))
    .sort((a, b) => labelForStatus(a).localeCompare(labelForStatus(b)));
  const priorityFilters = Array.from(new Set(normalizedOrders.map((order) => order.priority).filter(Boolean) as string[]))
    .sort((a, b) => labelForPriority(a).localeCompare(labelForPriority(b)));
  const orders = normalizedOrders
    .filter((order) => !selectedProductGroup || order.productType === selectedProductGroup)
    .filter((order) => !selectedStatus || order.supplierStatus === selectedStatus)
    .filter((order) => !selectedPriority || order.priority === selectedPriority)
    .filter((order) => !searchTitle || order.productTitle.toLowerCase().includes(searchTitle.toLowerCase()))
    .sort((a, b) => {
      if (sortBy === "titleAsc") return a.productTitle.localeCompare(b.productTitle);
      if (sortBy === "titleDesc") return b.productTitle.localeCompare(a.productTitle);
      if (sortBy === "orderDateAsc") return a.createdAt.getTime() - b.createdAt.getTime();
      return b.createdAt.getTime() - a.createdAt.getTime();
    });
  const selectedPackingList = packingId
    ? packingLists.find((list) => list.id === packingId) ?? null
    : null;
  const productResults = page === "packing" && selectedPackingList && productSearch.trim().length >= 2
    ? await searchShopifyProducts(productSearch)
    : [];

  // Collect all unique size names across all orders, sorted logically
  const sizeOrder = ["XS","S","S/M","M","M/L","L","L/XL","XL","2XL","3XL","4XL","ONE SIZE"];
  const allSizes = [...new Set(orders.flatMap((o) => o.lines.map((l) => l.variantTitle)))];
  allSizes.sort((a, b) => {
    const ai = sizeOrder.indexOf(a.toUpperCase());
    const bi = sizeOrder.indexOf(b.toUpperCase());
    if (ai === -1 && bi === -1) return a.localeCompare(b);
    if (ai === -1) return 1;
    if (bi === -1) return -1;
    return ai - bi;
  });

  return {
    orders,
    sizes: allSizes,
    productGroups,
    selectedProductGroup,
    selectedStatus,
    selectedPriority,
    searchTitle,
    statusFilters,
    priorityFilters,
    sortBy,
    page,
    columnWidths: normalizeColumnWidths(columnWidthsSetting?.value),
    packingColumnWidths: normalizeColumnWidths(packingColumnWidthsSetting?.value),
    packingLists,
    selectedPackingList,
    productSearch,
    packingSearchLineId,
    productResults,
    loginRequired,
    users,
    currentUser,
    activeUsers,
    loginBlocked: loginRequired && users.length > 0 && !currentUser,
  };
};

export const action = async ({ request }: ActionFunctionArgs) => {
  const form = await request.formData();
  const intent = String(form.get("intent"));
  const orderId = Number(form.get("orderId"));
  const usersSetting = await prisma.portalSetting.findUnique({
    where: { key: PORTAL_USERS_KEY },
    select: { value: true },
  });
  const loginRequiredSetting = await prisma.portalSetting.findUnique({
    where: { key: PORTAL_LOGIN_REQUIRED_KEY },
    select: { value: true },
  });
  const users = normalizePortalUsers(usersSetting?.value);
  const loginRequired = normalizeBooleanSetting(loginRequiredSetting?.value);
  const currentUser = getCurrentPortalUser(request, users);
  const canManageUsers = !loginRequired || users.length === 0 || currentUser?.admin;

  const updates: Record<string, unknown> = {};

  if (intent === "portal_login") {
    const userId = String(form.get("userId") ?? "");
    const user = users.find((item) => item.id === userId && item.active);
    if (!user) return null;
    return new Response(null, {
      status: 303,
      headers: {
        Location: "/portal",
        "Set-Cookie": `${PORTAL_USER_COOKIE}=${encodeURIComponent(user.id)}; Path=/; SameSite=Lax; Max-Age=2592000`,
      },
    });
  }

  if (intent === "portal_logout") {
    return new Response(null, {
      status: 303,
      headers: {
        Location: "/portal",
        "Set-Cookie": `${PORTAL_USER_COOKIE}=; Path=/; SameSite=Lax; Max-Age=0`,
      },
    });
  }

  if (intent === "update_login_required") {
    if (!canManageUsers) return null;
    await prisma.portalSetting.upsert({
      where: { key: PORTAL_LOGIN_REQUIRED_KEY },
      create: { key: PORTAL_LOGIN_REQUIRED_KEY, value: form.get("value") === "on" },
      update: { value: form.get("value") === "on" },
    });
    return null;
  }

  if (intent === "add_portal_user") {
    if (!canManageUsers) return null;
    const name = String(form.get("name") ?? "").trim();
    if (!name) return null;
    const nextUsers = [
      ...users,
      {
        id: crypto.randomUUID(),
        name,
        admin: form.get("admin") === "on",
        active: true,
      },
    ];
    await savePortalUsers(nextUsers);
    return null;
  }

  if (intent === "remove_portal_user") {
    if (!canManageUsers) return null;
    const userId = String(form.get("userId") ?? "");
    await savePortalUsers(users.filter((user) => user.id !== userId));
    return null;
  }

  if (loginRequired && users.length > 0 && !currentUser) {
    return null;
  }

  if (intent === "create_packing_list") {
    const title = String(form.get("title") ?? "").trim() || `Shipment ${formatPortalDate(new Date())}`;
    const invoiceNumber = String(form.get("invoiceNumber") ?? "").trim();
    const expectedLeaveFactoryDate = parsePortalDate(String(form.get("expectedLeaveFactoryDate") ?? ""));
    const packingList = await prisma.packingList.create({
      data: {
        title,
        invoiceNumber: invoiceNumber || null,
        shipmentDate: expectedLeaveFactoryDate,
        expectedLeaveFactoryDate,
        status: "still_packing",
        lines: {
          create: Array.from({ length: DEFAULT_PACKING_ROWS }, (_, index) => ({
            productTitle: "",
            isCustom: true,
            sortOrder: index + 1,
          })),
        },
      },
    });
    return new Response(null, {
      status: 303,
      headers: { Location: `/portal?page=packing&packingId=${packingList.id}` },
    });
  }

  if (intent === "update_packing_list") {
    const packingId = Number(form.get("packingId"));
    const field = String(form.get("field") ?? "");
    const value = String(form.get("value") ?? "");
    const data: Record<string, unknown> = {};
    if (field === "title") data.title = value.trim() || "Untitled shipment";
    if (field === "invoiceNumber") data.invoiceNumber = value.trim() || null;
    if (field === "status") data.status = value || "still_packing";
    if (field === "shipmentDate") {
      const parsedDate = value ? parsePortalDate(value) : null;
      if (value && !parsedDate) return null;
      data.shipmentDate = parsedDate;
    }
    if (field === "expectedLeaveFactoryDate") {
      const parsedDate = value ? parsePortalDate(value) : null;
      if (value && !parsedDate) return null;
      data.expectedLeaveFactoryDate = parsedDate;
      data.shipmentDate = parsedDate;
    }
    if (field === "notes") data.notes = value || null;
    if (packingId && Object.keys(data).length) {
      await prisma.packingList.update({ where: { id: packingId }, data });
    }
    return null;
  }

  if (intent === "add_custom_packing_line") {
    const packingId = Number(form.get("packingId"));
    const maxLine = await prisma.packingListLine.findFirst({
      where: { packingListId: packingId },
      orderBy: { sortOrder: "desc" },
      select: { sortOrder: true },
    });
    await prisma.packingListLine.create({
      data: {
        packingListId: packingId,
        productTitle: "",
        isCustom: true,
        sortOrder: (maxLine?.sortOrder ?? 0) + 1,
      },
    });
    return null;
  }

  if (intent === "add_product_packing_line") {
    const packingId = Number(form.get("packingId"));
    const product = JSON.parse(String(form.get("product") ?? "{}")) as ShopifySearchProduct;
    const maxLine = await prisma.packingListLine.findFirst({
      where: { packingListId: packingId },
      orderBy: { sortOrder: "desc" },
      select: { sortOrder: true },
    });
    await prisma.packingListLine.create({
      data: {
        packingListId: packingId,
        productId: product.id || null,
        productTitle: product.title || "Untitled product",
        productImageUrl: product.imageUrl || null,
        sku: product.skus?.filter(Boolean).join("\n") || null,
        qtys: Object.fromEntries((product.sizes ?? []).map((size) => [size, 0])),
        sortOrder: (maxLine?.sortOrder ?? 0) + 1,
      },
    });
    return null;
  }

  if (intent === "apply_product_to_packing_line") {
    const lineId = Number(form.get("lineId"));
    const product = JSON.parse(String(form.get("product") ?? "{}")) as ShopifySearchProduct;
    if (!lineId || !product?.id) return null;
    await prisma.packingListLine.update({
      where: { id: lineId },
      data: {
        productId: product.id,
        productTitle: product.title || "Untitled product",
        productImageUrl: product.imageUrl || null,
        sku: product.skus?.filter(Boolean).join("\n") || null,
        isCustom: false,
        qtys: Object.fromEntries((product.sizes ?? []).map((size) => [size, 0])),
      },
    });
    return null;
  }

  if (intent === "update_packing_line") {
    const lineId = Number(form.get("lineId"));
    const field = String(form.get("field") ?? "");
    const value = String(form.get("value") ?? "");
    const data: Record<string, unknown> = {};
    if (field === "boxNumber") data.boxNumber = value || null;
    if (field === "productTitle") data.productTitle = value;
    if (field === "sku") data.sku = value || null;
    if (field === "priceRupees") data.priceRupees = value ? Number(value) || 0 : null;
    if (field === "weight") data.weight = value ? Number(value) || 0 : null;
    if (field === "notes") data.notes = value || null;
    if (field === "fabricImageData") data.fabricImageData = value || null;
    if (lineId && Object.keys(data).length) {
      await prisma.packingListLine.update({ where: { id: lineId }, data });
    }
    return null;
  }

  if (intent === "update_packing_qty") {
    const lineId = Number(form.get("lineId"));
    const size = String(form.get("size") ?? "");
    const value = Math.max(0, Number(form.get("value") ?? 0) || 0);
    const line = await prisma.packingListLine.findUnique({ where: { id: lineId }, select: { qtys: true } });
    if (!line || !size) return null;
    const qtys = normalizeQtys(line.qtys);
    qtys[size] = value;
    await prisma.packingListLine.update({ where: { id: lineId }, data: { qtys } });
    return null;
  }

  if (intent === "duplicate_packing_line") {
    const lineId = Number(form.get("lineId"));
    const line = await prisma.packingListLine.findUnique({ where: { id: lineId } });
    if (!line) return null;
    await prisma.packingListLine.create({
      data: {
        packingListId: line.packingListId,
        boxNumber: line.boxNumber,
        productId: line.productId,
        productTitle: line.productTitle,
        productImageUrl: line.productImageUrl,
        fabricImageData: line.fabricImageData,
        sku: line.sku,
        isCustom: line.isCustom,
        qtys: line.qtys ?? {},
        priceRupees: line.priceRupees,
        weight: line.weight,
        notes: line.notes,
        sortOrder: line.sortOrder + 1,
      },
    });
    return null;
  }

  if (intent === "delete_packing_line") {
    const lineId = Number(form.get("lineId"));
    if (lineId) await prisma.packingListLine.delete({ where: { id: lineId } });
    return null;
  }

  if (intent === "delete_order") {
    await prisma.supplierOrder.delete({ where: { id: orderId } });
    return null;
  }

  if (intent === "update_column_widths") {
    let columnWidths: Record<string, number>;
    try {
      columnWidths = normalizeColumnWidths(JSON.parse(String(form.get("value") ?? "{}")));
    } catch {
      return null;
    }

    await prisma.portalSetting.upsert({
      where: { key: COLUMN_WIDTHS_KEY },
      create: { key: COLUMN_WIDTHS_KEY, value: columnWidths },
      update: { value: columnWidths },
    });
    return null;
  }

  if (intent === "update_packing_column_widths") {
    let columnWidths: Record<string, number>;
    try {
      columnWidths = normalizeColumnWidths(JSON.parse(String(form.get("value") ?? "{}")));
    } catch {
      return null;
    }

    await prisma.portalSetting.upsert({
      where: { key: PACKING_COLUMN_WIDTHS_KEY },
      create: { key: PACKING_COLUMN_WIDTHS_KEY, value: columnWidths },
      update: { value: columnWidths },
    });
    return null;
  }

  if (intent === "update_status")        updates.supplierStatus = form.get("value");
  if (intent === "update_priority")      updates.priority = form.get("value");
  if (intent === "update_product_type")  updates.productType = normalizeProductGroup(String(form.get("value") ?? "")) || null;
  if (intent === "update_factory_notes") updates.factoryNotes = form.get("value");
  if (intent === "update_notes")         updates.notes = form.get("value");
  if (intent === "update_eta") {
    const raw = String(form.get("value") ?? "");
    const parsedDate = raw ? parsePortalDate(raw) : null;
    if (raw && !parsedDate) return null;
    updates.eta = parsedDate;
  }

  if (intent === "update_qty") {
    const size = String(form.get("size") ?? "");
    const qtyOrdered = Math.max(0, Number(form.get("value") ?? 0) || 0);

    await prisma.$transaction(async (tx) => {
      const lines = await tx.orderLine.findMany({
        where: { orderId, variantTitle: size },
        orderBy: { id: "asc" },
        select: { id: true },
      });

      if (!lines.length) return;

      await tx.orderLine.update({
        where: { id: lines[0].id },
        data: { qtyOrdered },
      });

      if (lines.length > 1) {
        await tx.orderLine.updateMany({
          where: { id: { in: lines.slice(1).map((line) => line.id) } },
          data: { qtyOrdered: 0 },
        });
      }

      const allLines = await tx.orderLine.findMany({
        where: { orderId },
        select: { qtyOrdered: true },
      });
      await tx.supplierOrder.update({
        where: { id: orderId },
        data: { totalQty: allLines.reduce((sum, line) => sum + line.qtyOrdered, 0) },
      });
    });

    return null;
  }

  if (Object.keys(updates).length) {
    await prisma.supplierOrder.update({ where: { id: orderId }, data: updates });
  }
  return null;
};

// ─── Constants ───────────────────────────────────────────────────────────────

const STATUS_OPTIONS = [
  { value: "on_order",       label: "On Order" },
  { value: "on_production",  label: "On Production" },
  { value: "in_shipment",    label: "In Shipment" },
  { value: "arrived",        label: "Arrived" },
  { value: "arrived_loaded", label: "Arrived and Loaded" },
  { value: "cancelled",      label: "Cancelled" },
  { value: "ready_to_send",  label: "Ready To Send" },
];

const STATUS_COLORS: Record<string, string> = {
  on_order:       "#fef9c3",
  on_production:  "#dbeafe",
  in_shipment:    "#dcfce7",
  arrived:        "#bbf7d0",
  arrived_loaded: "#4ade80",
  cancelled:      "#fee2e2",
  ready_to_send:  "#ede9fe",
};

const PRIORITY_OPTIONS = [
  { value: "low",       label: "LOW",       bg: "#3b82f6", color: "#fff" },
  { value: "high",      label: "HIGH",      bg: "#7c3aed", color: "#fff" },
  { value: "urgent",    label: "URGENT",    bg: "#dc2626", color: "#fff" },
  { value: "cancelled", label: "Cancelled", bg: "#d97706", color: "#fff" },
];
const PACKING_STATUS_OPTIONS = [
  { value: "still_packing", label: "Still packing" },
  { value: "on_the_way", label: "On the way" },
  { value: "arrived", label: "Arrived" },
  { value: "loaded", label: "Loaded" },
];
const PACKING_SIZES = ["XS", "S", "M", "L", "XL", "2XL", "3XL", "S/M", "M/L", "L/XL"];
const DEFAULT_PACKING_ROWS = 5;
const PACKING_COLUMNS = [
  { id: "box", label: "Box", width: 70, center: true },
  { id: "picture", label: "Picture", width: 150, center: true },
  { id: "fabric", label: "Fabric Image", width: 150, center: true },
  { id: "name", label: "Name", width: 320 },
  { id: "sku", label: "SKU", width: 220 },
  ...PACKING_SIZES.map((size) => ({ id: `qty:${size}`, label: size, width: 76, center: true })),
  { id: "total", label: "Total", width: 82, center: true },
  { id: "price", label: "Price ₹", width: 92, center: true },
  { id: "value", label: "Value ₹", width: 96, center: true },
  { id: "weight", label: "Weight", width: 90, center: true },
  { id: "actions", label: "Actions", width: 112, center: true },
];
const PRODUCT_GROUP_RENAMES: Record<string, string> = {
  "Short Sleeve Dresses": "Dresses",
};

function normalizeProductGroup(value?: string | null) {
  const trimmed = value?.trim() ?? "";
  return PRODUCT_GROUP_RENAMES[trimmed] ?? trimmed;
}

function formatPortalDate(value?: string | Date | null) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = String(date.getFullYear()).slice(-2);
  return `${day}/${month}/${year}`;
}

function parsePortalDate(value: string) {
  const trimmed = value.trim();
  const auMatch = trimmed.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2}|\d{4})$/);
  const isoMatch = trimmed.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);

  const parts = auMatch
    ? { day: Number(auMatch[1]), month: Number(auMatch[2]), year: Number(auMatch[3]) }
    : isoMatch
      ? { day: Number(isoMatch[3]), month: Number(isoMatch[2]), year: Number(isoMatch[1]) }
      : null;

  if (!parts) return null;

  const fullYear = parts.year < 100 ? 2000 + parts.year : parts.year;
  const date = new Date(Date.UTC(fullYear, parts.month - 1, parts.day));
  const valid =
    date.getUTCFullYear() === fullYear &&
    date.getUTCMonth() === parts.month - 1 &&
    date.getUTCDate() === parts.day;

  return valid ? date : null;
}

function labelForStatus(value: string) {
  return STATUS_OPTIONS.find((option) => option.value === value)?.label ?? value;
}

function labelForPriority(value: string) {
  return PRIORITY_OPTIONS.find((option) => option.value === value)?.label ?? value;
}

function labelForPackingStatus(value: string) {
  return PACKING_STATUS_OPTIONS.find((option) => option.value === value)?.label ?? value;
}

const COLUMN_WIDTHS_KEY = "supplier-portal-column-widths-v1";
const PACKING_COLUMN_WIDTHS_KEY = "supplier-portal-packing-column-widths-v1";
const DELETE_CONFIRM_SKIP_KEY = "supplier-portal-delete-confirm-skip-until";
const PORTAL_LOGIN_REQUIRED_KEY = "supplier-portal-login-required-v1";
const PORTAL_USERS_KEY = "supplier-portal-users-v1";
const PORTAL_ACTIVE_USERS_KEY = "supplier-portal-active-users-v1";
const PORTAL_USER_COOKIE = "supplier_portal_user";
const ACTIVE_USER_WINDOW_MS = 5 * 60 * 1000;
const MIN_COLUMN_WIDTH = 52;
const FOCUSABLE_CELL_SELECTOR = [
  "input:not([type='hidden'])",
  "select",
  "textarea",
  "button",
  "[tabindex]:not([tabindex='-1'])",
].join(",");
const DEFAULT_COLUMN_WIDTHS: Record<string, number> = {
  factoryNotes: 190,
  orderDate: 92,
  picture: 88,
  name: 260,
  sku: 115,
  total: 80,
  status: 210,
  notes: 150,
  priority: 160,
  eta: 145,
  delete: 82,
};

type ColumnDef = { id: string; label: string; center?: boolean };
type PortalUser = { id: string; name: string; admin: boolean; active: boolean };
type ActivePortalUser = PortalUser & { initials: string; lastSeen: number };
type ShopifySearchProduct = {
  id: string;
  title: string;
  imageUrl: string | null;
  skus: string[];
  sizes: string[];
};

function normalizePortalUsers(value: unknown): PortalUser[] {
  if (!Array.isArray(value)) return [];
  return value
    .map((item) => {
      if (!item || typeof item !== "object") return null;
      const user = item as Record<string, unknown>;
      const id = String(user.id ?? "");
      const name = String(user.name ?? "").trim();
      if (!id || !name) return null;
      return {
        id,
        name,
        admin: Boolean(user.admin),
        active: user.active !== false,
      };
    })
    .filter(Boolean) as PortalUser[];
}

function normalizeBooleanSetting(value: unknown) {
  return value === true;
}

function getCookieValue(request: Request, key: string) {
  const cookie = request.headers.get("Cookie") ?? "";
  return cookie
    .split(";")
    .map((part) => part.trim())
    .find((part) => part.startsWith(`${key}=`))
    ?.slice(key.length + 1);
}

function getCurrentPortalUser(request: Request, users: PortalUser[]) {
  const userId = decodeURIComponent(getCookieValue(request, PORTAL_USER_COOKIE) ?? "");
  return users.find((user) => user.id === userId && user.active) ?? null;
}

function initialsForName(name: string) {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  return (parts.length > 1 ? `${parts[0][0]}${parts[parts.length - 1][0]}` : name.slice(0, 2)).toUpperCase();
}

function normalizeActiveUsers(value: unknown): Record<string, number> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  return Object.fromEntries(
    Object.entries(value)
      .map(([key, timestamp]) => [key, Number(timestamp) || 0] as const)
      .filter(([, timestamp]) => timestamp > 0),
  );
}

async function recordAndGetActiveUsers(currentUser: PortalUser | null, users: PortalUser[], value: unknown) {
  const now = Date.now();
  const activeMap = normalizeActiveUsers(value);
  if (currentUser) activeMap[currentUser.id] = now;
  const freshActiveMap = Object.fromEntries(
    Object.entries(activeMap).filter(([, timestamp]) => now - timestamp <= ACTIVE_USER_WINDOW_MS),
  );
  if (currentUser) {
    await prisma.portalSetting.upsert({
      where: { key: PORTAL_ACTIVE_USERS_KEY },
      create: { key: PORTAL_ACTIVE_USERS_KEY, value: freshActiveMap },
      update: { value: freshActiveMap },
    });
  }
  return users
    .filter((user) => user.active && freshActiveMap[user.id])
    .map((user) => ({ ...user, initials: initialsForName(user.name), lastSeen: freshActiveMap[user.id] }))
    .sort((a, b) => b.lastSeen - a.lastSeen);
}

async function savePortalUsers(users: PortalUser[]) {
  await prisma.portalSetting.upsert({
    where: { key: PORTAL_USERS_KEY },
    create: { key: PORTAL_USERS_KEY, value: users },
    update: { value: users },
  });
}

function normalizeQtys(value: unknown): Record<string, number> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  return Object.fromEntries(
    Object.entries(value)
      .map(([size, qty]) => [size, Math.max(0, Number(qty) || 0)] as const)
      .filter(([size]) => Boolean(size)),
  );
}

function packingTotal(qtys: Record<string, number>) {
  return Object.values(qtys).reduce((sum, qty) => sum + qty, 0);
}

function packingListTotal(list: PackingListWithLines | null) {
  if (!list) return 0;
  return list.lines.reduce((sum, line) => sum + packingTotal(normalizeQtys(line.qtys)), 0);
}

async function searchShopifyProducts(query: string): Promise<ShopifySearchProduct[]> {
  const trimmedQuery = query.trim();
  if (trimmedQuery.length < 2) return [];

  const session = await prisma.session.findFirst({
    where: {
      accessToken: { not: "" },
    },
    orderBy: { isOnline: "asc" },
  });
  if (!session?.shop || !session.accessToken) return [];

  const escapedQuery = trimmedQuery.replace(/[\\"]/g, "\\$&");
  const graphqlQuery = `#graphql
    query PackingProductSearch($query: String) {
      products(first: 20, query: $query, sortKey: TITLE) {
        edges {
          node {
            id
            title
            featuredImage { url }
            variants(first: 50) {
              edges {
                node {
                  title
                  sku
                }
              }
            }
          }
        }
      }
    }
  `;

  const mapProducts = (json: any) => (json.data?.products?.edges ?? []).map((edge: any) => {
    const variants = edge.node.variants.edges.map((variantEdge: any) => variantEdge.node);
    return {
      id: edge.node.id,
      title: edge.node.title,
      imageUrl: edge.node.featuredImage?.url ?? null,
      skus: variants.map((variant: any) => variant.sku).filter(Boolean),
      sizes: Array.from(new Set(variants.map((variant: any) => variant.title).filter(Boolean))),
    };
  });

  const matchesLocally = (product: ShopifySearchProduct) => {
    const needle = trimmedQuery.toLowerCase();
    return product.title.toLowerCase().includes(needle) || product.skus.some((sku) => sku.toLowerCase().includes(needle));
  };
  const directFetch = async (shopifyQuery: string | null) => {
    try {
      const response = await fetch(`https://${session.shop}/admin/api/2025-10/graphql.json`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "X-Shopify-Access-Token": session.accessToken,
        },
        body: JSON.stringify({
          query: graphqlQuery,
          variables: { query: shopifyQuery },
        }),
      });
      if (!response.ok) return [];
      return mapProducts(await response.json()) as ShopifySearchProduct[];
    } catch {
      return [];
    }
  };

  try {
    const { admin } = await unauthenticated.admin(session.shop);
    const response = await admin.graphql(graphqlQuery, {
      variables: { query: `title:*${escapedQuery}* OR sku:*${escapedQuery}*` },
    });
    const json = await response.json();
    const products = mapProducts(json);
    if (products.length) return products;

    const fallbackResponse = await admin.graphql(graphqlQuery, { variables: { query: null } });
    const fallbackJson = await fallbackResponse.json();
    return mapProducts(fallbackJson).filter(matchesLocally).slice(0, 8);
  } catch (error) {
    console.error("Shopify product search failed", error);
    const products = await directFetch(`title:*${escapedQuery}* OR sku:*${escapedQuery}*`);
    if (products.length) return products;
    return (await directFetch(null)).filter(matchesLocally).slice(0, 8);
  }
}

function sizeColumnId(size: string) {
  return `size:${size}`;
}

function defaultColumnWidth(columnId: string) {
  return columnId.startsWith("size:") ? 58 : DEFAULT_COLUMN_WIDTHS[columnId] ?? 110;
}

function defaultPackingColumnWidth(columnId: string) {
  return PACKING_COLUMNS.find((column) => column.id === columnId)?.width ?? 110;
}

function normalizeColumnWidths(value: unknown): Record<string, number> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};

  return Object.fromEntries(
    Object.entries(value)
      .map(([key, width]) => [key, Math.max(MIN_COLUMN_WIDTH, Number(width) || 0)] as const)
      .filter(([, width]) => width >= MIN_COLUMN_WIDTH),
  );
}

// ─── Main component ───────────────────────────────────────────────────────────

type Order = Awaited<ReturnType<typeof loader>>["orders"][number];

export default function PortalDashboard() {
  const {
    orders,
    sizes,
    productGroups,
    selectedProductGroup,
    selectedStatus,
    selectedPriority,
    searchTitle,
    statusFilters,
    priorityFilters,
    sortBy,
    page,
    columnWidths: savedColumnWidths,
    packingColumnWidths,
    packingLists,
    selectedPackingList,
    productSearch,
    packingSearchLineId,
    productResults,
    loginRequired,
    users,
    currentUser,
    activeUsers,
    loginBlocked,
  } = useLoaderData<typeof loader>();
  const [searchParams, setSearchParams] = useSearchParams();
  const columnWidthsFetcher = useFetcher();
  const [columnWidths, setColumnWidths] = useState<Record<string, number>>(savedColumnWidths);
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [searchTitleInput, setSearchTitleInput] = useState(searchTitle);
  const columns: ColumnDef[] = [
    { id: "factoryNotes", label: "Factory Notes" },
    { id: "orderDate", label: "Order Date" },
    { id: "picture", label: "Picture" },
    { id: "name", label: "Name" },
    { id: "sku", label: "SKU" },
    ...sizes.map((size) => ({ id: sizeColumnId(size), label: size, center: true })),
    { id: "total", label: "Total", center: true },
    { id: "status", label: "Status" },
    { id: "notes", label: "Notes" },
    { id: "priority", label: "Priority" },
    { id: "eta", label: "ETA" },
    { id: "delete", label: "Delete", center: true },
  ];

  const widthFor = (columnId: string) => columnWidths[columnId] ?? defaultColumnWidth(columnId);
  const tableWidth = columns.reduce((sum, column) => sum + widthFor(column.id), 0);
  const updateParams = (updates: Record<string, string>) => {
    const next = new URLSearchParams(searchParams);
    for (const [key, value] of Object.entries(updates)) {
      if (value) next.set(key, value);
      else next.delete(key);
    }
    setSearchParams(next, { replace: true, preventScrollReset: true });
  };
  useEffect(() => {
    setSearchTitleInput(searchTitle);
  }, [searchTitle]);
  useEffect(() => {
    const timer = window.setTimeout(() => {
      if (searchTitleInput !== searchTitle) updateParams({ q: searchTitleInput });
    }, 350);
    return () => window.clearTimeout(timer);
  }, [searchTitleInput, searchTitle]);
  const activePageTitle = page === "dashboard"
    ? "Dashboard"
    : page === "fabric"
      ? "Fabric in stock"
      : page === "settings"
        ? "Settings"
        : page === "packing"
          ? "Packing Lists"
        : selectedProductGroup || "Existing Products Restock";

  if (loginBlocked) {
    return <PortalLogin users={users} />;
  }

  const startResize = (columnId: string, event: React.MouseEvent<HTMLSpanElement>) => {
    event.preventDefault();
    const startX = event.clientX;
    const startWidth = widthFor(columnId);
    let nextColumnWidths = columnWidths;
    const previousCursor = document.body.style.cursor;
    const previousUserSelect = document.body.style.userSelect;

    document.body.style.cursor = "col-resize";
    document.body.style.userSelect = "none";

    const handleMove = (moveEvent: MouseEvent) => {
      const nextWidth = Math.max(MIN_COLUMN_WIDTH, startWidth + moveEvent.clientX - startX);
      nextColumnWidths = { ...nextColumnWidths, [columnId]: nextWidth };
      setColumnWidths(nextColumnWidths);
    };
    const handleUp = () => {
      document.body.style.cursor = previousCursor;
      document.body.style.userSelect = previousUserSelect;
      document.removeEventListener("mousemove", handleMove);
      document.removeEventListener("mouseup", handleUp);
      const formData = new FormData();
      formData.set("intent", "update_column_widths");
      formData.set("value", JSON.stringify(nextColumnWidths));
      columnWidthsFetcher.submit(formData, { method: "post" });
    };

    document.addEventListener("mousemove", handleMove);
    document.addEventListener("mouseup", handleUp);
  };

  const handleGridKeyDown = (event: React.KeyboardEvent<HTMLTableElement>) => {
    if (!["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"].includes(event.key)) return;
    if (event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) return;

    const currentCell = (event.target as HTMLElement).closest<HTMLElement>("[data-grid-row][data-grid-col]");
    if (!currentCell) return;

    const row = Number(currentCell.dataset.gridRow);
    const col = Number(currentCell.dataset.gridCol);
    const next = {
      ArrowUp: [row - 1, col],
      ArrowDown: [row + 1, col],
      ArrowLeft: [row, col - 1],
      ArrowRight: [row, col + 1],
    }[event.key]!;
    const [nextRow, nextCol] = next;
    const nextCell = event.currentTarget.querySelector<HTMLElement>(
      `[data-grid-row="${nextRow}"][data-grid-col="${nextCol}"]`,
    );

    if (!nextCell) return;

    event.preventDefault();
    const focusTarget = nextCell.querySelector<HTMLElement>(FOCUSABLE_CELL_SELECTOR) ?? nextCell;
    focusTarget.focus();

    if (focusTarget instanceof HTMLInputElement) {
      focusTarget.select();
    } else if (focusTarget instanceof HTMLTextAreaElement) {
      focusTarget.select();
    }
  };

  return (
    <div style={s.appShell}>
      <aside style={{ ...s.sidebar, ...(sidebarCollapsed ? s.sidebarCollapsed : {}) }}>
        <div style={sidebarCollapsed ? s.sidebarTopCollapsed : s.sidebarTop}>
          {!sidebarCollapsed && <div style={s.sidebarTitle}>Supplier Portal</div>}
          <button
            type="button"
            onClick={() => setSidebarCollapsed((current) => !current)}
            style={s.collapseButton}
            aria-label={sidebarCollapsed ? "Expand sidebar" : "Collapse sidebar"}
          >
            {sidebarCollapsed ? ">" : "<"}
          </button>
        </div>
        {sidebarCollapsed ? (
          <>
            <nav style={s.iconNav}>
              <a title="Dashboard" href="/portal?page=dashboard" style={{ ...s.iconNavItem, ...(page === "dashboard" ? s.iconNavItemActive : {}) }}>D</a>
              <a title="Existing Products Restock" href="/portal" style={{ ...s.iconNavItem, ...(page === "restock" && !selectedProductGroup ? s.iconNavItemActive : {}) }}>R</a>
              <a title="Fabric in stock" href="/portal?page=fabric" style={{ ...s.iconNavItem, ...(page === "fabric" ? s.iconNavItemActive : {}) }}>F</a>
              <a title="Packing Lists" href="/portal?page=packing" style={{ ...s.iconNavItem, ...(page === "packing" ? s.iconNavItemActive : {}) }}>P</a>
            </nav>
            <a title="Settings" href="/portal?page=settings" style={{ ...s.iconNavItem, ...(page === "settings" ? s.iconNavItemActive : {}), ...s.settingsLink }}>S</a>
          </>
        ) : (
          <>
            <nav style={s.nav}>
              <a href="/portal?page=dashboard" style={{ ...s.navItem, ...(page === "dashboard" ? s.navItemActive : {}) }}>Dashboard</a>
              <a href="/portal" style={{ ...s.navItem, ...(page === "restock" && !selectedProductGroup ? s.navItemActive : {}) }}>Existing Products Restock</a>
              <a href="/portal?page=fabric" style={{ ...s.navItem, ...(page === "fabric" ? s.navItemActive : {}) }}>Fabric in stock</a>
              <a href="/portal?page=packing" style={{ ...s.navItem, ...(page === "packing" ? s.navItemActive : {}) }}>Packing Lists</a>
            </nav>
            <a href="/portal?page=settings" style={{ ...s.navItem, ...(page === "settings" ? s.navItemActive : {}), ...s.settingsLink }}>Settings</a>
          </>
        )}
      </aside>

      <main style={s.main}>
        <header style={s.pageHeader}>
          <div>
            <h1 style={s.pageTitle}>{activePageTitle}</h1>
          </div>
          {page === "restock" && (
            <div style={s.headerControls}>
              <div style={s.utilityBar}>
                <label style={s.filterLabel}>
                  Search
                  <input
                    type="search"
                    value={searchTitleInput}
                    onChange={(event) => setSearchTitleInput(event.currentTarget.value)}
                    style={s.searchInput}
                    placeholder="Product title"
                  />
                </label>
                <div style={s.activeUsers} title="Currently active">
                  <span style={s.activeUsersLabel}>Active</span>
                  {activeUsers.length ? activeUsers.map((user) => (
                    <span key={user.id} style={s.activeUserBadge} title={user.name}>{user.initials}</span>
                  )) : <span style={s.activeUserEmpty}>No active users</span>}
                </div>
              </div>
              <div style={s.filters}>
                <label style={s.filterLabel}>
                  Product group
                  <select
                    value={selectedProductGroup}
                    onChange={(event) => updateParams({ productGroup: event.currentTarget.value, productType: "" })}
                    style={s.productTypeFilter}
                  >
                    <option value="">All groups</option>
                    {productGroups.map((group) => (
                      <option key={group} value={group}>{group}</option>
                    ))}
                  </select>
                </label>
                <label style={s.filterLabel}>
                  Sort
                  <select value={sortBy} onChange={(event) => updateParams({ sortBy: event.currentTarget.value })} style={s.productTypeFilter}>
                    <option value="orderDateDesc">Order date newest</option>
                    <option value="orderDateAsc">Order date oldest</option>
                    <option value="titleAsc">Product title A-Z</option>
                    <option value="titleDesc">Product title Z-A</option>
                  </select>
                </label>
                <label style={s.filterLabel}>
                  Status
                  <select value={selectedStatus} onChange={(event) => updateParams({ status: event.currentTarget.value })} style={s.productTypeFilter}>
                    <option value="">All statuses</option>
                    {statusFilters.map((status) => (
                      <option key={status} value={status}>{labelForStatus(status)}</option>
                    ))}
                  </select>
                </label>
                <label style={s.filterLabel}>
                  Priority
                  <select value={selectedPriority} onChange={(event) => updateParams({ priority: event.currentTarget.value })} style={s.productTypeFilter}>
                    <option value="">All priorities</option>
                    {priorityFilters.map((priority) => (
                      <option key={priority} value={priority}>{labelForPriority(priority)}</option>
                    ))}
                  </select>
                </label>
              </div>
            </div>
          )}
        </header>

        {page === "settings" ? (
          <SettingsPanel
            users={users}
            currentUser={currentUser}
            loginRequired={loginRequired}
          />
        ) : page === "packing" ? (
          <PackingListsPanel
            packingLists={packingLists}
            selectedPackingList={selectedPackingList}
            savedPackingColumnWidths={packingColumnWidths}
            productSearch={productSearch}
            packingSearchLineId={packingSearchLineId}
            productResults={productResults}
            updateParams={updateParams}
          />
        ) : page !== "restock" ? (
          <div style={s.empty}>{activePageTitle} will be set up here.</div>
        ) : orders.length === 0 ? (
          <div style={s.empty}>No open orders at the moment.</div>
        ) : (
          <div style={s.tableWrap}>
            <table style={{ ...s.table, width: tableWidth }} onKeyDown={handleGridKeyDown}>
              <colgroup>
                {columns.map((column) => (
                  <col key={column.id} style={{ width: widthFor(column.id) }} />
                ))}
              </colgroup>
              <thead>
                <tr style={s.headerRow}>
                  {columns.map((column) => (
                    <Th
                      key={column.id}
                      center={column.center}
                      onResizeStart={(event) => startResize(column.id, event)}
                    >
                      {column.label}
                    </Th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {orders.map((order, rowIndex) => (
                  <OrderRow key={order.id} order={order} rowIndex={rowIndex} sizes={sizes} />
                ))}
              </tbody>
            </table>
          </div>
        )}
      </main>
    </div>
  );
}

function PortalLogin({ users }: { users: PortalUser[] }) {
  return (
    <div style={s.loginShell}>
      <form method="post" style={s.loginCard}>
        <input type="hidden" name="intent" value="portal_login" />
        <h1 style={s.loginTitle}>Supplier Portal</h1>
        <p style={s.loginText}>Select your name to enter the portal.</p>
        <select name="userId" required style={s.loginSelect}>
          <option value="">Select your name</option>
          {users.filter((user) => user.active).map((user) => (
            <option key={user.id} value={user.id}>{user.name}</option>
          ))}
        </select>
        <button type="submit" style={s.loginButton}>Enter portal</button>
      </form>
    </div>
  );
}

function SettingsPanel({
  users,
  currentUser,
  loginRequired,
}: {
  users: PortalUser[];
  currentUser: PortalUser | null;
  loginRequired: boolean;
}) {
  const settingsFetcher = useFetcher();
  const canManageUsers = !loginRequired || users.length === 0 || currentUser?.admin;

  return (
    <div style={s.settingsPanel}>
      <section style={s.settingsCard}>
        <div style={s.settingsHeader}>
          <div>
            <h2 style={s.settingsTitle}>Portal access</h2>
            <p style={s.settingsHint}>Use lightweight name access to see who is working in the portal.</p>
          </div>
          {currentUser && (
            <form method="post">
              <input type="hidden" name="intent" value="portal_logout" />
              <button type="submit" style={s.secondaryButton}>Log out {currentUser.name}</button>
            </form>
          )}
        </div>

        <label style={s.switchRow}>
          <input
            type="checkbox"
            checked={loginRequired}
            disabled={!canManageUsers}
            onChange={(event) => submitPortalCell(settingsFetcher, {
              intent: "update_login_required",
              value: event.currentTarget.checked ? "on" : "off",
            })}
          />
          Require users to select their name before entering
        </label>

        {!canManageUsers && (
          <div style={s.settingsWarning}>Only an admin user can add/remove names or change portal access.</div>
        )}
      </section>

      <section style={s.settingsCard}>
        <h2 style={s.settingsTitle}>Allowed users</h2>
        <div style={s.userList}>
          {users.length ? users.map((user) => (
            <div key={user.id} style={s.userRow}>
              <span style={s.activeUserBadge}>{initialsForName(user.name)}</span>
              <span style={s.userName}>{user.name}</span>
              {user.admin && <span style={s.adminPill}>Admin</span>}
              {canManageUsers && (
                <button
                  type="button"
                  style={s.removeUserButton}
                  onClick={() => submitPortalCell(settingsFetcher, {
                    intent: "remove_portal_user",
                    userId: user.id,
                  })}
                >
                  Remove
                </button>
              )}
            </div>
          )) : (
            <div style={s.settingsHint}>No names added yet.</div>
          )}
        </div>

        {canManageUsers && (
          <settingsFetcher.Form method="post" style={s.addUserForm}>
            <input type="hidden" name="intent" value="add_portal_user" />
            <input name="name" required placeholder="First and last name" style={s.addUserInput} />
            <label style={s.adminCheckbox}>
              <input type="checkbox" name="admin" />
              Admin
            </label>
            <button type="submit" style={s.loginButton}>Add user</button>
          </settingsFetcher.Form>
        )}
      </section>
    </div>
  );
}

type PackingListWithLines = Awaited<ReturnType<typeof loader>>["packingLists"][number];

function PackingListsPanel({
  packingLists,
  selectedPackingList,
  savedPackingColumnWidths,
  productSearch,
  packingSearchLineId,
  productResults,
  updateParams,
}: {
  packingLists: PackingListWithLines[];
  selectedPackingList: PackingListWithLines | null;
  savedPackingColumnWidths: Record<string, number>;
  productSearch: string;
  packingSearchLineId: number | null;
  productResults: ShopifySearchProduct[];
  updateParams: (updates: Record<string, string>) => void;
}) {
  const fetcher = useFetcher();

  return (
    <div style={s.packingLayout}>
      {!selectedPackingList ? (
        <PackingListsOverview packingLists={packingLists} fetcher={fetcher} />
      ) : (
        <section style={s.packingDetail}>
          <PackingListDetail
            packingList={selectedPackingList}
            savedPackingColumnWidths={savedPackingColumnWidths}
            productSearch={productSearch}
            packingSearchLineId={packingSearchLineId}
            productResults={productResults}
            updateParams={updateParams}
          />
        </section>
      )}
    </div>
  );
}

function PackingListsOverview({
  packingLists,
  fetcher,
}: {
  packingLists: PackingListWithLines[];
  fetcher: ReturnType<typeof useFetcher>;
}) {
  return (
    <div style={s.packingOverview}>
      <section style={s.packingOverviewCreate}>
        <div>
          <h2 style={s.settingsTitle}>Create packing list</h2>
          <p style={s.settingsHint}>Create one shipment, then open it to add products and box quantities.</p>
        </div>
        <fetcher.Form method="post" style={s.packingCreateForm}>
          <input type="hidden" name="intent" value="create_packing_list" />
          <input name="title" placeholder="Shipment name" style={s.packingInput} />
          <input name="invoiceNumber" placeholder="Invoice number" style={s.packingInput} />
          <input name="expectedLeaveFactoryDate" placeholder="Leave factory dd/mm/yy" style={s.packingInput} />
          <button type="submit" style={s.loginButton}>New packing list</button>
        </fetcher.Form>
      </section>

      <section style={s.packingOverviewTableWrap}>
        <table style={{ ...s.table, width: "100%" }}>
          <thead>
            <tr style={s.headerRow}>
              {["Packing list", "Invoice", "Total qty", "Leave factory", "Status", "Created", "Open"].map((heading) => (
                <th key={heading} style={{ ...s.th, textAlign: heading === "Total qty" || heading === "Open" ? "center" : "left" }}>
                  <span style={s.thContent}>{heading}</span>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {packingLists.length ? packingLists.map((list) => (
              <tr key={list.id} style={s.row}>
                <td style={s.td}><strong style={s.productName}>{list.title}</strong></td>
                <td style={s.td}>{list.invoiceNumber || "—"}</td>
                <td style={{ ...s.td, textAlign: "center" }}><span style={s.total}>{packingListTotal(list)}</span></td>
                <td style={s.td}>{formatPortalDate(list.expectedLeaveFactoryDate ?? list.shipmentDate) || "—"}</td>
                <td style={s.td}>{labelForPackingStatus(list.status)}</td>
                <td style={s.td}>{formatPortalDate(list.createdAt)}</td>
                <td style={{ ...s.td, textAlign: "center" }}>
                  <a href={`/portal?page=packing&packingId=${list.id}`} style={s.smallLinkButton}>Open</a>
                </td>
              </tr>
            )) : (
              <tr style={s.row}>
                <td colSpan={7} style={{ ...s.td, textAlign: "center", padding: 40 }}>No packing lists yet.</td>
              </tr>
            )}
          </tbody>
        </table>
      </section>
    </div>
  );
}

function PackingListDetail({
  packingList,
  savedPackingColumnWidths,
  productSearch,
  packingSearchLineId,
  productResults,
  updateParams,
}: {
  packingList: PackingListWithLines;
  savedPackingColumnWidths: Record<string, number>;
  productSearch: string;
  packingSearchLineId: number | null;
  productResults: ShopifySearchProduct[];
  updateParams: (updates: Record<string, string>) => void;
}) {
  const fetcher = useFetcher();
  const columnWidthsFetcher = useFetcher();
  const [packingColumnWidths, setPackingColumnWidths] = useState<Record<string, number>>(savedPackingColumnWidths);
  const packingWidthFor = (columnId: string) => packingColumnWidths[columnId] ?? defaultPackingColumnWidth(columnId);
  const packingTableWidth = PACKING_COLUMNS.reduce((sum, column) => sum + packingWidthFor(column.id), 0);

  useEffect(() => {
    setPackingColumnWidths(savedPackingColumnWidths);
  }, [savedPackingColumnWidths]);

  const startPackingResize = (columnId: string, event: React.MouseEvent<HTMLSpanElement>) => {
    event.preventDefault();
    const startX = event.clientX;
    const startWidth = packingWidthFor(columnId);
    let nextColumnWidths = packingColumnWidths;
    const previousCursor = document.body.style.cursor;
    const previousUserSelect = document.body.style.userSelect;

    document.body.style.cursor = "col-resize";
    document.body.style.userSelect = "none";

    const handleMove = (moveEvent: MouseEvent) => {
      const nextWidth = Math.max(MIN_COLUMN_WIDTH, startWidth + moveEvent.clientX - startX);
      nextColumnWidths = { ...nextColumnWidths, [columnId]: nextWidth };
      setPackingColumnWidths(nextColumnWidths);
    };
    const handleUp = () => {
      document.body.style.cursor = previousCursor;
      document.body.style.userSelect = previousUserSelect;
      document.removeEventListener("mousemove", handleMove);
      document.removeEventListener("mouseup", handleUp);
      const formData = new FormData();
      formData.set("intent", "update_packing_column_widths");
      formData.set("value", JSON.stringify(nextColumnWidths));
      columnWidthsFetcher.submit(formData, { method: "post" });
    };

    document.addEventListener("mousemove", handleMove);
    document.addEventListener("mouseup", handleUp);
  };

  return (
    <div style={s.packingDetailInner}>
      <div style={s.packingTop}>
        <a href="/portal?page=packing" style={s.secondaryButton}>Back to packing lists</a>
        <div style={s.packingMeta}>
          <label style={s.filterLabel}>
            Shipment
            <input
              defaultValue={packingList.title}
              onBlur={(event) => submitPortalCell(fetcher, {
                intent: "update_packing_list",
                packingId: packingList.id,
                field: "title",
                value: event.currentTarget.value,
              })}
              style={s.packingInput}
            />
          </label>
          <label style={s.filterLabel}>
            Invoice number
            <input
              defaultValue={packingList.invoiceNumber ?? ""}
              onBlur={(event) => submitPortalCell(fetcher, {
                intent: "update_packing_list",
                packingId: packingList.id,
                field: "invoiceNumber",
                value: event.currentTarget.value,
              })}
              placeholder="Invoice number"
              style={s.packingInput}
            />
          </label>
          <label style={s.filterLabel}>
            Leave factory
            <input
              defaultValue={formatPortalDate(packingList.expectedLeaveFactoryDate ?? packingList.shipmentDate)}
              onBlur={(event) => submitPortalCell(fetcher, {
                intent: "update_packing_list",
                packingId: packingList.id,
                field: "expectedLeaveFactoryDate",
                value: event.currentTarget.value,
              })}
              placeholder="dd/mm/yy"
              style={s.packingInput}
            />
          </label>
          <label style={s.filterLabel}>
            Status
            <select
              value={packingList.status}
              onChange={(event) => submitPortalCell(fetcher, {
                intent: "update_packing_list",
                packingId: packingList.id,
                field: "status",
                value: event.currentTarget.value,
              })}
              style={s.productTypeFilter}
            >
              {PACKING_STATUS_OPTIONS.map((option) => (
                <option key={option.value} value={option.value}>{option.label}</option>
              ))}
            </select>
          </label>
          <div style={s.packingTotalPill}>
            Total quantity <strong>{packingListTotal(packingList)}</strong>
          </div>
        </div>
      </div>

      <div style={s.packingTableWrap}>
        <table style={{ ...s.table, width: packingTableWidth, minWidth: "100%" }}>
          <colgroup>
            {PACKING_COLUMNS.map((column) => (
              <col key={column.id} style={{ width: packingWidthFor(column.id) }} />
            ))}
          </colgroup>
          <thead>
            <tr style={s.headerRow}>
              {PACKING_COLUMNS.map((column) => (
                <Th
                  key={column.id}
                  center={column.center}
                  onResizeStart={(event) => startPackingResize(column.id, event)}
                >
                  {column.label}
                </Th>
              ))}
            </tr>
          </thead>
          <tbody>
            {packingList.lines.map((line) => (
              <PackingListLineRow
                key={line.id}
                line={line}
                activeSearchLineId={packingSearchLineId}
                productSearch={productSearch}
                productResults={productResults}
                updateParams={updateParams}
              />
            ))}
          </tbody>
        </table>
      </div>

      <div style={s.packingFooterActions}>
        <button
          type="button"
          style={s.smallButton}
          onClick={() => submitPortalCell(fetcher, {
            intent: "add_custom_packing_line",
            packingId: packingList.id,
          })}
        >
          Add row
        </button>
      </div>
    </div>
  );
}

function PackingListLineRow({
  line,
  activeSearchLineId,
  productSearch,
  productResults,
  updateParams,
}: {
  line: PackingListWithLines["lines"][number];
  activeSearchLineId: number | null;
  productSearch: string;
  productResults: ShopifySearchProduct[];
  updateParams: (updates: Record<string, string>) => void;
}) {
  const fetcher = useFetcher();
  const qtys = normalizeQtys(line.qtys);
  const total = packingTotal(qtys);
  const price = line.priceRupees ?? 0;
  const value = total * price;

  return (
    <tr style={s.row}>
      <td style={s.td}><PackingTextInput lineId={line.id} field="boxNumber" value={line.boxNumber ?? ""} /></td>
      <td style={{ ...s.td, textAlign: "center" }}>{line.productImageUrl ? <img src={line.productImageUrl} alt="" style={s.packingThumb} /> : <div style={s.noImg}>—</div>}</td>
      <td style={{ ...s.td, textAlign: "center" }}><FabricImageCell lineId={line.id} value={line.fabricImageData ?? ""} /></td>
      <td style={{ ...s.td, ...s.dropdownTd }}>
        <PackingProductNameCell
          line={line}
          isActiveSearch={activeSearchLineId === line.id}
          productSearch={productSearch}
          productResults={productResults}
          updateParams={updateParams}
        />
      </td>
      <td style={s.td}><PackingSkuCell lineId={line.id} value={line.sku ?? ""} /></td>
      {PACKING_SIZES.map((size) => (
        <td key={size} style={{ ...s.td, textAlign: "center" }}>
          <input
            type="text"
            inputMode="numeric"
            defaultValue={qtys[size] || ""}
            onChange={(event) => { event.currentTarget.value = event.currentTarget.value.replace(/\D/g, ""); }}
            onBlur={(event) => submitPortalCell(fetcher, {
              intent: "update_packing_qty",
              lineId: line.id,
              size,
              value: event.currentTarget.value,
            })}
            style={s.qtyInput}
          />
        </td>
      ))}
      <td style={{ ...s.td, textAlign: "center" }}><span style={s.total}>{total}</span></td>
      <td style={{ ...s.td, textAlign: "center" }}><PackingTextInput lineId={line.id} field="priceRupees" value={line.priceRupees?.toString() ?? ""} center /></td>
      <td style={{ ...s.td, textAlign: "center" }}><span style={s.total}>{value ? Math.round(value) : ""}</span></td>
      <td style={{ ...s.td, textAlign: "center" }}><PackingTextInput lineId={line.id} field="weight" value={line.weight?.toString() ?? ""} center /></td>
      <td style={{ ...s.td, textAlign: "center" }}>
        <div style={s.rowActions}>
          <button type="button" style={s.smallButton} onClick={() => submitPortalCell(fetcher, { intent: "duplicate_packing_line", lineId: line.id })}>Duplicate</button>
          <button type="button" style={s.removeUserButton} onClick={() => submitPortalCell(fetcher, { intent: "delete_packing_line", lineId: line.id })}>Delete</button>
        </div>
      </td>
    </tr>
  );
}

function PackingProductNameCell({
  line,
  isActiveSearch,
  productSearch,
  productResults,
  updateParams,
}: {
  line: PackingListWithLines["lines"][number];
  isActiveSearch: boolean;
  productSearch: string;
  productResults: ShopifySearchProduct[];
  updateParams: (updates: Record<string, string>) => void;
}) {
  const fetcher = useFetcher();
  const inputRef = useRef<HTMLInputElement | null>(null);
  const displayValue = line.isCustom && line.productTitle === "Custom item" ? "" : line.productTitle;
  const [value, setValue] = useState(displayValue);
  const [isFocused, setIsFocused] = useState(false);
  const [isProductSelected, setIsProductSelected] = useState(false);
  const [dropdownRect, setDropdownRect] = useState<DOMRect | null>(null);
  const [canPortalDropdown, setCanPortalDropdown] = useState(false);
  const canSearch = !isProductSelected && (isFocused || isActiveSearch);
  const shouldShowResults = canSearch && value.trim().length >= 2;
  const dropdownHeight = value.trim() !== productSearch || !productResults.length
    ? 48
    : Math.min(320, productResults.length * 62 + 12);
  const updateDropdownRect = () => {
    if (!inputRef.current) return;
    setDropdownRect(inputRef.current.getBoundingClientRect());
  };

  useEffect(() => {
    setValue(displayValue);
    setIsProductSelected(false);
  }, [displayValue, line.id]);

  useEffect(() => {
    setCanPortalDropdown(true);
  }, []);

  useEffect(() => {
    if (!shouldShowResults) {
      setDropdownRect(null);
      return;
    }
    updateDropdownRect();
    window.addEventListener("resize", updateDropdownRect);
    window.addEventListener("scroll", updateDropdownRect, true);
    return () => {
      window.removeEventListener("resize", updateDropdownRect);
      window.removeEventListener("scroll", updateDropdownRect, true);
    };
  }, [shouldShowResults, value, line.id]);

  useEffect(() => {
    if (!canSearch) return;
    const timer = window.setTimeout(() => {
      const trimmed = value.trim();
      updateParams({
        productSearch: trimmed.length >= 2 ? trimmed : "",
        packingSearchLineId: trimmed.length >= 2 ? String(line.id) : "",
      });
    }, 450);
    return () => window.clearTimeout(timer);
  }, [value, canSearch, line.id]);

  const applyProduct = (product: ShopifySearchProduct) => {
    setValue(product.title);
    setIsFocused(false);
    setIsProductSelected(true);
    setDropdownRect(null);
    inputRef.current?.blur();
    submitPortalCell(fetcher, {
      intent: "apply_product_to_packing_line",
      lineId: line.id,
      product: JSON.stringify(product),
    });
    updateParams({ productSearch: "", packingSearchLineId: "" });
  };

  return (
    <div style={s.productCellSearch}>
      <input
        ref={inputRef}
        type="search"
        value={value}
        onFocus={() => {
          setIsFocused(true);
          setIsProductSelected(false);
          if (value.trim().length >= 2) {
            updateParams({ productSearch: value.trim(), packingSearchLineId: String(line.id) });
          }
        }}
        onChange={(event) => setValue(event.currentTarget.value)}
        onBlur={(event) => {
          if (isProductSelected) return;
          setIsFocused(false);
          submitPortalCell(fetcher, {
            intent: "update_packing_line",
            lineId: line.id,
            field: "productTitle",
            value: event.currentTarget.value,
          });
        }}
        placeholder="Search or type product"
        style={s.packingCellInput}
      />
      {shouldShowResults && canPortalDropdown && dropdownRect && createPortal(
        <div
          style={{
            ...s.productCellResults,
            top: dropdownRect.bottom + 8,
            left: dropdownRect.left,
            width: Math.max(dropdownRect.width, 460),
            height: dropdownHeight,
          }}
        >
          {value.trim() !== productSearch ? (
            <div style={s.productCellResultEmpty}>Searching...</div>
          ) : productResults.length ? productResults.map((product) => (
            <button
              key={product.id}
              type="button"
              style={s.productCellResult}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => applyProduct(product)}
            >
              {product.imageUrl ? <img src={product.imageUrl} alt="" style={s.productCellResultImage} /> : <span style={s.productCellNoImage}>—</span>}
              <span style={s.productCellResultText}>
                <strong>{product.title}</strong>
                <span>{product.skus.slice(0, 3).join(", ") || "No SKU"}</span>
              </span>
            </button>
          )) : (
            <div style={s.productCellResultEmpty}>No products found. Keep typed text for a custom row.</div>
          )}
        </div>,
        document.body,
      )}
    </div>
  );
}

function PackingTextInput({ lineId, field, value, multiline, center }: { lineId: number; field: string; value: string; multiline?: boolean; center?: boolean }) {
  const fetcher = useFetcher();
  const common = {
    defaultValue: value,
    onBlur: (event: React.FocusEvent<HTMLInputElement | HTMLTextAreaElement>) => submitPortalCell(fetcher, {
      intent: "update_packing_line",
      lineId,
      field,
      value: event.currentTarget.value,
    }),
    style: { ...(multiline ? s.packingTextarea : s.packingCellInput), ...(center ? { textAlign: "center" as const } : {}) },
  };
  return multiline ? <textarea rows={3} {...common} /> : <input type="text" {...common} />;
}

function PackingSkuCell({ lineId, value }: { lineId: number; value: string }) {
  const fetcher = useFetcher();

  return (
    <textarea
      rows={7}
      defaultValue={value}
      onBlur={(event) => submitPortalCell(fetcher, {
        intent: "update_packing_line",
        lineId,
        field: "sku",
        value: event.currentTarget.value,
      })}
      style={s.packingSkuTextarea}
    />
  );
}

function FabricImageCell({ lineId, value }: { lineId: number; value: string }) {
  const fetcher = useFetcher();

  const saveFile = (file: File) => {
    const reader = new FileReader();
    reader.onload = () => submitPortalCell(fetcher, {
      intent: "update_packing_line",
      lineId,
      field: "fabricImageData",
      value: String(reader.result ?? ""),
    });
    reader.readAsDataURL(file);
  };

  return (
    <div
      tabIndex={0}
      style={s.fabricImageDrop}
      onPaste={(event) => {
        const file = Array.from(event.clipboardData.files).find((item) => item.type.startsWith("image/"));
        if (file) saveFile(file);
      }}
      title="Paste or upload fabric image"
    >
      {value ? <img src={value} alt="" style={s.fabricThumb} /> : <span>Paste image</span>}
      <input
        type="file"
        accept="image/*"
        style={s.hiddenFileInput}
        onChange={(event) => {
          const file = event.currentTarget.files?.[0];
          if (file) saveFile(file);
        }}
      />
    </div>
  );
}

// ─── Row ─────────────────────────────────────────────────────────────────────

function OrderRow({ order, rowIndex, sizes }: { order: Order; rowIndex: number; sizes: string[] }) {
  const qtyBySize = order.lines.reduce<Record<string, number>>((acc, line) => {
    acc[line.variantTitle] = (acc[line.variantTitle] ?? 0) + line.qtyOrdered;
    return acc;
  }, {});
  const allSkus = order.lines.map((l) => l.sku).filter(Boolean).join("\n");
  const etaValue = formatPortalDate(order.eta);
  const orderDate = formatPortalDate(order.createdAt);
  const totalCol = 5 + sizes.length;
  const statusCol = totalCol + 1;
  const notesCol = totalCol + 2;
  const priorityCol = totalCol + 3;
  const etaCol = totalCol + 4;
  const deleteCol = totalCol + 5;

  return (
    <tr style={s.row}>
      {/* Factory notes */}
      <Td rowIndex={rowIndex} colIndex={0}><NotesCell orderId={order.id} field="factory_notes" value={order.factoryNotes ?? ""} /></Td>

      {/* Order date */}
      <Td rowIndex={rowIndex} colIndex={1} center><span style={s.dateText}>{orderDate}</span></Td>

      {/* Picture */}
      <Td rowIndex={rowIndex} colIndex={2} center>
        <div style={s.imageCell}>
          {order.productImageUrl
            ? <img src={order.productImageUrl} alt="" style={s.thumb} />
            : <div style={s.noImg}>—</div>}
        </div>
      </Td>

      {/* Name */}
      <Td rowIndex={rowIndex} colIndex={3}><span style={s.productName}>{order.productTitle}</span></Td>

      {/* SKU */}
      <Td rowIndex={rowIndex} colIndex={4}><span style={s.sku}>{allSkus || "—"}</span></Td>

      {/* Size columns */}
      {sizes.map((sz, sizeIndex) => (
        <Td key={sz} rowIndex={rowIndex} colIndex={5 + sizeIndex} center>
          <QtyCell orderId={order.id} size={sz} value={qtyBySize[sz] ?? 0} />
        </Td>
      ))}

      {/* Total */}
      <Td rowIndex={rowIndex} colIndex={totalCol} center><span style={s.total}>{order.totalQty}</span></Td>

      {/* Status */}
      <Td rowIndex={rowIndex} colIndex={statusCol}><StatusCell orderId={order.id} value={order.supplierStatus} /></Td>

      {/* Notes (from order) */}
      <Td rowIndex={rowIndex} colIndex={notesCol}><NotesCell orderId={order.id} field="notes" value={order.notes ?? ""} /></Td>

      {/* Priority */}
      <Td rowIndex={rowIndex} colIndex={priorityCol}><PriorityCell orderId={order.id} value={order.priority ?? ""} /></Td>

      {/* ETA */}
      <Td rowIndex={rowIndex} colIndex={etaCol}><EtaCell orderId={order.id} value={etaValue} /></Td>

      {/* Delete */}
      <Td rowIndex={rowIndex} colIndex={deleteCol} center><DeleteCell orderId={order.id} /></Td>
    </tr>
  );
}

// ─── Editable cells ───────────────────────────────────────────────────────────

function submitPortalCell(
  fetcher: ReturnType<typeof useFetcher>,
  fields: Record<string, string | number>,
) {
  const formData = new FormData();
  for (const [key, value] of Object.entries(fields)) {
    formData.set(key, String(value));
  }
  fetcher.submit(formData, { method: "post" });
}

function StatusCell({ orderId, value }: { orderId: number; value: string }) {
  const fetcher = useFetcher();
  const current = fetcher.formData ? String(fetcher.formData.get("value")) : value;
  const bg = STATUS_COLORS[current] ?? "#f3f4f6";

  return (
    <select
      value={current}
      onChange={(e) => submitPortalCell(fetcher, {
        intent: "update_status",
        orderId,
        value: e.currentTarget.value,
      })}
      style={{ ...s.select, background: bg }}
    >
      {STATUS_OPTIONS.map((o) => (
        <option key={o.value} value={o.value}>{o.label}</option>
      ))}
    </select>
  );
}

function PriorityCell({ orderId, value }: { orderId: number; value: string }) {
  const fetcher = useFetcher();
  const current = fetcher.formData ? String(fetcher.formData.get("value")) : value;
  const opt = PRIORITY_OPTIONS.find((o) => o.value === current);

  return (
    <select
      value={current}
      onChange={(e) => submitPortalCell(fetcher, {
        intent: "update_priority",
        orderId,
        value: e.currentTarget.value,
      })}
      style={{
        ...s.select,
        background: opt?.bg ?? "#f3f4f6",
        color: opt?.color ?? "#374151",
        fontWeight: 700,
      }}
    >
      <option value="">— Priority —</option>
      {PRIORITY_OPTIONS.map((o) => (
        <option key={o.value} value={o.value}>{o.label}</option>
      ))}
    </select>
  );
}

function NotesCell({ orderId, field, value }: { orderId: number; field: string; value: string }) {
  const fetcher = useFetcher();
  return (
    <textarea
      defaultValue={value}
      onBlur={(e) => submitPortalCell(fetcher, {
        intent: `update_${field}`,
        orderId,
        value: e.currentTarget.value,
      })}
      rows={2}
      style={s.textarea}
      placeholder="Add note…"
    />
  );
}

function EtaCell({ orderId, value }: { orderId: number; value: string }) {
  const fetcher = useFetcher();
  return (
    <input
      type="text"
      defaultValue={value}
      onBlur={(e) => submitPortalCell(fetcher, {
        intent: "update_eta",
        orderId,
        value: e.currentTarget.value,
      })}
      style={s.dateInput}
      placeholder="dd/mm/yy"
    />
  );
}

function QtyCell({ orderId, size, value }: { orderId: number; size: string; value: number }) {
  const fetcher = useFetcher();
  const current = fetcher.formData ? String(fetcher.formData.get("value")) : String(value);
  const numericCurrent = Number(current) || 0;
  const normalizeQty = (input: HTMLInputElement) => {
    input.value = input.value.replace(/\D/g, "");
  };

  return (
    <input
      type="text"
      inputMode="numeric"
      pattern="[0-9]*"
      defaultValue={value}
      onChange={(e) => normalizeQty(e.currentTarget)}
      onBlur={(e) => submitPortalCell(fetcher, {
        intent: "update_qty",
        orderId,
        size,
        value: e.currentTarget.value,
      })}
      style={{
        ...s.qtyInput,
        ...(numericCurrent > 0 ? s.qtyInputActive : s.qtyInputZero),
      }}
    />
  );
}

function DeleteCell({ orderId }: { orderId: number }) {
  const fetcher = useFetcher();
  const confirmedRef = useRef(false);
  const [confirmOpen, setConfirmOpen] = useState(false);

  const skipConfirmForToday = () => {
    window.localStorage.setItem(
      DELETE_CONFIRM_SKIP_KEY,
      String(Date.now() + 24 * 60 * 60 * 1000),
    );
  };

  const shouldSkipConfirm = () => {
    const skipUntil = Number(window.localStorage.getItem(DELETE_CONFIRM_SKIP_KEY) ?? 0);
    return skipUntil > Date.now();
  };

  return (
    <fetcher.Form
      method="post"
      onSubmit={(e) => {
        if (confirmedRef.current || shouldSkipConfirm()) {
          confirmedRef.current = false;
          return;
        }

        e.preventDefault();
        setConfirmOpen(true);
      }}
    >
      <input type="hidden" name="intent" value="delete_order" />
      <input type="hidden" name="orderId" value={orderId} />
      <button type="submit" style={s.deleteButton}>Delete</button>
      {confirmOpen && (
        <div style={s.deleteConfirm}>
          <div style={s.deleteConfirmCard}>
            <div style={s.deleteConfirmTitle}>Delete order?</div>
            <div style={s.deleteConfirmText}>Are you sure you want to delete this order?</div>
            <div style={s.deleteConfirmActions}>
              <button
                type="submit"
                style={{ ...s.deleteConfirmButton, ...s.deleteConfirmDanger }}
                onClick={() => { confirmedRef.current = true; }}
              >
                Yes, delete
              </button>
              <button type="button" style={s.deleteConfirmButton} onClick={() => setConfirmOpen(false)}>
                No
              </button>
              <button
                type="submit"
                style={s.deleteConfirmButton}
                onClick={() => {
                  skipConfirmForToday();
                  confirmedRef.current = true;
                }}
              >
                Don’t ask me again for a day
              </button>
            </div>
          </div>
        </div>
      )}
    </fetcher.Form>
  );
}

// ─── Table helpers ────────────────────────────────────────────────────────────

function Th({
  children,
  center,
  onResizeStart,
}: {
  children: React.ReactNode;
  center?: boolean;
  onResizeStart: (event: React.MouseEvent<HTMLSpanElement>) => void;
}) {
  return (
    <th style={{ ...s.th, textAlign: center ? "center" : "left" }}>
      <span style={s.thContent}>{children}</span>
      <span
        role="separator"
        aria-orientation="vertical"
        aria-label={`Resize ${String(children)} column`}
        onMouseDown={onResizeStart}
        style={s.resizeHandle}
      />
    </th>
  );
}
function Td({
  children,
  center,
  rowIndex,
  colIndex,
}: {
  children: React.ReactNode;
  center?: boolean;
  rowIndex: number;
  colIndex: number;
}) {
  return (
    <td
      data-grid-row={rowIndex}
      data-grid-col={colIndex}
      tabIndex={0}
      style={{ ...s.td, textAlign: center ? "center" : "left" }}
    >
      {children}
    </td>
  );
}

// ─── Styles ───────────────────────────────────────────────────────────────────

const s: Record<string, React.CSSProperties> = {
  appShell: {
    minHeight: "100vh",
    background: "#f3f4f6",
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
    display: "flex",
  },
  sidebar: {
    width: 230,
    flexShrink: 0,
    background: "#111827",
    color: "#fff",
    borderRight: "1px solid #0f172a",
    padding: "18px 14px",
    display: "flex",
    flexDirection: "column",
    position: "relative",
    transition: "width 160ms ease, padding 160ms ease",
  },
  sidebarCollapsed: { width: 68, padding: "18px 10px", alignItems: "center" },
  sidebarTop: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 10,
    marginBottom: 22,
  },
  sidebarTopCollapsed: {
    display: "flex",
    justifyContent: "center",
    marginBottom: 22,
    width: "100%",
  },
  collapseButton: {
    width: 32,
    height: 32,
    borderRadius: 999,
    border: "1px solid #334155",
    background: "#1f2937",
    color: "#f8fafc",
    fontWeight: 800,
    cursor: "pointer",
    boxShadow: "inset 0 1px 0 rgba(255,255,255,0.08)",
  },
  sidebarTitle: { fontSize: 17, fontWeight: 800 },
  nav: { display: "flex", flexDirection: "column", gap: 8 },
  iconNav: { display: "flex", flexDirection: "column", gap: 10, alignItems: "center", width: "100%" },
  navItem: {
    display: "block",
    color: "#cbd5e1",
    textDecoration: "none",
    borderRadius: 8,
    padding: "10px 12px",
    fontSize: 13,
    fontWeight: 700,
  },
  navSubItem: {
    display: "block",
    color: "#cbd5e1",
    textDecoration: "none",
    borderRadius: 8,
    padding: "8px 12px 8px 24px",
    fontSize: 12,
    fontWeight: 700,
  },
  navItemActive: { background: "#fff", color: "#111827" },
  iconNavItem: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: 40,
    height: 40,
    color: "#cbd5e1",
    textDecoration: "none",
    borderRadius: 10,
    fontSize: 13,
    fontWeight: 800,
    border: "1px solid rgba(148,163,184,0.18)",
    background: "rgba(15,23,42,0.35)",
  },
  iconNavItemActive: { background: "#fff", color: "#111827", borderColor: "#fff" },
  settingsLink: { marginTop: "auto" },
  count: { fontSize: 13, color: "#6b7280" },
  main: { flex: 1, minWidth: 0, padding: "24px 16px" },
  pageHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 16,
    marginBottom: 18,
  },
  pageTitle: { margin: 0, fontSize: 24, color: "#111827", lineHeight: 1.2 },
  headerControls: {
    display: "flex",
    flexDirection: "column",
    alignItems: "flex-end",
    gap: 8,
  },
  utilityBar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "flex-end",
    flexWrap: "wrap",
    gap: 12,
    padding: "6px 8px",
    border: "1px solid #dbe3ee",
    borderRadius: 10,
    background: "#f8fafc",
  },
  filters: { display: "flex", alignItems: "center", justifyContent: "flex-end", flexWrap: "wrap", gap: 10 },
  filterLabel: { display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontWeight: 700, color: "#374151" },
  productTypeFilter: {
    border: "1px solid #b6c0cc",
    borderRadius: 6,
    padding: "7px 10px",
    minWidth: 150,
    background: "#fff",
    fontSize: 13,
    fontWeight: 600,
  },
  searchInput: {
    border: "1px solid #b6c0cc",
    borderRadius: 6,
    padding: "7px 10px",
    width: 180,
    background: "#fff",
    fontSize: 13,
    fontWeight: 600,
    outline: "none",
  },
  activeUsers: {
    display: "flex",
    alignItems: "center",
    gap: 6,
    minHeight: 32,
  },
  activeUsersLabel: { color: "#4b5563", fontSize: 12, fontWeight: 800 },
  activeUserBadge: {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    width: 32,
    height: 32,
    borderRadius: 8,
    background: "#111827",
    color: "#fff",
    fontSize: 12,
    fontWeight: 800,
    letterSpacing: "0.04em",
  },
  activeUserEmpty: { color: "#6b7280", fontSize: 12, fontWeight: 700 },
  loginShell: {
    minHeight: "100vh",
    background: "linear-gradient(135deg, #f8fafc 0%, #e5edf6 100%)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: 24,
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
  },
  loginCard: {
    width: "min(420px, 100%)",
    background: "#fff",
    border: "1px solid #cbd5e1",
    borderRadius: 18,
    padding: 28,
    boxShadow: "0 28px 80px rgba(15,23,42,0.18)",
    display: "flex",
    flexDirection: "column",
    gap: 14,
  },
  loginTitle: { margin: 0, fontSize: 26, color: "#111827" },
  loginText: { margin: 0, color: "#6b7280", fontSize: 14, fontWeight: 600 },
  loginSelect: {
    border: "1px solid #b6c0cc",
    borderRadius: 8,
    padding: "10px 12px",
    fontSize: 15,
    fontWeight: 700,
    background: "#fff",
  },
  loginButton: {
    border: "1px solid #111827",
    borderRadius: 8,
    padding: "10px 14px",
    background: "#111827",
    color: "#fff",
    fontSize: 14,
    fontWeight: 800,
    cursor: "pointer",
  },
  settingsPanel: { display: "grid", gap: 16, maxWidth: 880 },
  settingsCard: {
    background: "#fff",
    border: "1px solid #cbd5e1",
    borderRadius: 14,
    padding: 18,
    boxShadow: "0 1px 2px rgba(15,23,42,0.08)",
  },
  settingsHeader: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 12 },
  settingsTitle: { margin: 0, fontSize: 18, color: "#111827" },
  settingsHint: { margin: "6px 0 0", color: "#6b7280", fontSize: 13, fontWeight: 600 },
  switchRow: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    marginTop: 16,
    fontSize: 14,
    fontWeight: 800,
    color: "#374151",
  },
  settingsWarning: {
    marginTop: 12,
    padding: 10,
    borderRadius: 8,
    background: "#fef3c7",
    color: "#92400e",
    fontSize: 13,
    fontWeight: 700,
  },
  secondaryButton: {
    border: "1px solid #d1d5db",
    borderRadius: 8,
    padding: "8px 10px",
    background: "#fff",
    color: "#374151",
    fontSize: 13,
    fontWeight: 800,
    cursor: "pointer",
    textDecoration: "none",
  },
  userList: { display: "grid", gap: 8, marginTop: 14 },
  userRow: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    padding: 10,
    borderRadius: 10,
    background: "#f8fafc",
    border: "1px solid #e5e7eb",
  },
  userName: { fontSize: 14, fontWeight: 800, color: "#111827", flex: 1 },
  adminPill: {
    padding: "3px 7px",
    borderRadius: 999,
    background: "#dcfce7",
    color: "#166534",
    fontSize: 11,
    fontWeight: 800,
  },
  removeUserButton: {
    border: "1px solid #fecaca",
    borderRadius: 7,
    padding: "6px 8px",
    background: "#fee2e2",
    color: "#991b1b",
    fontSize: 12,
    fontWeight: 800,
    cursor: "pointer",
  },
  addUserForm: { display: "flex", alignItems: "center", flexWrap: "wrap", gap: 10, marginTop: 16 },
  addUserInput: {
    border: "1px solid #b6c0cc",
    borderRadius: 8,
    padding: "10px 12px",
    minWidth: 260,
    fontSize: 14,
    fontWeight: 700,
  },
  adminCheckbox: { display: "flex", alignItems: "center", gap: 6, fontSize: 13, fontWeight: 800, color: "#374151" },
  packingLayout: {
    display: "grid",
    gridTemplateColumns: "minmax(0, 1fr)",
    gap: 10,
  },
  packingOverview: { display: "grid", gap: 14 },
  packingOverviewCreate: {
    display: "flex",
    alignItems: "flex-end",
    justifyContent: "space-between",
    gap: 16,
    flexWrap: "wrap",
    background: "#fff",
    border: "1px solid #cbd5e1",
    borderRadius: 12,
    padding: 14,
  },
  packingOverviewTableWrap: {
    overflow: "auto",
    background: "#fff",
    border: "1px solid #cbd5e1",
    boxShadow: "0 1px 2px rgba(15,23,42,0.08)",
  },
  packingToolbar: {
    background: "#fff",
    border: "1px solid #cbd5e1",
    borderRadius: 12,
    padding: "10px 12px",
    display: "flex",
    alignItems: "flex-end",
    flexWrap: "wrap",
    gap: 12,
  },
  packingCreateForm: { display: "flex", alignItems: "flex-end", flexWrap: "wrap", gap: 8 },
  packingInput: {
    border: "1px solid #b6c0cc",
    borderRadius: 7,
    padding: "8px 10px",
    fontSize: 13,
    fontWeight: 700,
    background: "#fff",
    minWidth: 130,
  },
  packingSelect: {
    border: "1px solid #b6c0cc",
    borderRadius: 7,
    padding: "8px 10px",
    fontSize: 13,
    fontWeight: 800,
    background: "#fff",
    minWidth: 360,
  },
  packingListNav: { display: "grid", gap: 8 },
  packingListLink: {
    display: "grid",
    gap: 3,
    padding: 10,
    borderRadius: 9,
    border: "1px solid #e5e7eb",
    background: "#f8fafc",
    color: "#374151",
    textDecoration: "none",
    fontSize: 13,
  },
  packingListLinkActive: { background: "#111827", color: "#fff", borderColor: "#111827" },
  packingDetail: { minWidth: 0 },
  packingDetailInner: { display: "grid", gap: 10 },
  packingTop: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-end",
    gap: 12,
    flexWrap: "wrap",
    background: "#fff",
    border: "1px solid #cbd5e1",
    borderRadius: 12,
    padding: 12,
  },
  packingMeta: { display: "flex", alignItems: "center", flexWrap: "wrap", gap: 10 },
  packingTotalPill: {
    alignSelf: "flex-end",
    border: "1px solid #cbd5e1",
    borderRadius: 999,
    padding: "8px 12px",
    background: "#f8fafc",
    color: "#374151",
    fontSize: 13,
    fontWeight: 800,
  },
  packingActions: { display: "flex", gap: 8 },
  productSearchPanel: {
    background: "#fff",
    border: "1px solid #cbd5e1",
    borderRadius: 12,
    padding: 12,
    display: "grid",
    gap: 10,
  },
  productSearchInput: {
    border: "1px solid #b6c0cc",
    borderRadius: 7,
    padding: "8px 10px",
    width: 280,
    fontSize: 13,
    fontWeight: 700,
  },
  productResults: { display: "grid", gap: 8 },
  productResult: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    padding: 8,
    borderRadius: 9,
    border: "1px solid #e5e7eb",
    background: "#f8fafc",
  },
  productResultImage: { width: 42, height: 56, objectFit: "cover", borderRadius: 4 },
  productResultText: { display: "grid", gap: 3, flex: 1, fontSize: 13, color: "#374151" },
  packingTableWrap: {
    maxHeight: "calc(100vh - 185px)",
    overflowX: "auto",
    overflowY: "visible",
    background: "#fff",
    border: "1px solid #cbd5e1",
    boxShadow: "0 1px 2px rgba(15,23,42,0.08)",
  },
  packingFooterActions: {
    display: "flex",
    justifyContent: "flex-start",
    alignItems: "center",
    gap: 8,
    padding: "0 0 4px",
  },
  packingThumb: { width: 108, height: 144, objectFit: "cover", borderRadius: 3 },
  fabricThumb: { width: 120, height: 120, objectFit: "cover", borderRadius: 3 },
  fabricImageDrop: {
    position: "relative",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    minHeight: 120,
    border: "1px dashed #94a3b8",
    borderRadius: 6,
    color: "#64748b",
    fontSize: 12,
    fontWeight: 800,
    overflow: "hidden",
    cursor: "pointer",
    background: "#f8fafc",
  },
  hiddenFileInput: {
    position: "absolute",
    inset: 0,
    opacity: 0,
    cursor: "pointer",
  },
  packingCellInput: {
    width: "100%",
    border: "1px solid transparent",
    background: "transparent",
    outline: "none",
    fontSize: 13,
    fontWeight: 700,
    color: "#111827",
    boxSizing: "border-box",
    textAlign: "inherit",
  },
  packingTextarea: {
    width: "100%",
    border: "1px solid transparent",
    background: "transparent",
    outline: "none",
    fontSize: 12,
    fontWeight: 700,
    color: "#4b5563",
    resize: "vertical",
    boxSizing: "border-box",
    fontFamily: "inherit",
  },
  packingSkuTextarea: {
    width: "100%",
    minHeight: 142,
    border: "1px solid transparent",
    background: "transparent",
    outline: "none",
    fontSize: 13,
    lineHeight: 1.35,
    fontWeight: 800,
    color: "#4b5563",
    resize: "none",
    boxSizing: "border-box",
    fontFamily: "monospace",
    whiteSpace: "pre-line",
    overflow: "hidden",
  },
  productCellSearch: {
    position: "relative",
    minWidth: 0,
  },
  productCellResults: {
    position: "fixed",
    top: 0,
    left: 0,
    width: 460,
    maxHeight: 320,
    overflow: "auto",
    background: "#ffffff",
    border: "1px solid #cbd5e1",
    borderRadius: 10,
    boxShadow: "0 22px 50px rgba(15,23,42,0.32)",
    zIndex: 2147483647,
    padding: 6,
    isolation: "isolate",
  },
  productCellResult: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    width: "100%",
    border: 0,
    borderRadius: 8,
    padding: 7,
    background: "#fff",
    color: "#111827",
    textAlign: "left",
    cursor: "pointer",
  },
  productCellResultImage: { width: 34, height: 46, objectFit: "cover", borderRadius: 4, flex: "0 0 auto" },
  productCellNoImage: {
    width: 34,
    height: 46,
    display: "grid",
    placeItems: "center",
    borderRadius: 4,
    background: "#f3f4f6",
    color: "#9ca3af",
    flex: "0 0 auto",
  },
  productCellResultText: { display: "grid", gap: 2, fontSize: 12, lineHeight: 1.25 },
  productCellResultEmpty: { padding: 10, fontSize: 12, fontWeight: 700, color: "#6b7280" },
  rowActions: { display: "grid", gap: 6 },
  smallLinkButton: {
    display: "inline-block",
    border: "1px solid #d1d5db",
    borderRadius: 6,
    padding: "5px 9px",
    background: "#fff",
    color: "#111827",
    fontSize: 11,
    fontWeight: 800,
    textDecoration: "none",
  },
  smallButton: {
    border: "1px solid #d1d5db",
    borderRadius: 6,
    padding: "5px 7px",
    background: "#fff",
    color: "#374151",
    fontSize: 11,
    fontWeight: 800,
    cursor: "pointer",
  },
  empty: { background: "#fff", borderRadius: 12, padding: 40, textAlign: "center", color: "#6b7280" },
  tableWrap: {
    maxHeight: "calc(100vh - 118px)",
    overflow: "auto",
    background: "#fff",
    border: "1px solid #cbd5e1",
    boxShadow: "0 1px 2px rgba(15,23,42,0.08)",
  },
  table: {
    borderCollapse: "separate",
    borderSpacing: 0,
    fontSize: 13,
    minWidth: 900,
    tableLayout: "fixed",
  },
  headerRow: { background: "#eef2f7" },
  th: {
    padding: "8px 10px",
    fontWeight: 700,
    fontSize: 11,
    textTransform: "uppercase",
    letterSpacing: "0.05em",
    color: "#4b5563",
    border: "1px solid #cbd5e1",
    whiteSpace: "nowrap",
    background: "#eef2f7",
    position: "sticky",
    top: 0,
    zIndex: 50,
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  thContent: { display: "block", overflow: "hidden", textOverflow: "ellipsis" },
  resizeHandle: {
    position: "absolute",
    top: 0,
    right: -3,
    width: 8,
    height: "100%",
    cursor: "col-resize",
    zIndex: 2,
    touchAction: "none",
  },
  row: { background: "#fff" },
  td: {
    padding: "8px 10px",
    verticalAlign: "middle",
    color: "#374151",
    border: "1px solid #d1d5db",
    background: "#fff",
    overflow: "hidden",
  },
  dropdownTd: {
    position: "relative",
  },
  imageCell: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: "100%",
  },
  thumb: { width: 86, height: 115, objectFit: "cover", borderRadius: 2, display: "block", margin: "0 auto" },
  noImg: { color: "#d1d5db", textAlign: "center" },
  productName: { fontWeight: 600, color: "#111827", whiteSpace: "normal", overflowWrap: "anywhere", lineHeight: 1.35 },
  sku: { fontFamily: "monospace", fontSize: 11, color: "#6b7280", whiteSpace: "pre-line" },
  qty: { fontWeight: 700, color: "#111827" },
  qtyZero: { color: "#d1d5db" },
  dateText: { color: "#374151", fontWeight: 600, whiteSpace: "nowrap" },
  total: { fontWeight: 700, fontSize: 14, color: "#111827" },
  noteText: { fontSize: 12, color: "#6b7280", maxWidth: 160, display: "block", whiteSpace: "pre-wrap" },
  select: {
    border: "1px solid #b6c0cc",
    borderRadius: 3,
    padding: "5px 8px",
    fontSize: 12,
    fontWeight: 600,
    cursor: "pointer",
    outline: "none",
    width: "100%",
  },
  textarea: {
    border: "1px solid #cbd5e1",
    borderRadius: 3,
    padding: "6px 8px",
    fontSize: 12,
    resize: "vertical",
    width: "100%",
    boxSizing: "border-box",
    fontFamily: "inherit",
    outline: "none",
    color: "#374151",
  },
  dateInput: {
    border: "1px solid #b6c0cc",
    borderRadius: 3,
    padding: "5px 8px",
    fontSize: 12,
    outline: "none",
    color: "#374151",
    width: "100%",
    boxSizing: "border-box",
  },
  qtyInput: {
    display: "block",
    width: "100%",
    border: "1px solid transparent",
    borderRadius: 3,
    padding: "4px 0",
    fontSize: 13,
    fontWeight: 700,
    textAlign: "center",
    outline: "none",
    background: "transparent",
    boxSizing: "border-box",
  },
  qtyInputActive: { color: "#111827" },
  qtyInputZero: { color: "#d1d5db" },
  deleteButton: {
    border: "1px solid #fecaca",
    borderRadius: 3,
    padding: "5px 8px",
    background: "#fee2e2",
    color: "#991b1b",
    fontSize: 12,
    fontWeight: 700,
    cursor: "pointer",
  },
  deleteConfirm: {
    position: "fixed",
    inset: 0,
    zIndex: 1000,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: 24,
    background: "rgba(15, 23, 42, 0.38)",
  },
  deleteConfirmCard: {
    width: "min(420px, 100%)",
    padding: 22,
    border: "1px solid #fecaca",
    borderRadius: 14,
    background: "#fff",
    textAlign: "left",
    boxShadow: "0 24px 70px rgba(15, 23, 42, 0.28)",
  },
  deleteConfirmTitle: {
    marginBottom: 8,
    color: "#111827",
    fontSize: 18,
    fontWeight: 800,
  },
  deleteConfirmText: {
    marginBottom: 18,
    color: "#4b5563",
    fontSize: 14,
    fontWeight: 600,
    lineHeight: 1.45,
  },
  deleteConfirmActions: {
    display: "flex",
    flexWrap: "wrap",
    justifyContent: "flex-end",
    gap: 10,
  },
  deleteConfirmButton: {
    border: "1px solid #d1d5db",
    borderRadius: 8,
    padding: "9px 12px",
    background: "#fff",
    color: "#374151",
    fontSize: 13,
    fontWeight: 700,
    cursor: "pointer",
  },
  deleteConfirmDanger: {
    borderColor: "#fca5a5",
    background: "#fee2e2",
    color: "#991b1b",
  },
};
