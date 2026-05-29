import bcrypt from "bcryptjs";
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import type { ActionFunctionArgs, LoaderFunctionArgs, ShouldRevalidateFunction } from "react-router";
import { isRouteErrorResponse, useActionData, useFetcher, useLoaderData, useRevalidator, useRouteError, useSearchParams, useSubmit } from "react-router";
import prisma from "../db.server";
import { fabricStockSheets as initialFabricStockSheets, type FabricStockSheet } from "../fabric-stock-data";
import { syncOrderNoteMessages, syncSampleIterationMessages } from "../portal-messages.server";
import { VisionBoardV2Panel } from "../portal-vision-board";
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
  const serverSearchTitle = page === "restock" ? "" : searchTitle;
  const messageOrderId = Number(url.searchParams.get("messageOrderId") ?? 0) || null;
  const packingId = Number(url.searchParams.get("packingId") ?? 0) || null;
  const productSearch = url.searchParams.get("productSearch") ?? "";
  const restockProductSearch = url.searchParams.get("restockProductSearch") ?? "";
  const packingSearchLineId = Number(url.searchParams.get("packingSearchLineId") ?? 0) || null;
  const sortBy = url.searchParams.get("sortBy") ?? "orderDateDesc";
  try {
  const SETTING_KEYS = [
    COLUMN_WIDTHS_KEY,
    PACKING_COLUMN_WIDTHS_KEY,
    TABLE_HEADER_LABELS_KEY,
    TABLE_CUSTOM_COLUMNS_KEY,
    TABLE_CUSTOM_CELLS_KEY,
    TABLE_ROW_HEIGHTS_KEY,
    FABRIC_CUSTOM_SHEETS_KEY,
    FABRIC_MANUAL_SHEETS_KEY,
    RESTOCK_SETTINGS_KEY,
    UNIVERSAL_SETTINGS_KEY,
    FABRIC_SETTINGS_KEY,
    PRODUCT_INFO_KEY,
    PORTAL_USERS_KEY,
    PORTAL_ACTIVE_USERS_KEY,
    PORTAL_NAV_ORDER_KEY,
  ];
  const needsOrders = page === "restock" || page === "packing";
  const needsPackingLists = page === "packing" || packingId !== null;
  const needsActivityLogs = page !== "visionboard" && page !== "samples";

  const [settingsRows, allOrders, packingLists] = await retryAsync(() => Promise.all([
      prisma.portalSetting.findMany({ where: { key: { in: SETTING_KEYS } }, select: { key: true, value: true } }),
      needsOrders
        ? prisma.supplierOrder.findMany({
            where: { status: "open" },
            include: { lines: { orderBy: { id: "asc" } } },
            orderBy: { createdAt: "desc" },
          })
        : (Promise.resolve([]) as ReturnType<typeof prisma.supplierOrder.findMany<{ include: { lines: true } }>>),
      needsPackingLists
        ? prisma.packingList.findMany({
            orderBy: { createdAt: "desc" },
            include: { lines: { orderBy: [{ sortOrder: "asc" }, { id: "asc" }] } },
          })
        : (Promise.resolve([]) as ReturnType<typeof prisma.packingList.findMany<{ include: { lines: true } }>>),
    ]),
    "portal base data",
  );
  const settingsMap = new Map(settingsRows.map((row) => [row.key, row.value as unknown]));
  const wrap = (key: string) => {
    const value = settingsMap.get(key);
    return value === undefined ? null : { value };
  };
  const columnWidthsSetting = wrap(COLUMN_WIDTHS_KEY);
  const packingColumnWidthsSetting = wrap(PACKING_COLUMN_WIDTHS_KEY);
  const headerLabelsSetting = wrap(TABLE_HEADER_LABELS_KEY);
  const customColumnsSetting = wrap(TABLE_CUSTOM_COLUMNS_KEY);
  const customCellsSetting = wrap(TABLE_CUSTOM_CELLS_KEY);
  const rowHeightsSetting = wrap(TABLE_ROW_HEIGHTS_KEY);
  const fabricCustomSheetsSetting = wrap(FABRIC_CUSTOM_SHEETS_KEY);
  const fabricManualSheetsSetting = wrap(FABRIC_MANUAL_SHEETS_KEY);
  const restockSettingsSetting = wrap(RESTOCK_SETTINGS_KEY);
  const universalSettingsSetting = wrap(UNIVERSAL_SETTINGS_KEY);
  const fabricSettingsSetting = wrap(FABRIC_SETTINGS_KEY);
  const productInfoSetting = wrap(PRODUCT_INFO_KEY);
  const usersSetting = wrap(PORTAL_USERS_KEY);
  const activeUsersSetting = wrap(PORTAL_ACTIVE_USERS_KEY);
  const navOrderSetting = wrap(PORTAL_NAV_ORDER_KEY);
  const samples = page === "samples"
    ? await (async () => {
        try {
          // Raw SQL: ship NO image bytes in the loader — only metadata + image
          // count. Thumbnails are batch-fetched after first paint.
          type RawSample = { id: number; sortOrder: number; name: string; createdAt: Date; updatedAt: Date };
          type RawIterSlim = {
            id: number; sampleId: number; version: number; name: string | null; notes: string | null;
            fabricType: string | null; sampleSize: string | null; buttonType: string | null; factoryCost: string | null; status: string;
            imageCount: number; hasThumbnail: boolean; taggedUsers: unknown;
            createdAt: Date; updatedAt: Date;
          };
          const rawSamples = await prisma.$queryRaw<RawSample[]>`SELECT id, "sortOrder", name, "createdAt", "updatedAt" FROM "Sample" ORDER BY "sortOrder" ASC, "createdAt" DESC`;
          const rawIters = await prisma.$queryRaw<RawIterSlim[]>`
            SELECT id, "sampleId", version, name, notes, "fabricType", "sampleSize", "buttonType", "factoryCost", status,
              CASE WHEN jsonb_typeof(images) = 'array'
                THEN jsonb_array_length(images) ELSE 0
              END AS "imageCount",
              (thumbnail IS NOT NULL) AS "hasThumbnail",
              "taggedUsers", "createdAt", "updatedAt"
            FROM "SampleIteration"
            ORDER BY version ASC
          `;
          return rawSamples.map((s) => ({
            ...s,
            iterations: rawIters.filter((it) => it.sampleId === s.id).map((it) => ({
              id: it.id, sampleId: it.sampleId, version: it.version,
              name: it.name, notes: it.notes,
              fabricType: it.fabricType, sampleSize: it.sampleSize, buttonType: it.buttonType, factoryCost: it.factoryCost,
              status: it.status,
              images: [] as string[],
              imageCount: Number(it.imageCount),
              hasThumbnail: Boolean(it.hasThumbnail),
              taggedUsers: it.taggedUsers,
              createdAt: it.createdAt, updatedAt: it.updatedAt,
            })),
          }));
        } catch {
          // Fallback path 1: legacy schema without fabricType/sampleSize/buttonType columns
          try {
            type RawSample = { id: number; sortOrder: number; name: string; createdAt: Date; updatedAt: Date };
            type RawIter1 = { id: number; sampleId: number; version: number; name: string | null; notes: string | null; status: string; imageCount: number; taggedUsers: unknown; createdAt: Date; updatedAt: Date };
            const rawSamples = await prisma.$queryRaw<RawSample[]>`SELECT id, "sortOrder", name, "createdAt", "updatedAt" FROM "Sample" ORDER BY "sortOrder" ASC, "createdAt" DESC`;
            const rawIters = await prisma.$queryRaw<RawIter1[]>`
              SELECT id, "sampleId", version, name, notes, status,
                CASE WHEN jsonb_typeof(images) = 'array'
                  THEN jsonb_array_length(images) ELSE 0
                END AS "imageCount",
                "taggedUsers", "createdAt", "updatedAt"
              FROM "SampleIteration" ORDER BY version ASC
            `;
            return rawSamples.map((s) => ({
              ...s,
              iterations: rawIters.filter((it) => it.sampleId === s.id).map((it) => ({
                id: it.id, sampleId: it.sampleId, version: it.version, name: it.name, notes: it.notes,
                fabricType: null, sampleSize: null, buttonType: null, factoryCost: null, status: it.status,
                images: [] as string[], imageCount: Number(it.imageCount),
                taggedUsers: it.taggedUsers, createdAt: it.createdAt, updatedAt: it.updatedAt,
              })),
            }));
          } catch {
            // Fallback path 2: legacy schema without name/taggedUsers (pre-000001)
            try {
              type RawSample = { id: number; sortOrder: number; name: string; createdAt: Date; updatedAt: Date };
              type RawIter0 = { id: number; sampleId: number; version: number; notes: string | null; status: string; imageCount: number; createdAt: Date; updatedAt: Date };
              const rawSamples = await prisma.$queryRaw<RawSample[]>`SELECT id, "sortOrder", name, "createdAt", "updatedAt" FROM "Sample" ORDER BY "sortOrder" ASC, "createdAt" DESC`;
              const rawIters = await prisma.$queryRaw<RawIter0[]>`
                SELECT id, "sampleId", version, notes, status,
                  CASE WHEN jsonb_typeof(images) = 'array'
                    THEN jsonb_array_length(images) ELSE 0
                  END AS "imageCount",
                  "createdAt", "updatedAt"
                FROM "SampleIteration" ORDER BY version ASC
              `;
              return rawSamples.map((s) => ({
                ...s,
                iterations: rawIters.filter((it) => it.sampleId === s.id).map((it) => ({
                  id: it.id, sampleId: it.sampleId, version: it.version, name: null, notes: it.notes,
                  fabricType: null, sampleSize: null, buttonType: null, factoryCost: null, status: it.status,
                  images: [] as string[], imageCount: Number(it.imageCount),
                  taggedUsers: [], createdAt: it.createdAt, updatedAt: it.updatedAt,
                })),
              }));
            } catch {
              return [];
            }
          }
        }
      })()
    : [];
  // Vision Board V2 — slim loader that only ships the active board's items
  // (id/name/sortOrder/imageCount/hasThumbnail/updatedAt). Drawer fetches the
  // full item on open. One-time copy from the original VisionBoard tables
  // happens automatically on first visit; idempotent thereafter.
  const visionBoardData = page === "visionboard"
    ? await (async () => {
        const v2Count = await prisma.visionBoardV2.count();
        if (v2Count === 0) {
          const v1Count = await prisma.visionBoard.count();
          if (v1Count > 0) {
            const v1Boards = await prisma.visionBoard.findMany({
              orderBy: [{ sortOrder: "asc" }, { createdAt: "asc" }],
              include: { items: { orderBy: [{ sortOrder: "asc" }, { createdAt: "asc" }] } },
            });
            for (const b of v1Boards) {
              const newBoard = await prisma.visionBoardV2.create({
                data: { name: b.name, sortOrder: b.sortOrder },
              });
              if (b.items.length === 0) continue;
              await prisma.visionBoardV2Item.createMany({
                data: b.items.map((it) => ({
                  boardId: newBoard.id,
                  name: it.name,
                  sortOrder: it.sortOrder,
                  images: it.images as object,
                  thumbnail: it.thumbnail,
                  fields: it.fields as object,
                  notes: it.notes,
                })),
              });
            }
          }
        }

        const requestedBoardId = Number(url.searchParams.get("boardId") ?? 0) || null;
        const boards = await prisma.visionBoardV2.findMany({
          orderBy: [{ sortOrder: "asc" }, { createdAt: "asc" }],
          select: { id: true, name: true, sortOrder: true },
        });
        const activeBoardId = requestedBoardId && boards.some((b) => b.id === requestedBoardId)
          ? requestedBoardId
          : boards[0]?.id ?? null;

        let items: Array<{
          id: number; name: string; sortOrder: number;
          imageCount: number; hasThumbnail: boolean; updatedAt: Date;
        }> = [];
        if (activeBoardId) {
          const raw = await prisma.$queryRaw<Array<{
            id: number; name: string; sortOrder: number;
            imageCount: number; hasThumbnail: boolean; updatedAt: Date;
          }>>`
            SELECT
              id, name, "sortOrder",
              CASE WHEN jsonb_typeof(images) = 'array'
                THEN jsonb_array_length(images)
                ELSE 0
              END AS "imageCount",
              (thumbnail IS NOT NULL) AS "hasThumbnail",
              "updatedAt"
            FROM "VisionBoardV2Item"
            WHERE "boardId" = ${activeBoardId}
            ORDER BY "sortOrder" ASC, "createdAt" ASC
          `;
          items = raw.map((it) => ({
            ...it,
            imageCount: Number(it.imageCount),
            hasThumbnail: Boolean(it.hasThumbnail),
          }));
        }
        return { boards, activeBoardId, items };
      })()
    : { boards: [] as Array<{ id: number; name: string; sortOrder: number }>, activeBoardId: null as number | null, items: [] as Array<{ id: number; name: string; sortOrder: number; imageCount: number; hasThumbnail: boolean; updatedAt: Date }> };
  // Collections listing — slim projection (no rows JSON, no thumbnail bytes).
  // The drawer fetches the full collection (rows + columns) on open.
  const collections = page === "collections"
    ? await prisma.$queryRawUnsafe<Array<{
        id: number; name: string; sortOrder: number;
        hasThumbnail: boolean; rowCount: number;
        createdAt: Date; updatedAt: Date;
      }>>(`
        SELECT id, name, "sortOrder",
          (thumbnail IS NOT NULL) AS "hasThumbnail",
          CASE WHEN jsonb_typeof(rows) = 'array' THEN jsonb_array_length(rows) ELSE 0 END AS "rowCount",
          "createdAt", "updatedAt"
        FROM "Collection"
        ORDER BY "sortOrder" ASC, "createdAt" ASC
      `).then((rows) => rows.map((r) => ({
        id: r.id, name: r.name, sortOrder: r.sortOrder,
        hasThumbnail: Boolean(r.hasThumbnail),
        rowCount: Number(r.rowCount),
        createdAt: r.createdAt, updatedAt: r.updatedAt,
      }))).catch(() => [] as Array<{ id: number; name: string; sortOrder: number; hasThumbnail: boolean; rowCount: number; createdAt: Date; updatedAt: Date }>)
    : [];
  const ninetyDaysAgo = new Date(Date.now() - 90 * 24 * 60 * 60 * 1000);
  const activityLogs = needsActivityLogs
    ? await prisma.activityLog.findMany({
        where: { createdAt: { gte: ninetyDaysAgo } },
        orderBy: { createdAt: "desc" },
        take: 1000,
      }).catch(() => [] as { id: number; userName: string; action: string; entity: string; entityId: string | null; entityName: string | null; field: string | null; toValue: string | null; createdAt: Date }[])
    : [] as { id: number; userName: string; action: string; entity: string; entityId: string | null; entityName: string | null; field: string | null; toValue: string | null; createdAt: Date }[];
  const users = normalizePortalUsers(usersSetting?.value);
  const customColumns = normalizeTableCustomColumns(customColumnsSetting?.value);
  const customCells = normalizeTableCustomCells(customCellsSetting?.value);
  const rowHeights = normalizeTableRowHeights(rowHeightsSetting?.value);
  const restockSettings = normalizeRestockSettings(restockSettingsSetting?.value);
  const universalSettings = normalizeUniversalSettings(universalSettingsSetting?.value);
  let fabricSettings = normalizeFabricSettings(fabricSettingsSetting?.value);
  const productInfo = normalizeProductInfo(productInfoSetting?.value);
  const usersWithSeed = await ensureSuperAdmin(users);
  const currentUser = getCurrentPortalUser(request, usersWithSeed);
  const activeUsers = await recordAndGetActiveUsers(currentUser, usersWithSeed, activeUsersSetting?.value)
    .catch((error) => {
      console.error("Active user tracking failed", error);
      return [] as ActivePortalUser[];
    });
  const normalizedOrders = allOrders.map((order) => ({
    ...order,
    productType: normalizeProductGroup(order.productType) || null,
  }));
  const productGroups = Array.from(new Set(normalizedOrders.map((order) => order.productType).filter(Boolean) as string[]))
    .sort((a, b) => a.localeCompare(b));
  const statusFilters = Array.from(new Set(normalizedOrders.map((order) => order.supplierStatus).filter(Boolean)))
    .sort((a, b) => labelForOption(restockSettings.statusOptions, a).localeCompare(labelForOption(restockSettings.statusOptions, b)));
  const statusFilterCounts = normalizedOrders
    .filter((order) => !selectedProductGroup || order.productType === selectedProductGroup)
    .filter((order) => !selectedPriority || order.priority === selectedPriority)
    .reduce<Record<string, number>>((counts, order) => {
      if (order.supplierStatus) counts[order.supplierStatus] = (counts[order.supplierStatus] ?? 0) + 1;
      return counts;
    }, {});
  const priorityFilters = Array.from(new Set(normalizedOrders.map((order) => order.priority).filter(Boolean) as string[]))
    .sort((a, b) => labelForOption(restockSettings.priorityOptions, a).localeCompare(labelForOption(restockSettings.priorityOptions, b)));
  const filteredOrders = normalizedOrders
    .filter((order) => !selectedProductGroup || order.productType === selectedProductGroup)
    .filter((order) => !selectedStatus || order.supplierStatus === selectedStatus)
    .filter((order) => !selectedPriority || order.priority === selectedPriority)
    .filter((order) => !serverSearchTitle || order.productTitle.toLowerCase().includes(serverSearchTitle.toLowerCase()))
    .filter((order) => !messageOrderId || order.id === messageOrderId)
    .sort((a, b) => {
      if (sortBy === "titleAsc") return a.productTitle.localeCompare(b.productTitle);
      if (sortBy === "titleDesc") return b.productTitle.localeCompare(a.productTitle);
      if (sortBy === "orderDateAsc") return a.createdAt.getTime() - b.createdAt.getTime();
      return b.createdAt.getTime() - a.createdAt.getTime();
    });
  const orders = page === "restock"
    ? await enrichOrdersWithShopifyVariants(filteredOrders)
    : filteredOrders;
  // Defensive: hydrate shippingMethod from the DB in case the Prisma client
  // wasn't regenerated on this environment to know about the field. The
  // ADD COLUMN IF NOT EXISTS makes this safe on a stale schema too.
  if ((page === "packing" || page === "restock") && packingLists.length) {
    try {
      await prisma.$executeRawUnsafe(`ALTER TABLE "PackingList" ADD COLUMN IF NOT EXISTS "shippingMethod" TEXT`);
      const shippingRows = await prisma.$queryRawUnsafe<Array<{ id: number; shippingMethod: string | null }>>(
        `SELECT id, "shippingMethod" FROM "PackingList"`,
      );
      const map = new Map(shippingRows.map((row) => [row.id, row.shippingMethod]));
      for (const list of packingLists) {
        (list as { shippingMethod: string | null }).shippingMethod = map.get(list.id) ?? null;
      }
      const nonNullCount = shippingRows.filter((r) => r.shippingMethod).length;
      if (nonNullCount > 0) {
        console.log(`[shippingMethod] hydrated ${nonNullCount}/${shippingRows.length} packing lists with a value`);
      }
    } catch (e) {
      console.warn("[shippingMethod] hydration failed:", e);
    }
  }
  const selectedPackingList = packingId
    ? packingLists.find((list) => list.id === packingId) ?? null
    : null;
  // Shop domain for "open in Shopify admin" links and per-row inventory loads.
  const shopDomain = page === "packing" && selectedPackingList
    ? (await prisma.session.findFirst({ where: { accessToken: { not: "" } }, orderBy: { isOnline: "asc" }, select: { shop: true } }))?.shop ?? null
    : null;
  const productResults = page === "packing" && selectedPackingList && productSearch.trim().length >= 2
    ? await searchShopifyProducts(productSearch)
    : [];
  const restockProductResults = page === "restock" && restockProductSearch.trim().length >= 2
    ? await searchShopifyProducts(restockProductSearch)
    : [];
  const messages = currentUser
    ? await retryAsync(() => prisma.portalMessage.findMany({
        where: { userId: currentUser.id, readAt: null },
        orderBy: { createdAt: "desc" },
        take: 25,
      }), "portal messages").catch((error) => {
        console.error("Portal messages failed", error);
        return [] as PortalMessageItem[];
      })
    : [];
  const needsFabricSheets = page === "fabric" || page === "restock";
  const fabricCellOverridesSetting = needsFabricSheets
    ? await prisma.portalSetting.findUnique({
        where: { key: FABRIC_CELL_OVERRIDES_KEY },
        select: { value: true },
      })
    : null;
  const fabricCustomRowsSetting = needsFabricSheets
    ? await prisma.portalSetting.findUnique({
        where: { key: FABRIC_CUSTOM_ROWS_KEY },
        select: { value: true },
      })
    : null;
  const fabricDeletedRowsSetting = needsFabricSheets
    ? await prisma.portalSetting.findUnique({
        where: { key: FABRIC_DELETED_ROWS_KEY },
        select: { value: true },
      })
    : null;
  const fabricDeletedSheetsSetting = needsFabricSheets
    ? await prisma.portalSetting.findUnique({
        where: { key: FABRIC_DELETED_SHEETS_KEY },
        select: { value: true },
      })
    : null;
  const inrToAudRate = page === "fabric" ? await getInrToAudRate() : null;
  const manualFabricSheets = needsFabricSheets
    ? await getManualFabricSheets({
        savedValue: fabricManualSheetsSetting?.value,
        customSheetsValue: fabricCustomSheetsSetting?.value,
        customRowsValue: fabricCustomRowsSetting?.value,
        deletedRowsValue: fabricDeletedRowsSetting?.value,
        overridesValue: fabricCellOverridesSetting?.value,
        deletedSheetsValue: fabricDeletedSheetsSetting?.value,
      })
    : [];
  const fabricStockIndex: FabricStockEntry[] = page === "restock"
    ? buildFabricStockIndex(manualFabricSheets)
    : [];
  if (page === "fabric") {
    fabricSettings = {
      ...fabricSettings,
      fabricTypeOptions: ensureFabricTypeChipsForRows(fabricSettings.fabricTypeOptions, manualFabricSheets),
    };
  }
  // Cache buster for the fabric image URLs we generate below. Using the
  // blob's updatedAt means every edit invalidates every image URL, which is
  // wasteful but correct — browsers refetch images via HTTP cache.
  const fabricSheets = page === "fabric"
    ? getFabricSheets(
        undefined,
        undefined,
        undefined,
        fabricSettings.tileOrder,
        [],
        customColumns.fabric,
        {},
        manualFabricSheets,
      )
        .filter((sheet) => isCombinedFabricSource(sheet))
        .map(padCombinedFabricSheet)
    : [];

  // Collect all unique size names across all orders, sorted logically
  const sizeOrder = ["XS","S","M","L","XL","2XL","3XL","S-M","M-L","L-XL","S/M","M/L","L/XL","4XL","ONE SIZE"];
  const allSizes = [...new Set([
    ...orders.flatMap((o) => o.lines.map((l) => l.variantTitle)),
    ...restockProductResults.flatMap((product) => product.variants.map((variant) => variant.title)),
  ])];
  allSizes.sort((a, b) => {
    const ai = sizeOrder.indexOf(a.toUpperCase());
    const bi = sizeOrder.indexOf(b.toUpperCase());
    if (ai === -1 && bi === -1) return a.localeCompare(b);
    if (ai === -1) return 1;
    if (bi === -1) return -1;
    return ai - bi;
  });

  const rawNavOrder = Array.isArray(navOrderSetting?.value) ? navOrderSetting.value as string[] : DEFAULT_NAV_ORDER;
  const navOrder = DEFAULT_NAV_ORDER.map((id) => id as NavItemId)
    .sort((a, b) => {
      const ai = rawNavOrder.indexOf(a);
      const bi = rawNavOrder.indexOf(b);
      return (ai === -1 ? 999 : ai) - (bi === -1 ? 999 : bi);
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
    statusFilterCounts,
    priorityFilters,
    sortBy,
    page,
    columnWidths: normalizeColumnWidths(columnWidthsSetting?.value),
    packingColumnWidths: normalizeColumnWidths(packingColumnWidthsSetting?.value),
    tableHeaderLabels: normalizeTableHeaderLabels(headerLabelsSetting?.value),
    customColumns,
    customCells,
    rowHeights,
    restockSettings,
    universalSettings,
    fabricSettings,
    productInfo,
    packingLists,
    selectedPackingList,
    shopDomain,
    productSearch,
    restockProductSearch,
    packingSearchLineId,
    productResults,
    restockProductResults,
    users: usersWithSeed,
    currentUser,
    activeUsers,
    messages,
    messageOrderId,
    loginBlocked: !currentUser,
    activityLogs,
    navOrder,
    fabricSheets,
    inrToAudRate,
    samples,
    visionBoardData,
    collections,
    fabricStockIndex,
  };
  } catch (error) {
    console.error("Portal loader error:", error);
    throw new Response("Failed to load portal data. Please refresh the page.", { status: 503 });
  }
};

// Return a raw JSON Response (bypasses React Router's single-fetch turbo-stream
// encoding) so plain `fetch()` callers can JSON.parse the body. useFetcher
// callers also work because they read .json() under the hood.
function jsonResponse<T>(body: T): Response {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { "Content-Type": "application/json; charset=utf-8" },
  });
}

export const action = async ({ request }: ActionFunctionArgs) => {
  try {
  const form = await request.formData();
  const intent = String(form.get("intent"));
  const orderId = Number(form.get("orderId"));
  const usersSetting = await prisma.portalSetting.findUnique({
    where: { key: PORTAL_USERS_KEY },
    select: { value: true },
  });
  const rawUsers = normalizePortalUsers(usersSetting?.value);
  const users = await ensureSuperAdmin(rawUsers);
  const currentUser = getCurrentPortalUser(request, users);
  const canManageUsers = currentUser?.role === "superadmin" || currentUser?.role === "admin";
  const canLoadPackingInventory = canPortalUserLoadPackingInventory(users, currentUser);

  const updates: Record<string, unknown> = {};

  if (intent === "portal_login") {
    const username = String(form.get("username") ?? "").trim().toLowerCase();
    const password = String(form.get("password") ?? "");
    const user = users.find((u) => u.name.toLowerCase() === username && u.active);
    if (!user || !user.passwordHash) return { loginError: "Invalid username or password" };
    const valid = await bcrypt.compare(password, user.passwordHash);
    if (!valid) return { loginError: "Invalid username or password" };
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

  if (intent === "add_portal_user") {
    if (!canManageUsers) return null;
    const name = String(form.get("name") ?? "").trim();
    const password = String(form.get("password") ?? "");
    const roleRaw = String(form.get("role") ?? "user");
    if (!name || !password) return { userError: "Name and password are required" };
    if (users.some((u) => u.name.toLowerCase() === name.toLowerCase())) return { userError: "A user with that name already exists" };
    const allowedRoles: PortalUserRole[] = currentUser?.role === "superadmin" ? ["superadmin", "admin", "user"] : ["admin", "user"];
    const role = allowedRoles.includes(roleRaw as PortalUserRole) ? (roleRaw as PortalUserRole) : "user";
    const passwordHash = await bcrypt.hash(password, 10);
    const isAdmin = role === "superadmin" || role === "admin";
    await savePortalUsers([...users, { id: crypto.randomUUID(), name, username: name.toLowerCase(), passwordHash, role, admin: isAdmin, canLoadInventory: isAdmin, active: true, pageAccess: {} }]);
    return null;
  }

  if (intent === "update_portal_user") {
    if (!canManageUsers) return null;
    const userId = String(form.get("userId") ?? "");
    const target = users.find((u) => u.id === userId);
    if (!target) return null;
    if (target.role === "superadmin" && currentUser?.role !== "superadmin") return null;
    const updated: PortalUser = { ...target };
    const newName = String(form.get("name") ?? "").trim();
    if (newName) {
      if (newName.toLowerCase() !== target.name.toLowerCase() && users.some((u) => u.id !== userId && u.name.toLowerCase() === newName.toLowerCase())) return { userError: "A user with that name already exists" };
      updated.name = newName;
      updated.username = newName.toLowerCase();
    }
    const newPassword = String(form.get("password") ?? "");
    if (newPassword) updated.passwordHash = await bcrypt.hash(newPassword, 10);
    const roleRaw = String(form.get("role") ?? "");
    if (roleRaw) {
      const allowedRoles: PortalUserRole[] = currentUser?.role === "superadmin" ? ["superadmin", "admin", "user"] : ["admin", "user"];
      if (allowedRoles.includes(roleRaw as PortalUserRole)) {
        updated.role = roleRaw as PortalUserRole;
        updated.admin = updated.role === "superadmin" || updated.role === "admin";
      }
    }
    if (form.has("pageAccess")) {
      try { updated.pageAccess = JSON.parse(String(form.get("pageAccess"))); } catch { /* keep existing */ }
    }
    if (form.has("canLoadInventory")) updated.canLoadInventory = form.get("canLoadInventory") === "on";
    await savePortalUsers(users.map((u) => u.id === userId ? updated : u));
    return null;
  }

  if (intent === "remove_portal_user") {
    if (!canManageUsers) return null;
    const userId = String(form.get("userId") ?? "");
    const target = users.find((u) => u.id === userId);
    if (!target) return null;
    if (target.role === "superadmin" && currentUser?.role !== "superadmin") return null;
    if (target.id === currentUser?.id) return null;
    await savePortalUsers(users.filter((u) => u.id !== userId));
    return null;
  }

  // Sample delete: move before auth guard and use raw SQL to avoid Prisma
  // column introspection failing when migration 000002 hasn't applied yet
  if (intent === "delete_sample") {
    const sampleId = Number(form.get("sampleId"));
    if (!sampleId) return null;
    await prisma.$executeRaw`DELETE FROM "SampleIteration" WHERE "sampleId" = ${sampleId}`;
    await prisma.$executeRaw`DELETE FROM "Sample" WHERE "id" = ${sampleId}`;
    return null;
  }

  if (!currentUser) return null;

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
    await logActivity(currentUser?.name ?? "Unknown", "Created", "Packing List", {
      entityId: String(packingList.id),
      entityName: invoiceNumber || packingList.title,
    });
    return new Response(null, {
      status: 303,
      headers: { Location: `/portal?page=packing&packingId=${packingList.id}` },
    });
  }

  if (intent === "import_supplier_packing_csv") {
    const csvFile = form.get("packingCsv");
    if (!csvFile || typeof csvFile === "string" || typeof csvFile.text !== "function") {
      return null;
    }
    const parsed = parseSupplierPackingCsv(await csvFile.text());
    const productCache = new Map<string, ShopifySearchProduct | null>();
    const importedLines = [];
    for (const [index, line] of parsed.lines.entries()) {
      const product = await findShopifyProductForSupplierLine(line.styleName, line.colorName, productCache);
      importedLines.push({
        boxNumber: line.boxNumber,
        productId: product?.id ?? null,
        productTitle: product?.title ?? line.productTitle,
        productImageUrl: product?.imageUrl ?? null,
        sku: product?.skus.join("\n") ?? line.supplierCode,
        isCustom: !product,
        qtys: line.qtys,
        sortOrder: index + 1,
      });
    }
    const packingList = await prisma.packingList.create({
      data: {
        title: parsed.title,
        invoiceNumber: parsed.invoiceNumber,
        shipmentDate: parsed.shipmentDate,
        expectedLeaveFactoryDate: parsed.shipmentDate,
        status: "still_packing",
        lines: {
          create: importedLines.length ? importedLines : Array.from({ length: DEFAULT_PACKING_ROWS }, (_, index) => ({
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
    if (field === "status") {
      const nextStatus = value || "still_packing";
      // Status is a one-way gate: once it leaves "still_packing" it can't go
      // back (which would re-open quantity editing) — admins exempt.
      if (nextStatus === "still_packing" && packingId && !currentUser?.admin) {
        const current = await prisma.packingList.findUnique({ where: { id: packingId }, select: { status: true } });
        if ((current?.status ?? "still_packing") !== "still_packing") return null;
      }
      data.status = nextStatus;
    }
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
    // Handle shippingMethod via raw SQL so the feature works even if the
    // Prisma client / DB migration haven't caught up on this environment.
    // `ADD COLUMN IF NOT EXISTS` is idempotent.
    if (packingId && field === "shippingMethod") {
      const normalised = value.toLowerCase();
      const finalValue: string | null = (normalised === "sea" || normalised === "air") ? normalised : null;
      try {
        await prisma.$executeRawUnsafe(`ALTER TABLE "PackingList" ADD COLUMN IF NOT EXISTS "shippingMethod" TEXT`);
      } catch (e) {
        console.warn("[shippingMethod] ALTER TABLE failed (column may already exist):", e);
      }
      try {
        const rows = await prisma.$executeRawUnsafe(
          `UPDATE "PackingList" SET "shippingMethod" = $1 WHERE id = $2`,
          finalValue,
          packingId,
        );
        console.log(`[shippingMethod] saved packingId=${packingId} value=${finalValue} rowsAffected=${rows}`);
      } catch (e) {
        console.error("[shippingMethod] UPDATE failed:", e);
      }
      const existing = await prisma.packingList.findUnique({ where: { id: packingId }, select: { invoiceNumber: true, title: true } });
      if (existing) {
        await logActivity(currentUser?.name ?? "Unknown", "Updated", "Packing List", {
          entityId: String(packingId),
          entityName: existing.invoiceNumber || existing.title || `#${packingId}`,
          field: "Shipping method",
          toValue: finalValue ? finalValue.charAt(0).toUpperCase() + finalValue.slice(1) : "—",
        });
      }
      return null;
    }
    if (packingId && Object.keys(data).length) {
      const existing = await prisma.packingList.findUnique({ where: { id: packingId }, select: { invoiceNumber: true, title: true } });
      await prisma.packingList.update({ where: { id: packingId }, data });
      const logField = field === "invoiceNumber" ? "Invoice number"
        : field === "status" ? "Status"
        : field === "expectedLeaveFactoryDate" ? "Estimated arrival"
        : field;
      const logValue = field === "status"
        ? (PACKING_STATUS_OPTIONS.find((o) => o.value === value)?.label ?? value)
        : field === "expectedLeaveFactoryDate"
          ? formatPortalDate(parsePortalDate(value))
          : value;
      if (logField && logValue) {
        await logActivity(currentUser?.name ?? "Unknown", "Updated", "Packing List", {
          entityId: String(packingId),
          entityName: existing?.invoiceNumber || existing?.title || `#${packingId}`,
          field: logField,
          toValue: logValue,
        });
      }
    }
    return null;
  }

  if (intent === "set_packing_list_hidden") {
    const packingId = Number(form.get("packingId"));
    const hidden = String(form.get("hidden") ?? "") === "true";
    if (packingId) {
      await prisma.packingList.update({
        where: { id: packingId },
        data: { hiddenAt: hidden ? new Date() : null },
      });
    }
    return null;
  }

  if (intent === "delete_packing_list") {
    const packingId = Number(form.get("packingId"));
    if (packingId) {
      const existing = await prisma.packingList.findUnique({ where: { id: packingId }, select: { invoiceNumber: true, title: true } });
      await prisma.packingList.deleteMany({ where: { id: packingId } });
      await logActivity(currentUser?.name ?? "Unknown", "Deleted", "Packing List", {
        entityId: String(packingId),
        entityName: existing?.invoiceNumber || existing?.title || `#${packingId}`,
      });
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

  if (intent === "move_packing_line") {
    const lineId = Number(form.get("lineId"));
    const direction = String(form.get("direction") ?? "");
    const line = await prisma.packingListLine.findUnique({
      where: { id: lineId },
      select: { id: true, packingListId: true, sortOrder: true },
    });
    if (!line || !["up", "down"].includes(direction)) return null;
    const sibling = await prisma.packingListLine.findFirst({
      where: {
        packingListId: line.packingListId,
        sortOrder: direction === "up" ? { lt: line.sortOrder } : { gt: line.sortOrder },
      },
      orderBy: { sortOrder: direction === "up" ? "desc" : "asc" },
      select: { id: true, sortOrder: true },
    });
    if (!sibling) return null;
    await prisma.$transaction([
      prisma.packingListLine.update({ where: { id: line.id }, data: { sortOrder: sibling.sortOrder } }),
      prisma.packingListLine.update({ where: { id: sibling.id }, data: { sortOrder: line.sortOrder } }),
    ]);
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
    // Existing packed quantities (and any manual "loaded" marks) must survive
    // linking — the row already represents real packed goods. shopifyLoadedQtys
    // is cleared because the row is now pointing at a different product.
    // We do seed any of the product's sizes that aren't yet present with 0,
    // so the size column appears in the table even before the user fills it
    // in — needed for non-baseline sizes like "Free Size" or "XL-2XL" that
    // are derived from the line data.
    const existing = await prisma.packingListLine.findUnique({
      where: { id: lineId },
      select: { sku: true, productImageUrl: true, qtys: true },
    });
    const mergedQtys = { ...normalizeQtys(existing?.qtys) };
    for (const size of product.sizes ?? []) {
      if (size && !(size in mergedQtys)) mergedQtys[size] = 0;
    }
    await prisma.packingListLine.update({
      where: { id: lineId },
      data: {
        productId: product.id,
        productTitle: product.title || "Untitled product",
        productImageUrl: existing?.productImageUrl || product.imageUrl || null,
        sku: existing?.sku || product.skus?.filter(Boolean).join("\n") || null,
        isCustom: false,
        qtys: mergedQtys,
        shopifyLoadedQtys: {},
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
    if (field === "productImageUrl") data.productImageUrl = value || null;
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
    const line = await prisma.packingListLine.findUnique({
      where: { id: lineId },
      select: { qtys: true, shopifyLoadedQtys: true, manuallyLoadedQtys: true, productTitle: true, packingList: { select: { status: true } } },
    });
    if (!line || !size) return null;
    // Quantities are frozen once the list moves past "still_packing" — admins exempt.
    if ((line.packingList?.status ?? "still_packing") !== "still_packing" && !currentUser?.admin) return null;
    const qtys = normalizeQtys(line.qtys);
    const shopifyLoadedQtys = normalizeQtys(line.shopifyLoadedQtys);
    const manuallyLoadedQtys = normalizeQtys(line.manuallyLoadedQtys);
    const previousQty = qtys[size] ?? 0;
    qtys[size] = value;
    // Changing the qty invalidates any prior "loaded" marker for that size.
    delete shopifyLoadedQtys[size];
    delete manuallyLoadedQtys[size];
    await prisma.packingListLine.update({ where: { id: lineId }, data: { qtys, shopifyLoadedQtys, manuallyLoadedQtys } });
    if (value !== previousQty) {
      await logActivity(currentUser?.name ?? "Unknown", "Updated", "Packing List Line", {
        entityId: String(lineId),
        entityName: line.productTitle || `Line #${lineId}`,
        field: `Qty (${size})`,
        toValue: String(value),
      });
    }
    return null;
  }

  if (intent === "load_packing_inventory" || intent === "load_packing_inventory_for_product") {
    if (!canLoadPackingInventory) return null;
    const packingId = Number(form.get("packingId"));
    const onlyProductId = intent === "load_packing_inventory_for_product"
      ? String(form.get("productId") ?? "")
      : null;
    if (intent === "load_packing_inventory_for_product" && !onlyProductId) return null;
    const skipWords = String(form.get("skipWords") ?? "")
      .split(",")
      .map((word) => word.trim().toLowerCase())
      .filter(Boolean);
    const packingList = await prisma.packingList.findUnique({
      where: { id: packingId },
      include: { lines: { orderBy: [{ sortOrder: "asc" }, { id: "asc" }] } },
    });
    const session = await prisma.session.findFirst({
      where: { accessToken: { not: "" } },
      orderBy: { isOnline: "asc" },
    });
    if (!packingList || !session?.shop || !session.accessToken) return null;

    const variantCache = new Map<string, ShopifyInventoryVariantInfo[]>();
    const getVariants = async (productId: string) => {
      if (!variantCache.has(productId)) {
        variantCache.set(productId, await getShopifyInventoryVariants(session.shop, productId));
      }
      return variantCache.get(productId) ?? [];
    };

    for (const line of packingList.lines) {
      const title = line.productTitle.toLowerCase();
      if (!line.productId || line.isCustom || skipWords.some((word) => title.includes(word))) continue;
      if (onlyProductId && line.productId !== onlyProductId) continue;

      const qtys = normalizeQtys(line.qtys);
      const loadedQtys = normalizeQtys(line.shopifyLoadedQtys);
      const manualQtys = normalizeQtys(line.manuallyLoadedQtys);
      const variants = await getVariants(line.productId);
      const changes: ShopifyInventoryChange[] = [];
      const nextLoadedQtys: Record<string, number> = { ...loadedQtys };

      for (const [size, qty] of Object.entries(qtys)) {
        // Skip if already pushed to Shopify OR manually marked as loaded.
        if (qty <= 0 || loadedQtys[size] === qty || manualQtys[size] === qty) continue;
        const variant = matchingVariantForSize(variants, size);
        if (!variant?.inventoryItemId) continue;
        changes.push({ size, qty, inventoryItemId: variant.inventoryItemId });
      }

      if (!changes.length) continue;
      const loadedSizes = await addShopifyInventory(session.shop, session.accessToken, changes);
      if (!loadedSizes.length) continue;
      for (const size of loadedSizes) {
        nextLoadedQtys[size] = qtys[size] ?? 0;
      }
      await prisma.packingListLine.update({
        where: { id: line.id },
        data: { shopifyLoadedQtys: nextLoadedQtys },
      });
    }

    // Master button is one-time-use: record when it was pressed so the UI
    // hides it afterward. Per-product loads don't set this.
    if (intent === "load_packing_inventory" && !packingList.masterInventoryLoadedAt) {
      await prisma.packingList.update({
        where: { id: packingId },
        data: { masterInventoryLoadedAt: new Date() },
      });
    }

    return null;
  }

  if (intent === "toggle_packing_qty_manual_loaded") {
    if (!currentUser?.admin) return null;
    const lineId = Number(form.get("lineId"));
    const size = String(form.get("size") ?? "");
    const line = await prisma.packingListLine.findUnique({
      where: { id: lineId },
      select: { qtys: true, manuallyLoadedQtys: true },
    });
    if (!line || !size) return null;
    const qtys = normalizeQtys(line.qtys);
    const manualQtys = normalizeQtys(line.manuallyLoadedQtys);
    const qty = qtys[size] ?? 0;
    if (manualQtys[size] === qty && qty > 0) {
      delete manualQtys[size];
    } else if (qty > 0) {
      manualQtys[size] = qty;
    }
    await prisma.packingListLine.update({
      where: { id: lineId },
      data: { manuallyLoadedQtys: manualQtys },
    });
    return null;
  }

  if (intent === "toggle_packing_qty_manual_loaded_for_product") {
    if (!currentUser?.admin) return null;
    const packingId = Number(form.get("packingId"));
    const productId = String(form.get("productId") ?? "");
    const size = String(form.get("size") ?? "");
    if (!packingId || !productId || !size) return null;
    const lines = await prisma.packingListLine.findMany({
      where: { packingListId: packingId, productId },
      select: { id: true, qtys: true, manuallyLoadedQtys: true },
    });
    // If every contributing line is already fully marked for this size, the
    // toggle removes the mark from all of them. Otherwise, mark every line.
    const allMarked = lines.every((line) => {
      const qty = normalizeQtys(line.qtys)[size] ?? 0;
      if (qty <= 0) return true;
      return normalizeQtys(line.manuallyLoadedQtys)[size] === qty;
    });
    for (const line of lines) {
      const qtys = normalizeQtys(line.qtys);
      const qty = qtys[size] ?? 0;
      if (qty <= 0) continue;
      const manual = normalizeQtys(line.manuallyLoadedQtys);
      if (allMarked) delete manual[size]; else manual[size] = qty;
      await prisma.packingListLine.update({
        where: { id: line.id },
        data: { manuallyLoadedQtys: manual },
      });
    }
    return null;
  }

  if (intent === "bulk_set_packing_qty_manual_loaded") {
    if (!currentUser?.admin) return null;
    const packingId = Number(form.get("packingId"));
    const action = String(form.get("action") ?? "");
    if (!packingId || (action !== "mark" && action !== "unmark")) return null;
    let cells: { lineId: number; size: string }[] = [];
    try {
      const raw = JSON.parse(String(form.get("cells") ?? "[]"));
      if (Array.isArray(raw)) {
        cells = raw
          .map((entry) => ({
            lineId: Number(entry?.lineId ?? 0),
            size: String(entry?.size ?? ""),
          }))
          .filter((entry) => Number.isFinite(entry.lineId) && entry.lineId > 0 && entry.size);
      }
    } catch {
      return null;
    }
    if (!cells.length) return null;
    const lineIds = Array.from(new Set(cells.map((cell) => cell.lineId)));
    const lines = await prisma.packingListLine.findMany({
      where: { packingListId: packingId, id: { in: lineIds } },
      select: { id: true, qtys: true, manuallyLoadedQtys: true },
    });
    const sizesByLine = new Map<number, Set<string>>();
    for (const cell of cells) {
      let set = sizesByLine.get(cell.lineId);
      if (!set) { set = new Set(); sizesByLine.set(cell.lineId, set); }
      set.add(cell.size);
    }
    for (const line of lines) {
      const sizes = sizesByLine.get(line.id);
      if (!sizes || !sizes.size) continue;
      const qtys = normalizeQtys(line.qtys);
      const manual = normalizeQtys(line.manuallyLoadedQtys);
      let changed = false;
      for (const size of sizes) {
        const qty = qtys[size] ?? 0;
        if (qty <= 0) continue;
        if (action === "mark") {
          if (manual[size] !== qty) { manual[size] = qty; changed = true; }
        } else {
          if (manual[size] !== undefined) { delete manual[size]; changed = true; }
        }
      }
      if (changed) {
        await prisma.packingListLine.update({
          where: { id: line.id },
          data: { manuallyLoadedQtys: manual },
        });
      }
    }
    return null;
  }

  if (intent === "relink_packing_lines_to_shopify") {
    const packingId = Number(form.get("packingId"));
    if (!packingId) return jsonResponse({ relinked: 0, scanned: 0, error: "missing_packing_id" });
    const lines = await prisma.packingListLine.findMany({
      where: {
        packingListId: packingId,
        OR: [{ productId: null }, { productId: "" }],
      },
      select: { id: true, productTitle: true, productImageUrl: true, isCustom: true },
    });
    let relinked = 0;
    // Cache by normalized title so duplicate titles in the same list only hit
    // Shopify once.
    const cache = new Map<string, { id: string; imageUrl: string | null } | null>();
    for (const line of lines) {
      const title = (line.productTitle ?? "").trim();
      if (!title) continue;
      const cacheKey = title.toLowerCase();
      let matched: { id: string; imageUrl: string | null } | null | undefined = cache.get(cacheKey);
      if (matched === undefined) {
        const results = await searchShopifyProducts(title);
        const exact = results.filter((p) => p.title.trim().toLowerCase() === cacheKey);
        matched = exact.length === 1
          ? { id: exact[0].id, imageUrl: exact[0].imageUrl ?? null }
          : null;
        cache.set(cacheKey, matched);
      }
      if (!matched) continue;
      await prisma.packingListLine.update({
        where: { id: line.id },
        data: {
          productId: matched.id,
          ...(line.isCustom ? { isCustom: false } : {}),
          ...(line.productImageUrl ? {} : { productImageUrl: matched.imageUrl ?? undefined }),
        },
      });
      relinked += 1;
    }
    return jsonResponse({ relinked, scanned: lines.length });
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
    await prisma.portalMessage.deleteMany({ where: { orderId } });
    await prisma.supplierOrder.delete({ where: { id: orderId } });
    return null;
  }

  if (intent === "duplicate_order") {
    const order = await prisma.supplierOrder.findUnique({
      where: { id: orderId },
      include: { lines: { orderBy: { id: "asc" } } },
    });
    if (!order) return null;
    const shopifyVariants = await getShopifyProductVariants(order.shop, order.productId);
    const linesByVariantId = new Map(order.lines.map((line) => [line.variantId, line]));
    const linesToCreate = shopifyVariants.length
      ? shopifyVariants.map((variant) => {
          const sourceLine = linesByVariantId.get(variant.id);
          return {
            variantId: variant.id,
            variantTitle: variant.title,
            sku: variant.sku ?? sourceLine?.sku ?? null,
            qtyOrdered: 0,
            costPrice: sourceLine?.costPrice ?? null,
          };
        })
      : order.lines.map((line) => ({
          variantId: line.variantId,
          variantTitle: line.variantTitle,
          sku: line.sku,
          qtyOrdered: 0,
          costPrice: line.costPrice,
        }));

    await prisma.supplierOrder.create({
      data: {
        shop: order.shop,
        poNumber: order.poNumber,
        supplier: order.supplier,
        productId: order.productId,
        productTitle: order.productTitle,
        productType: order.productType,
        status: "open",
        supplierStatus: order.supplierStatus,
        priority: order.priority,
        productImageUrl: order.productImageUrl,
        eta: order.eta,
        totalQty: 0,
        lines: { create: linesToCreate },
      },
    });
    return null;
  }

  if (intent === "create_restock_order_from_portal") {
    let product: ShopifySearchProduct;
    let qtys: Record<string, number>;
    try {
      product = JSON.parse(String(form.get("product") ?? "{}")) as ShopifySearchProduct;
      qtys = normalizeQtys(JSON.parse(String(form.get("qtys") ?? "{}")));
    } catch {
      return null;
    }

    if (!product?.id || !product.title) return null;
    const fallbackSession = product.shop
      ? null
      : await prisma.session.findFirst({ where: { accessToken: { not: "" } }, orderBy: { isOnline: "asc" } });
    const shop = product.shop ?? fallbackSession?.shop ?? "";
    if (!shop) return null;

    const variants = product.variants?.length
      ? product.variants
      : shop ? await getShopifyProductVariants(shop, product.id) : [];
    if (!variants.length) return null;

    const totalQty = variants.reduce((sum, variant) => sum + (qtys[variant.title] ?? 0), 0);
    if (totalQty <= 0) return null;

    const etaRaw = String(form.get("eta") ?? "");
    const eta = etaRaw ? parsePortalDate(etaRaw) : null;
    if (etaRaw && !eta) return null;

    const notes = String(form.get("notes") ?? "").trim();
    const createdOrder = await prisma.supplierOrder.create({
      data: {
        shop,
        poNumber: null,
        supplier: "Portal",
        productId: product.id,
        productTitle: product.title,
        productType: normalizeProductGroup(String(form.get("productType") ?? "")) || null,
        status: "open",
        supplierStatus: String(form.get("supplierStatus") ?? "") || "on_order",
        priority: String(form.get("priority") ?? "") || null,
        productImageUrl: product.imageUrl,
        eta,
        notes: notes || null,
        totalQty,
        lines: {
          create: variants.map((variant) => ({
            variantId: variant.id,
            variantTitle: variant.title,
            sku: variant.sku,
            qtyOrdered: qtys[variant.title] ?? 0,
          })),
        },
      },
    });

    if (notes) {
      await syncOrderNoteMessages({
        orderId: createdOrder.id,
        field: "notes",
        text: notes,
        fromName: currentUser?.name ?? null,
      });
    }
    return null;
  }

  if (intent === "mark_message_read") {
    const messageId = Number(form.get("messageId"));
    if (messageId && currentUser) {
      await prisma.portalMessage.updateMany({
        where: { id: messageId, userId: currentUser.id },
        data: { readAt: new Date() },
      });
    }
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

  if (intent === "update_table_header") {
    const key = String(form.get("key") ?? "");
    const value = String(form.get("value") ?? "").trim();
    if (!key) return null;
    const setting = await prisma.portalSetting.findUnique({
      where: { key: TABLE_HEADER_LABELS_KEY },
      select: { value: true },
    });
    const labels = normalizeTableHeaderLabels(setting?.value);
    if (value) labels[key] = value;
    else delete labels[key];
    await prisma.portalSetting.upsert({
      where: { key: TABLE_HEADER_LABELS_KEY },
      create: { key: TABLE_HEADER_LABELS_KEY, value: labels },
      update: { value: labels },
    });
    return null;
  }

  if (intent === "update_table_custom_cell") {
    const key = String(form.get("key") ?? "");
    const value = String(form.get("value") ?? "");
    if (!key) return null;
    const setting = await prisma.portalSetting.findUnique({ where: { key: TABLE_CUSTOM_CELLS_KEY }, select: { value: true } });
    const cells = normalizeTableCustomCells(setting?.value);
    cells[key] = value;
    await prisma.portalSetting.upsert({
      where: { key: TABLE_CUSTOM_CELLS_KEY },
      create: { key: TABLE_CUSTOM_CELLS_KEY, value: cells },
      update: { value: cells },
    });
    return null;
  }

  if (intent === "add_table_column" || intent === "remove_table_column") {
    const table = String(form.get("table") ?? "");
    const gid = String(form.get("gid") ?? "");
    const columnId = String(form.get("columnId") ?? "");
    const label = String(form.get("label") ?? "New Column").trim() || "New Column";
    const setting = await prisma.portalSetting.findUnique({ where: { key: TABLE_CUSTOM_COLUMNS_KEY }, select: { value: true } });
    const columns = normalizeTableCustomColumns(setting?.value);
    if (intent === "add_table_column") {
      const column = { id: columnId.startsWith("custom_") ? columnId : `custom_${Date.now()}`, label };
      if (table === "restock") columns.restock = [...columns.restock, column];
      if (table === "packing") columns.packing = [...columns.packing, column];
      if (table === "fabric" && gid) columns.fabric[gid] = [...(columns.fabric[gid] ?? []), column];
    } else if (columnId) {
      if (table === "restock") columns.restock = columns.restock.filter((column) => column.id !== columnId);
      if (table === "packing") columns.packing = columns.packing.filter((column) => column.id !== columnId);
      if (table === "fabric" && gid) columns.fabric[gid] = (columns.fabric[gid] ?? []).filter((column) => column.id !== columnId);
    }
    await prisma.portalSetting.upsert({
      where: { key: TABLE_CUSTOM_COLUMNS_KEY },
      create: { key: TABLE_CUSTOM_COLUMNS_KEY, value: columns },
      update: { value: columns },
    });
    return null;
  }

  if (intent === "update_row_height") {
    const key = String(form.get("key") ?? "");
    const height = Math.min(420, Math.max(34, Number(form.get("height")) || 0));
    if (!key || !height) return null;
    const setting = await prisma.portalSetting.findUnique({ where: { key: TABLE_ROW_HEIGHTS_KEY }, select: { value: true } });
    const heights = normalizeTableRowHeights(setting?.value);
    heights[key] = height;
    await prisma.portalSetting.upsert({
      where: { key: TABLE_ROW_HEIGHTS_KEY },
      create: { key: TABLE_ROW_HEIGHTS_KEY, value: heights },
      update: { value: heights },
    });
    return null;
  }

  if (intent === "create_fabric_sheet") {
    const name = String(form.get("name") ?? "").trim();
    const gid = String(form.get("gid") ?? "").trim();
    if (!name) return null;
    const sheets = await loadManualFabricSheetsForAction();
    sheets.push({
      gid: gid.startsWith("custom_") ? gid : `custom_${Date.now()}`,
      name,
      kind: "stock",
      headers: DEFAULT_FABRIC_HEADERS,
      rows: Array.from({ length: 3 }, () => Array.from({ length: DEFAULT_FABRIC_HEADERS.length }, () => "")),
      rowCount: 3,
      totalQuantity: 3,
    });
    await saveManualFabricSheets(sheets);
    return null;
  }

  if (intent === "delete_fabric_sheet") {
    const gid = String(form.get("gid") ?? "");
    if (!gid) return null;
    await saveManualFabricSheets((await loadManualFabricSheetsForAction()).filter((sheet) => sheet.gid !== gid));
    return null;
  }

  if (intent === "restore_fabric_sheet") {
    const gid = String(form.get("gid") ?? "");
    const sheetJson = String(form.get("sheet") ?? "");
    if (!gid) return null;
    const sheets = await loadManualFabricSheetsForAction();
    try {
      const restoredSheet = normalizeFabricCustomSheets([JSON.parse(sheetJson)])[0];
      if (restoredSheet && !sheets.some((sheet) => sheet.gid === restoredSheet.gid)) {
        sheets.push(toManualFabricSheet(restoredSheet));
      }
    } catch { /* ignore malformed undo payload */ }
    await saveManualFabricSheets(sheets);
    return null;
  }

  if (intent === "update_restock_settings") {
    let settings: RestockSettings;
    try {
      settings = normalizeRestockSettings(JSON.parse(String(form.get("value") ?? "{}")));
    } catch {
      return null;
    }

    await prisma.portalSetting.upsert({
      where: { key: RESTOCK_SETTINGS_KEY },
      create: { key: RESTOCK_SETTINGS_KEY, value: settings },
      update: { value: settings },
    });
    return null;
  }

  if (intent === "update_universal_settings") {
    let settings: UniversalSettings;
    try {
      settings = normalizeUniversalSettings(JSON.parse(String(form.get("value") ?? "{}")));
    } catch {
      return null;
    }

    await prisma.portalSetting.upsert({
      where: { key: UNIVERSAL_SETTINGS_KEY },
      create: { key: UNIVERSAL_SETTINGS_KEY, value: settings },
      update: { value: settings },
    });
    return null;
  }

  if (intent === "update_fabric_settings") {
    let settings: FabricSettings;
    try {
      settings = normalizeFabricSettings(JSON.parse(String(form.get("value") ?? "{}")));
    } catch {
      return null;
    }

    await prisma.portalSetting.upsert({
      where: { key: FABRIC_SETTINGS_KEY },
      create: { key: FABRIC_SETTINGS_KEY, value: settings },
      update: { value: settings },
    });
    return null;
  }

  if (intent === "update_nav_order") {
    try {
      const order = JSON.parse(String(form.get("value") ?? "[]"));
      if (Array.isArray(order)) {
        await prisma.portalSetting.upsert({
          where: { key: PORTAL_NAV_ORDER_KEY },
          create: { key: PORTAL_NAV_ORDER_KEY, value: order },
          update: { value: order },
        });
      }
    } catch { /* ignore */ }
    return null;
  }

  if (intent === "add_product_category") {
    const name = String(form.get("name") ?? "").trim();
    if (!name) return null;
    const productInfo = await loadProductInfoForAction();
    productInfo.categories.push({
      id: `category_${Date.now()}`,
      name,
      styles: [],
    });
    await saveProductInfo(productInfo);
    return null;
  }

  if (intent === "delete_product_category") {
    const categoryId = String(form.get("categoryId") ?? "");
    if (!categoryId) return null;
    const productInfo = await loadProductInfoForAction();
    productInfo.categories = productInfo.categories.filter((category) => category.id !== categoryId);
    await saveProductInfo(productInfo);
    return null;
  }

  if (intent === "add_product_style") {
    const categoryId = String(form.get("categoryId") ?? "");
    const name = String(form.get("name") ?? "").trim();
    if (!categoryId || !name) return null;
    const productInfo = await loadProductInfoForAction();
    const category = productInfo.categories.find((item) => item.id === categoryId);
    if (!category) return null;
    category.styles.push({
      id: `style_${Date.now()}`,
      name,
      hidden: false,
      imageUrl: "",
    });
    await saveProductInfo(productInfo);
    return null;
  }

  if (intent === "update_product_style_image") {
    const categoryId = String(form.get("categoryId") ?? "");
    const styleId = String(form.get("styleId") ?? "");
    const imageUrl = String(form.get("imageUrl") ?? "").trim();
    if (!categoryId || !styleId) return null;
    const productInfo = await loadProductInfoForAction();
    const category = productInfo.categories.find((item) => item.id === categoryId);
    const style = category?.styles.find((item) => item.id === styleId);
    if (!style) return null;
    style.imageUrl = imageUrl;
    await saveProductInfo(productInfo);
    return null;
  }

  if (intent === "update_product_style_details") {
    const categoryId = String(form.get("categoryId") ?? "");
    const styleId = String(form.get("styleId") ?? "");
    if (!categoryId || !styleId) return null;
    const productInfo = await loadProductInfoForAction();
    const category = productInfo.categories.find((item) => item.id === categoryId);
    const style = category?.styles.find((item) => item.id === styleId);
    if (!style) return null;
    const readNumber = (key: string) => {
      const raw = String(form.get(key) ?? "").trim();
      if (!raw) return undefined;
      const value = Number(raw);
      return Number.isFinite(value) ? value : undefined;
    };
    style.averageMeters = readNumber("averageMeters");
    style.averageTrimMeters = readNumber("averageTrimMeters");
    style.stitchingCost = readNumber("stitchingCost");
    style.zipButtonsCost = readNumber("zipButtonsCost");
    style.sheetCount = readNumber("sheetCount");
    style.zipButtonType = String(form.get("zipButtonType") ?? "").trim();
    style.costingNotes = String(form.get("costingNotes") ?? "").trim();
    await saveProductInfo(productInfo);
    return null;
  }

  if (intent === "update_product_info_grid") {
    const gridColumns = Number(form.get("gridColumns"));
    if (gridColumns !== 3 && gridColumns !== 4) return null;
    const productInfo = await loadProductInfoForAction();
    productInfo.gridColumns = gridColumns;
    await saveProductInfo(productInfo);
    return null;
  }

  if (intent === "reorder_product_styles") {
    const categoryId = String(form.get("categoryId") ?? "");
    if (!categoryId) return null;
    let styleIds: string[] = [];
    try {
      const parsed = JSON.parse(String(form.get("styleIds") ?? "[]"));
      if (Array.isArray(parsed)) styleIds = parsed.map((item) => String(item)).filter(Boolean);
    } catch {
      styleIds = [];
    }
    if (!styleIds.length) return null;
    const productInfo = await loadProductInfoForAction();
    const category = productInfo.categories.find((item) => item.id === categoryId);
    if (!category) return null;
    const styleById = new Map(category.styles.map((style) => [style.id, style]));
    const ordered = styleIds.map((id) => styleById.get(id)).filter(Boolean) as ProductInfoStyle[];
    const orderedIds = new Set(ordered.map((style) => style.id));
    category.styles = [...ordered, ...category.styles.filter((style) => !orderedIds.has(style.id))];
    await saveProductInfo(productInfo);
    return null;
  }

  if (intent === "hide_product_style") {
    const categoryId = String(form.get("categoryId") ?? "");
    const styleId = String(form.get("styleId") ?? "");
    if (!categoryId || !styleId) return null;
    const productInfo = await loadProductInfoForAction();
    const category = productInfo.categories.find((item) => item.id === categoryId);
    const style = category?.styles.find((item) => item.id === styleId);
    if (!style) return null;
    style.hidden = true;
    await saveProductInfo(productInfo);
    return null;
  }

  if (intent === "unhide_product_style") {
    const categoryId = String(form.get("categoryId") ?? "");
    const styleId = String(form.get("styleId") ?? "");
    if (!categoryId || !styleId) return null;
    const productInfo = await loadProductInfoForAction();
    const category = productInfo.categories.find((item) => item.id === categoryId);
    const style = category?.styles.find((item) => item.id === styleId);
    if (!style) return null;
    style.hidden = false;
    await saveProductInfo(productInfo);
    return null;
  }

  if (intent === "delete_product_style") {
    const categoryId = String(form.get("categoryId") ?? "");
    const styleId = String(form.get("styleId") ?? "");
    if (!categoryId || !styleId) return null;
    const productInfo = await loadProductInfoForAction();
    const category = productInfo.categories.find((item) => item.id === categoryId);
    if (!category) return null;
    category.styles = category.styles.filter((style) => style.id !== styleId);
    await saveProductInfo(productInfo);
    return null;
  }

  if (intent === "add_sample") {
    const name = String(form.get("name") ?? "").trim();
    if (!name) return null;
    await prisma.sample.create({ data: { name } });
    return null;
  }
  if (intent === "get_sample_full") {
    const sampleId = Number(form.get("sampleId"));
    if (!sampleId) return jsonResponse({ sample: null });
    const sample = await prisma.sample.findUnique({
      where: { id: sampleId },
      include: { iterations: { orderBy: { version: "asc" } } },
    }).catch(() => null);
    return jsonResponse({ sample });
  }
  if (intent === "rename_sample") {
    const sampleId = Number(form.get("sampleId"));
    const name = String(form.get("name") ?? "").trim();
    if (!sampleId || !name) return null;
    await prisma.sample.update({ where: { id: sampleId }, data: { name } });
    return null;
  }
  if (intent === "reorder_samples") {
    const ids = JSON.parse(String(form.get("sampleIds") ?? "[]")) as number[];
    await Promise.all(ids.map((id, index) => prisma.sample.update({ where: { id }, data: { sortOrder: index } })));
    return null;
  }
  if (intent === "add_sample_iteration") {
    const sampleId = Number(form.get("sampleId"));
    if (!sampleId) return null;
    const existing = await prisma.sampleIteration.findMany({ where: { sampleId }, select: { version: true } });
    const nextVersion = existing.length > 0 ? Math.max(...existing.map((i) => i.version)) + 1 : 1;
    await prisma.sampleIteration.create({ data: { sampleId, version: nextVersion, status: "under_consideration" } });
    return null;
  }
  if (intent === "update_sample_iteration") {
    const iterationId = Number(form.get("iterationId"));
    if (!iterationId) return null;
    const iteration = await prisma.sampleIteration.findUnique({ where: { id: iterationId } });
    if (!iteration) return null;
    const updates: Record<string, unknown> = { updatedAt: new Date() };
    if (form.has("notes")) {
      const notesText = String(form.get("notes") ?? "");
      updates.notes = notesText;
      const sample = await prisma.sample.findUnique({ where: { id: iteration.sampleId }, select: { name: true } });
      await syncSampleIterationMessages({ iterationId, sampleName: sample?.name ?? "Sample", text: notesText });
    }
    if (form.has("name")) updates.name = String(form.get("name") ?? "") || null;
    if (form.has("fabricType")) updates.fabricType = String(form.get("fabricType") ?? "") || null;
    if (form.has("sampleSize")) updates.sampleSize = String(form.get("sampleSize") ?? "") || null;
    if (form.has("buttonType")) updates.buttonType = String(form.get("buttonType") ?? "") || null;
    if (form.has("factoryCost")) updates.factoryCost = String(form.get("factoryCost") ?? "") || null;
    if (form.has("status")) updates.status = String(form.get("status") ?? "under_consideration");
    if (form.has("addImage")) {
      const currentImages = Array.isArray(iteration.images) ? iteration.images as string[] : [];
      updates.images = [...currentImages, String(form.get("addImage"))];
    }
    if (form.has("addImages")) {
      try {
        const newImgs = JSON.parse(String(form.get("addImages"))) as string[];
        if (Array.isArray(newImgs) && newImgs.length > 0) {
          const base = Array.isArray(updates.images) ? (updates.images as string[]) : (Array.isArray(iteration.images) ? iteration.images as string[] : []);
          updates.images = [...base, ...newImgs];
        }
      } catch { /* ignore */ }
    }
    if (form.has("imagesReplace")) {
      try {
        const replaced = JSON.parse(String(form.get("imagesReplace"))) as string[];
        if (Array.isArray(replaced)) updates.images = replaced;
      } catch { /* ignore */ }
    }
    if (form.has("thumbnailImage")) {
      const currentImages = Array.isArray(iteration.images) ? iteration.images as string[] : [];
      updates.images = [String(form.get("thumbnailImage")), ...currentImages.slice(1)];
    }
    if (form.has("removeImageIndex")) {
      const idx = Number(form.get("removeImageIndex"));
      const currentImages = Array.isArray(iteration.images) ? iteration.images as string[] : [];
      updates.images = currentImages.filter((_, i) => i !== idx);
    }
    // Optional client-supplied small thumbnail (~5–15 KB) for the card.
    if (form.has("thumbnail")) {
      const thumb = String(form.get("thumbnail") ?? "");
      updates.thumbnail = thumb || null;
    } else if ("images" in updates) {
      // Images changed and client didn't supply a new thumbnail — clear it so
      // the next viewing round generates a fresh one from the new first image.
      updates.thumbnail = null;
    }
    await prisma.sampleIteration.update({ where: { id: iterationId }, data: updates });
    return null;
  }
  if (intent === "delete_sample_iteration") {
    const iterationId = Number(form.get("iterationId"));
    if (!iterationId) return null;
    await prisma.sampleIteration.delete({ where: { id: iterationId } });
    return null;
  }

  // ─── Vision Board V2 intents ──────────────────────────────────────────────
  if (intent === "vb_add_board") {
    if (currentUser?.role !== "superadmin") return null;
    const name = String(form.get("name") ?? "").trim() || "New Board";
    const existing = await prisma.visionBoardV2.count();
    await prisma.visionBoardV2.create({ data: { name, sortOrder: existing } });
    return null;
  }
  if (intent === "vb_rename_board") {
    if (currentUser?.role !== "superadmin") return null;
    const id = Number(form.get("boardId"));
    const name = String(form.get("name") ?? "").trim();
    if (!id || !name) return null;
    await prisma.visionBoardV2.update({ where: { id }, data: { name } });
    return null;
  }
  if (intent === "vb_delete_board") {
    if (currentUser?.role !== "superadmin") return null;
    const id = Number(form.get("boardId"));
    if (!id) return null;
    await prisma.visionBoardV2.delete({ where: { id } });
    return null;
  }
  if (intent === "vb_reorder_boards") {
    if (currentUser?.role !== "superadmin") return null;
    try {
      const ids = JSON.parse(String(form.get("ids") ?? "[]")) as number[];
      await Promise.all(ids.map((id, i) => prisma.visionBoardV2.update({ where: { id }, data: { sortOrder: i } })));
    } catch { /* ignore */ }
    return null;
  }
  if (intent === "vb_add_item") {
    if (currentUser?.role !== "superadmin") return null;
    const boardId = Number(form.get("boardId"));
    if (!boardId) return null;
    const name = String(form.get("name") ?? "").trim() || "Untitled";
    const existing = await prisma.visionBoardV2Item.count({ where: { boardId } });
    const image = form.get("image");
    const thumb = form.get("thumbnail");
    const images = typeof image === "string" && image.length > 0 ? [image] : [];
    const thumbnail = typeof thumb === "string" && thumb.length > 0 ? thumb : null;
    await prisma.visionBoardV2Item.create({ data: { boardId, name, sortOrder: existing, images, thumbnail } });
    return null;
  }
  if (intent === "vb_rename_item") {
    if (currentUser?.role !== "superadmin") return null;
    const id = Number(form.get("itemId"));
    if (!id) return null;
    await prisma.visionBoardV2Item.update({ where: { id }, data: { name: String(form.get("name") ?? "") } });
    return null;
  }
  if (intent === "vb_delete_item") {
    if (currentUser?.role !== "superadmin") return null;
    const id = Number(form.get("itemId"));
    if (!id) return null;
    await prisma.visionBoardV2Item.delete({ where: { id } });
    return null;
  }
  if (intent === "vb_reorder_items") {
    if (currentUser?.role !== "superadmin") return null;
    try {
      const ids = JSON.parse(String(form.get("ids") ?? "[]")) as number[];
      await Promise.all(ids.map((id, i) => prisma.visionBoardV2Item.update({ where: { id }, data: { sortOrder: i } })));
    } catch { /* ignore */ }
    return null;
  }
  if (intent === "vb_get_item") {
    const id = Number(form.get("itemId"));
    if (!id) return jsonResponse({ item: null });
    const item = await prisma.visionBoardV2Item.findUnique({ where: { id } });
    return jsonResponse({ item });
  }
  if (intent === "vb_update_item") {
    if (currentUser?.role !== "superadmin") return null;
    const id = Number(form.get("itemId"));
    if (!id) return null;
    const data: Record<string, unknown> = {};
    if (form.has("name")) data.name = String(form.get("name") ?? "");
    if (form.has("notes")) data.notes = String(form.get("notes") ?? "") || null;
    if (form.has("fields")) {
      try { data.fields = JSON.parse(String(form.get("fields") ?? "[]")); } catch { /* ignore */ }
    }
    if (form.has("imagesReplace")) {
      try {
        const next = JSON.parse(String(form.get("imagesReplace") ?? "[]")) as unknown[];
        if (Array.isArray(next)) data.images = next;
      } catch { /* ignore */ }
    }
    if (form.has("thumbnail")) {
      const t = String(form.get("thumbnail") ?? "");
      data.thumbnail = t || null;
    }
    if (Object.keys(data).length === 0) return null;
    await prisma.visionBoardV2Item.update({ where: { id }, data });
    return null;
  }
  if (intent === "vb_append_item_image") {
    if (currentUser?.role !== "superadmin") return null;
    const id = Number(form.get("itemId"));
    const image = String(form.get("image") ?? "");
    if (!id || !image) return null;
    const cur = await prisma.visionBoardV2Item.findUnique({ where: { id }, select: { images: true, thumbnail: true } });
    if (!cur) return null;
    const arr = Array.isArray(cur.images) ? (cur.images as unknown[]).slice() : [];
    arr.push(image);
    const data: Record<string, unknown> = { images: arr };
    if (!cur.thumbnail && form.has("thumbnail")) {
      const t = String(form.get("thumbnail") ?? "");
      if (t) data.thumbnail = t;
    }
    await prisma.visionBoardV2Item.update({ where: { id }, data });
    return null;
  }
  if (intent === "vb_remove_item_image") {
    if (currentUser?.role !== "superadmin") return null;
    const id = Number(form.get("itemId"));
    const index = Number(form.get("index"));
    if (!id || !Number.isInteger(index) || index < 0) return null;
    const cur = await prisma.visionBoardV2Item.findUnique({ where: { id }, select: { images: true, thumbnail: true } });
    if (!cur) return null;
    const arr = Array.isArray(cur.images) ? (cur.images as unknown[]).slice() : [];
    if (index >= arr.length) return null;
    const wasFirst = index === 0;
    arr.splice(index, 1);
    const data: Record<string, unknown> = { images: arr };
    if (wasFirst) data.thumbnail = null;
    await prisma.visionBoardV2Item.update({ where: { id }, data });
    return null;
  }
  if (intent === "get_sample_iteration_thumbnails") {
    let ids: number[] = [];
    try {
      const parsed = JSON.parse(String(form.get("ids") ?? "[]"));
      if (Array.isArray(parsed)) ids = parsed.filter((n): n is number => typeof n === "number" && Number.isFinite(n) && Number.isInteger(n));
    } catch { /* ignore */ }
    if (ids.length === 0) return jsonResponse({ thumbs: {} as Record<string, string> });
    const small = await prisma.sampleIteration.findMany({
      where: { id: { in: ids } },
      select: { id: true, thumbnail: true },
    }).catch(() => [] as Array<{ id: number; thumbnail: string | null }>);
    const thumbs: Record<string, string> = {};
    const missingIds: number[] = [];
    for (const row of small) {
      if (row.thumbnail) thumbs[String(row.id)] = row.thumbnail;
      else missingIds.push(row.id);
    }
    if (missingIds.length > 0) {
      const fallback = await prisma.$queryRawUnsafe<Array<{ id: number; firstImage: string | null }>>(
        `SELECT id,
           CASE WHEN jsonb_typeof(images) = 'array' AND jsonb_array_length(images) > 0
             THEN images ->> 0 ELSE NULL
           END AS "firstImage"
         FROM "SampleIteration"
         WHERE id IN (${missingIds.join(",")})`
      ).catch(() => [] as Array<{ id: number; firstImage: string | null }>);
      for (const row of fallback) if (row.firstImage) thumbs[String(row.id)] = row.firstImage;
    }
    return jsonResponse({ thumbs });
  }

  // ─── Collections ─────────────────────────────────────────────────────────────
  if (intent === "add_collection") {
    const name = String(form.get("name") ?? "").trim() || "Untitled collection";
    const existing = await prisma.collection.count();
    const initialThumb = form.get("thumbnail");
    const thumbnail = typeof initialThumb === "string" && initialThumb.length > 0 ? initialThumb : null;
    await prisma.collection.create({ data: { name, sortOrder: existing, thumbnail } });
    return null;
  }
  if (intent === "rename_collection") {
    const id = Number(form.get("collectionId"));
    const name = String(form.get("name") ?? "").trim();
    if (!id || !name) return null;
    await prisma.collection.update({ where: { id }, data: { name } });
    return null;
  }
  if (intent === "delete_collection") {
    const id = Number(form.get("collectionId"));
    if (!id) return null;
    await prisma.collection.delete({ where: { id } });
    return null;
  }
  if (intent === "reorder_collections") {
    const ids = JSON.parse(String(form.get("collectionIds") ?? "[]")) as number[];
    await Promise.all(ids.map((id, index) => prisma.collection.update({ where: { id }, data: { sortOrder: index } })));
    return null;
  }
  if (intent === "get_collection_full") {
    const id = Number(form.get("collectionId"));
    if (!id) return jsonResponse({ collection: null });
    const collection = await prisma.collection.findUnique({ where: { id } }).catch(() => null);
    return jsonResponse({ collection });
  }
  if (intent === "update_collection") {
    const id = Number(form.get("collectionId"));
    if (!id) return null;
    const data: Record<string, unknown> = { updatedAt: new Date() };
    if (form.has("rows")) {
      try {
        const rows = JSON.parse(String(form.get("rows")));
        if (Array.isArray(rows)) data.rows = rows;
      } catch { /* keep existing */ }
    }
    if (form.has("columns")) {
      try {
        const columns = JSON.parse(String(form.get("columns")));
        if (Array.isArray(columns)) data.columns = columns;
      } catch { /* keep existing */ }
    }
    if (form.has("thumbnail")) {
      const thumb = String(form.get("thumbnail") ?? "");
      data.thumbnail = thumb || null;
    }
    if (form.has("name")) {
      const name = String(form.get("name") ?? "").trim();
      if (name) data.name = name;
    }
    await prisma.collection.update({ where: { id }, data });
    return null;
  }

  if (intent === "update_fabric_cell") {
    const gid = String(form.get("gid") ?? "");
    const rowIndex = Number(form.get("rowIndex"));
    const colIndex = Number(form.get("colIndex"));
    const value = String(form.get("value") ?? "");
    if (!gid || !Number.isInteger(rowIndex) || !Number.isInteger(colIndex) || rowIndex < 0 || colIndex < 0) {
      return null;
    }

    const sheets = await loadManualFabricSheetsForAction();
    const sheet = sheets.find((item) => item.gid === gid);
    if (!sheet?.rows[rowIndex]) return null;
    while (sheet.rows[rowIndex].length <= colIndex) sheet.rows[rowIndex].push("");
    sheet.rows[rowIndex][colIndex] = value;
    await saveManualFabricSheets(sheets);
    return null;
  }

  if (intent === "upload_fabric_image") {
    const gid = String(form.get("gid") ?? "");
    const rowIndex = Number(form.get("rowIndex"));
    const colIndex = Number(form.get("colIndex"));
    const imageFile = form.get("image");
    if (
      !gid ||
      !Number.isInteger(rowIndex) ||
      !Number.isInteger(colIndex) ||
      rowIndex < 0 ||
      colIndex < 0 ||
      !imageFile ||
      typeof imageFile === "string" ||
      typeof imageFile.arrayBuffer !== "function"
    ) {
      return null;
    }
    if (!imageFile.type.startsWith("image/") || imageFile.size > 5 * 1024 * 1024) return null;

    const buffer = Buffer.from(await imageFile.arrayBuffer());
    const dataUrl = `data:${imageFile.type};base64,${buffer.toString("base64")}`;
    const sheets = await loadManualFabricSheetsForAction();
    const sheet = sheets.find((item) => item.gid === gid);
    if (!sheet?.rows[rowIndex]) return null;
    while (sheet.rows[rowIndex].length <= colIndex) sheet.rows[rowIndex].push("");
    sheet.rows[rowIndex][colIndex] = dataUrl;
    await saveManualFabricSheets(sheets);
    return null;
  }

  if (intent === "delete_fabric_row" || intent === "move_fabric_row" || intent === "duplicate_fabric_row") {
    const gid = String(form.get("gid") ?? "");
    const rowIndex = Number(form.get("rowIndex"));
    const targetGid = String(form.get("targetGid") ?? "");
    const sheets = await loadManualFabricSheetsForAction();
    const sourceSheet = sheets.find((item) => item.gid === gid);
    const targetSheet = sheets.find((item) => item.gid === targetGid);
    if (!sourceSheet || !Number.isInteger(rowIndex) || rowIndex < 0) return null;
    if (intent === "move_fabric_row" && (!targetSheet || targetSheet.gid === sourceSheet.gid || isHiddenFabricSheet(targetSheet.name))) return null;

    if ((intent === "move_fabric_row" && targetSheet) || intent === "duplicate_fabric_row") {
      const row = sourceSheet.rows[rowIndex]?.map((value) => value);
      if (!row) return null;
      if (intent === "duplicate_fabric_row") {
        sourceSheet.rows.push(row);
      } else if (targetSheet) {
        const sourceHeaderMap = new Map<string, number>();
        for (const [index, header] of sourceSheet.headers.entries()) {
          sourceHeaderMap.set(normalizeFabricHeader(header), index);
          const role = fabricHeaderRole(header);
          if (!sourceHeaderMap.has(role)) sourceHeaderMap.set(role, index);
        }
        const movedRow = targetSheet.headers.map((header) => {
          const sourceIndex = sourceHeaderMap.get(normalizeFabricHeader(header)) ?? sourceHeaderMap.get(fabricHeaderRole(header));
          return sourceIndex == null ? "" : row[sourceIndex] ?? "";
        });
        targetSheet.rows.push(movedRow);
      }
    }

    if (intent !== "duplicate_fabric_row") {
      sourceSheet.rows.splice(rowIndex, 1);
    }

    await saveManualFabricSheets(sheets);
    return null;
  }

  if (intent === "add_fabric_row") {
    const gid = String(form.get("gid") ?? "");
    const sheets = await loadManualFabricSheetsForAction();
    const sheet = sheets.find((item) => item.gid === gid);
    if (!sheet) return null;
    sheet.rows.push(Array.from({ length: sheet.headers.length }, () => ""));
    await saveManualFabricSheets(sheets);
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
    const orderForVariant = await prisma.supplierOrder.findUnique({
      where: { id: orderId },
      select: { shop: true, productId: true, productTitle: true },
    });
    const matchingVariant = orderForVariant
      ? matchingVariantForSize(await getShopifyProductVariants(orderForVariant.shop, orderForVariant.productId), size)
      : null;
    const existingLines = await prisma.orderLine.findMany({
      where: { orderId, variantTitle: size },
      orderBy: { id: "asc" },
      select: { id: true, variantId: true, qtyOrdered: true },
    });

    const previousQty = existingLines.length ? existingLines[0].qtyOrdered : 0;

    await prisma.$transaction(async (tx) => {
      const lines = existingLines;

      if (!lines.length) {
        await tx.orderLine.create({
          data: {
            orderId,
            variantId: matchingVariant?.id ?? `${orderId}:${size}`,
            variantTitle: matchingVariant?.title ?? size,
            sku: matchingVariant?.sku ?? null,
            qtyOrdered,
          },
        });
      } else {
        await tx.orderLine.update({
          where: { id: lines[0].id },
          data: {
            qtyOrdered,
            ...(matchingVariant ? {
              variantId: matchingVariant.id,
              variantTitle: matchingVariant.title,
              sku: matchingVariant.sku,
            } : {}),
          },
        });

        if (lines.length > 1) {
          await tx.orderLine.updateMany({
            where: { id: { in: lines.slice(1).map((line) => line.id) } },
            data: { qtyOrdered: 0 },
          });
        }
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

    if (qtyOrdered !== previousQty) {
      await logActivity(currentUser?.name ?? "Unknown", "Updated", "Restock Order", {
        entityId: String(orderId),
        entityName: orderForVariant?.productTitle ?? `Order #${orderId}`,
        field: `Qty (${size})`,
        toValue: String(qtyOrdered),
      });
    }

    return null;
  }

  if (Object.keys(updates).length) {
    await prisma.supplierOrder.update({ where: { id: orderId }, data: updates });
    if (intent === "update_factory_notes" || intent === "update_notes") {
      await syncOrderNoteMessages({
        orderId,
        field: intent === "update_factory_notes" ? "factory_notes" : "notes",
        text: String(form.get("value") ?? ""),
        fromName: currentUser?.name ?? null,
      });
    }
    const loggableIntents: Record<string, string> = {
      update_status: "Status",
      update_priority: "Priority",
      update_factory_notes: "Factory notes",
      update_notes: "Notes",
      update_eta: "ETA",
    };
    const logField = loggableIntents[intent];
    if (logField && orderId) {
      const order = await prisma.supplierOrder.findUnique({ where: { id: orderId }, select: { productTitle: true } });
      await logActivity(currentUser?.name ?? "Unknown", "Updated", "Restock Order", {
        entityId: String(orderId),
        entityName: order?.productTitle ?? `Order #${orderId}`,
        field: logField,
        toValue: String(form.get("value") ?? ""),
      });
    }
  }
  return null;
  } catch (error) {
    console.error("Portal action error:", error);
    return null;
  }
};

// Reorder is handled optimistically in the UI — skip the expensive full loader reload
export const shouldRevalidate: ShouldRevalidateFunction = ({ formData, defaultShouldRevalidate }) => {
  // Background callers (e.g. the silent image-compression hook) opt out of
  // loader revalidation by setting noRevalidate=1. This stops dozens of
  // sequential background POSTs from each triggering a full loader run.
  if (formData?.get("noRevalidate") === "1") return false;
  const intent = formData?.get("intent") as string | null;
  if (intent === "reorder_samples" || intent === "rename_sample" || intent === "update_sample_iteration") return false;
  // Vision Board: only skip revalidation for in-drawer notes/fields saves
  // (they don't affect the card grid). Anything that changes name, image
  // count, ordering or board metadata must refresh the grid so the panel
  // sees the new items prop.
  if (intent === "vb_update_item" && !formData?.has("name") && !formData?.has("thumbnail")) return false;
  if (intent === "update_collection" || intent === "rename_collection" || intent === "reorder_collections") return false;
  if (intent === "update_column_widths" || intent === "update_packing_column_widths") return false;
  // Read-only fetchers used by drawers and background hooks don't change
  // any data the page renders, so the loader doesn't need to re-run.
  if (intent === "vb_get_item" || intent === "get_sample_full"
    || intent === "get_sample_iteration_thumbnails"
    || intent === "get_collection_full") return false;
  return defaultShouldRevalidate;
};

// ─── Activity log helper ──────────────────────────────────────────────────────

async function logActivity(
  userName: string,
  action: string,
  entity: string,
  opts?: { entityId?: string; entityName?: string; field?: string; toValue?: string }
) {
  try {
    await prisma.activityLog.create({
      data: {
        userName,
        action,
        entity,
        entityId: opts?.entityId ?? null,
        entityName: opts?.entityName ?? null,
        field: opts?.field ?? null,
        toValue: opts?.toValue ?? null,
      },
    });
  } catch {
    // silently ignore if table not yet migrated
  }
}

// ─── Constants ───────────────────────────────────────────────────────────────

const DEFAULT_STATUS_OPTIONS = [
  { value: "on_order",       label: "On Order" },
  { value: "on_production",  label: "On Production" },
  { value: "in_shipment",    label: "In Shipment" },
  { value: "packed",         label: "Packed" },
  { value: "arrived",        label: "Arrived" },
  { value: "arrived_loaded", label: "Arrived and Loaded" },
  { value: "cancelled",      label: "Cancelled" },
  { value: "ready_to_send",  label: "Ready To Send" },
];

const DEFAULT_STATUS_COLORS: Record<string, { bg: string; color: string }> = {
  on_order:       { bg: "#fef9c3", color: "#374151" },
  on_production:  { bg: "#dbeafe", color: "#374151" },
  in_shipment:    { bg: "#dcfce7", color: "#374151" },
  packed:         { bg: "#fed7aa", color: "#7c2d12" },
  arrived:        { bg: "#bbf7d0", color: "#14532d" },
  arrived_loaded: { bg: "#4ade80", color: "#052e16" },
  cancelled:      { bg: "#fee2e2", color: "#991b1b" },
  ready_to_send:  { bg: "#ede9fe", color: "#4c1d95" },
};

const DEFAULT_PRIORITY_OPTIONS = [
  { value: "low",       label: "LOW",       bg: "#3b82f6", color: "#fff" },
  { value: "high",      label: "HIGH",      bg: "#7c3aed", color: "#fff" },
  { value: "urgent",    label: "URGENT",    bg: "#dc2626", color: "#fff" },
  { value: "cancelled", label: "Cancelled", bg: "#d97706", color: "#fff" },
];
const FABRIC_CHIP_COLORS = [
  { bg: "#dbeafe", color: "#1e3a8a" },
  { bg: "#dcfce7", color: "#14532d" },
  { bg: "#fef3c7", color: "#92400e" },
  { bg: "#ede9fe", color: "#4c1d95" },
  { bg: "#fee2e2", color: "#991b1b" },
  { bg: "#ccfbf1", color: "#134e4a" },
  { bg: "#fce7f3", color: "#831843" },
  { bg: "#e0f2fe", color: "#075985" },
];
const PACKING_STATUS_OPTIONS = [
  { value: "still_packing", label: "Still packing" },
  { value: "on_the_way", label: "On the way" },
  { value: "arrived", label: "Arrived" },
  { value: "partially_loaded", label: "Partially loaded" },
  { value: "loaded", label: "Inventory loaded" },
];
const PACKING_SIZES = ["XS", "S", "M", "L", "XL", "2XL", "3XL", "S/M", "M/L", "L/XL"];
const DEFAULT_PACKING_ROWS = 5;
const PACKING_COLUMNS_BEFORE_SIZES = [
  { id: "box", label: "Box", width: 70, center: true },
  { id: "picture", label: "Picture", width: 150, center: true },
  { id: "fabric", label: "Fabric Image", width: 150, center: true },
  { id: "name", label: "Name", width: 320 },
  { id: "sku", label: "SKU", width: 220 },
];
const PACKING_COLUMNS_AFTER_SIZES = [
  { id: "total", label: "Total", width: 82, center: true },
  { id: "price", label: "Price ₹", width: 92, center: true },
  { id: "value", label: "Value ₹", width: 96, center: true },
  { id: "weight", label: "Weight", width: 90, center: true },
  { id: "shopify", label: "Shopify", width: 84, center: true },
];
// Builds the size column list for a given packing list, merging the
// baseline (XS..3XL plus the slash combos) with any extra size keys
// present in the lines themselves (Free Size, XL-2XL, etc). Baseline
// columns come first in their canonical order; extras come after in the
// order they first appear in the data — keeps the layout stable for
// standard lists while still surfacing one-off sizes.
function derivePackingSizes(lines: Array<{ qtys?: unknown }>): string[] {
  const baselineSet = new Set(PACKING_SIZES);
  const extras: string[] = [];
  const seenExtras = new Set<string>();
  for (const line of lines) {
    const qtys = normalizeQtys(line.qtys);
    for (const key of Object.keys(qtys)) {
      const trimmed = key.trim();
      if (!trimmed || baselineSet.has(trimmed) || seenExtras.has(trimmed)) continue;
      seenExtras.add(trimmed);
      extras.push(trimmed);
    }
  }
  return [...PACKING_SIZES, ...extras];
}
function packingColumnsForSizes(sizes: string[]) {
  return [
    ...PACKING_COLUMNS_BEFORE_SIZES,
    ...sizes.map((size) => ({ id: `qty:${size}`, label: size, width: 76, center: true })),
    ...PACKING_COLUMNS_AFTER_SIZES,
  ];
}
const PACKING_COLUMNS = packingColumnsForSizes(PACKING_SIZES);
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

async function retryAsync<T>(operation: () => Promise<T>, label: string, attempts = 2): Promise<T> {
  let lastError: unknown;
  for (let attempt = 0; attempt <= attempts; attempt += 1) {
    try {
      return await operation();
    } catch (error) {
      lastError = error;
      if (attempt >= attempts) break;
      await new Promise((resolve) => setTimeout(resolve, 180 * (attempt + 1)));
    }
  }
  console.error(`${label} failed after retry`, lastError);
  throw lastError;
}

function labelForOption(options: RestockOption[], value: string) {
  return options.find((option) => option.value === value)?.label ?? value;
}

function labelForPackingStatus(value: string) {
  return PACKING_STATUS_OPTIONS.find((option) => option.value === value)?.label ?? value;
}

function normalizeTableHeaderLabels(value: unknown): Record<string, string> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  return Object.fromEntries(
    Object.entries(value as Record<string, unknown>)
      .map(([key, label]) => [key, String(label ?? "").trim()])
      .filter(([, label]) => label),
  );
}

function headerLabel(labels: Record<string, string>, key: string, fallback: string) {
  return labels[key] || fallback;
}

function normalizeTableCustomColumns(value: unknown): TableCustomColumns {
  const raw = value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : {};
  const normalizeList = (list: unknown): TableCustomColumn[] => Array.isArray(list)
    ? list.map((item) => {
        if (!item || typeof item !== "object") return null;
        const column = item as Record<string, unknown>;
        const id = String(column.id ?? "").trim();
        const label = String(column.label ?? "").trim();
        return id && label ? { id, label } : null;
      }).filter(Boolean) as TableCustomColumn[]
    : [];
  const fabric = raw.fabric && typeof raw.fabric === "object" && !Array.isArray(raw.fabric)
    ? Object.fromEntries(Object.entries(raw.fabric as Record<string, unknown>).map(([gid, list]) => [gid, normalizeList(list)]))
    : {};
  return { restock: normalizeList(raw.restock), packing: normalizeList(raw.packing), fabric };
}

function normalizeTableCustomCells(value: unknown): Record<string, string> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  return Object.fromEntries(Object.entries(value as Record<string, unknown>).map(([key, cellValue]) => [key, String(cellValue ?? "")]));
}

function normalizeTableRowHeights(value: unknown): Record<string, number> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  return Object.fromEntries(
    Object.entries(value as Record<string, unknown>)
      .map(([key, height]) => [key, Math.min(420, Math.max(34, Number(height) || 0))] as const)
      .filter(([, height]) => height > 0),
  );
}

function normalizeFabricCustomSheets(value: unknown): FabricCustomSheet[] {
  if (!Array.isArray(value)) return [];
  return value.map((item) => {
    if (!item || typeof item !== "object") return null;
    const sheet = item as Record<string, unknown>;
    const gid = String(sheet.gid ?? "").trim();
    const name = String(sheet.name ?? "").trim();
    const headers = Array.isArray(sheet.headers) ? sheet.headers.map((header) => String(header || "Column")) : [];
    const rows = Array.isArray(sheet.rows)
      ? sheet.rows.filter((row): row is unknown[] => Array.isArray(row)).map((row) => row.map((cell) => String(cell ?? "")))
      : [];
    return gid && name ? { gid, name, headers: headers.length ? headers : DEFAULT_FABRIC_HEADERS, rows } : null;
  }).filter(Boolean) as FabricCustomSheet[];
}

function normalizeFabricDeletedSheets(value: unknown): Record<string, boolean> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  return Object.fromEntries(
    Object.entries(value as Record<string, unknown>)
      .filter(([gid, deleted]) => gid && Boolean(deleted))
      .map(([gid]) => [gid, true]),
  );
}

const COLUMN_WIDTHS_KEY = "supplier-portal-column-widths-v1";
const PACKING_COLUMN_WIDTHS_KEY = "supplier-portal-packing-column-widths-v1";
const TABLE_HEADER_LABELS_KEY = "production-portal-table-header-labels-v1";
const TABLE_CUSTOM_COLUMNS_KEY = "production-portal-table-custom-columns-v1";
const TABLE_CUSTOM_CELLS_KEY = "production-portal-table-custom-cells-v1";
const TABLE_ROW_HEIGHTS_KEY = "production-portal-table-row-heights-v1";
const RESTOCK_SETTINGS_KEY = "supplier-portal-restock-settings-v1";
const UNIVERSAL_SETTINGS_KEY = "production-portal-universal-settings-v1";
const PORTAL_NAV_ORDER_KEY = "production-portal-nav-order-v1";
const FABRIC_SETTINGS_KEY = "production-portal-fabric-settings-v1";
const FABRIC_CELL_OVERRIDES_KEY = "production-portal-fabric-cell-overrides-v1";
const FABRIC_CUSTOM_ROWS_KEY = "production-portal-fabric-custom-rows-v1";
const FABRIC_DELETED_ROWS_KEY = "production-portal-fabric-deleted-rows-v1";
const FABRIC_CUSTOM_SHEETS_KEY = "production-portal-fabric-custom-sheets-v1";
const FABRIC_DELETED_SHEETS_KEY = "production-portal-fabric-deleted-sheets-v1";
const FABRIC_MANUAL_SHEETS_KEY = "production-portal-fabric-manual-sheets-v1";
const PRODUCT_INFO_KEY = "production-portal-product-info-v2";
const DEFAULT_FABRIC_HEADERS = ["Supplier", "Fabric Type", "Fabric", "Name", "Cost per Meter", "Meters in Stock", "Cut Pieces", "Received / Date", "Products", "Notes"];
const PRODUCT_STYLE_IMAGES: Record<string, string> = {
  "Acacia Maxi Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/acacia-maxi-dress-spruced-up-teal-rayon-summer-maxi-dress-womens-karma-east_6572.jpg?v=1764067484",
  "Alice Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/Alice-Dress-Midnight-Reverie-retro-cotton-dress-with-pockets.jpg?v=1693956103",
  "August Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/August-Dress-Dahlia-long-cotton-dress-with-pockets-and-lining.jpg?v=1723600463",
  "Billie Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/billie-dress-daisy-dress-pockets-sage-floral-print-karma-east.jpg_5255.jpg?v=1768303381",
  "Blythe Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/blythe-dress-daisy-100-cotton-sage-based-daisy-print-pockets-karma-east_0042.jpg?v=1772595403",
  "Boho Tiered Maxi Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/Boho-Tiered-Maxi-Dress-Rose-100-percent-cotton-pink-floral-sundress-with-pockets_6131.jpg?v=1767960651",
  "Chelsea Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/chelsea-dress-daisy-100-cotton-sun-dress-pockets-sage-based-floral-karma-east0188.jpg?v=1770206506",
  "Claudia Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/Claudia-Dress-Black-100-percent-cotton-maxi-dress.jpg?v=1710475889",
  "Ember Midi Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/ember-midi-dress-blue-mango-midi-dress-100-cotton-navy-floral-print-karma-east.jpg_9316.jpg?v=1777029682",
  "Etta Dress": "https://cdn.shopify.com/s/files/1/1204/4848/products/EttaDressDapple_1104.jpg?v=1665727569",
  "Francesca Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/Francessca-Dress-Lilly-Lane-navy-floral-rayon-sundress-maxi.jpg?v=1727854086",
  "Frankie Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/frankie-dress-amla-midi-dress-100-cotton-navy-floral-print-karma-east.jpg_9373.jpg?v=1777096362",
  "Harper Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/HarperPincordDressTabascoOrangeRedCottonCordDress.jpg?v=1776219896",
  "Kari Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/kari-dress-nocturne-dark-chocolate-black-cotton-womens-dress-with-pockets-karma-east0493.jpg?v=1761705710",
  "Katie Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/katie-dress-rose-100-cotton-cream-based-floral-buttoned-bodice-karma-east_2144.jpg?v=1772796568",
  "Mabel Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/MabelTangerineDreams2.jpg?v=1776305383",
  "Maddison Dress": "https://cdn.shopify.com/s/files/1/1204/4848/products/MaddisonBirdofParadise5.jpg?v=1668577902",
  "Nakita Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/nakita-dress-spruced-up-teal-rayon-lightweight-sundress-womens-karma-east_6673.jpg?v=1764115693",
  "Nina Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/nina-dress-rose-100-cotton-cream-based-rose-floral-fit-flare-cap-sleeves-karma-east_0120.jpg?v=1772797333",
  "Pippa Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/Pippa-Dress-Longer-Riley-Dress-Ochre-Front.jpg?v=1776765867",
  "Promenade Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/promenade-dress-blue-mango-midi-dress-100-cotton-navy-floral-print-karma-east.jpg_9424.jpg?v=1777179679",
  "Riley Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/Riley-Dress-Dapple-navy-cotton-midi-dress-with-pockets-polkadots_4132.jpg?v=1752628341",
  "Rita Dress": "https://cdn.shopify.com/s/files/1/1204/4848/products/RitaDressViola_1656.jpg?v=1664518477",
  "Scarlett Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/scarlett-dress-camellia-dress-100-cotton-black-floral-print-karma-east.jpg_6761.jpg?v=1767096580",
  "Tully Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/Tully-Dress-Tibetian-Red-100-percentt-cotton-sundress.jpg?v=1732676529",
  "Tulsi Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/tulsi-dress-amla-dress-100-cotton-navy-floral-print-karma-east.jpg_9539.jpg?v=1777097015",
  "Ursula Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/ursula-dress-rose-100-cotton-cream-based-rose-floral-pockets-buttoned-bodice-karma-east_0155.jpg?v=1772797657",
  "Vivien Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/Vivien-Dress-Dapple-50s-style-cotton-button-through-midi-length-pockets-2.jpg?v=1758091942",
  "Willow Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/willow-dress-sakura-pink-plum-floral-cotton-womens-dress-karma-east_8206.jpg?v=1761964279",
  "Avery Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/Avery-Dress-Black-100-cotton-midi_6654.jpg?v=1752025759",
  "Briar Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/Briar-Dress-Cascade-lined-cotton.jpg?v=1737512029",
  "Ella Wrap Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/Etta-Wrap-Dress-Midnight-Reverie-Front.jpg?v=1773808001",
  "Jamie Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/Jamie-Dress-Pomegranate-knee-length-blue-base-fruit-print.jpg?v=1689726170",
  "Olivia Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/Olivia-Dress-Kaveri-100-cotton-maxi-with-sleeves.jpg?v=1744695045",
  "Savannah Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/karma-east-savannah-dress-spruced-up-teal-rayon-sundress-lightweight-summer-dress-with-pockets_0067.jpg?v=1763615601",
  "Tiered Maxi Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/tiered-maxi-dress-blue-mango-maxi-dress-100-cotton-navy-floral-print-karma-east.jpg_9505.jpg?v=1777180192",
  "Tiered Midi Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/tiered-midi-dress-blue-mango-midi-dress-100-cotton-navy-floral-print-karma-east.jpg_9272.jpg?v=1777026109",
  "Tilda Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/tilda-dress-amla-dress-100-cotton-navy-floral-print-karma-east.jpg_9595.jpg?v=1777095763",
  "Zoey Dress": "https://cdn.shopify.com/s/files/1/1204/4848/files/Zoey-Dress-Black-100-percent-double-cotton-gauze-long-sleeve-dress.jpg?v=1719463574",
  "Aria Top": "https://cdn.shopify.com/s/files/1/1204/4848/files/aria-top-amla-top-100-cotton-navy-floral-print-karma-east.jpg_1052.jpg?v=1777094608",
  "Aubrey Top": "https://cdn.shopify.com/s/files/1/1204/4848/files/aubrey-top-camellia-top-100-cotton-black-floral-print-karma-east.jpg_6546.jpg?v=1766572419",
  "Boxy Top": "https://cdn.shopify.com/s/files/1/1204/4848/files/Boxy-Top-Bird-of-Paradise-navy-cotton-floral-womens-top_425.jpg?v=1752129885",
  "Camille Top": "https://cdn.shopify.com/s/files/1/1204/4848/files/camille-top-natural-white-cotton-womens-top-karma-east_9648.jpg?v=1761098932",
  "Chloe Top": "https://cdn.shopify.com/s/files/1/1204/4848/files/Chloe-Top-Natural-white-cotton-womens.jpg?v=1739337037",
  "Eden Top": "https://cdn.shopify.com/s/files/1/1204/4848/files/Eden-Top-Neela-loose-cotton-womens-blouse.jpg?v=1725001453",
  "Neesha Top": "https://cdn.shopify.com/s/files/1/1204/4848/products/NeeshaTopDeepDiveCottonGauzeShortSleeve.jpg?v=1676526279",
  "Pia Top": "https://cdn.shopify.com/s/files/1/1204/4848/files/Pia-Top-Queen-Protea-black-based-floral-cotton-1.jpg?v=1758717947",
  "Quinn Top": "https://cdn.shopify.com/s/files/1/1204/4848/files/Quinn-Natural-White-100-percent-cotton-blouse.jpg?v=1760782606",
  "Shell Top": "https://cdn.shopify.com/s/files/1/1204/4848/products/ShellPeonya.jpg?v=1603953508",
  "Sky Blouse": "https://cdn.shopify.com/s/files/1/1204/4848/files/sky-blouse-amla-top-100-cotton-navy-floral-print-karma-east.jpg_0999.jpg?v=1777032120",
  "Sylvia Top": "https://cdn.shopify.com/s/files/1/1204/4848/files/Sylvia-Top-Sweet-Spot-100-percent-cotton_4793.jpg?v=1752808606",
  "Tillie Top": "https://cdn.shopify.com/s/files/1/1204/4848/files/tillie-top-rose-100-cotton-top-cream-based-floral-karma-east9411.jpg?v=1770375766",
  "Tulsi Short Sleeve Top": "https://cdn.shopify.com/s/files/1/1204/4848/files/tulsi-short-sleeve-top-amla-top-100-cotton-navy-floral-print-karma-east.jpg_1108.jpg?v=1777180817",
  "Yasmin Top": "https://cdn.shopify.com/s/files/1/1204/4848/files/yasmin-top-rose-100-cotton-top-cream-based-floral-karma-east9349.jpg?v=1770376178",
  "Zali Top": "https://cdn.shopify.com/s/files/1/1204/4848/files/zali-top-black-100-percent-cotton-summer-top_6666.jpg?v=1774324401",
  "Demi Top": "https://cdn.shopify.com/s/files/1/1204/4848/files/Demi-Top-Natural-White-pure-cotton-blouse.jpg?v=1753763334",
  "Florence Top": "https://cdn.shopify.com/s/files/1/1204/4848/files/florence-top-blue-mango-top-100-cotton-navy-floral-print-karma-east.jpg_0925.jpg?v=1777028824",
  "Isla Top": "https://cdn.shopify.com/s/files/1/1204/4848/files/Isla-Top-Midnight-Reverie-navy-blue-cotton-floral-womans-top.jpg?v=1693972613",
  "Leia Top": "https://cdn.shopify.com/s/files/1/1204/4848/files/Leia-Top-Ochre-Front.jpg?v=1776765271",
  "Sophia Blouse": "https://cdn.shopify.com/s/files/1/1204/4848/files/Sophia-Top-Cranberry-cotton-blouse.jpg?v=1744427994",
  "Tulsi Long Sleeve Top": "https://cdn.shopify.com/s/files/1/1204/4848/files/tulsi-long-sleeve-top-blue-mango-top-100-cotton-navy-floral-print-karma-east.jpg_0878.jpg?v=1777028016",
  "Aalia Skirt": "https://cdn.shopify.com/s/files/1/1204/4848/files/aalia-skirt-sakura-pink-plum-floral-cotton-womens-skirt-karma-east_9477.jpg?v=1761967853",
  "Belt Loop Skirt": "https://cdn.shopify.com/s/files/1/1204/4848/products/ShadedSpruce1.jpg?v=1652670216",
  "Bridgette Skirt": "https://cdn.shopify.com/s/files/1/1204/4848/files/bridgette-skirt-mistlerose-red-floral-cotton-summer-skirt-womens-karma-east_6113.jpg?v=1763638301",
  "Reversible Skirt": "https://cdn.shopify.com/s/files/1/1204/4848/files/Reversible-Skirt-Poppy-Queen-Protea-floral-cotton-zip-2.jpg?v=1758890691",
  "Ruby Skirt": "https://cdn.shopify.com/s/files/1/1204/4848/files/ruby-skirt-tallowwood-100-cotton-retro-floral-a-line-button-through-knee-length-pockets-karma-east_0347.jpg?v=1772799566",
  "Zarah Skirt": "https://cdn.shopify.com/s/files/1/1204/4848/files/zarah-skirt-daisy-skirt-100-cotton-sage-floral-print-karma-east.jpg_6434.jpg?v=1767419347",
  "Cora Skirt": "https://cdn.shopify.com/s/files/1/1204/4848/products/CoraSkirtProtea_KE_871-Edit.jpg?v=1662372122",
  "Dawn Maxi Skirt": "https://cdn.shopify.com/s/files/1/1204/4848/files/dawn-maxi-amla-skirt-100-cotton-navy-floral-print-karma-east.jpg_0607.jpg?v=1777181724",
  "Acapulco Pant": "https://cdn.shopify.com/s/files/1/1204/4848/files/Acapulco-Pant-White-cotton-harem-pockets-3.jpg?v=1760669768",
  "Flower Pants": "https://cdn.shopify.com/s/files/1/1204/4848/files/flower-pants-moon-flower-pants-100-cotton-moon-floral-print-karma-east.jpg_0451.jpg?v=1776299535",
  "Greta Pant": "https://cdn.shopify.com/s/files/1/1204/4848/files/Greta-Pant-Black-100-percent-cotton-womens-pant6305.jpg?v=1758771997",
  "Janis Pant": "https://cdn.shopify.com/s/files/1/1204/4848/products/JanisPantPowderBlue.jpg?v=1636614663",
  "Jaya Pant": "https://cdn.shopify.com/s/files/1/1204/4848/files/Jaya-Pant-Evergreen-100-percent-cotton-green-womans-pants-wide-leg-with-pocket.jpg?v=1694655280",
  "Nora Pant": "https://cdn.shopify.com/s/files/1/1204/4848/files/nora-pants-nocturne-pants-100-cotton-dark-floral-print-karma-east.jpg_0376.jpg?v=1774928258",
  "Pilot Pant": "https://cdn.shopify.com/s/files/1/1204/4848/files/pilot-pants-amla-pants-100-cotton-navy-floral-print-karma-east.jpg_0713.jpg?v=1777093677",
  "Remi Pant": "https://cdn.shopify.com/s/files/1/1204/4848/files/Remi-Pants-Queen-Protea-black-cotton-women_s-pant5793.jpg?v=1758888894",
  "Tenzin Pant": "https://cdn.shopify.com/s/files/1/1204/4848/files/tenzin-pant-nocturne-pants-100-cotton-dark-floral-print-karma-east.jpg_0820.jpg?v=1774009212",
  "Umbrella Pant": "https://cdn.shopify.com/s/files/1/1204/4848/files/umbrella-pant-jasmine-floral-rayon-womens-pant-karma-east9998.jpg?v=1761625795",
  "Wide Leg Stretch Pocket Pants": "https://cdn.shopify.com/s/files/1/1204/4848/products/2M4A9621.jpg?v=1603784887",
  "Aalia Corduroy Skirt": "https://cdn.shopify.com/s/files/1/1204/4848/files/aalia-corduroy-skirt-maritime-blue-100-cotton-knee-length-pockets-karma-east_2049_641405da-7511-4949-9bfd-c0625b68f79e.jpg?v=1773800943",
  "Belt Loop Corduroy Skirt": "https://cdn.shopify.com/s/files/1/1204/4848/files/Belt-Loop-Corduroy-Skirt-Douglas-Fir-green-midi-length.jpg?v=1741825931",
  "Corduroy Jacket": "https://cdn.shopify.com/s/files/1/1204/4848/files/Cord-Jacket-Rain-Forest-Green-with-breast-pockets.jpg?v=1684822814",
  "Corduroy Overalls": "https://cdn.shopify.com/s/files/1/1204/4848/files/Corduroy-Overalls-Parachute-Purple-cotton-pocket-overalls.jpg?v=1712035154",
  "Corduroy Pinafore": "https://cdn.shopify.com/s/files/1/1204/4848/files/Corduroy-Pinafore-Zinfandel-maroon-cotton-layer-over-leggings.jpg?v=1710811127",
  "Esta Corduroy Pants": "https://cdn.shopify.com/s/files/1/1204/4848/files/Corduroy-Overalls-Parachute-Purple-cotton-pocket-overalls.jpg?v=1712035154",
  "Jamie Corduroy Dress": "https://cdn.shopify.com/s/files/1/1204/4848/products/JamieCordCherryc.jpg?v=1620800356",
  "Nora Corduroy Pants": "https://cdn.shopify.com/s/files/1/1204/4848/files/nora-corduroy-pants-black-100-cotton-pockets-zip-high-waist-karma-east_1409.jpg?v=1773294116",
  "Polly Pocket Corduroy Tunic": "https://cdn.shopify.com/s/files/1/1204/4848/files/polly-pocket-corduroy-sleeveless-tunic-black-100-cotton-pockets-karma-east_2609_1430bb10-1339-438e-8dfe-80c16f2d61dd.jpg?v=1773284527",
};
const PRODUCT_STYLE_COSTING: Record<string, ProductStyleCosting> = {
  "Acacia Maxi Dress": { sheetCount: 8, averageMeters: 4.47, stitchingCost: 105, fabricCost: 385.88, zipButtonsCost: 6, totalCost: 494.38, zipButtonType: "6 Button" },
  "Alice Dress": { sheetCount: 24, averageMeters: 2.7, stitchingCost: 100, fabricCost: 310.88, zipButtonsCost: 40.29, totalCost: 451.17, zipButtonType: "24 inch" },
  "August Dress": { sheetCount: 16, averageMeters: 3.44, averageTrimMeters: 1.79, stitchingCost: 75, fabricCost: 437.63, zipButtonsCost: 2.44, totalCost: 515.06, zipButtonType: "3 Button" },
  "Billie Dress": { sheetCount: 28, averageMeters: 2.49, stitchingCost: 90, fabricCost: 297.41, zipButtonsCost: 28.04, totalCost: 415.41, zipButtonType: "14 Inch" },
  "Boho Tiered Maxi Dress": { sheetCount: 15, averageMeters: 3.74, stitchingCost: 95.67, fabricCost: 391, zipButtonsCost: 5.87, totalCost: 492.53, zipButtonType: "3 Button" },
  "Chelsea Dress": { sheetCount: 21, averageMeters: 3.6, stitchingCost: 79.52, fabricCost: 352.95, zipButtonsCost: 14.1, totalCost: 445.62, zipButtonType: "14 Button" },
  "Claudia Dress": { sheetCount: 24, averageMeters: 2.65, averageTrimMeters: 2, stitchingCost: 82.08, fabricCost: 316.7, zipButtonsCost: 3.75, totalCost: 402.65, zipButtonType: "2 Button" },
  "Ember Midi Dress": { sheetCount: 13, averageMeters: 2.85, stitchingCost: 85, fabricCost: 297.54, zipButtonsCost: 5.38, totalCost: 387.92, zipButtonType: "5 Button" },
  "Etta Dress": { sheetCount: 26, averageMeters: 2.39, stitchingCost: 75, fabricCost: 270.6, totalCost: 345.6 },
  "Francesca Dress": { sheetCount: 9, averageMeters: 3.89, stitchingCost: 75, fabricCost: 332.22, totalCost: 407.22 },
  "Frankie Dress": { sheetCount: 37, averageMeters: 2.9, averageTrimMeters: 3.46, stitchingCost: 80, fabricCost: 383.08, totalCost: 463.08 },
  "Harper Dress": { sheetCount: 16, averageMeters: 2.44, stitchingCost: 80, fabricCost: 298.33, zipButtonsCost: 4.63, totalCost: 382.93, zipButtonType: "2 Button" },
  "Kari Dress": { sheetCount: 13, averageMeters: 2.32, stitchingCost: 100, fabricCost: 292.75, zipButtonsCost: 35.77, totalCost: 428.25, zipButtonType: "22 Inch" },
  "Katie Dress": { sheetCount: 15, averageMeters: 3.08, stitchingCost: 100, fabricCost: 306.4, zipButtonsCost: 7.6, totalCost: 414, zipButtonType: "7 bouton" },
  "Mabel Dress": { sheetCount: 47, averageMeters: 2.08, stitchingCost: 56.91, fabricCost: 245.59, zipButtonsCost: 2.38, totalCost: 304.91, zipButtonType: "2 Button" },
  "Maddison Dress": { sheetCount: 14, averageMeters: 2.4, stitchingCost: 75, fabricCost: 316.77, zipButtonsCost: 4.36, totalCost: 396.31, zipButtonType: "2 Button" },
  "Nakita Dress": { sheetCount: 8, averageMeters: 3.66, stitchingCost: 75, fabricCost: 317.13, totalCost: 392.13 },
  "Nina Dress": { sheetCount: 15, averageMeters: 2.44, stitchingCost: 90, fabricCost: 271.27, zipButtonsCost: 28, totalCost: 389.27, zipButtonType: "14 Inch" },
  "Pippa Dress": { sheetCount: 1, averageMeters: 3.7, stitchingCost: 80, fabricCost: 354, zipButtonsCost: 12, totalCost: 446, zipButtonType: "12 Button" },
  "Promenade Dress": { sheetCount: 23, averageMeters: 3.66, stitchingCost: 95, fabricCost: 352.78, zipButtonsCost: 13, totalCost: 460.78, zipButtonType: "13 Botton" },
  "Riley Dress": { sheetCount: 16, averageMeters: 3.35, stitchingCost: 80, fabricCost: 318.47, zipButtonsCost: 12, totalCost: 410.47, zipButtonType: "12 Button" },
  "Rita Dress": { sheetCount: 3, averageMeters: 2.62, stitchingCost: 70, fabricCost: 213, zipButtonsCost: 11, totalCost: 294, zipButtonType: "11 Button" },
  "Scarlett Dress": { sheetCount: 20, averageMeters: 3.37, averageTrimMeters: 12, stitchingCost: 130, fabricCost: 393.1, zipButtonsCost: 37.55, totalCost: 560.65, zipButtonType: "22 Inch" },
  "Tully Dress": { sheetCount: 25, averageMeters: 2.46, stitchingCost: 84.8, fabricCost: 280.76, zipButtonsCost: 3.96, totalCost: 369.52, zipButtonType: "2 Button" },
  "Tulsi Dress": { sheetCount: 29, averageMeters: 2.68, averageTrimMeters: 1.29, stitchingCost: 80, fabricCost: 323.03, totalCost: 403.03 },
  "Ursula Dress": { sheetCount: 9, averageMeters: 2.76, stitchingCost: 100, fabricCost: 276.44, zipButtonsCost: 7.22, totalCost: 383.67, zipButtonType: "5 Button" },
  "Vivien Dress": { sheetCount: 40, averageMeters: 3.9, stitchingCost: 110, fabricCost: 425.85, zipButtonsCost: 22.35, totalCost: 558, zipButtonType: "12 Butten" },
  "Willow Dress": { sheetCount: 7, averageMeters: 2.35, stitchingCost: 70, fabricCost: 234.33, zipButtonsCost: 3, totalCost: 307.33, zipButtonType: "3 Button" },
  "Avery Dress": { sheetCount: 18, averageMeters: 2.59, stitchingCost: 68.33, fabricCost: 332.39, totalCost: 400.72 },
  "Briar Dress": { sheetCount: 9, averageMeters: 3.32, averageTrimMeters: 1.4, stitchingCost: 75, fabricCost: 360.89, zipButtonsCost: 2, totalCost: 437.89, zipButtonType: "2 Button" },
  "Ella Wrap Dress": { sheetCount: 15, averageMeters: 4.67, stitchingCost: 80, fabricCost: 535.33, totalCost: 615.33 },
  "Jamie Dress": { sheetCount: 21, averageMeters: 2.28, stitchingCost: 69.29, fabricCost: 264.43, totalCost: 333.71 },
  "Olivia Dress": { sheetCount: 6, averageMeters: 3.23, stitchingCost: 80, fabricCost: 340, zipButtonsCost: 2, totalCost: 422, zipButtonType: "2 Button" },
  "Savannah Dress": { sheetCount: 3, averageMeters: 3.47, stitchingCost: 75, fabricCost: 283, zipButtonsCost: 6, totalCost: 364, zipButtonType: "6 Button" },
  "Tiered Maxi Dress": { sheetCount: 33, averageMeters: 4.4, stitchingCost: 95.76, fabricCost: 430.09, zipButtonsCost: 6.52, totalCost: 532.36, zipButtonType: "4 Botton" },
  "Tiered Midi Dress": { sheetCount: 19, averageMeters: 3.55, stitchingCost: 85, fabricCost: 315.95, zipButtonsCost: 6.84, totalCost: 407.79, zipButtonType: "4 Botton" },
  "Tilda Dress": { sheetCount: 24, averageMeters: 3.06, averageTrimMeters: 1.82, stitchingCost: 89.79, fabricCost: 379.63, totalCost: 469.42 },
  "Zoey Dress": { sheetCount: 7, averageMeters: 2.51, stitchingCost: 75, fabricCost: 401, totalCost: 476 },
};

function productInfoStyles(names: string[]): ProductInfoStyle[] {
  return names.map((name) => ({
    id: slugForOption(name),
    name,
    imageUrl: PRODUCT_STYLE_IMAGES[name] ?? "",
    ...(PRODUCT_STYLE_COSTING[name] ?? {}),
  }));
}

const DEFAULT_PRODUCT_INFO: ProductInfo = {
  gridColumns: 4,
  categories: [
    {
      id: "short_sleeve_dresses",
      name: "Short Sleeve Dresses",
      styles: productInfoStyles([
        "Acacia Maxi Dress",
        "Alice Dress",
        "August Dress",
        "Billie Dress",
        "Blythe Dress",
        "Boho Tiered Maxi Dress",
        "Brixton Dress",
        "Chelsea Dress",
        "Claudia Dress",
        "Ember Midi Dress",
        "Etta Dress",
        "Francesca Dress",
        "Frankie Dress",
        "Harper Dress",
        "Kari Dress",
        "Katie Dress",
        "Mabel Dress",
        "Maddison Dress",
        "Nakita Dress",
        "Nina Dress",
        "Pippa Dress",
        "Promenade Dress",
        "Riley Dress",
        "Rita Dress",
        "Scarlett Dress",
        "Tully Dress",
        "Tulsi Dress",
        "Ursula Dress",
        "Vivien Dress",
        "Willow Dress",
      ]),
    },
    {
      id: "long_sleeve_dresses",
      name: "Long Sleeve Dresses",
      styles: productInfoStyles([
        "Avery Dress",
        "Briar Dress",
        "Ella Wrap Dress",
        "Jamie Dress",
        "Olivia Dress",
        "Savannah Dress",
        "Tiered Maxi Dress",
        "Tiered Midi Dress",
        "Tilda Dress",
        "Zoey Dress",
      ]),
    },
    {
      id: "short_sleeve_tops",
      name: "Short Sleeve Tops",
      styles: productInfoStyles([
        "Aria Top",
        "Aubrey Top",
        "Boxy Top",
        "Camille Top",
        "Chloe Top",
        "Eden Top",
        "Neesha Top",
        "Pia Top",
        "Quinn Top",
        "Shell Top",
        "Sky Blouse",
        "Sylvia Top",
        "Tillie Top",
        "Tulsi Short Sleeve Top",
        "Yasmin Top",
        "Zali Top",
      ]),
    },
    {
      id: "long_sleeve_tops",
      name: "Long Sleeve Tops",
      styles: productInfoStyles([
        "Demi Top",
        "Florence Top",
        "Isla Top",
        "Leia Top",
        "Sophia Blouse",
        "Tulsi Long Sleeve Top",
      ]),
    },
    {
      id: "mid_length_skirts",
      name: "Mid Length Skirts",
      styles: productInfoStyles([
        "Aalia Skirt",
        "Belt Loop Skirt",
        "Bridgette Skirt",
        "Reversible Skirt",
        "Ruby Skirt",
        "Zarah Skirt",
      ]),
    },
    {
      id: "long_skirts",
      name: "Long Skirts",
      styles: productInfoStyles([
        "Cora Skirt",
        "Dawn Maxi Skirt",
      ]),
    },
    {
      id: "pants",
      name: "Pants",
      styles: productInfoStyles([
        "Acapulco Pant",
        "Flower Pants",
        "Greta Pant",
        "Janis Pant",
        "Jaya Pant",
        "Nora Pant",
        "Pilot Pant",
        "Remi Pant",
        "Tenzin Pant",
        "Umbrella Pant",
        "Wide Leg Stretch Pocket Pants",
      ]),
    },
    {
      id: "corduroy",
      name: "Corduroy",
      styles: productInfoStyles([
        "Aalia Corduroy Skirt",
        "Belt Loop Corduroy Skirt",
        "Corduroy Jacket",
        "Corduroy Overalls",
        "Corduroy Pinafore",
        "Esta Corduroy Pants",
        "Jamie Corduroy Dress",
        "Nora Corduroy Pants",
        "Polly Pocket Corduroy Tunic",
      ]),
    },
  ],
};
const HIDDEN_FABRIC_SHEET_NAMES = new Set(["new fabric on order", "fabric on order"]);
const FABRIC_TOTAL_EXCLUDED_NAMES = new Set([
  "on order",
  "fabric samples under consideration",
  "new fabric on order",
  "random fabric bits and bobs",
  "fabric on order",
]);
const ALL_NAV_ITEMS = [
  { id: "dashboard", label: "Dashboard", href: "/portal?page=dashboard" },
  { id: "restock", label: "Existing Products Restock", href: "/portal" },
  { id: "fabric", label: "Fabric in stock", href: "/portal?page=fabric" },
  { id: "packing", label: "Packing Lists", href: "/portal?page=packing" },
  { id: "pricelist", label: "Price List", href: "/portal?page=pricelist" },
  { id: "productinfo", label: "Product Information", href: "/portal?page=productinfo" },
  { id: "samples", label: "Samples", href: "/portal?page=samples" },
  { id: "newproduct", label: "New Product Orders", href: "/portal?page=newproduct" },
  { id: "visionboard", label: "Vision Board", href: "/portal?page=visionboard", superadminOnly: true },
  { id: "collections", label: "Collections", href: "/portal?page=collections" },
] as const;
type NavItemId = typeof ALL_NAV_ITEMS[number]["id"];
const DEFAULT_NAV_ORDER: NavItemId[] = ["dashboard", "restock", "fabric", "packing", "pricelist", "productinfo", "samples", "newproduct", "visionboard", "collections"];
type FabricSheetData = FabricStockSheet & { originalRows?: string[][]; rowKeys?: number[]; totalCost?: number | null; error?: string };
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
  fabricStock: 170,
  delete: 104,
};

type ColumnDef = { id: string; label: string; center?: boolean };
type PortalUserRole = "superadmin" | "admin" | "user";
type PortalUser = {
  id: string;
  name: string;
  username: string;
  passwordHash: string;
  role: PortalUserRole;
  admin: boolean; // derived: true for superadmin/admin
  active: boolean;
  canLoadInventory: boolean;
  pageAccess: Record<string, boolean>;
};
type ActivePortalUser = PortalUser & { initials: string; lastSeen: number };
type RestockOption = { value: string; label: string; bg: string; color: string };
type RestockSettings = {
  statusOptions: RestockOption[];
  priorityOptions: RestockOption[];
  quantityFontSize: number;
  quantityFontColor: string;
  inventoryArrowColor: string;
};
type FabricSettings = {
  supplierOptions: RestockOption[];
  fabricTypeOptions: RestockOption[];
  tileOrder: string[];
  combinedColumnOrder: string[];
  combinedColumnWidths: Record<string, number>;
  imagesCompactedV1: boolean;
};
type TableCustomColumn = { id: string; label: string };
type TableCustomColumns = {
  restock: TableCustomColumn[];
  packing: TableCustomColumn[];
  fabric: Record<string, TableCustomColumn[]>;
};
type FabricCustomSheet = {
  gid: string;
  name: string;
  headers: string[];
  rows: string[][];
};
type ProductInfoStyle = {
  id: string;
  name: string;
  imageUrl?: string;
  averageMeters?: number;
  averageTrimMeters?: number;
  zipButtonType?: string;
  stitchingCost?: number;
  fabricCost?: number;
  zipButtonsCost?: number;
  totalCost?: number;
  sheetCount?: number;
  costingNotes?: string;
  hidden?: boolean;
};
type ProductInfoCategory = {
  id: string;
  name: string;
  styles: ProductInfoStyle[];
};
type ProductInfo = {
  gridColumns?: 3 | 4;
  categories: ProductInfoCategory[];
};
type ProductStyleCosting = Pick<
  ProductInfoStyle,
  "averageMeters" | "averageTrimMeters" | "zipButtonType" | "stitchingCost" | "fabricCost" | "zipButtonsCost" | "totalCost" | "sheetCount" | "costingNotes"
>;
type UniversalSettings = {
  primaryButtonBg: string;
  primaryButtonColor: string;
  tableTextSize: number;
  tableTextColor: string;
  headingTextSize: number;
  headingTextColor: string;
  panelTextSize: number;
  inventoryFontSize: number;
  menuBg: string;
  menuTextColor: string;
  pageBg: string;
  logoUrl: string;
};
type ShopifySearchProduct = {
  id: string;
  shop?: string;
  title: string;
  imageUrl: string | null;
  skus: string[];
  sizes: string[];
  variants: ShopifyVariantInfo[];
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
      const role = (["superadmin", "admin", "user"].includes(String(user.role)) ? user.role : "user") as PortalUserRole;
      const isAdmin = role === "superadmin" || role === "admin";
      return {
        id,
        name,
        username: name.toLowerCase(),
        passwordHash: String(user.passwordHash ?? ""),
        role,
        admin: isAdmin,
        canLoadInventory: "canLoadInventory" in user ? Boolean(user.canLoadInventory) : isAdmin,
        active: user.active !== false,
        pageAccess: (user.pageAccess && typeof user.pageAccess === "object" && !Array.isArray(user.pageAccess))
          ? user.pageAccess as Record<string, boolean>
          : {},
      };
    })
    .filter(Boolean) as PortalUser[];
}

function normalizeBooleanSetting(value: unknown) {
  return value === true;
}

function canPortalUserLoadPackingInventory(users: PortalUser[], currentUser: PortalUser | null) {
  if (users.length === 0) return true;
  return Boolean(currentUser?.canLoadInventory);
}

function slugForOption(label: string) {
  return label
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function normalizeHexColor(value: unknown, fallback: string) {
  const color = String(value ?? "").trim();
  return /^#[0-9a-f]{6}$/i.test(color) ? color : fallback;
}

function normalizeRestockOptions(
  value: unknown,
  defaults: Array<{ value: string; label: string; bg?: string; color?: string }>,
  defaultColors?: Record<string, { bg: string; color: string }>,
): RestockOption[] {
  const usingDefaults = !Array.isArray(value);
  const rawItems = usingDefaults ? defaults : value;
  const seen = new Set<string>();
  const items = rawItems
    .map((item) => {
      if (!item || typeof item !== "object") return null;
      const option = item as Record<string, unknown>;
      const label = String(option.label ?? "").trim();
      const value = slugForOption(String(option.value ?? label));
      if (!label || !value || seen.has(value)) return null;
      seen.add(value);
      const fallback = (defaultColors?.[value] ?? defaults.find((defaultOption) => defaultOption.value === value) ?? {}) as { bg?: string; color?: string };
      return {
        value,
        label,
        bg: normalizeHexColor(option.bg, fallback.bg ?? "#f3f4f6"),
        color: normalizeHexColor(option.color, fallback.color ?? "#374151"),
      };
    })
    .filter(Boolean) as RestockOption[];

  if (usingDefaults) {
    for (const defaultOption of defaults) {
      if (seen.has(defaultOption.value)) continue;
      const fallback = defaultColors?.[defaultOption.value] ?? defaultOption;
      items.push({
        value: defaultOption.value,
        label: defaultOption.label,
        bg: normalizeHexColor(defaultOption.bg, fallback.bg ?? "#f3f4f6"),
        color: normalizeHexColor(defaultOption.color, fallback.color ?? "#374151"),
      });
    }
  }

  return items.length ? items : normalizeRestockOptions(defaults, defaults, defaultColors);
}

function normalizeRestockSettings(value: unknown): RestockSettings {
  const settings = value && typeof value === "object" && !Array.isArray(value)
    ? value as Record<string, unknown>
    : {};
  const quantityFontSize = Math.min(32, Math.max(10, Number(settings.quantityFontSize) || 13));

  return {
    statusOptions: normalizeRestockOptions(settings.statusOptions, DEFAULT_STATUS_OPTIONS, DEFAULT_STATUS_COLORS),
    priorityOptions: normalizeRestockOptions(settings.priorityOptions, DEFAULT_PRIORITY_OPTIONS),
    quantityFontSize,
    quantityFontColor: normalizeHexColor(settings.quantityFontColor, "#111827"),
    inventoryArrowColor: normalizeHexColor(settings.inventoryArrowColor, "#4b5563"),
  };
}

function fabricHeaderIndex(headers: string[], pattern: RegExp) {
  return headers.findIndex((header) => pattern.test(header));
}

const FABRIC_TYPE_ALIASES: Array<{ canonical: string; aliases: string[] }> = [
  { canonical: "40x40 Printed", aliases: ["40x40", "40x40 printed"] },
  { canonical: "60x60 Printed", aliases: ["60x60", "60x60 printed"] },
  { canonical: "Plain 40x40", aliases: ["40x40 lattha", "40x40 latha", "plain 40x40"] },
];

function canonicalizeFabricType(raw: string): string {
  const trimmed = raw.trim();
  if (!trimmed) return trimmed;
  const lower = trimmed.toLowerCase();
  for (const entry of FABRIC_TYPE_ALIASES) {
    if (entry.aliases.includes(lower) || entry.canonical.toLowerCase() === lower) {
      return entry.canonical;
    }
  }
  return trimmed;
}

function ensureFabricTypeChipsForRows(options: RestockOption[], sheets: FabricStockSheet[]): RestockOption[] {
  const byLabel = new Map<string, RestockOption>(options.map((opt) => [canonicalizeFabricType(opt.label), opt]));
  let nextIndex = byLabel.size;
  const addLabel = (rawLabel: string) => {
    const canonical = canonicalizeFabricType(rawLabel.trim());
    if (!canonical || byLabel.has(canonical)) return;
    const colors = FABRIC_CHIP_COLORS[nextIndex % FABRIC_CHIP_COLORS.length];
    byLabel.set(canonical, { value: slugForOption(canonical), label: canonical, ...colors });
    nextIndex++;
  };
  for (const sheet of sheets) {
    if (!isCombinedFabricSource(sheet)) continue;
    const fabricTypeIdx = sheet.headers.findIndex((h) => /^fabric\s*type$/i.test(h.trim()) || /^type$/i.test(h.trim()));
    for (const row of sheet.rows) {
      if (fabricTypeIdx >= 0) {
        const raw = String(row[fabricTypeIdx] ?? "").trim();
        if (raw) addLabel(raw);
      } else {
        addLabel(sheet.name);
      }
    }
  }
  return [...byLabel.values()];
}

function canonicalizeFabricTypeOptions(options: RestockOption[]): RestockOption[] {
  const byLabel = new Map<string, RestockOption>();
  for (const option of options) {
    const canonicalLabel = canonicalizeFabricType(option.label);
    if (byLabel.has(canonicalLabel)) continue;
    byLabel.set(canonicalLabel, { ...option, label: canonicalLabel, value: slugForOption(canonicalLabel) });
  }
  for (const [index, alias] of FABRIC_TYPE_ALIASES.entries()) {
    if (byLabel.has(alias.canonical)) continue;
    const colors = FABRIC_CHIP_COLORS[(byLabel.size + index) % FABRIC_CHIP_COLORS.length];
    byLabel.set(alias.canonical, { value: slugForOption(alias.canonical), label: alias.canonical, ...colors });
  }
  return [...byLabel.values()];
}

function buildFabricDefaults(pattern: RegExp): RestockOption[] {
  const labels = new Set<string>();
  for (const sheet of initialFabricStockSheets) {
    const index = fabricHeaderIndex(sheet.headers, pattern);
    if (index < 0) continue;
    for (const row of sheet.rows) {
      const label = String(row[index] ?? "").trim();
      if (label && label !== "—" && label !== "-") labels.add(label);
    }
  }
  return [...labels].sort((a, b) => a.localeCompare(b)).map((label, index) => {
    const colors = FABRIC_CHIP_COLORS[index % FABRIC_CHIP_COLORS.length];
    return { value: slugForOption(label), label, ...colors };
  });
}

function normalizeFabricSettings(value: unknown): FabricSettings {
  const settings = value && typeof value === "object" && !Array.isArray(value)
    ? value as Record<string, unknown>
    : {};
  return {
    supplierOptions: normalizeRestockOptions(settings.supplierOptions, buildFabricDefaults(/^supplier$/i)),
    fabricTypeOptions: canonicalizeFabricTypeOptions(
      normalizeRestockOptions(settings.fabricTypeOptions, buildFabricDefaults(/fabric\s*type/i)),
    ),
    tileOrder: Array.isArray(settings.tileOrder)
      ? settings.tileOrder.map((item) => String(item)).filter(Boolean)
      : [],
    combinedColumnOrder: Array.isArray(settings.combinedColumnOrder)
      ? settings.combinedColumnOrder.map((item) => String(item)).filter(Boolean)
      : [],
    combinedColumnWidths: settings.combinedColumnWidths && typeof settings.combinedColumnWidths === "object"
      ? Object.fromEntries(
          Object.entries(settings.combinedColumnWidths as Record<string, unknown>)
            .map(([key, value]) => [key, Number(value)])
            .filter(([, value]) => Number.isFinite(value) && (value as number) >= 40),
        )
      : {},
    imagesCompactedV1: Boolean(settings.imagesCompactedV1),
  };
}

function normalizeUniversalSettings(value: unknown): UniversalSettings {
  const settings = value && typeof value === "object" && !Array.isArray(value)
    ? value as Record<string, unknown>
    : {};

  return {
    primaryButtonBg: normalizeHexColor(settings.primaryButtonBg, "#111827"),
    primaryButtonColor: normalizeHexColor(settings.primaryButtonColor, "#ffffff"),
    tableTextSize: Math.min(20, Math.max(10, Number(settings.tableTextSize) || 13)),
    tableTextColor: normalizeHexColor(settings.tableTextColor, "#374151"),
    headingTextSize: Math.min(34, Math.max(14, Number(settings.headingTextSize) || 24)),
    headingTextColor: normalizeHexColor(settings.headingTextColor, "#111827"),
    panelTextSize: Math.min(22, Math.max(11, Number(settings.panelTextSize) || 13)),
    inventoryFontSize: Math.min(32, Math.max(9, Number(settings.inventoryFontSize) || 13)),
    menuBg: normalizeHexColor(settings.menuBg, "#111827"),
    menuTextColor: normalizeHexColor(settings.menuTextColor, "#cbd5e1"),
    pageBg: normalizeHexColor(settings.pageBg, "#f3f4f6"),
    logoUrl: typeof settings.logoUrl === "string" ? settings.logoUrl : "",
  };
}

function normalizeProductInfo(value: unknown): ProductInfo {
  const rawCategories = value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>).categories
    : null;
  if (!Array.isArray(rawCategories)) return structuredClone(DEFAULT_PRODUCT_INFO);

  const categories = rawCategories.map((item) => {
    if (!item || typeof item !== "object") return null;
    const category = item as Record<string, unknown>;
    const id = String(category.id ?? "").trim();
    const name = String(category.name ?? "").trim();
    const styles = Array.isArray(category.styles)
      ? category.styles.map((styleItem) => {
          if (!styleItem || typeof styleItem !== "object") return null;
          const style = styleItem as Record<string, unknown>;
          const styleId = String(style.id ?? "").trim();
          const styleName = String(style.name ?? "").trim();
          const imageUrl = String(style.imageUrl ?? PRODUCT_STYLE_IMAGES[styleName] ?? "").trim();
          const defaults = PRODUCT_STYLE_COSTING[styleName] ?? {};
          return styleId && styleName ? {
            id: styleId,
            name: styleName,
            imageUrl,
            averageMeters: Number(style.averageMeters) || defaults.averageMeters,
            averageTrimMeters: Number(style.averageTrimMeters) || defaults.averageTrimMeters,
            zipButtonType: String(style.zipButtonType ?? defaults.zipButtonType ?? "").trim(),
            stitchingCost: Number(style.stitchingCost) || defaults.stitchingCost,
            fabricCost: Number(style.fabricCost) || defaults.fabricCost,
            zipButtonsCost: Number(style.zipButtonsCost) || defaults.zipButtonsCost,
            totalCost: Number(style.totalCost) || defaults.totalCost,
            sheetCount: Number(style.sheetCount) || defaults.sheetCount,
            costingNotes: String(style.costingNotes ?? defaults.costingNotes ?? "").trim(),
            hidden: style.hidden === true,
          } : null;
        }).filter(Boolean) as ProductInfoStyle[]
      : [];
    return id && name ? { id, name, styles } : null;
  }).filter(Boolean) as ProductInfoCategory[];

  const gridColumns = (value as Record<string, unknown>).gridColumns === 3 ? 3 : 4;
  return categories.length ? { gridColumns, categories } : structuredClone(DEFAULT_PRODUCT_INFO);
}

async function loadProductInfoForAction() {
  const setting = await prisma.portalSetting.findUnique({
    where: { key: PRODUCT_INFO_KEY },
    select: { value: true },
  });
  return normalizeProductInfo(setting?.value);
}

async function saveProductInfo(productInfo: ProductInfo) {
  const value = normalizeProductInfo(productInfo);
  await prisma.portalSetting.upsert({
    where: { key: PRODUCT_INFO_KEY },
    create: { key: PRODUCT_INFO_KEY, value },
    update: { value },
  });
}

function fabricCellKey(gid: string, rowIndex: number, colIndex: number) {
  return `${gid}:${rowIndex}:${colIndex}`;
}

function fabricRowKey(gid: string, rowIndex: number) {
  return `${gid}:${rowIndex}`;
}

function normalizeFabricName(name: string) {
  return name.trim().toLowerCase().replace(/\s+/g, " ");
}

function normalizeFabricHeader(header: string) {
  return header.trim().toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function fabricHeaderRole(header: string) {
  const normalized = header.trim().toLowerCase();
  if (/picture|image/.test(normalized) || normalized === "fabric" || normalized === "print") return "image";
  if (/supplier/.test(normalized)) return "supplier";
  if (/fabric\s*type|^type$/.test(normalized)) return "fabricType";
  if (/name/.test(normalized)) return "name";
  if (/planned.*date|date\s*ordered|order\s*date/.test(normalized)) return "date";
  if (/(quantity\s*ordered|meters?\s*in\s*stock|meters?\s*available|^meters?$|^qty$)/.test(normalized)) return "quantity";
  if (/quantity\s*received/.test(normalized)) return "received";
  if (/status/.test(normalized)) return "status";
  if (/eta/.test(normalized)) return "eta";
  if (/cost|price/.test(normalized)) return "cost";
  return normalizeFabricHeader(header);
}

function isLockedFabricCalculationHeader(header: string) {
  const normalized = header.trim().toLowerCase();
  return /cost\s*per\s*meter|price\s*per\s*meter|cost\/meter|price\/meter/.test(normalized)
    || /meters?\s*in\s*stock|meters?\s*available/.test(normalized);
}

function isHiddenFabricSheet(name: string) {
  return HIDDEN_FABRIC_SHEET_NAMES.has(normalizeFabricName(name));
}

const COMBINED_FABRIC_ON_ORDER_GID = "759049382";

function isCombinedFabricSource(sheet: { gid: string; kind: string; name: string }) {
  if (isHiddenFabricSheet(sheet.name)) return false;
  if (sheet.gid === COMBINED_FABRIC_ON_ORDER_GID) return true;
  return sheet.kind === "stock" || sheet.kind === "simple-stock";
}

// Columns we want available for editing on every row in the combined view,
// even when a source sheet doesn't include them. Order matters — appended
// columns become stable colIndexes for the override store, so don't reorder
// or remove entries; only add new ones at the end.
const COMBINED_FABRIC_PAD_COLUMNS: Array<{ header: string; regex: RegExp }> = [
  { header: "Collection", regex: /^collection$/ },
  { header: "Cost per Meter", regex: /cost\s*per\s*meter|price\s*per\s*meter|^price$/ },
  { header: "Cut Pieces", regex: /^cut\s*pieces?$/ },
  { header: "Received / Date", regex: /received|^order\s*date$/ },
  { header: "Products", regex: /^products?$/ },
  { header: "Order Date", regex: /^order\s*date$/ },
  { header: "Meters in Stock", regex: /meters?\s*in\s*stock|meters?\s*available|^meters?$/ },
  { header: "Quantity Ordered", regex: /quantity\s*ordered/ },
];

function padCombinedFabricSheet(sheet: FabricSheetData): FabricSheetData {
  const existing = sheet.headers.map((h) => h.trim().toLowerCase());
  const missing: string[] = [];
  for (const col of COMBINED_FABRIC_PAD_COLUMNS) {
    if (existing.some((h) => col.regex.test(h))) continue;
    missing.push(col.header);
  }
  if (!missing.length) return sheet;
  const newLength = sheet.headers.length + missing.length;
  // Pad only as needed: if a row already has a value at the pad position (because it
  // was saved into a pad column before the column was persisted into headers), keep it.
  const padRow = (row: string[]) => {
    if (row.length >= newLength) return row;
    const result = [...row];
    while (result.length < newLength) result.push("");
    return result;
  };
  return {
    ...sheet,
    headers: [...sheet.headers, ...missing],
    rows: sheet.rows.map(padRow),
    originalRows: sheet.originalRows?.map(padRow),
  };
}

function isFabricTotalsExcluded(name: string) {
  return FABRIC_TOTAL_EXCLUDED_NAMES.has(normalizeFabricName(name));
}

function shouldRestoreLostPeacockRow(gid: string, rowIndex: number) {
  return gid === "1829736341" && (rowIndex === 1 || rowIndex === 2);
}

function fabricCellByRole(sheet: FabricStockSheet, row: string[], role: string) {
  const index = sheet.headers.findIndex((header) => fabricHeaderRole(header) === role);
  return index < 0 ? "" : String(row[index] ?? "").trim();
}

function fabricMoveMatchKey(sheet: FabricStockSheet, row: string[]) {
  const supplier = fabricCellByRole(sheet, row, "supplier").toLowerCase();
  const fabricType = fabricCellByRole(sheet, row, "fabricType").toLowerCase();
  if (!supplier || !fabricType) return "";
  return `${supplier}|${fabricType}`;
}

function isBrokenMovedOnOrderRow(sheet: FabricStockSheet, row: string[], rowIndex: number) {
  if (sheet.gid !== "759049382" || rowIndex < sheet.rows.length) return false;
  const pictureIndex = sheet.headers.findIndex((header) => fabricHeaderRole(header) === "image");
  const quantityIndex = sheet.headers.findIndex((header) => fabricHeaderRole(header) === "quantity");
  const supplierIndex = sheet.headers.findIndex((header) => fabricHeaderRole(header) === "supplier");
  const fabricTypeIndex = sheet.headers.findIndex((header) => fabricHeaderRole(header) === "fabricType");
  const hasSupplierAndType = Boolean(row[supplierIndex]?.trim()) && Boolean(row[fabricTypeIndex]?.trim());
  const missingMovedFields = !row[pictureIndex]?.trim() && !row[quantityIndex]?.trim();
  return hasSupplierAndType && missingMovedFields && row.slice(2).every((cell) => !String(cell ?? "").trim());
}

function recoveredBrokenMoveRows(customRows: Record<string, string[][]>, deletedRows: Record<string, boolean>) {
  const recovered = new Set<string>();
  const onOrderSheet = initialFabricStockSheets.find((sheet) => sheet.gid === "759049382");
  if (!onOrderSheet) return recovered;

  const brokenMoveKeys = new Set(
    (customRows[onOrderSheet.gid] ?? [])
      .map((row, index) => ({ row, rowIndex: onOrderSheet.rows.length + index }))
      .filter(({ row, rowIndex }) => isBrokenMovedOnOrderRow(onOrderSheet, row, rowIndex))
      .map(({ row }) => fabricMoveMatchKey(onOrderSheet, row))
      .filter(Boolean),
  );
  if (!brokenMoveKeys.size) return recovered;

  for (const sheet of initialFabricStockSheets) {
    if (sheet.gid === onOrderSheet.gid) continue;
    for (const [rowIndex, row] of sheet.rows.entries()) {
      const key = fabricRowKey(sheet.gid, rowIndex);
      if (!deletedRows[key]) continue;
      if (brokenMoveKeys.has(fabricMoveMatchKey(sheet, row))) recovered.add(key);
    }
  }

  return recovered;
}

function normalizeFabricCellOverrides(value: unknown): Record<string, string> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  return Object.fromEntries(
    Object.entries(value)
      .filter(([key, cellValue]) => /^\S+:\d+:\d+$/.test(key) && typeof cellValue === "string"),
  ) as Record<string, string>;
}

function normalizeFabricDeletedRows(value: unknown): Record<string, boolean> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  return Object.fromEntries(
    Object.entries(value)
      .filter(([key, rowValue]) => /^\S+:\d+$/.test(key) && rowValue === true),
  ) as Record<string, boolean>;
}

function normalizeFabricCustomRows(value: unknown): Record<string, string[][]> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  return Object.fromEntries(
    Object.entries(value)
      .map(([gid, rows]) => [
        gid,
        Array.isArray(rows)
          ? rows
              .filter((row): row is unknown[] => Array.isArray(row))
              .map((row) => row.map((cell) => String(cell ?? "")))
          : [],
      ] as const)
      .filter(([gid]) => Boolean(gid)),
  );
}

function sumFabricRows(headers: string[], rows: string[][]) {
  const quantityIndexes = headers
    .map((header, index) => ({ header: header.toLowerCase(), index }))
    .filter(({ header }) => /(meter|qty|quantity|ordered|stock|available)/.test(header) && !/(date|cost|price)/.test(header))
    .map(({ index }) => index);
  if (!quantityIndexes.length) return rows.length || null;
  const total = rows.reduce((sum, row) => (
    sum + quantityIndexes.reduce((cellSum, index) => cellSum + (Number(String(row[index] ?? "").replace(/,/g, "")) || 0), 0)
  ), 0);
  return total > 0 ? total : rows.length || null;
}

function sumFabricCost(headers: string[], rows: string[][]) {
  const costIndex = headers.findIndex((header) => /(cost|price)/i.test(header));
  const quantityIndex = headers.findIndex((header) => (
    /(meters?\s+in\s+stock|meters?\s+available|quantity\s+ordered|meters?|stock|available|ordered|qty)/i.test(header) &&
    !/(date|cost|price|cut)/i.test(header)
  ));
  if (costIndex < 0 || quantityIndex < 0) return null;
  const total = rows.reduce((sum, row) => {
    const cost = Number(String(row[costIndex] ?? "").replace(/[^0-9.-]/g, "")) || 0;
    const quantity = Number(String(row[quantityIndex] ?? "").replace(/[^0-9.-]/g, "")) || 0;
    return sum + (cost * quantity);
  }, 0);
  return total > 0 ? total : null;
}

function orderFabricSheets(sheets: FabricSheetData[], tileOrder: string[] = []) {
  if (!tileOrder.length) return sheets;
  const order = new Map(tileOrder.map((gid, index) => [gid, index]));
  return [...sheets].sort((a, b) => {
    const ai = order.get(a.gid);
    const bi = order.get(b.gid);
    if (ai == null && bi == null) return 0;
    if (ai == null) return 1;
    if (bi == null) return -1;
    return ai - bi;
  });
}

function getFabricSheetSummaries(sheets: FabricStockSheet[], tileOrder: string[] = []): FabricSheetData[] {
  return orderFabricSheets(
    sheets
      .filter((sheet) => !isHiddenFabricSheet(sheet.name))
      .map((sheet) => ({
        ...sheet,
        rows: [],
        rowCount: sheet.rows.length,
        totalQuantity: sumFabricRows(sheet.headers, sheet.rows),
        totalCost: sumFabricCost(sheet.headers, sheet.rows),
      })),
    tileOrder,
  );
}

function toManualFabricSheet(sheet: FabricStockSheet | FabricCustomSheet): FabricStockSheet {
  return {
    gid: sheet.gid,
    name: sheet.name,
    kind: "kind" in sheet ? sheet.kind : "stock",
    headers: sheet.headers.map((header) => String(header || "Column")),
    rows: sheet.rows.map((row) => row.map((cell) => String(cell ?? ""))),
    rowCount: sheet.rows.length,
    totalQuantity: sumFabricRows(sheet.headers, sheet.rows),
  };
}

function normalizeManualFabricSheets(value: unknown): FabricStockSheet[] {
  if (!Array.isArray(value)) return [];
  return value.map((item) => {
    if (!item || typeof item !== "object") return null;
    const sheet = item as Record<string, unknown>;
    const gid = String(sheet.gid ?? "").trim();
    const name = String(sheet.name ?? "").trim();
    const headers = Array.isArray(sheet.headers)
      ? sheet.headers.map((header) => String(header || "Column"))
      : [];
    const rows = Array.isArray(sheet.rows)
      ? sheet.rows.filter((row): row is unknown[] => Array.isArray(row)).map((row) => row.map((cell) => String(cell ?? "")))
      : [];
    if (!gid || !name || !headers.length) return null;
    return {
      gid,
      name,
      kind: String(sheet.kind ?? "stock"),
      headers,
      rows,
      rowCount: rows.length,
      totalQuantity: sumFabricRows(headers, rows),
    };
  }).filter(Boolean) as FabricStockSheet[];
}

async function saveManualFabricSheets(sheets: FabricStockSheet[]) {
  const value = sheets.map(toManualFabricSheet);
  await prisma.portalSetting.upsert({
    where: { key: FABRIC_MANUAL_SHEETS_KEY },
    create: { key: FABRIC_MANUAL_SHEETS_KEY, value },
    update: { value },
  });
}

// Build an index of fabric stock entries: one entry per row in any stock sheet.
// Used by the restock page to show fabric availability per product based on the
// fabric name appearing in the product title. Entries are kept separate (not summed)
// so the same fabric name appearing in multiple sheets shows each individually.
type FabricStockEntry = { name: string; meters: number; sheetName: string; kind: "stock" | "order" };

function buildFabricStockIndex(sheets: FabricStockSheet[]): FabricStockEntry[] {
  const out: FabricStockEntry[] = [];
  for (const sheet of sheets) {
    const isStock = sheet.kind === "stock";
    const isOrder = sheet.kind === "order";
    if (!isStock && !isOrder) continue;
    const nameIdx = sheet.headers.findIndex((h) => /^name$/i.test(h));
    const metersIdx = isStock
      ? sheet.headers.findIndex((h) => /meters?\s*in\s*stock/i.test(h))
      : sheet.headers.findIndex((h) => /quantity\s*ordered|meters?\s*ordered/i.test(h));
    if (nameIdx < 0 || metersIdx < 0) continue;
    for (const row of sheet.rows) {
      const name = (row[nameIdx] ?? "").trim();
      if (!name || name.length < 2) continue;
      const cleaned = (row[metersIdx] ?? "").toString().split(/[^0-9.]/)[0];
      const m = Number(cleaned);
      if (!Number.isFinite(m)) continue;
      // Skip zero on-order rows so we don't clutter the cell with empty entries
      if (isOrder && m === 0) continue;
      out.push({ name, meters: m, sheetName: sheet.name, kind: isStock ? "stock" : "order" });
    }
  }
  return out;
}

// Find every stock entry whose fabric name appears as a whole word in the product title.
// Multiple sheets may have the same fabric name (e.g. Morialta in 60x60 and Voil) — return
// them all so the UI can show them separately rather than misleadingly summing.
function findFabricStockMatches(title: string, index: FabricStockEntry[]): FabricStockEntry[] {
  if (!title) return [];
  const lower = title.toLowerCase();
  const names = Array.from(new Set(index.map((i) => i.name))).sort((a, b) => b.length - a.length);
  for (const name of names) {
    if (name.length < 3) continue;
    const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const re = new RegExp(`\\b${escaped}\\b`, "i");
    if (re.test(lower)) {
      return index.filter((i) => i.name.toLowerCase() === name.toLowerCase());
    }
  }
  return [];
}

async function getManualFabricSheets({
  savedValue,
  customSheetsValue,
  customRowsValue,
  deletedRowsValue,
  overridesValue,
  deletedSheetsValue,
}: {
  savedValue?: unknown;
  customSheetsValue?: unknown;
  customRowsValue?: unknown;
  deletedRowsValue?: unknown;
  overridesValue?: unknown;
  deletedSheetsValue?: unknown;
}) {
  const savedSheets = normalizeManualFabricSheets(savedValue);
  if (savedSheets.length) return savedSheets;

  const legacySheets = getFabricSheets(
    overridesValue,
    customRowsValue,
    deletedRowsValue,
    [],
    normalizeFabricCustomSheets(customSheetsValue),
    {},
    normalizeFabricDeletedSheets(deletedSheetsValue),
    initialFabricStockSheets,
    { includeHidden: true },
  ).map(toManualFabricSheet);
  await saveManualFabricSheets(legacySheets);
  return legacySheets;
}

async function loadManualFabricSheetsForAction() {
  const [
    manualSheetsSetting,
    customSheetsSetting,
    customRowsSetting,
    deletedRowsSetting,
    overridesSetting,
    deletedSheetsSetting,
  ] = await Promise.all([
    prisma.portalSetting.findUnique({ where: { key: FABRIC_MANUAL_SHEETS_KEY }, select: { value: true } }),
    prisma.portalSetting.findUnique({ where: { key: FABRIC_CUSTOM_SHEETS_KEY }, select: { value: true } }),
    prisma.portalSetting.findUnique({ where: { key: FABRIC_CUSTOM_ROWS_KEY }, select: { value: true } }),
    prisma.portalSetting.findUnique({ where: { key: FABRIC_DELETED_ROWS_KEY }, select: { value: true } }),
    prisma.portalSetting.findUnique({ where: { key: FABRIC_CELL_OVERRIDES_KEY }, select: { value: true } }),
    prisma.portalSetting.findUnique({ where: { key: FABRIC_DELETED_SHEETS_KEY }, select: { value: true } }),
  ]);
  return getManualFabricSheets({
    savedValue: manualSheetsSetting?.value,
    customSheetsValue: customSheetsSetting?.value,
    customRowsValue: customRowsSetting?.value,
    deletedRowsValue: deletedRowsSetting?.value,
    overridesValue: overridesSetting?.value,
    deletedSheetsValue: deletedSheetsSetting?.value,
  });
}

async function getInrToAudRate() {
  const cached = getInrToAudRate as typeof getInrToAudRate & { cachedRate?: number; cachedAt?: number };
  if (cached.cachedRate && cached.cachedAt && Date.now() - cached.cachedAt < 3 * 60 * 60 * 1000) {
    return cached.cachedRate;
  }
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 700);
  try {
    const response = await fetch("https://open.er-api.com/v6/latest/INR", {
      headers: { Accept: "application/json" },
      signal: controller.signal,
    });
    if (!response.ok) return null;
    const data = await response.json() as { rates?: Record<string, number> };
    const rate = Number(data.rates?.AUD);
    if (!Number.isFinite(rate) || rate <= 0) return null;
    cached.cachedRate = rate;
    cached.cachedAt = Date.now();
    return rate;
  } catch {
    return null;
  } finally {
    clearTimeout(timeout);
  }
}

function getFabricSheets(
  overridesValue?: unknown,
  customRowsValue?: unknown,
  deletedRowsValue?: unknown,
  tileOrder: string[] = [],
  customSheets: FabricCustomSheet[] = [],
  customColumns: Record<string, TableCustomColumn[]> = {},
  deletedSheets: Record<string, boolean> = {},
  baseSheets: FabricStockSheet[] = initialFabricStockSheets,
  options: { includeHidden?: boolean } = {},
): FabricSheetData[] {
  const overrides = normalizeFabricCellOverrides(overridesValue);
  const customRows = normalizeFabricCustomRows(customRowsValue);
  const deletedRows = normalizeFabricDeletedRows(deletedRowsValue);
  const recoveredRows = recoveredBrokenMoveRows(customRows, deletedRows);
  const sourceSheets: FabricStockSheet[] = [
    ...baseSheets.filter((sheet) => (options.includeHidden || !isHiddenFabricSheet(sheet.name)) && !deletedSheets[sheet.gid]),
    ...customSheets.map((sheet) => ({
      ...sheet,
      kind: "stock",
      rowCount: sheet.rows.length,
      totalQuantity: sumFabricRows(sheet.headers, sheet.rows),
    })).filter((sheet) => (options.includeHidden || !isHiddenFabricSheet(sheet.name)) && !deletedSheets[sheet.gid]),
  ];
  const sheets = sourceSheets.map((sheet) => {
    const extraHeaders = (customColumns[sheet.gid] ?? []).map((column) => column.label);
    const headers = [...sheet.headers, ...extraHeaders];
    const originalRows = [...sheet.rows, ...(customRows[sheet.gid] ?? [])];
    const rowEntries = originalRows
      .map((row, rowIndex) => ({ row, rowIndex }))
      .filter(({ row, rowIndex }) => !isBrokenMovedOnOrderRow(sheet, row, rowIndex))
      .filter(({ rowIndex }) => {
        const rowKey = fabricRowKey(sheet.gid, rowIndex);
        return !deletedRows[rowKey] || recoveredRows.has(rowKey) || shouldRestoreLostPeacockRow(sheet.gid, rowIndex);
      });
    const rows = rowEntries.map(({ row, rowIndex }) => {
      // Preserve cells beyond headers.length so values written into not-yet-persisted
      // pad columns (e.g. "Quantity Ordered", "Order Date") aren't dropped on reload.
      const length = Math.max(headers.length, row.length);
      return Array.from({ length }, (_, colIndex) => overrides[fabricCellKey(sheet.gid, rowIndex, colIndex)] ?? row[colIndex] ?? "");
    });
    return {
      ...sheet,
      headers,
      originalRows,
      rowKeys: rowEntries.map(({ rowIndex }) => rowIndex),
      rows,
      rowCount: rows.length,
      totalQuantity: sumFabricRows(sheet.headers, rows),
      totalCost: sumFabricCost(sheet.headers, rows),
    };
  });
  return orderFabricSheets(sheets, tileOrder);
}

// Replace any `data:image/...;base64,...` cell values in the fabric sheets
// with a URL pointing at the /portal/fabric-image route. The route reads the
// single cell from the stored JSONB blob and streams it as binary, so the
// loader response itself no longer carries multi-megabyte base64 strings.
// The cache-buster comes from the blob's PortalSetting.updatedAt — any edit
// to any cell mints fresh URLs that browsers will refetch.
function replaceFabricImagesWithUrls(sheets: FabricSheetData[], version: number): FabricSheetData[] {
  const buildUrl = (gid: string, row: number, col: number) =>
    `/portal/fabric-image/${encodeURIComponent(gid)}/${row}/${col}.png?v=${version}`;
  return sheets.map((sheet) => {
    const swap = (cell: unknown, row: number, col: number): string => {
      if (typeof cell !== "string") return String(cell ?? "");
      return /^data:image\//i.test(cell) ? buildUrl(sheet.gid, row, col) : cell;
    };
    const rows = sheet.rows.map((row, displayIdx) => {
      const sourceIdx = sheet.rowKeys?.[displayIdx] ?? displayIdx;
      return row.map((cell, cIdx) => swap(cell, sourceIdx, cIdx));
    });
    const originalRows = sheet.originalRows
      ? sheet.originalRows.map((row, sourceIdx) => row.map((cell, cIdx) => swap(cell, sourceIdx, cIdx)))
      : sheet.originalRows;
    return { ...sheet, rows, originalRows };
  });
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

async function ensureSuperAdmin(users: PortalUser[]) {
  if (users.some((u) => u.role === "superadmin")) return users;
  const passwordHash = await bcrypt.hash("Koku", 10);
  const seed: PortalUser = {
    id: crypto.randomUUID(),
    name: "Koku",
    username: "koku",
    passwordHash,
    role: "superadmin",
    admin: true,
    active: true,
    canLoadInventory: true,
    pageAccess: {},
  };
  const next = [...users, seed];
  await savePortalUsers(next);
  return next;
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

function packingListBoxCount(list: PackingListWithLines | null) {
  if (!list) return 0;
  const boxes = new Set<string>();
  for (const line of list.lines) {
    const box = (line.boxNumber ?? "").trim();
    if (box) boxes.add(box);
  }
  return boxes.size;
}

function packingLineMatchesSearch(line: PackingListWithLines["lines"][number], search: string) {
  const qtys = normalizeQtys(line.qtys);
  const searchable = [
    line.boxNumber,
    line.productTitle,
    line.sku,
    line.priceRupees,
    line.weight,
    packingTotal(qtys),
    ...Object.entries(qtys).flatMap(([size, qty]) => [size, qty]),
  ]
    .filter((value) => value !== null && value !== undefined && value !== "")
    .join(" ")
    .toLowerCase();
  return searchable.includes(search);
}

function csvCell(value: unknown) {
  const text = String(value ?? "");
  return /[",\n\r]/.test(text) ? `"${text.replace(/"/g, "\"\"")}"` : text;
}

type SupplierPackingLine = {
  boxNumber: string;
  styleName: string;
  supplierCode: string;
  colorName: string;
  productTitle: string;
  qtys: Record<string, number>;
};

function parseCsvRows(text: string) {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (char === "\"") {
      if (quoted && next === "\"") {
        cell += "\"";
        index += 1;
      } else {
        quoted = !quoted;
      }
      continue;
    }
    if (char === "," && !quoted) {
      row.push(cell.trim());
      cell = "";
      continue;
    }
    if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(cell.trim());
      if (row.some(Boolean)) rows.push(row);
      row = [];
      cell = "";
      continue;
    }
    cell += char;
  }

  row.push(cell.trim());
  if (row.some(Boolean)) rows.push(row);
  return rows;
}

function humanizeSupplierColour(value: string) {
  return value
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeSupplierText(value: string) {
  return humanizeSupplierColour(value)
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeSupplierSize(value: string) {
  const normalized = value.trim().toUpperCase().replace(/\s+/g, "").replace(/-/g, "/");
  const sizeAliases: Record<string, string> = {
    SM: "S/M",
    ML: "M/L",
    LXL: "L/XL",
    "1XL": "XL",
    XXL: "2XL",
    XXXL: "3XL",
  };
  return sizeAliases[normalized] ?? normalized;
}

function parseSupplierDate(value: string) {
  const match = value.match(/(\d{1,2})\s+([a-z]{3,})\s+(\d{4})/i);
  if (!match) return null;
  const months: Record<string, number> = {
    jan: 0,
    feb: 1,
    mar: 2,
    apr: 3,
    may: 4,
    jun: 5,
    jul: 6,
    aug: 7,
    sep: 8,
    oct: 9,
    nov: 10,
    dec: 11,
  };
  const month = months[match[2].slice(0, 3).toLowerCase()];
  if (month === undefined) return null;
  return new Date(Number(match[3]), month, Number(match[1]));
}

function parseSupplierPackingCsv(text: string) {
  const rows = parseCsvRows(text);
  let currentBox = "";
  let invoiceNumber: string | null = null;
  let shipmentDate: Date | null = null;
  const grouped = new Map<string, SupplierPackingLine>();

  for (const row of rows) {
    const joined = row.filter(Boolean).join(" ");
    const invoiceMatch = joined.match(/INV\s*NO\.?\s*(.+?)(?:\s{2,}|$)/i);
    if (invoiceMatch && !invoiceNumber) invoiceNumber = invoiceMatch[1].replace(/\s+/g, " ").trim();
    if (/DATE\s*:/i.test(joined) && !shipmentDate) shipmentDate = parseSupplierDate(joined);

    const boxMatch = joined.match(/\bBox\s*NO\.?\s*([A-Za-z0-9-]+)/i);
    if (boxMatch) {
      currentBox = boxMatch[1];
      continue;
    }

    const styleName = row[1]?.trim() ?? "";
    const supplierCode = row[2]?.trim() ?? "";
    const rawSize = row[3]?.trim() ?? "";
    const rawColor = row[4]?.trim() ?? "";
    const quantity = Number(row[5] ?? 0);
    const size = normalizeSupplierSize(rawSize);
    if (!styleName || !rawColor || !PACKING_SIZES.includes(size) || !Number.isFinite(quantity) || quantity <= 0) continue;

    const colorName = humanizeSupplierColour(rawColor);
    const key = `${currentBox}::${normalizeSupplierText(styleName)}::${normalizeSupplierText(colorName)}`;
    const existing = grouped.get(key) ?? {
      boxNumber: currentBox,
      styleName,
      supplierCode,
      colorName,
      productTitle: `${styleName} ${colorName}`,
      qtys: {},
    };
    existing.qtys[size] = (existing.qtys[size] ?? 0) + quantity;
    grouped.set(key, existing);
  }

  const title = invoiceNumber ? `Packing list ${invoiceNumber}` : `Packing list ${formatPortalDate(new Date())}`;
  return { title, invoiceNumber, shipmentDate, lines: Array.from(grouped.values()) };
}

function scoreSupplierProductMatch(product: ShopifySearchProduct, styleName: string, colorName: string) {
  const title = normalizeSupplierText(product.title);
  const style = normalizeSupplierText(styleName);
  const color = normalizeSupplierText(colorName);
  const styleWords = style.split(" ").filter(Boolean);
  const colorWords = color.split(" ").filter(Boolean);
  let score = 0;
  if (title.includes(style)) score += 4;
  if (title.includes(color)) score += 4;
  if (styleWords.every((word) => title.includes(word))) score += 2;
  if (colorWords.every((word) => title.includes(word))) score += 2;
  return score;
}

async function findShopifyProductForSupplierLine(
  styleName: string,
  colorName: string,
  cache: Map<string, ShopifySearchProduct | null>,
) {
  const query = `${styleName} ${humanizeSupplierColour(colorName)}`.trim();
  const cacheKey = normalizeSupplierText(query);
  if (cache.has(cacheKey)) return cache.get(cacheKey) ?? null;

  const results = await searchShopifyProducts(query);
  const best = results
    .map((product) => ({ product, score: scoreSupplierProductMatch(product, styleName, colorName) }))
    .sort((a, b) => b.score - a.score)[0];
  const matched = best && best.score >= 6 ? best.product : null;
  cache.set(cacheKey, matched);
  return matched;
}

async function searchShopifyProducts(query: string): Promise<ShopifySearchProduct[]> {
  const trimmedQuery = query.trim();
  if (trimmedQuery.length < 2) return [];

  const session = await retryAsync(() => prisma.session.findFirst({
      where: {
        accessToken: { not: "" },
      },
      orderBy: { isOnline: "asc" },
    }),
    "Shopify product search session",
  ).catch((error) => {
    console.error("Shopify product search session failed", error);
    return null;
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
            variants(first: 100) {
              edges {
                node {
                  id
                  title
                  sku
                  selectedOptions { name value }
                }
              }
            }
          }
        }
      }
    }
  `;

  const mapProducts = (json: any): ShopifySearchProduct[] => (json?.data?.products?.edges ?? []).map((edge: any) => {
    const rawVariants: any[] = (edge.node.variants?.edges ?? []).map((variantEdge: any) => variantEdge.node);

    const seen = new Set<string>();
    const sizeVariants = rawVariants
      .map((v: any) => ({ raw: v, sizeValue: extractShopifySizeLabel(v.selectedOptions, rawVariants.length) }))
      .filter(({ sizeValue }) => {
        if (!sizeValue || seen.has(sizeValue)) return false;
        seen.add(sizeValue);
        return true;
      })
      .map(({ raw: v, sizeValue }) => ({
        id: String(v.id ?? ""),
        title: sizeValue as string,
        sku: v.sku ? String(v.sku) : null,
        availableInventory: null,
      }))
      .filter((v: ShopifyVariantInfo) => v.id && v.title);

    return {
      id: edge.node.id,
      shop: session.shop,
      title: edge.node.title,
      imageUrl: edge.node.featuredImage?.url ?? null,
      skus: Array.from(new Set(sizeVariants.map((v: ShopifyVariantInfo) => v.sku).filter(Boolean))),
      sizes: sizeVariants.map((v: ShopifyVariantInfo) => v.title),
      variants: sizeVariants,
    };
  });

  const needle = trimmedQuery.toLowerCase();
  const matchesLocally = (product: ShopifySearchProduct) =>
    product.title.toLowerCase().includes(needle)
    || product.skus.some((sku) => sku.toLowerCase().includes(needle));

  const shopifyQuery = `title:*${escapedQuery}* OR sku:*${escapedQuery}*`;

  const directFetch = async (q: string | null) => {
    try {
      const response = await fetch(`https://${session.shop}/admin/api/2025-10/graphql.json`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "X-Shopify-Access-Token": session.accessToken,
        },
        body: JSON.stringify({ query: graphqlQuery, variables: { query: q } }),
      });
      if (!response.ok) return [];
      return mapProducts(await response.json());
    } catch {
      return [];
    }
  };

  try {
    const { admin } = await unauthenticated.admin(session.shop);
    const response = await admin.graphql(graphqlQuery, { variables: { query: shopifyQuery } });
    const products = mapProducts(await response.json());
    if (products.length) return products;

    const fallbackResponse = await admin.graphql(graphqlQuery, { variables: { query: null } });
    return mapProducts(await fallbackResponse.json()).filter(matchesLocally).slice(0, 8);
  } catch (error) {
    console.error("Shopify product search failed", error);
    const products = await directFetch(shopifyQuery);
    if (products.length) return products;
    return (await directFetch(null)).filter(matchesLocally).slice(0, 8);
  }
}

type ShopifyVariantInfo = { id: string; title: string; sku: string | null; availableInventory: number | null };
type ShopifyInventoryVariantInfo = ShopifyVariantInfo & { inventoryItemId: string | null };
type ShopifyInventoryChange = { size: string; qty: number; inventoryItemId: string };

function shopifyVariantAvailableInventory(variant: any): number | null {
  const inventoryLevels = variant.inventoryItem?.inventoryLevels?.nodes ?? [];
  let totalAvailable = 0;
  let hasAvailableQuantity = false;

  for (const level of inventoryLevels) {
    for (const quantity of level.quantities ?? []) {
      if (quantity?.name !== "available") continue;
      const available = Number(quantity.quantity);
      if (!Number.isFinite(available)) continue;
      totalAvailable += available;
      hasAvailableQuantity = true;
    }
  }

  return hasAvailableQuantity ? totalAvailable : null;
}

async function getShopifyProductVariants(shop: string, productId: string): Promise<ShopifyVariantInfo[]> {
  const session = await retryAsync(() => prisma.session.findFirst({
      where: { shop, accessToken: { not: "" } },
      orderBy: { isOnline: "asc" },
    }),
    "Shopify variant session",
  ).catch((error) => {
    console.error("Shopify variant session failed", error);
    return null;
  });
  if (!session) return [];

  const graphqlQuery = `
    query ProductVariants($id: ID!) {
      product(id: $id) {
        variants(first: 100) {
          nodes {
            id
            title
            sku
            inventoryItem {
              inventoryLevels(first: 20) {
                nodes {
                  quantities(names: ["available"]) {
                    name
                    quantity
                  }
                }
              }
            }
          }
        }
      }
    }
  `;
  const mapVariants = (json: any): ShopifyVariantInfo[] =>
    (json.data?.product?.variants?.nodes ?? [])
      .map((variant: any) => ({
        id: String(variant.id ?? ""),
        title: String(variant.title ?? ""),
        sku: variant.sku ? String(variant.sku) : null,
        availableInventory: shopifyVariantAvailableInventory(variant),
      }))
      .filter((variant: ShopifyVariantInfo) => variant.id && variant.title);

  const directFetch = async () => {
    try {
      const response = await fetch(`https://${session.shop}/admin/api/2025-10/graphql.json`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "X-Shopify-Access-Token": session.accessToken,
        },
        body: JSON.stringify({ query: graphqlQuery, variables: { id: productId } }),
      });
      if (!response.ok) return [];
      return mapVariants(await response.json());
    } catch {
      return [];
    }
  };

  try {
    const { admin } = await unauthenticated.admin(session.shop);
    const response = await admin.graphql(graphqlQuery, { variables: { id: productId } });
    const variants = mapVariants(await response.json());
    if (variants.length) return variants;
    return directFetch();
  } catch {
    return directFetch();
  }
}

function normalizeVariantSizeLabel(value: string) {
  return value.trim().toLowerCase().replace(/\s+/g, "").replace(/-/g, "/");
}

// Derive a packing-list-friendly size label for a Shopify variant.
// Returns the explicit "Size" option value when one exists (so XL-2XL, Free,
// or any custom size label is preserved). When a product has exactly one
// variant with no Size option, returns the "Free Size" sentinel so the
// packing list can still target it. Multi-option products without a Size
// option (e.g. Color-only) are deliberately skipped to avoid pushing
// inventory to an ambiguous variant.
function extractShopifySizeLabel(
  selectedOptions: { name?: string | null; value?: string | null }[] | undefined | null,
  totalVariantCount: number,
): string | null {
  const sizeOption = (selectedOptions ?? []).find(
    (option) => (option?.name ?? "").trim().toLowerCase() === "size",
  );
  if (sizeOption?.value && sizeOption.value.trim()) return sizeOption.value.trim();
  if (totalVariantCount === 1) return "Free Size";
  return null;
}

function matchingVariantForSize(variants: ShopifyVariantInfo[], size: string) {
  const normalizedSize = normalizeVariantSizeLabel(size);
  return variants.find((variant) => normalizeVariantSizeLabel(variant.title) === normalizedSize) ?? null;
}

async function shopifyGraphql<T>(
  shop: string,
  accessToken: string,
  query: string,
  variables?: Record<string, unknown>,
): Promise<T | null> {
  try {
    const response = await fetch(`https://${shop}/admin/api/2025-10/graphql.json`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-Shopify-Access-Token": accessToken,
      },
      body: JSON.stringify({ query, variables }),
    });
    if (!response.ok) return null;
    return await response.json() as T;
  } catch {
    return null;
  }
}

async function getShopifyInventoryVariants(shop: string, productId: string): Promise<ShopifyInventoryVariantInfo[]> {
  const session = await retryAsync(() => prisma.session.findFirst({
      where: { shop, accessToken: { not: "" } },
      orderBy: { isOnline: "asc" },
    }),
    "Shopify inventory session",
  ).catch((error) => {
    console.error("Shopify inventory session failed", error);
    return null;
  });
  if (!session) return [];

  const graphqlQuery = `
    query ProductInventoryVariants($id: ID!) {
      product(id: $id) {
        variants(first: 100) {
          nodes {
            id
            title
            sku
            selectedOptions { name value }
            inventoryItem {
              id
              inventoryLevels(first: 20) {
                nodes {
                  quantities(names: ["available"]) {
                    name
                    quantity
                  }
                }
              }
            }
          }
        }
      }
    }
  `;
  const json = await shopifyGraphql<any>(session.shop, session.accessToken, graphqlQuery, { id: productId });

  const nodes: any[] = json?.data?.product?.variants?.nodes ?? [];
  // Relabel the lone default-title variant of a single-variant product to the
  // "Free Size" sentinel so matchingVariantForSize("Free Size", …) finds it.
  // For everything else, preserve the original variant title so existing
  // size-only and multi-option products keep matching the way they did
  // before (size-only: title is already the size; multi-option: title stays
  // "Red / M" and is intentionally skipped by the matcher).
  const isFreeSize = nodes.length === 1 && (() => {
    const opts = (nodes[0]?.selectedOptions ?? []) as { name?: string | null; value?: string | null }[];
    if (!opts.length) return true;
    const hasSize = opts.some((o) => (o?.name ?? "").trim().toLowerCase() === "size");
    if (hasSize) return false;
    return opts.every((o) => (o?.name ?? "") === "Title" && (o?.value ?? "") === "Default Title");
  })();

  return nodes
    .map((variant: any) => ({
      id: String(variant.id ?? ""),
      title: isFreeSize ? "Free Size" : String(variant.title ?? ""),
      sku: variant.sku ? String(variant.sku) : null,
      availableInventory: shopifyVariantAvailableInventory(variant),
      inventoryItemId: variant.inventoryItem?.id ? String(variant.inventoryItem.id) : null,
    }))
    .filter((variant: ShopifyInventoryVariantInfo) => variant.id && variant.title);
}

async function getPrimaryShopifyLocationId(shop: string, accessToken: string) {
  const json = await shopifyGraphql<any>(shop, accessToken, `
    query PrimaryInventoryLocation {
      locations(first: 10) {
        nodes { id isActive }
      }
    }
  `);
  const locations = json?.data?.locations?.nodes ?? [];
  return (locations.find((location: any) => location.isActive) ?? locations[0])?.id ?? null;
}

async function addShopifyInventory(shop: string, accessToken: string, changes: ShopifyInventoryChange[]) {
  const locationId = await getPrimaryShopifyLocationId(shop, accessToken);
  if (!locationId || !changes.length) return [];

  const json = await shopifyGraphql<any>(shop, accessToken, `
    mutation AddPackingListInventory($input: InventoryAdjustQuantitiesInput!) {
      inventoryAdjustQuantities(input: $input) {
        userErrors { field message }
        inventoryAdjustmentGroup { id }
      }
    }
  `, {
    input: {
      name: "available",
      reason: "correction",
      changes: changes.map((change) => ({
        delta: change.qty,
        inventoryItemId: change.inventoryItemId,
        locationId,
      })),
    },
  });

  const userErrors = json?.data?.inventoryAdjustQuantities?.userErrors ?? [];
  if (userErrors.length || !json?.data?.inventoryAdjustQuantities?.inventoryAdjustmentGroup) return [];
  return changes.map((change) => change.size);
}

async function enrichOrdersWithShopifyVariants<T extends {
  id: number;
  shop: string;
  productId: string;
  createdAt: Date;
  lines: Array<{
    id: number;
    orderId: number;
    variantId: string;
    variantTitle: string;
    sku: string | null;
    qtyOrdered: number;
    qtyReceived: number;
    costPrice: number | null;
    createdAt: Date;
    availableInventory?: number | null;
  }>;
}>(orders: T[]): Promise<T[]> {
  const variantEntries = await Promise.all(
    Array.from(new Set(orders.map((order) => `${order.shop}|||${order.productId}`))).map(async (key) => {
      const [shop, productId] = key.split("|||");
      return [key, await getShopifyProductVariants(shop, productId)] as const;
    }),
  );
  const variantsByProduct = new Map(variantEntries);

  return orders.map((order) => {
    const variants = variantsByProduct.get(`${order.shop}|||${order.productId}`) ?? [];
    if (!variants.length) return order;

    const linesByVariantId = new Map(order.lines.map((line) => [line.variantId, line]));
    const linesByTitle = new Map(order.lines.map((line) => [line.variantTitle.trim().toLowerCase(), line]));
    const usedLineIds = new Set<number>();
    const nextLines: Array<T["lines"][number]> = [];

    for (const variant of variants) {
      const exactLine = linesByVariantId.get(variant.id);
      if (exactLine) {
        usedLineIds.add(exactLine.id);
        nextLines.push({ ...exactLine, availableInventory: variant.availableInventory });
        continue;
      }
      const titleMatch = linesByTitle.get(variant.title.trim().toLowerCase());
      if (titleMatch) {
        usedLineIds.add(titleMatch.id);
        nextLines.push({
          ...titleMatch,
          id: -Math.abs(titleMatch.id),
          variantId: variant.id,
          variantTitle: variant.title,
          sku: variant.sku ?? titleMatch.sku,
          availableInventory: variant.availableInventory,
        });
        continue;
      }

      nextLines.push({
        id: -Number(`${order.id}${nextLines.length + 1}`),
        orderId: order.id,
        variantId: variant.id,
        variantTitle: variant.title,
        sku: variant.sku,
        qtyOrdered: 0,
        qtyReceived: 0,
        costPrice: null,
        createdAt: order.createdAt,
        availableInventory: variant.availableInventory,
      });
    }

    for (const line of order.lines) {
      if (usedLineIds.has(line.id)) continue;
      nextLines.push(line);
    }

    return {
      ...order,
      lines: nextLines,
    };
  });
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

function handleTableGridKeyDown(event: React.KeyboardEvent<HTMLTableElement>) {
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
}

// ─── Main component ───────────────────────────────────────────────────────────

type Order = Awaited<ReturnType<typeof loader>>["orders"][number];
type PortalMessageItem = Awaited<ReturnType<typeof loader>>["messages"][number];
type RowMenuAction = {
  label: string;
  onClick?: () => void;
  danger?: boolean;
  disabled?: boolean;
  options?: { label: string; value: string }[];
  onSelect?: (value: string) => void;
};
type PortalUndoEntry = {
  label: string;
  fields: Record<string, string | number>;
};

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
    statusFilterCounts,
    priorityFilters,
    sortBy,
    page,
    columnWidths: savedColumnWidths,
    packingColumnWidths,
    tableHeaderLabels,
    customColumns,
    customCells,
    rowHeights,
    restockSettings,
    universalSettings,
    fabricSettings,
    productInfo,
    packingLists,
    selectedPackingList,
    shopDomain,
    productSearch,
    restockProductSearch,
    packingSearchLineId,
    productResults,
    restockProductResults,
    users,
    currentUser,
    activeUsers,
    messages,
    messageOrderId,
    loginBlocked,
    activityLogs,
    navOrder,
    fabricSheets,
    inrToAudRate,
    samples,
    visionBoardData,
    collections,
    fabricStockIndex,
  } = useLoaderData<typeof loader>();
  const [searchParams, setSearchParams] = useSearchParams();
  const submit = useSubmit();
  const columnWidthsFetcher = useFetcher();
  const undoFetcher = useFetcher();
  const [addRowNonce, setAddRowNonce] = useState(0);
  // Reset the restock table's inner scroll only when the user navigates TO
  // the restock page (or when the scroll node first mounts). Inline callback
  // refs re-fire on every render, which was forcing scrollTop=0 on every
  // state update and bouncing the page back to the top mid-edit.
  const restockTableScrollRef = useRef<HTMLDivElement | null>(null);
  const setRestockTableScrollRef = useCallback((node: HTMLDivElement | null) => {
    restockTableScrollRef.current = node;
    if (node) node.scrollTop = 0;
  }, []);
  useLayoutEffect(() => {
    if (page !== "restock") return;
    const el = restockTableScrollRef.current;
    if (el) el.scrollTop = 0;
  }, [page]);
  const [showCategoryDropdown, setShowCategoryDropdown] = useState(false);
  const [selectedAddCategory, setSelectedAddCategory] = useState<string | null>(null);
  const [newCategoryInput, setNewCategoryInput] = useState("");
  const [showNewCategoryInput, setShowNewCategoryInput] = useState(false);
  const [historyMenu, setHistoryMenu] = useState<{ x: number; y: number; entity: string; entityId: string; field: string; entityName: string } | null>(null);
  // Mirror universal-settings CSS vars onto :root so portaled overlays (drawers, modals)
  // can read them (the root container's style scope only covers in-tree descendants).
  useEffect(() => {
    if (typeof document === "undefined") return;
    const root = document.documentElement;
    root.style.setProperty("--portal-panel-font-size", `${universalSettings.panelTextSize}px`);
    root.style.setProperty("--portal-primary-button-bg", universalSettings.primaryButtonBg);
    root.style.setProperty("--portal-primary-button-color", universalSettings.primaryButtonColor);
    root.style.setProperty("--portal-table-font-size", `${universalSettings.tableTextSize}px`);
    root.style.setProperty("--portal-table-text-color", universalSettings.tableTextColor);
    root.style.setProperty("--portal-heading-font-size", `${universalSettings.headingTextSize}px`);
    root.style.setProperty("--portal-heading-text-color", universalSettings.headingTextColor);
    root.style.setProperty("--portal-inventory-font-size", `${universalSettings.inventoryFontSize}px`);
  }, [universalSettings]);
  useEffect(() => {
    const handler = (e: Event) => {
      const detail = (e as CustomEvent).detail as { x: number; y: number; entity: string; entityId: string; field: string; entityName: string };
      setHistoryMenu(detail);
    };
    document.addEventListener("show-cell-history", handler);
    return () => document.removeEventListener("show-cell-history", handler);
  }, []);
  useEffect(() => {
    const handleUndoKey = (event: KeyboardEvent) => {
      if (!(event.metaKey || event.ctrlKey) || event.shiftKey || event.key.toLowerCase() !== "z") return;
      const activeElement = document.activeElement as HTMLElement | null;
      const isEditingText = activeElement instanceof HTMLInputElement
        || activeElement instanceof HTMLTextAreaElement
        || activeElement?.isContentEditable;
      if (isEditingText) return;
      const undone = submitLastPortalUndo(undoFetcher);
      if (undone) event.preventDefault();
    };
    window.addEventListener("keydown", handleUndoKey);
    return () => window.removeEventListener("keydown", handleUndoKey);
  }, [undoFetcher]);
  const [columnWidths, setColumnWidths] = useState<Record<string, number>>(savedColumnWidths);
  const [searchTitleInput, setSearchTitleInput] = useState(searchTitle);
  const [isSearchFocused, setIsSearchFocused] = useState(false);
  const canLoadPackingInventory = canPortalUserLoadPackingInventory(users, currentUser);
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
    { id: "fabricStock", label: "Fabric in stock", center: true },
    ...customColumns.restock.map((column) => ({ id: column.id, label: column.label })),
  ];

  const widthFor = (columnId: string) => columnWidths[columnId] ?? defaultColumnWidth(columnId);
  const tableWidth = columns.reduce((sum, column) => sum + widthFor(column.id), 48);
  const updateParams = (updates: Record<string, string>) => {
    const next = new URLSearchParams(typeof window === "undefined" ? searchParams : window.location.search);
    for (const [key, value] of Object.entries(updates)) {
      if (value) next.set(key, value);
      else next.delete(key);
    }
    setSearchParams(next, { replace: true, preventScrollReset: true });
  };
  useEffect(() => {
    if (!isSearchFocused) setSearchTitleInput(searchTitle);
  }, [searchTitle]);
  useEffect(() => {
    if (!isSearchFocused) return;
    if (page === "restock") return;
    const timer = window.setTimeout(() => {
      if (searchTitleInput !== searchTitle) updateParams({ q: searchTitleInput });
    }, 350);
    return () => window.clearTimeout(timer);
  }, [searchTitleInput, isSearchFocused, page]);
  const restockSearch = page === "restock" ? searchTitleInput.trim().toLowerCase() : "";
  const visibleOrders = page === "restock" && restockSearch
    ? orders.filter((order) => order.productTitle.toLowerCase().includes(restockSearch))
    : orders;
  const activePageTitle = page === "dashboard" ? "Dashboard"
    : page === "fabric" ? "Fabric in stock"
    : page === "settings" ? "Settings"
    : page === "packing" ? "Packing Lists"
    : page === "pricelist" ? "Price List"
    : page === "productinfo" ? "Product Information"
    : page === "samples" ? "Samples"
    : page === "visionboard" ? "Vision Board"
    : page === "collections" ? "Collections"
    : page === "newproduct" ? "New Product Orders"
    : selectedProductGroup || "Existing Products Restock";
  const orderedNavItems = navOrder
    .map((id) => ALL_NAV_ITEMS.find((item) => item.id === id))
    .filter(Boolean)
    .filter((item) => {
      if ((item as { superadminOnly?: boolean }).superadminOnly) return currentUser?.role === "superadmin";
      return currentUser?.role === "superadmin" || Boolean(currentUser?.pageAccess[item!.id]);
    }) as typeof ALL_NAV_ITEMS[number][];

  if (loginBlocked) {
    return <PortalLogin />;
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
      submitPortalCell(
        columnWidthsFetcher,
        { intent: "update_column_widths", value: JSON.stringify(nextColumnWidths) },
        { label: "Undo column width", fields: { intent: "update_column_widths", value: JSON.stringify(columnWidths) } },
      );
    };

    document.addEventListener("mousemove", handleMove);
    document.addEventListener("mouseup", handleUp);
  };

  return (
    <div
      style={{
        ...s.appShell,
        background: universalSettings.pageBg,
        "--portal-primary-button-bg": universalSettings.primaryButtonBg,
        "--portal-primary-button-color": universalSettings.primaryButtonColor,
        "--portal-table-font-size": `${universalSettings.tableTextSize}px`,
        "--portal-table-text-color": universalSettings.tableTextColor,
        "--portal-heading-font-size": `${universalSettings.headingTextSize}px`,
        "--portal-heading-text-color": universalSettings.headingTextColor,
        "--portal-panel-font-size": `${universalSettings.panelTextSize}px`,
        "--portal-inventory-font-size": `${universalSettings.inventoryFontSize}px`,
      } as React.CSSProperties}
    >
      <style>
        {`
          .no-number-spinner::-webkit-outer-spin-button,
          .no-number-spinner::-webkit-inner-spin-button {
            -webkit-appearance: none;
            margin: 0;
          }

          .no-number-spinner {
            appearance: textfield;
            -moz-appearance: textfield;
          }

          /* Always-visible horizontal scrollbar on table wrappers (macOS overlay scrollbars auto-hide otherwise) */
          .portal-table-scroll::-webkit-scrollbar { height: 12px; width: 12px; }
          .portal-table-scroll::-webkit-scrollbar-track { background: #f1f5f9; }
          .portal-table-scroll::-webkit-scrollbar-thumb { background: #94a3b8; border-radius: 6px; border: 2px solid #f1f5f9; }
          .portal-table-scroll::-webkit-scrollbar-thumb:hover { background: #64748b; }
          .portal-table-scroll { scrollbar-color: #94a3b8 #f1f5f9; }
        `}
      </style>
      <aside style={{ ...s.sidebar, background: universalSettings.menuBg, color: universalSettings.menuTextColor }}>
        {universalSettings.logoUrl && (
          <div style={{ padding: "2px 14px 0" }}>
            <img src={universalSettings.logoUrl} alt="Logo" style={{ maxWidth: "100%", display: "block", borderRadius: 6 }} />
          </div>
        )}
        <div style={{ ...s.sidebarTop, ...(universalSettings.logoUrl ? { marginTop: 14 } : {}) }}>
          <div style={s.sidebarTitle}>Production Portal</div>
        </div>
        <nav style={s.nav}>
          {orderedNavItems.map((item) => {
            const isActive = item.id === "restock" ? (page === "restock" && !selectedProductGroup) : page === item.id;
            return (
              <a key={item.id} href={item.href} style={{ ...s.navItem, ...(isActive ? s.navItemActive : {}) }}>{item.label}</a>
            );
          })}
        </nav>
        <a href="/portal?page=settings" style={{ ...s.navItem, ...(page === "settings" ? s.navItemActive : {}), ...s.settingsLink }}>Settings</a>
      </aside>

      <main style={s.main}>
        <header style={s.pageHeader}>
          <div>
            <h1 style={s.pageTitle}>{activePageTitle}</h1>
          </div>
          <div style={s.headerControls}>
            <div style={s.utilityBar}>
              {(page === "restock" || page === "packing" || page === "fabric" || page === "productinfo" || page === "samples") && (
                <label style={s.filterLabel}>
                  Search
                  <input
                    type="search"
                    value={searchTitleInput}
                    onChange={(event) => setSearchTitleInput(event.currentTarget.value)}
                    onFocus={() => setIsSearchFocused(true)}
                    onBlur={() => {
                      setIsSearchFocused(false);
                      if (page !== "restock" && searchTitleInput !== searchTitle) updateParams({ q: searchTitleInput });
                    }}
                    style={s.searchInput}
                    placeholder={page === "fabric" ? "Fabric name" : page === "packing" ? "Invoice / list title" : page === "productinfo" ? "Style name" : page === "samples" ? "Sample name" : "Product title"}
                  />
                </label>
              )}
              <MessagesMenu messages={messages} />
              <div style={s.activeUsers} title="Currently active">
                <span style={s.activeUsersLabel}>Active</span>
                {activeUsers.length ? activeUsers.map((user) => (
                  <ActiveUserBadge
                    key={user.id}
                    user={user}
                    isSelf={currentUser?.id === user.id}
                    onLogout={() => submit({ intent: "portal_logout" }, { method: "post" })}
                  />
                )) : <span style={s.activeUserEmpty}>No active users</span>}
              </div>
            </div>
            {page === "restock" && (
              <div style={{ ...s.filters, position: "relative" }}>
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
                      <option key={status} value={status}>{labelForOption(restockSettings.statusOptions, status)} ({statusFilterCounts[status] ?? 0})</option>
                    ))}
                  </select>
                </label>
                <label style={s.filterLabel}>
                  Priority
                  <select value={selectedPriority} onChange={(event) => updateParams({ priority: event.currentTarget.value })} style={s.productTypeFilter}>
                    <option value="">All priorities</option>
                    {priorityFilters.map((priority) => (
                      <option key={priority} value={priority}>{labelForOption(restockSettings.priorityOptions, priority)}</option>
                    ))}
                  </select>
                </label>
              </div>
            )}
          </div>
        </header>

        {page === "settings" ? (
          <div style={{ display: "grid", gap: 16 }}>
            <SettingsPanel
              users={users}
              currentUser={currentUser}
              restockSettings={restockSettings}
              universalSettings={universalSettings}
            />
            <ActivityLogPanel activityLogs={activityLogs} />
          </div>
        ) : page === "packing" ? (
          <PackingListsPanel
            packingLists={packingLists}
            selectedPackingList={selectedPackingList}
            savedPackingColumnWidths={packingColumnWidths}
            tableHeaderLabels={tableHeaderLabels}
            customColumns={customColumns.packing}
            customCells={customCells}
            rowHeights={rowHeights}
            productSearch={productSearch}
            packingSearchLineId={packingSearchLineId}
            productResults={productResults}
            searchTitle={searchTitleInput}
            updateParams={updateParams}
            canLoadInventory={canLoadPackingInventory}
            canEditLockedQuantities={Boolean(currentUser?.admin)}
            isAdmin={Boolean(currentUser?.admin)}
            shopDomain={shopDomain}
          />
        ) : page === "fabric" ? (
          <CombinedFabricStockPanel
            sheets={fabricSheets}
            fabricSettings={fabricSettings}
            productInfo={productInfo}
            users={users}
            rowHeights={rowHeights}
            inrToAudRate={inrToAudRate}
            nameSearch={searchTitleInput}
          />
        ) : page === "productinfo" ? (
          <ProductInformationPanel
            productInfo={productInfo}
            selectedCategoryId={searchParams.get("category") ?? ""}
            search={searchTitleInput}
            updateParams={updateParams}
          />
        ) : page === "samples" ? (
          <SamplesPanel
            samples={samples}
            search={searchTitleInput}
            users={users}
            currentUser={currentUser}
          />
        ) : page === "visionboard" ? (
          <VisionBoardV2Panel
            boards={visionBoardData.boards}
            activeBoardId={visionBoardData.activeBoardId}
            items={visionBoardData.items}
          />
        ) : page === "collections" ? (
          <CollectionsPanel
            collections={collections}
          />
        ) : page !== "restock" ? (
          <div style={s.empty}>{activePageTitle} will be set up here.</div>
        ) : (
          <div className="portal-table-scroll" style={s.tableWrap} ref={setRestockTableScrollRef}>
            <table style={{ ...s.table, width: tableWidth }} onKeyDown={handleTableGridKeyDown}>
              <colgroup>
                <col style={{ width: 48 }} />
                {columns.map((column) => (
                  <col key={column.id} style={{ width: widthFor(column.id) }} />
                ))}
              </colgroup>
              <thead>
                <tr style={s.headerRow}>
                  <th style={{ ...s.th, ...s.rowNumberHeader }}>#</th>
                  {columns.map((column, colIdx) => {
                    const restockFrozenLeft = [
                      48,
                      48 + widthFor("factoryNotes"),
                      48 + widthFor("factoryNotes") + widthFor("orderDate"),
                      48 + widthFor("factoryNotes") + widthFor("orderDate") + widthFor("picture"),
                    ];
                    return (
                      <Th
                        key={column.id}
                        center={column.center}
                        headerKey={`restock:${column.id}`}
                        columnId={column.id}
                        onResizeStart={(event) => startResize(column.id, event)}
                        stickyLeft={colIdx < 4 ? restockFrozenLeft[colIdx] : undefined}
                        isLastFrozen={colIdx === 3}
                      >
                        {headerLabel(tableHeaderLabels, `restock:${column.id}`, column.label)}
                      </Th>
                    );
                  })}
                </tr>
              </thead>
              <tbody
                onContextMenu={(e) => {
                  const td = (e.target as HTMLElement).closest<HTMLElement>("[data-history-field]");
                  if (!td) return;
                  e.preventDefault();
                  const entity = td.dataset.historyEntity;
                  const entityId = td.dataset.historyEntityId;
                  const field = td.dataset.historyField;
                  const entityName = td.dataset.historyEntityName ?? "";
                  if (!entity || !entityId || !field) return;
                  setHistoryMenu({ x: e.clientX, y: e.clientY, entity, entityId, field, entityName });
                }}
              >
                {visibleOrders.map((order, rowIndex) => {
                  const restockFrozenOffsets = [
                    48,
                    48 + widthFor("factoryNotes"),
                    48 + widthFor("factoryNotes") + widthFor("orderDate"),
                    48 + widthFor("factoryNotes") + widthFor("orderDate") + widthFor("picture"),
                  ];
                  return (
                  <OrderRow
                    key={order.id}
                    order={order}
                    rowIndex={rowIndex + 1}
                    sizes={sizes}
                    users={users}
                    restockSettings={restockSettings}
                    customColumns={customColumns.restock}
                    customCells={customCells}
                    rowHeights={rowHeights}
                    frozenOffsets={restockFrozenOffsets}
                    fabricStockIndex={fabricStockIndex}
                  />
                  );
                })}
                {Array.from({ length: 10 }, (_, index) => (
                  <AddRestockOrderRow
                    key={`${addRowNonce}:${index}`}
                    rowIndex={visibleOrders.length + index + 1}
                    sizes={sizes}
                    initialProductGroup={selectedProductGroup}
                    productSearch={restockProductSearch}
                    productResults={restockProductResults}
                    updateParams={updateParams}
                    restockSettings={restockSettings}
                    onSaved={() => setAddRowNonce((current) => current + 1)}
                  />
                ))}
                {visibleOrders.length === 0 && (
                  <tr style={s.row}>
                    <td colSpan={columns.length + 1} style={{ ...s.td, textAlign: "center", color: "#6b7280", fontWeight: 700 }}>
                      {messageOrderId ? "That message is for an order that is no longer open." : restockSearch ? "No restock rows match this search." : "Search in the blank rows below to add products."}
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        )}
      </main>
      {historyMenu && typeof document !== "undefined" && (
        <CellHistoryMenu
          x={historyMenu.x}
          y={historyMenu.y}
          entity={historyMenu.entity}
          entityId={historyMenu.entityId}
          field={historyMenu.field}
          entityName={historyMenu.entityName}
          activityLogs={activityLogs}
          onClose={() => setHistoryMenu(null)}
        />
      )}
    </div>
  );
}

function ActiveUserBadge({
  user,
  isSelf,
  onLogout,
}: {
  user: ActivePortalUser;
  isSelf: boolean;
  onLogout: () => void;
}) {
  const ref = useRef<HTMLSpanElement | null>(null);
  const [hover, setHover] = useState(false);
  const [menu, setMenu] = useState<{ x: number; y: number } | null>(null);
  const [tipPos, setTipPos] = useState<{ left: number; top: number } | null>(null);

  useEffect(() => {
    if (!hover) { setTipPos(null); return; }
    const rect = ref.current?.getBoundingClientRect();
    if (!rect) return;
    setTipPos({ left: rect.left + rect.width / 2, top: rect.bottom + 8 });
  }, [hover]);

  useEffect(() => {
    if (!menu) return;
    const close = () => setMenu(null);
    document.addEventListener("mousedown", close);
    document.addEventListener("keydown", close);
    return () => {
      document.removeEventListener("mousedown", close);
      document.removeEventListener("keydown", close);
    };
  }, [menu]);

  return (
    <>
      <span
        ref={ref}
        style={{ ...s.activeUserBadge, ...(isSelf ? { cursor: "context-menu" } : {}) }}
        onMouseEnter={() => setHover(true)}
        onMouseLeave={() => setHover(false)}
        onContextMenu={isSelf ? (event) => {
          event.preventDefault();
          setHover(false);
          setMenu({ x: event.clientX, y: event.clientY });
        } : undefined}
      >
        {user.initials}
      </span>
      {tipPos && typeof document !== "undefined" && createPortal(
        <div style={{ ...s.activeUserTooltip, left: tipPos.left, top: tipPos.top }}>
          {user.name}
        </div>,
        document.body,
      )}
      {menu && typeof document !== "undefined" && createPortal(
        <div
          style={{ ...s.contextMenu, left: menu.x, top: menu.y }}
          onMouseDown={(event) => event.stopPropagation()}
        >
          <button
            type="button"
            style={s.contextMenuButton}
            onClick={() => {
              setMenu(null);
              onLogout();
            }}
          >
            Log out
          </button>
        </div>,
        document.body,
      )}
    </>
  );
}

function RowNumberCell({
  rowNumber,
  actions,
  heightKey,
}: {
  rowNumber: number | string;
  actions: RowMenuAction[];
  heightKey?: string;
}) {
  const fetcher = useFetcher();
  const [menu, setMenu] = useState<{ x: number; y: number } | null>(null);
  const startRowResize = (event: React.MouseEvent<HTMLSpanElement>) => {
    if (!heightKey) return;
    event.preventDefault();
    event.stopPropagation();
    const row = event.currentTarget.closest("tr");
    if (!row) return;
    const startY = event.clientY;
    const startHeight = row.getBoundingClientRect().height;
    const previousCursor = document.body.style.cursor;
    const previousUserSelect = document.body.style.userSelect;
    document.body.style.cursor = "row-resize";
    document.body.style.userSelect = "none";
    let nextHeight = startHeight;
    const handleMove = (moveEvent: MouseEvent) => {
      nextHeight = Math.min(420, Math.max(34, startHeight + moveEvent.clientY - startY));
      row.style.height = `${nextHeight}px`;
    };
    const handleUp = () => {
      document.body.style.cursor = previousCursor;
      document.body.style.userSelect = previousUserSelect;
      document.removeEventListener("mousemove", handleMove);
      document.removeEventListener("mouseup", handleUp);
      submitPortalCell(
        fetcher,
        { intent: "update_row_height", key: heightKey, height: Math.round(nextHeight) },
        { label: "Undo row height", fields: { intent: "update_row_height", key: heightKey, height: Math.round(startHeight) } },
      );
    };
    document.addEventListener("mousemove", handleMove);
    document.addEventListener("mouseup", handleUp);
  };
  useEffect(() => {
    if (!menu) return;
    const close = () => setMenu(null);
    document.addEventListener("mousedown", close);
    document.addEventListener("keydown", close);
    return () => {
      document.removeEventListener("mousedown", close);
      document.removeEventListener("keydown", close);
    };
  }, [menu]);
  return (
    <>
      <td
        tabIndex={0}
        style={s.rowNumberCell}
        onContextMenu={(event) => {
          event.preventDefault();
          setMenu({ x: event.clientX, y: event.clientY });
        }}
        title="Right click for row actions"
      >
        {rowNumber}
        {heightKey ? <span style={s.rowResizeHandle} onMouseDown={startRowResize} title="Drag to resize row" /> : null}
      </td>
      {menu && typeof document !== "undefined" && createPortal(
        <div
          style={{ ...s.contextMenu, left: menu.x, top: menu.y }}
          onMouseDown={(event) => event.stopPropagation()}
        >
          {actions.map((action) => action.options ? (
            <label key={action.label} style={s.contextMenuLabel}>
              <span>{action.label}</span>
              <select
                value=""
                disabled={action.disabled}
                style={s.contextMenuSelect}
                onChange={(event) => {
                  const nextValue = event.currentTarget.value;
                  if (!nextValue) return;
                  setMenu(null);
                  action.onSelect?.(nextValue);
                }}
              >
                <option value="">Select...</option>
                {action.options.map((option) => (
                  <option key={option.value} value={option.value}>{option.label}</option>
                ))}
              </select>
            </label>
          ) : (
            <button
              key={action.label}
              type="button"
              disabled={action.disabled}
              style={{ ...s.contextMenuButton, ...(action.danger ? s.contextMenuDanger : {}) }}
              onClick={() => {
                setMenu(null);
                action.onClick?.();
              }}
            >
              {action.label}
            </button>
          ))}
        </div>,
        document.body,
      )}
    </>
  );
}

function PortalLogin() {
  const actionData = useActionData<{ loginError?: string }>();
  return (
    <div style={s.loginShell}>
      <form method="post" style={s.loginCard}>
        <input type="hidden" name="intent" value="portal_login" />
        <h1 style={s.loginTitle}>Production Portal</h1>
        {actionData?.loginError && (
          <div style={{ background: "#fef2f2", color: "#dc2626", border: "1px solid #fca5a5", borderRadius: 6, padding: "8px 12px", fontSize: 13, marginBottom: 8 }}>
            {actionData.loginError}
          </div>
        )}
        <label style={s.loginLabel}>
          Name
          <input name="username" required autoComplete="username" style={s.loginInput} placeholder="Enter your name" />
        </label>
        <label style={s.loginLabel}>
          Password
          <input name="password" type="password" required autoComplete="current-password" style={s.loginInput} placeholder="Enter your password" />
        </label>
        <button type="submit" style={s.loginButton}>Sign in</button>
      </form>
    </div>
  );
}

function MessagesMenu({ messages }: { messages: PortalMessageItem[] }) {
  const [open, setOpen] = useState(false);
  const fetcher = useFetcher();

  return (
    <div style={s.messagesWrap}>
      <button
        type="button"
        onClick={() => setOpen((current) => !current)}
        style={messages.length ? s.messagesButtonActive : s.messagesButton}
        title="Messages"
        aria-label={messages.length ? `${messages.length} unread messages` : "Messages"}
      >
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
          <path d="M4 5.75C4 4.78 4.78 4 5.75 4h12.5C19.22 4 20 4.78 20 5.75v8.5c0 .97-.78 1.75-1.75 1.75H9.4L5.2 19.15A.75.75 0 0 1 4 18.55V5.75Z" stroke="currentColor" strokeWidth="2" strokeLinejoin="round" />
          <path d="M7.5 8.5h9M7.5 12h6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
        </svg>
        {messages.length > 0 && <span style={s.messageCount}>{messages.length}</span>}
      </button>
      {open && (
        <div style={s.messagesPopover}>
          <div style={s.messagesHeader}>Messages</div>
          {messages.length ? messages.map((message) => (
            <div key={message.id} style={s.messageItem}>
              <a
                href={message.field === "sample_notes" ? `/portal?page=samples` : `/portal?messageOrderId=${message.orderId}#order-${message.orderId}`}
                style={s.messageLink}
              >
                <strong>{message.productTitle || `Order #${message.orderId}`}</strong>
                <span>{message.field === "factory_notes" ? "Factory notes" : message.field === "sample_notes" ? "Sample notes" : "Notes"}</span>
                <span style={s.messageBody}>{message.body}</span>
              </a>
              <fetcher.Form method="post">
                <input type="hidden" name="intent" value="mark_message_read" />
                <input type="hidden" name="messageId" value={message.id} />
                <button type="submit" style={s.messageReadButton}>Done</button>
              </fetcher.Form>
            </div>
          )) : (
            <div style={s.messageEmpty}>No messages for you.</div>
          )}
        </div>
      )}
    </div>
  );
}

// ─── Custom Colour Picker ────────────────────────────────────────────────────

const _pickerRecentColors: string[] = [];

function _hexToHsv(hex: string) {
  const r = parseInt(hex.slice(1, 3), 16) / 255;
  const g = parseInt(hex.slice(3, 5), 16) / 255;
  const b = parseInt(hex.slice(5, 7), 16) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b), d = max - min;
  const v = max;
  const sv = max === 0 ? 0 : d / max;
  let h = 0;
  if (d !== 0) {
    if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
    else if (max === g) h = ((b - r) / d + 2) / 6;
    else h = ((r - g) / d + 4) / 6;
  }
  return { h: h * 360, s: sv * 100, v: v * 100 };
}

function _hsvToHex(h: number, s: number, v: number) {
  h /= 360; s /= 100; v /= 100;
  const i = Math.floor(h * 6), f = h * 6 - i;
  const p = v * (1 - s), q = v * (1 - f * s), t = v * (1 - (1 - f) * s);
  let r = 0, g = 0, b = 0;
  switch (i % 6) {
    case 0: r = v; g = t; b = p; break;
    case 1: r = q; g = v; b = p; break;
    case 2: r = p; g = v; b = t; break;
    case 3: r = p; g = q; b = v; break;
    case 4: r = t; g = p; b = v; break;
    case 5: r = v; g = p; b = q; break;
  }
  return "#" + [r, g, b].map((n) => Math.round(Math.max(0, Math.min(255, n * 255))).toString(16).padStart(2, "0")).join("");
}

function ColorPickerInput({
  value,
  disabled,
  onChange,
}: {
  value: string;
  disabled?: boolean;
  onChange: (hex: string) => void;
}) {
  const safeHex = /^#[0-9a-f]{6}$/i.test(value) ? value.toLowerCase() : "#000000";
  const [open, setOpen] = useState(false);
  const [panelPos, setPanelPos] = useState({ top: 0, left: 0 });
  const [hsv, setHsv] = useState(() => _hexToHsv(safeHex));
  const [hexText, setHexText] = useState(safeHex.slice(1));
  const triggerRef = useRef<HTMLButtonElement>(null);
  const panelRef = useRef<HTMLDivElement>(null);
  const gradRef = useRef<HTMLDivElement>(null);
  const hueRef = useRef<HTMLDivElement>(null);
  const dragging = useRef<"grad" | "hue" | null>(null);
  const hsvRef = useRef(hsv);
  hsvRef.current = hsv;

  // Position the portal panel below the trigger
  const openPanel = (e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (disabled) return;
    if (!open && triggerRef.current) {
      const r = triggerRef.current.getBoundingClientRect();
      const panelH = 340;
      const top = r.bottom + 6 + panelH > window.innerHeight ? r.top - panelH - 6 : r.bottom + 6;
      setPanelPos({ top, left: Math.min(r.left, window.innerWidth - 308) });
    }
    setOpen((o) => !o);
  };

  // Close on outside click (checks both trigger and panel)
  useEffect(() => {
    if (!open) return;
    const close = (e: MouseEvent) => {
      if (
        !triggerRef.current?.contains(e.target as Node) &&
        !panelRef.current?.contains(e.target as Node)
      ) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);

  // Drag handling for gradient and hue sliders
  useEffect(() => {
    const move = (e: MouseEvent) => {
      if (dragging.current === "grad" && gradRef.current) {
        const rect = gradRef.current.getBoundingClientRect();
        const ns = Math.round(Math.max(0, Math.min(100, ((e.clientX - rect.left) / rect.width) * 100)));
        const nv = Math.round(Math.max(0, Math.min(100, (1 - (e.clientY - rect.top) / rect.height) * 100)));
        const next = { ...hsvRef.current, s: ns, v: nv };
        setHsv(next);
        const hex = _hsvToHex(next.h, next.s, next.v);
        setHexText(hex.slice(1));
        onChange(hex);
      } else if (dragging.current === "hue" && hueRef.current) {
        const rect = hueRef.current.getBoundingClientRect();
        const nh = Math.round(Math.max(0, Math.min(360, ((e.clientY - rect.top) / rect.height) * 360)));
        const next = { ...hsvRef.current, h: nh };
        setHsv(next);
        const hex = _hsvToHex(next.h, next.s, next.v);
        setHexText(hex.slice(1));
        onChange(hex);
      }
    };
    const up = () => { dragging.current = null; };
    document.addEventListener("mousemove", move);
    document.addEventListener("mouseup", up);
    return () => { document.removeEventListener("mousemove", move); document.removeEventListener("mouseup", up); };
  }, [onChange]);

  const currentHex = _hsvToHex(hsv.h, hsv.s, hsv.v);
  const hueHex = _hsvToHex(hsv.h, 100, 100);

  const pickGrad = (e: React.MouseEvent<HTMLDivElement>) => {
    const rect = e.currentTarget.getBoundingClientRect();
    const ns = Math.round(Math.max(0, Math.min(100, ((e.clientX - rect.left) / rect.width) * 100)));
    const nv = Math.round(Math.max(0, Math.min(100, (1 - (e.clientY - rect.top) / rect.height) * 100)));
    const next = { ...hsv, s: ns, v: nv };
    setHsv(next); onChange(_hsvToHex(next.h, next.s, next.v));
    setHexText(_hsvToHex(next.h, next.s, next.v).slice(1));
  };
  const pickHue = (e: React.MouseEvent<HTMLDivElement>) => {
    const rect = e.currentTarget.getBoundingClientRect();
    const nh = Math.round(Math.max(0, Math.min(360, ((e.clientY - rect.top) / rect.height) * 360)));
    const next = { ...hsv, h: nh };
    setHsv(next); onChange(_hsvToHex(next.h, next.s, next.v));
    setHexText(_hsvToHex(next.h, next.s, next.v).slice(1));
  };
  const commitHex = () => {
    const clean = hexText.replace("#", "").toLowerCase();
    if (/^[0-9a-f]{6}$/.test(clean)) {
      const hex = "#" + clean;
      setHsv(_hexToHsv(hex));
      onChange(hex);
      if (!_pickerRecentColors.includes(hex)) {
        _pickerRecentColors.unshift(hex);
        if (_pickerRecentColors.length > 10) _pickerRecentColors.pop();
      }
    }
  };

  const panel = open ? (
    <div
      ref={panelRef}
      style={{
        position: "fixed", top: panelPos.top, left: panelPos.left, zIndex: 99999,
        background: "#fff", borderRadius: 14,
        boxShadow: "0 8px 32px rgba(0,0,0,0.22)", border: "1px solid #e5e7eb",
        padding: 12, width: 296,
      }}
      onMouseDown={(e) => e.stopPropagation()}
    >
      <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
        <div
          ref={gradRef}
          style={{
            flex: 1, height: 200, borderRadius: 8, position: "relative", cursor: "crosshair",
            background: `linear-gradient(to bottom, transparent, #000), linear-gradient(to right, #fff, ${hueHex})`,
            userSelect: "none",
          }}
          onMouseDown={(e) => { dragging.current = "grad"; pickGrad(e); }}
        >
          <div style={{
            position: "absolute", left: `${hsv.s}%`, top: `${100 - hsv.v}%`,
            width: 14, height: 14, borderRadius: "50%",
            border: "2px solid #fff", boxShadow: "0 0 0 1.5px rgba(0,0,0,0.35)",
            transform: "translate(-50%,-50%)", pointerEvents: "none",
          }} />
        </div>
        <div
          ref={hueRef}
          style={{
            width: 18, height: 200, borderRadius: 999, position: "relative",
            cursor: "ns-resize", flexShrink: 0, userSelect: "none",
            background: "linear-gradient(to bottom,#f00,#ff0,#0f0,#0ff,#00f,#f0f,#f00)",
          }}
          onMouseDown={(e) => { dragging.current = "hue"; pickHue(e); }}
        >
          <div style={{
            position: "absolute", left: "50%", top: `${(hsv.h / 360) * 100}%`,
            width: 24, height: 8, borderRadius: 999, background: hueHex,
            border: "2px solid #fff", boxShadow: "0 0 0 1.5px rgba(0,0,0,0.35)",
            transform: "translate(-50%,-50%)", pointerEvents: "none",
          }} />
        </div>
      </div>

      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
        <div style={{ width: 34, height: 34, borderRadius: 8, background: currentHex, border: "1px solid rgba(0,0,0,0.12)", flexShrink: 0 }} />
        <div style={{ display: "flex", alignItems: "center", gap: 4, flex: 1, border: "2px solid #2563eb", borderRadius: 8, padding: "6px 10px" }}>
          <span style={{ color: "#6b7280", fontSize: 13, fontWeight: 700 }}>#</span>
          <input
            value={hexText.toUpperCase()}
            maxLength={6}
            onChange={(e) => setHexText(e.target.value.replace("#", ""))}
            onBlur={commitHex}
            onKeyDown={(e) => e.key === "Enter" && commitHex()}
            style={{ border: "none", outline: "none", fontSize: 13, fontWeight: 700, width: "100%", background: "transparent" }}
          />
        </div>
      </div>

      {_pickerRecentColors.length > 0 && (
        <div>
          <div style={{ fontSize: 11, fontWeight: 700, color: "#6b7280", marginBottom: 6 }}>▸ Recently used</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
            {_pickerRecentColors.map((col) => (
              <button key={col} type="button" title={col}
                onClick={() => { setHsv(_hexToHsv(col)); setHexText(col.slice(1)); onChange(col); }}
                style={{
                  width: 28, height: 28, borderRadius: 6, background: col, cursor: "pointer",
                  border: col === currentHex ? "2px solid #2563eb" : "1px solid rgba(0,0,0,0.12)",
                }} />
            ))}
          </div>
        </div>
      )}
    </div>
  ) : null;

  return (
    <div style={{ display: "inline-block" }}>
      <button
        ref={triggerRef}
        type="button"
        disabled={disabled}
        onClick={openPanel}
        style={{
          display: "inline-flex", alignItems: "center", gap: 8,
          padding: "5px 10px 5px 6px",
          border: "1px solid #d1d5db", borderRadius: 999,
          background: "#f9fafb", cursor: disabled ? "not-allowed" : "pointer",
          opacity: disabled ? 0.5 : 1, fontSize: 13, fontWeight: 700, color: "#111827",
          minWidth: 130,
        }}
      >
        <span style={{
          width: 22, height: 22, borderRadius: "50%", background: safeHex, flexShrink: 0,
          border: "2px solid rgba(0,0,0,0.15)", boxShadow: "0 1px 2px rgba(0,0,0,0.2)",
        }} />
        <span>#{hexText.toUpperCase()}</span>
        <svg width="13" height="13" viewBox="0 0 16 16" style={{ marginLeft: "auto", opacity: 0.45 }}>
          <path d="M2 14h3l8-8-3-3-8 8v3zm13-11l-2-2-1.5 1.5 2 2L15 3z" fill="currentColor" />
        </svg>
      </button>
      {typeof document !== "undefined" && createPortal(panel, document.body)}
    </div>
  );
}

// ─── Cell Change History Menu ─────────────────────────────────────────────────

function CellHistoryMenu({
  x, y, entity, entityId, field, entityName, activityLogs, onClose,
}: {
  x: number; y: number;
  entity: string; entityId: string; field: string; entityName: string;
  activityLogs: { id: number; userName: string; action: string; entity: string; entityId: string | null; entityName: string | null; field: string | null; toValue: string | null; createdAt: Date | string }[];
  onClose: () => void;
}) {
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleClick = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) onClose();
    };
    const handleKey = (e: KeyboardEvent) => { if (e.key === "Escape") onClose(); };
    document.addEventListener("mousedown", handleClick);
    document.addEventListener("keydown", handleKey);
    return () => {
      document.removeEventListener("mousedown", handleClick);
      document.removeEventListener("keydown", handleKey);
    };
  }, [onClose]);

  const entries = activityLogs.filter(
    (log) => log.entity === entity && log.entityId === entityId && log.field === field,
  );

  const menuW = 320;
  const menuMaxH = 380;
  const left = x + menuW > window.innerWidth ? x - menuW : x;
  const top = y + menuMaxH > window.innerHeight ? Math.max(0, y - menuMaxH) : y;

  const panel = (
    <div
      ref={ref}
      style={{
        position: "fixed", top, left, zIndex: 99999,
        background: "#fff", borderRadius: 10,
        boxShadow: "0 8px 32px rgba(0,0,0,0.22)", border: "1px solid #e5e7eb",
        width: menuW, maxHeight: menuMaxH, overflowY: "auto",
      }}
    >
      <div style={{ padding: "12px 16px 10px", borderBottom: "1px solid #f3f4f6" }}>
        <div style={{ fontWeight: 700, fontSize: 13, color: "#111827" }}>Change History</div>
        <div style={{ fontSize: 11, color: "#6b7280", marginTop: 2 }}>{entityName} — {field}</div>
      </div>
      {entries.length === 0 ? (
        <div style={{ padding: 20, color: "#9ca3af", fontSize: 13, textAlign: "center" }}>
          No history recorded for this field.
        </div>
      ) : (
        entries.map((log) => (
          <div key={log.id} style={{ padding: "10px 16px", borderBottom: "1px solid #f9fafb" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 8 }}>
              <span style={{ fontWeight: 700, fontSize: 13, color: "#111827" }}>{log.userName}</span>
              <span style={{ fontSize: 11, color: "#9ca3af", flexShrink: 0 }}>
                {new Date(log.createdAt).toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" })}
              </span>
            </div>
            <div style={{ fontSize: 13, color: "#374151", marginTop: 3 }}>
              {"→"} {log.toValue || <span style={{ color: "#9ca3af" }}>(cleared)</span>}
            </div>
          </div>
        ))
      )}
    </div>
  );

  return createPortal(panel, document.body);
}

// ─── Samples ─────────────────────────────────────────────────────────────────

type SampleIterationType = {
  id: number;
  sampleId: number;
  version: number;
  name: string | null;
  notes: string | null;
  fabricType: string | null;
  sampleSize: string | null;
  buttonType: string | null;
  factoryCost: string | null;
  status: string;
  images: unknown;
  imageCount?: number;
  hasThumbnail?: boolean;
  taggedUsers: unknown;
  createdAt: Date;
  updatedAt: Date;
};
type SampleType = {
  id: number;
  name: string;
  sortOrder: number;
  createdAt: Date;
  updatedAt: Date;
  iterations: SampleIterationType[];
};

function sampleStatusLabel(status: string) {
  if (status === "approved") return "Approved";
  if (status === "approved_for_production") return "Approved on Production";
  if (status === "changes_requested") return "Changes Requested";
  if (status === "sample_in_production") return "Sample in Production";
  if (status === "given_to_factory") return "Given to Factory";
  if (status === "sent") return "Sent";
  if (status === "under_consideration") return "Under Consideration";
  if (status === "in_progress") return "In Progress";
  return "No versions yet";
}

function sampleStatusPillStyle(status: string, large?: boolean): React.CSSProperties {
  const base: React.CSSProperties = { borderRadius: 99, fontWeight: 700, whiteSpace: "nowrap", padding: large ? "4px 12px" : "2px 9px", fontSize: large ? 12 : 11 };
  if (status === "approved") return { ...base, background: "#dcfce7", color: "#166534" };
  if (status === "approved_for_production") return { ...base, background: "#bbf7d0", color: "#14532d" };
  if (status === "changes_requested") return { ...base, background: "#fef3c7", color: "#92400e" };
  if (status === "sample_in_production") return { ...base, background: "#ffedd5", color: "#9a3412" };
  if (status === "given_to_factory") return { ...base, background: "#dbeafe", color: "#1e40af" };
  if (status === "sent") return { ...base, background: "#cffafe", color: "#155e75" };
  if (status === "under_consideration") return { ...base, background: "#ede9fe", color: "#5b21b6" };
  if (status === "in_progress") return { ...base, background: "#dbeafe", color: "#1e40af" };
  return { ...base, background: "#f1f5f9", color: "#64748b" };
}

function SamplesPanel({
  samples,
  search,
  users,
  currentUser,
}: {
  samples: SampleType[];
  search: string;
  users: PortalUser[];
  currentUser: PortalUser | null;
}) {
  const fetcher = useFetcher();
  const [localSamples, setLocalSamples] = useState(samples);
  // Track IDs removed this session so loader re-runs can't restore them
  const deletedIds = useRef<Set<number>>(new Set());
  const [selectedSampleId, setSelectedSampleId] = useState<number | null>(null);
  const [gridColumns, setGridColumns] = useState<3 | 4 | 5 | 6>(6);
  const [dragSampleId, setDragSampleId] = useState<number | null>(null);
  const [dragOverSampleId, setDragOverSampleId] = useState<number | null>(null);
  const [addModalOpen, setAddModalOpen] = useState(false);
  const [addName, setAddName] = useState("");

  // Sync from server but never restore items the user has deleted this session
  useEffect(() => {
    setLocalSamples(samples.filter((s) => !deletedIds.current.has(s.id)));
  }, [samples]);

  // Auto-compress legacy uncompressed images in the background.
  useAutoCompressSampleIterations(localSamples);

  const selectedSample = localSamples.find((s) => s.id === selectedSampleId) ?? null;
  const normalizedSearch = search.trim().toLowerCase();
  const visibleSamples = normalizedSearch
    ? localSamples.filter((s) => s.name.toLowerCase().includes(normalizedSearch))
    : localSamples;

  const addSample = () => { setAddName(""); setAddModalOpen(true); };
  const submitAddSample = () => {
    if (!addName.trim()) return;
    fetcher.submit({ intent: "add_sample", name: addName.trim() }, { method: "post" });
    setAddModalOpen(false);
    setAddName("");
  };

  const handleDelete = (sampleId: number) => {
    deletedIds.current.add(sampleId);
    setLocalSamples((prev) => prev.filter((s) => s.id !== sampleId));
    if (selectedSampleId === sampleId) setSelectedSampleId(null);
  };

  // When the drawer mutates an iteration locally (status, notes, etc.),
  // mirror the change into our localSamples so the listing card reflects it
  // without waiting for a loader revalidation.
  const handleIterationLocalUpdate = (iterationId: number, patch: Partial<SampleIterationType>) => {
    setLocalSamples((prev) => prev.map((s) => ({
      ...s,
      iterations: s.iterations.map((it) => it.id === iterationId ? { ...it, ...patch } : it),
    })));
  };

  const reorderSamples = (targetId: number) => {
    if (!dragSampleId || dragSampleId === targetId || normalizedSearch) return;
    const fromIndex = localSamples.findIndex((s) => s.id === dragSampleId);
    const toIndex = localSamples.findIndex((s) => s.id === targetId);
    if (fromIndex < 0 || toIndex < 0) return;
    const next = [...localSamples];
    const [moved] = next.splice(fromIndex, 1);
    next.splice(toIndex, 0, moved);
    // Optimistic: reorder immediately, loader skipped (shouldRevalidate)
    setLocalSamples(next);
    fetcher.submit({ intent: "reorder_samples", sampleIds: JSON.stringify(next.map((s) => s.id)) }, { method: "post" });
  };

  return (
    <div style={s.productInfoPage}>
      <div style={s.productInfoToolbar}>
        <div style={s.productInfoToolbarLeft}>
          <div>
            <h2 style={s.productInfoHeading}>Samples</h2>
            <div style={s.productInfoMeta}>{visibleSamples.length} sample{visibleSamples.length !== 1 ? "s" : ""}</div>
          </div>
        </div>
        <div style={s.productInfoActions}>
          <div style={s.productInfoSegmented} aria-label="Cards per row">
            {([3, 4, 5, 6] as const).map((count) => (
              <button
                key={count}
                type="button"
                style={gridColumns === count ? { ...s.productInfoSegmentButton, ...s.productInfoSegmentButtonActive } : s.productInfoSegmentButton}
                onClick={() => setGridColumns(count)}
              >{count}</button>
            ))}
          </div>
          <button type="button" style={s.primaryActionButton} onClick={addSample}>
            Add Sample
          </button>
        </div>
      </div>

      <div style={{ ...s.productInfoList, gridTemplateColumns: `repeat(${gridColumns}, minmax(0, 1fr))` }}>
        {visibleSamples.map((sample) => (
          <SampleCard
            key={sample.id}
            sample={sample}
            isDragging={dragSampleId === sample.id}
            isDragOver={dragOverSampleId === sample.id && dragSampleId !== sample.id}
            draggable={!normalizedSearch}
            onOpen={() => setSelectedSampleId(sample.id)}
            onDragStart={(event) => {
              if (normalizedSearch) { event.preventDefault(); return; }
              setDragSampleId(sample.id);
              event.dataTransfer.effectAllowed = "move";
            }}
            onDragOver={(event) => {
              if (!dragSampleId || normalizedSearch) return;
              event.preventDefault();
              setDragOverSampleId(sample.id);
            }}
            onDragLeave={() => setDragOverSampleId((c) => c === sample.id ? null : c)}
            onDrop={(event) => {
              event.preventDefault();
              reorderSamples(sample.id);
              setDragSampleId(null);
              setDragOverSampleId(null);
            }}
            onDragEnd={() => { setDragSampleId(null); setDragOverSampleId(null); }}
            onDeleted={(id) => handleDelete(id)}
          />
        ))}
        {visibleSamples.length === 0 && (
          <div style={{ gridColumn: "1 / -1", padding: "48px 0", textAlign: "center", color: "#9ca3af", fontSize: 14 }}>
            {search ? "No samples match this search." : "No samples yet. Click Add Sample to create your first one."}
          </div>
        )}
      </div>

      {selectedSample && typeof document !== "undefined" && createPortal(
        <SampleDetailPanel sample={selectedSample} onClose={() => setSelectedSampleId(null)} users={users} currentUser={currentUser} onIterationLocalUpdate={handleIterationLocalUpdate} />,
        document.body,
      )}

      {addModalOpen && typeof document !== "undefined" && createPortal(
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.35)", zIndex: 1400, display: "flex", alignItems: "center", justifyContent: "center" }}
          onClick={(e) => { if (e.target === e.currentTarget) setAddModalOpen(false); }}
        >
          <div style={{ background: "#fff", borderRadius: 12, padding: "28px 28px 24px", width: 380, boxShadow: "0 20px 60px rgba(0,0,0,0.18)" }}>
            <div style={{ fontSize: 16, fontWeight: 700, color: "#111827", marginBottom: 16 }}>Add sample</div>
            <input
              autoFocus
              style={{ width: "100%", border: "1px solid #d1d5db", borderRadius: 7, padding: "9px 12px", fontSize: 14, outline: "none", boxSizing: "border-box" as const, fontFamily: "inherit" }}
              placeholder="Sample name"
              value={addName}
              onChange={(e) => setAddName(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter") submitAddSample(); if (e.key === "Escape") setAddModalOpen(false); }}
            />
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
              <button type="button" onClick={() => setAddModalOpen(false)} style={{ padding: "8px 18px", borderRadius: 7, border: "1px solid #d1d5db", background: "#fff", fontSize: 14, cursor: "pointer", color: "#374151" }}>Cancel</button>
              <button type="button" onClick={submitAddSample} style={{ padding: "8px 18px", borderRadius: 7, border: "none", background: "#111827", color: "#fff", fontSize: 14, fontWeight: 600, cursor: "pointer" }}>Add</button>
            </div>
          </div>
        </div>,
        document.body,
      )}
    </div>
  );
}

function SampleCard({
  sample,
  isDragging,
  isDragOver,
  draggable,
  onOpen,
  onDragStart,
  onDragOver,
  onDragLeave,
  onDrop,
  onDragEnd,
  onDeleted,
}: {
  sample: SampleType;
  isDragging: boolean;
  isDragOver: boolean;
  draggable: boolean;
  onOpen: () => void;
  onDragStart: (e: React.DragEvent<HTMLDivElement>) => void;
  onDragOver: (e: React.DragEvent<HTMLDivElement>) => void;
  onDragLeave: () => void;
  onDrop: (e: React.DragEvent<HTMLDivElement>) => void;
  onDragEnd: () => void;
  onDeleted: (id: number) => void;
}) {
  const deleteFetcher = useFetcher();
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [gone, setGone] = useState(false);

  if (gone) return null;

  const latestIteration = sample.iterations.length > 0 ? sample.iterations[sample.iterations.length - 1] : null;
  const images = Array.isArray(latestIteration?.images) ? latestIteration.images as string[] : [];
  const expectsImage = (latestIteration?.imageCount ?? images.length) > 0;
  const status = latestIteration?.status ?? "none";
  const lastUpdated = latestIteration?.updatedAt ?? null;

  return (
    <div
      style={{
        ...s.productStyleCard,
        ...(isDragging ? s.productStyleCardDragging : {}),
        ...(isDragOver ? s.productStyleCardDropTarget : {}),
        cursor: "pointer",
      }}
      onClick={onOpen}
      onDragOver={onDragOver}
      onDragLeave={onDragLeave}
      onDrop={onDrop}
    >
      {/* Drag handle — draggable only from here */}
      <span
        draggable={draggable}
        style={{ ...s.productStyleDragHandle, pointerEvents: draggable ? "auto" : "none", cursor: draggable ? "grab" : "default" }}
        title="Drag to reorder"
        onDragStart={onDragStart}
        onDragEnd={onDragEnd}
        onClick={(e) => e.stopPropagation()}
      >::</span>

      {/* × button — first click shows overlay, overlay has the real Delete button */}
      <button
        type="button"
        title="Delete sample"
        onClick={(e) => { e.stopPropagation(); setConfirmDelete(true); }}
        style={{ position: "absolute", top: 8, right: 8, width: 24, height: 24, borderRadius: "50%", border: "none", background: "rgba(0,0,0,0.12)", color: "#374151", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, lineHeight: 1, padding: 0, zIndex: 2 }}
      >×</button>

      {/* Confirmation overlay — appears on top of card */}
      {confirmDelete && (
        <div
          style={{ position: "absolute", inset: 0, background: "rgba(0,0,0,0.6)", zIndex: 10, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 10, borderRadius: 8 }}
          onClick={(e) => e.stopPropagation()}
        >
          <div style={{ color: "#fff", fontSize: 13, fontWeight: 600 }}>Delete this sample?</div>
          <div style={{ display: "flex", gap: 8 }}>
            <button
              type="button"
              onClick={(e) => {
                e.stopPropagation();
                setGone(true);
                onDeleted(sample.id);
                deleteFetcher.submit(
                  { intent: "delete_sample", sampleId: String(sample.id) },
                  { method: "post" },
                );
              }}
              style={{ padding: "7px 16px", borderRadius: 7, border: "none", background: "#ef4444", color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer" }}
            >Delete</button>
            <button
              type="button"
              onClick={(e) => { e.stopPropagation(); setConfirmDelete(false); }}
              style={{ padding: "7px 16px", borderRadius: 7, border: "1px solid rgba(255,255,255,0.3)", background: "transparent", color: "#fff", fontSize: 13, cursor: "pointer" }}
            >Cancel</button>
          </div>
        </div>
      )}

      {/* Image */}
      <div style={s.productStyleImageWrap}>
        {expectsImage && latestIteration ? (
          <img
            src={`/portal/thumbnail/sample/${latestIteration.id}?v=${new Date(latestIteration.updatedAt).getTime()}`}
            alt={sample.name}
            style={s.productStyleImage}
            loading="lazy"
            decoding="async"
          />
        ) : (
          <div style={s.productStyleImageEmpty}>No image yet</div>
        )}
      </div>

      {/* Card body */}
      <div style={{ ...s.productStyleCardBody, paddingBottom: 14 }}>
        <span style={s.productStyleTitle}>{sample.name}</span>
        <div style={{ display: "flex", gap: 6, alignItems: "center", marginTop: 4, flexWrap: "wrap", justifyContent: "center" }}>
          <span style={sampleStatusPillStyle(status, true)}>{sampleStatusLabel(status)}</span>
          {sample.iterations.length > 0 && (
            <span style={s.sampleVersionBadge}>v{sample.iterations.length}</span>
          )}
        </div>
        {lastUpdated && (
          <span style={{ fontSize: 11, color: "#9ca3af", marginTop: 2 }}>
            Updated {formatPortalDate(lastUpdated)}
          </span>
        )}
      </div>
    </div>
  );
}

function SampleDetailPanel({
  sample,
  onClose,
  users,
  currentUser,
  onIterationLocalUpdate,
}: {
  sample: SampleType;
  onClose: () => void;
  users: PortalUser[];
  currentUser: PortalUser | null;
  onIterationLocalUpdate: (iterationId: number, patch: Partial<SampleIterationType>) => void;
}) {
  const fetcher = useFetcher();
  const [isEditingName, setIsEditingName] = useState(false);
  const [nameDraft, setNameDraft] = useState(sample.name);
  const [iterations, setIterations] = useState<SampleIterationType[]>(sample.iterations);

  useEffect(() => {
    setNameDraft(sample.name);
  }, [sample.name]);

  useEffect(() => {
    // The loader ships slim iteration data (no full images, just metadata
    // and counts). Each iteration block renders images via URL routes, so we
    // just use whatever the prop gives us.
    setIterations(sample.iterations);
  }, [sample.iterations]);

  useEffect(() => {
    const handleKey = (event: KeyboardEvent) => { if (event.key === "Escape") onClose(); };
    window.addEventListener("keydown", handleKey);
    return () => window.removeEventListener("keydown", handleKey);
  }, [onClose]);

  const saveName = () => {
    setIsEditingName(false);
    if (nameDraft.trim() && nameDraft.trim() !== sample.name) {
      fetcher.submit({ intent: "rename_sample", sampleId: String(sample.id), name: nameDraft.trim() }, { method: "post" });
    } else {
      setNameDraft(sample.name);
    }
  };

  const addIteration = () => {
    fetcher.submit({ intent: "add_sample_iteration", sampleId: String(sample.id) }, { method: "post" });
  };

  const sortedIterations = [...iterations].sort((a, b) => b.version - a.version);

  return (
    <>
      <div style={s.samplePanelBackdrop} onClick={onClose} />
      <div style={s.samplePanel}>
        <div style={s.samplePanelHeader}>
          <div style={s.samplePanelNameWrap}>
            {isEditingName ? (
              <input
                autoFocus
                style={s.samplePanelNameInput}
                value={nameDraft}
                onChange={(e) => setNameDraft(e.target.value)}
                onBlur={saveName}
                onKeyDown={(e) => { if (e.key === "Enter") (e.target as HTMLInputElement).blur(); if (e.key === "Escape") { setNameDraft(sample.name); setIsEditingName(false); } }}
              />
            ) : (
              <h2 style={s.samplePanelName} onClick={() => setIsEditingName(true)} title="Click to rename">
                {sample.name}
              </h2>
            )}
            <span style={s.samplePanelVersionCount}>
              {iterations.length} version{iterations.length !== 1 ? "s" : ""}
            </span>
          </div>
          <button type="button" style={s.samplePanelClose} onClick={onClose} aria-label="Close">×</button>
        </div>

        <div style={s.samplePanelTopActions}>
          <button type="button" style={{ ...s.primaryActionButton, width: "100%" }} onClick={addIteration}>
            + Add new version
          </button>
        </div>

        <div style={s.samplePanelIterations}>
          {sortedIterations.length === 0 && (
            <div style={s.samplePanelEmpty}>
              No versions yet. Click &ldquo;Add new version&rdquo; to record the first sample.
            </div>
          )}
          {sortedIterations.map((iteration) => (
            <SampleIterationBlock key={iteration.id} iteration={iteration} users={users} currentUser={currentUser} onLocalUpdate={onIterationLocalUpdate} />
          ))}
        </div>
      </div>
    </>
  );
}

function SampleIterationBlock({
  iteration,
  users,
  currentUser,
  onLocalUpdate,
}: {
  iteration: SampleIterationType;
  users: PortalUser[];
  currentUser: PortalUser | null;
  onLocalUpdate: (iterationId: number, patch: Partial<SampleIterationType>) => void;
}) {
  const fetcher = useFetcher();
  const notesRef = useRef<HTMLTextAreaElement>(null);
  const [notes, setNotes] = useState(iteration.notes ?? "");
  const [nameEditing, setNameEditing] = useState(false);
  const [nameDraft, setNameDraft] = useState(iteration.name ?? "");
  const [lightboxIndex, setLightboxIndex] = useState<number | null>(null);
  const [uploadModalOpen, setUploadModalOpen] = useState(false);
  const [mentionQuery, setMentionQuery] = useState<string | null>(null);
  const [mentionStart, setMentionStart] = useState(-1);
  const [fabricType, setFabricType] = useState(iteration.fabricType ?? "");
  const [sampleSize, setSampleSize] = useState(iteration.sampleSize ?? "");
  const [buttonType, setButtonType] = useState(iteration.buttonType ?? "");
  const [factoryCost, setFactoryCost] = useState(iteration.factoryCost ?? "");
  const [status, setStatus] = useState(iteration.status);
  // Saved-image count + cache-buster version. Each image is rendered as
  // <img src="/portal/image/sample/<id>/<i>?v=<version>"> so the browser
  // handles parallel download and forever-caches binary content.
  const [savedCount, setSavedCount] = useState<number>(() => {
    if (typeof iteration.imageCount === "number") return iteration.imageCount;
    const imgs = Array.isArray(iteration.images) ? iteration.images as string[] : [];
    return imgs.length;
  });
  const [version, setVersion] = useState<number>(() => new Date(iteration.updatedAt).getTime());
  type PendingImage = { id: string; blobUrl: string };
  const [pendingImages, setPendingImages] = useState<PendingImage[]>([]);

  useEffect(() => { setNotes(iteration.notes ?? ""); }, [iteration.notes, iteration.id]);
  useEffect(() => { setNameDraft(iteration.name ?? ""); }, [iteration.name, iteration.id]);
  useEffect(() => { setFabricType(iteration.fabricType ?? ""); }, [iteration.fabricType, iteration.id]);
  useEffect(() => { setSampleSize(iteration.sampleSize ?? ""); }, [iteration.sampleSize, iteration.id]);
  useEffect(() => { setButtonType(iteration.buttonType ?? ""); }, [iteration.buttonType, iteration.id]);
  useEffect(() => { setFactoryCost(iteration.factoryCost ?? ""); }, [iteration.factoryCost, iteration.id]);
  useEffect(() => { setStatus(iteration.status); }, [iteration.status, iteration.id]);
  useEffect(() => {
    // Sync savedCount up if the loader (or background hook) has written
    // more images than we know about. Don't reset down; our optimistic
    // adds shouldn't be wiped by a slim loader payload.
    const propCount = typeof iteration.imageCount === "number"
      ? iteration.imageCount
      : (Array.isArray(iteration.images) ? (iteration.images as string[]).length : 0);
    setSavedCount((cur) => (propCount > cur ? propCount : cur));
  }, [iteration.imageCount, iteration.images, iteration.id]);

  // Revoke pending blob URLs when the iteration changes / unmounts.
  useEffect(() => () => {
    setPendingImages((cur) => {
      cur.forEach((p) => URL.revokeObjectURL(p.blobUrl));
      return [];
    });
  }, [iteration.id]);

  const submitUpdate = (fields: Record<string, string>) => {
    fetcher.submit({ intent: "update_sample_iteration", iterationId: String(iteration.id), ...fields }, { method: "post" });
  };

  const saveName = () => {
    setNameEditing(false);
    const trimmed = nameDraft.trim();
    if (trimmed !== (iteration.name ?? "")) submitUpdate({ name: trimmed });
  };

  const saveNotes = () => {
    const previous = iteration.notes ?? "";
    if (notes === previous) return;
    // If the user appended new text at the end of the existing note, prefix
    // that new text with their name so the next reader knows who wrote it.
    // If they edited the middle (or wholly rewrote), save as-is — auto-
    // prefixing would mangle their edits.
    let toSave = notes;
    const authorName = (currentUser?.name ?? "").trim().split(/\s+/)[0];
    if (authorName && notes.startsWith(previous)) {
      const addition = notes.slice(previous.length);
      const trimmedAddition = addition.replace(/^\s+/, "");
      if (trimmedAddition.length > 0) {
        const sep = previous.length === 0 ? "" : (previous.endsWith("\n") ? "" : "\n");
        toSave = previous + sep + `${authorName}: ${trimmedAddition}`;
        // Reflect the saved value in local state so subsequent typing is
        // treated as a further append (not a rewrite of the prefix).
        setNotes(toSave);
      }
    }
    submitUpdate({ notes: toSave });
  };

  const addImages = async (files: File[]) => {
    if (files.length === 0) return;
    const previews: PendingImage[] = files.map((f) => ({
      id: typeof crypto !== "undefined" && typeof crypto.randomUUID === "function" ? crypto.randomUUID() : `p_${Date.now()}_${Math.random()}`,
      blobUrl: URL.createObjectURL(f),
    }));
    setPendingImages((cur) => [...cur, ...previews]);

    const dataUrls = await Promise.all(files.map((file) => compressImageToDataUrl(file)));
    const wasEmpty = savedCount === 0;
    const payload: Record<string, string> = { addImages: JSON.stringify(dataUrls) };
    if (wasEmpty && dataUrls.length > 0) {
      const thumb = await generateThumbnail(dataUrls[0]);
      if (thumb) payload.thumbnail = thumb;
    }
    const result = await postPortalAction({
      intent: "update_sample_iteration",
      iterationId: String(iteration.id),
      ...payload,
    });
    if (!result.ok) return;
    const promotedIds = new Set(previews.map((p) => p.id));
    setSavedCount((c) => c + previews.length);
    setVersion(Date.now());
    setPendingImages((cur) => {
      cur.filter((p) => promotedIds.has(p.id)).forEach((p) => URL.revokeObjectURL(p.blobUrl));
      return cur.filter((p) => !promotedIds.has(p.id));
    });
  };

  const [pendingDeleteIndex, setPendingDeleteIndex] = useState<number | null>(null);
  const doRemoveSavedImage = async (index: number) => {
    setSavedCount((c) => Math.max(0, c - 1));
    await postPortalAction({
      intent: "update_sample_iteration",
      iterationId: String(iteration.id),
      removeImageIndex: String(index),
    });
    setVersion(Date.now());
  };
  const requestRemoveSavedImage = (index: number) => {
    if (shouldSkipDeleteConfirm()) {
      void doRemoveSavedImage(index);
      return;
    }
    setPendingDeleteIndex(index);
  };

  const removePendingImage = (id: string) => {
    setPendingImages((cur) => {
      const found = cur.find((p) => p.id === id);
      if (found) URL.revokeObjectURL(found.blobUrl);
      return cur.filter((p) => p.id !== id);
    });
  };

  const handleNotesChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    const val = e.target.value;
    setNotes(val);
    const pos = e.target.selectionStart ?? val.length;
    const textBefore = val.slice(0, pos);
    const match = textBefore.match(/@([a-zA-Z0-9._-]*)$/);
    if (match) {
      setMentionQuery(match[1]);
      setMentionStart(pos - match[0].length);
    } else {
      setMentionQuery(null);
      setMentionStart(-1);
    }
  };

  const selectMention = (user: PortalUser) => {
    const tag = user.name.trim().split(/\s+/)[0].toLowerCase();
    const before = notes.slice(0, mentionStart);
    const after = notes.slice(mentionStart + 1 + (mentionQuery ?? "").length);
    const inserted = `@${tag} `;
    const newNotes = before + inserted + after;
    setNotes(newNotes);
    setMentionQuery(null);
    setMentionStart(-1);
    setTimeout(() => {
      const ta = notesRef.current;
      if (ta) {
        const newPos = before.length + inserted.length;
        ta.focus();
        ta.setSelectionRange(newPos, newPos);
      }
    }, 0);
  };

  const filteredMentionUsers = mentionQuery !== null
    ? users.filter((u) => u.name.toLowerCase().startsWith(mentionQuery!.toLowerCase()) || mentionQuery === "")
    : [];

  const versionLabel = iteration.name
    ? `Version ${iteration.version} — ${iteration.name}`
    : `Version ${iteration.version}`;

  return (
    <div style={s.sampleIterationBlock}>
      {/* Header row */}
      <div style={s.sampleIterationHeader}>
        <div style={{ flex: 1, minWidth: 0 }}>
          {nameEditing ? (
            <input
              autoFocus
              style={{ fontSize: 13, fontWeight: 700, color: "#111827", border: "2px solid #2563eb", borderRadius: 5, padding: "2px 7px", outline: "none", width: "100%", fontFamily: "inherit" }}
              value={nameDraft}
              placeholder={`Version ${iteration.version} name (optional)`}
              onChange={(e) => setNameDraft(e.target.value)}
              onBlur={saveName}
              onKeyDown={(e) => { if (e.key === "Enter") (e.target as HTMLInputElement).blur(); if (e.key === "Escape") { setNameDraft(iteration.name ?? ""); setNameEditing(false); } }}
            />
          ) : (
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <span style={{ fontSize: 13, fontWeight: 700, color: "#111827" }}>{versionLabel}</span>
              <button
                type="button"
                onClick={() => setNameEditing(true)}
                style={{ background: "none", border: "none", padding: "2px 4px", cursor: "pointer", color: "#9ca3af", lineHeight: 1, borderRadius: 4, display: "flex", alignItems: "center" }}
                title="Edit version name"
              >
                <svg width="13" height="13" viewBox="0 0 20 20" fill="currentColor"><path d="M13.586 3.586a2 2 0 112.828 2.828l-.793.793-2.828-2.828.793-.793zM11.379 5.793L3 14.172V17h2.828l8.38-8.379-2.83-2.828z"/></svg>
              </button>
            </div>
          )}
        </div>
        <select
          value={status}
          onChange={(e) => {
            const next = e.target.value;
            setStatus(next); // optimistic — show immediately, no waiting on the POST
            onLocalUpdate(iteration.id, { status: next });
            submitUpdate({ status: next });
          }}
          style={{
            ...sampleStatusPillStyle(status),
            border: "1px solid rgba(0,0,0,0.12)",
            cursor: "pointer",
            appearance: "none" as const,
            paddingRight: 22,
            backgroundImage: `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6' viewBox='0 0 10 6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%236b7280'/%3E%3C/svg%3E")`,
            backgroundRepeat: "no-repeat",
            backgroundPosition: "right 6px center",
            backgroundSize: "8px",
          }}
        >
          <option value="under_consideration">Under Consideration</option>
          <option value="given_to_factory">Given to Factory</option>
          <option value="sent">Sent</option>
          <option value="sample_in_production">Sample in Production</option>
          <option value="changes_requested">Changes Requested</option>
          <option value="approved">Approved</option>
          <option value="approved_for_production">Approved on Production</option>
        </select>
        <span style={s.sampleIterationDate}>{formatPortalDate(iteration.createdAt)}</span>
        <button
          type="button"
          style={s.removeUserButton}
          onClick={() => {
            if (window.confirm("Delete this version?")) {
              fetcher.submit({ intent: "delete_sample_iteration", iterationId: String(iteration.id) }, { method: "post" });
            }
          }}
        >Delete version</button>
      </div>

      {/* Images */}
      <div style={s.sampleIterationImages}>
        {Array.from({ length: savedCount }).map((_, index) => (
          <div key={`saved-${index}`} style={s.sampleIterationImageWrap}>
            <img
              src={`/portal/image/sample/${iteration.id}/${index}?v=${version}`}
              alt={`v${iteration.version} image ${index + 1}`}
              style={{ ...s.sampleIterationImage, cursor: "zoom-in" }}
              loading="lazy"
              decoding="async"
              onClick={() => setLightboxIndex(index)}
            />
            <button type="button" style={s.sampleIterationImageRemove} onClick={() => requestRemoveSavedImage(index)} aria-label="Remove image">×</button>
          </div>
        ))}
        {pendingImages.map((p, i) => (
          <div key={`pending-${p.id}`} style={s.sampleIterationImageWrap}>
            <img
              src={p.blobUrl}
              alt={`uploading ${i + 1}`}
              style={{ ...s.sampleIterationImage, opacity: 0.6 }}
            />
            <button type="button" style={s.sampleIterationImageRemove} onClick={() => removePendingImage(p.id)} aria-label="Cancel upload">×</button>
          </div>
        ))}
        <button type="button" style={s.sampleIterationAddImage} onClick={() => setUploadModalOpen(true)}>
          <span>+ Add photos</span>
        </button>
      </div>

      {/* Fabric / Size / Button / Factory cost fields */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 0, borderTop: "1px solid #f1f5f9" }}>
        {([
          { label: "Fabric type", value: fabricType, setter: setFabricType, field: "fabricType" },
          { label: "Sample size", value: sampleSize, setter: setSampleSize, field: "sampleSize" },
          { label: "Button type", value: buttonType, setter: setButtonType, field: "buttonType" },
          { label: "Factory cost", value: factoryCost, setter: setFactoryCost, field: "factoryCost" },
        ] as const).map(({ label, value, setter, field }, i) => (
          <div key={field} style={{ borderRight: i < 3 ? "1px solid #f1f5f9" : "none", padding: "10px 14px" }}>
            <div style={{ fontSize: 11, fontWeight: 600, color: "#9ca3af", marginBottom: 4, textTransform: "uppercase" as const, letterSpacing: "0.04em" }}>{label}</div>
            <input
              style={{ width: "100%", border: "none", outline: "none", fontSize: "var(--portal-panel-font-size, 13px)", color: "#111827", background: "transparent", fontFamily: "inherit", padding: 0 }}
              value={value}
              placeholder={`Enter ${label.toLowerCase()}`}
              onChange={(e) => setter(e.target.value)}
              onBlur={(e) => { const v = e.target.value; if (v !== (iteration[field as keyof SampleIterationType] ?? "")) submitUpdate({ [field]: v }); }}
            />
          </div>
        ))}
      </div>

      <div style={{ position: "relative" }}>
        <div style={{ padding: "10px 16px 0", fontSize: 11, fontWeight: 600, color: "#9ca3af", textTransform: "uppercase", letterSpacing: "0.04em", borderTop: "1px solid #f1f5f9", background: "#fafafa" }}>
          Notes for version {iteration.version}
        </div>
        <textarea
          ref={notesRef}
          style={s.sampleIterationNotes}
          value={notes}
          placeholder="Notes, change requests, measurements… type @ to tag someone"
          onChange={handleNotesChange}
          onBlur={() => { setMentionQuery(null); saveNotes(); }}
          onKeyDown={(e) => {
            if (mentionQuery !== null && filteredMentionUsers.length > 0) {
              if (e.key === "Escape") { e.preventDefault(); setMentionQuery(null); }
            }
          }}
          rows={3}
        />
        {mentionQuery !== null && filteredMentionUsers.length > 0 && (
          <div style={{ position: "absolute", bottom: "calc(100% + 4px)", left: 0, background: "#fff", border: "1px solid #e5e7eb", borderRadius: 8, boxShadow: "0 4px 16px rgba(0,0,0,0.14)", zIndex: 300, minWidth: 180, padding: 4 }}>
            {filteredMentionUsers.map((user) => (
              <button
                key={user.id}
                type="button"
                onMouseDown={(e) => { e.preventDefault(); selectMention(user); }}
                style={{ display: "flex", alignItems: "center", gap: 8, width: "100%", padding: "7px 10px", background: "none", border: "none", cursor: "pointer", fontSize: 13, color: "#111827", borderRadius: 5, textAlign: "left" }}
              >
                <span style={{ fontWeight: 600, color: "#2563eb" }}>@{user.name.split(/\s+/)[0].toLowerCase()}</span>
                <span style={{ color: "#6b7280" }}>{user.name}</span>
              </button>
            ))}
          </div>
        )}
      </div>

      {lightboxIndex !== null && typeof document !== "undefined" && createPortal(
        <ImageLightbox
          images={[
            ...Array.from({ length: savedCount }, (_, i) => `/portal/image/sample/${iteration.id}/${i}?v=${version}`),
            ...pendingImages.map((p) => p.blobUrl),
          ]}
          initialIndex={lightboxIndex}
          onClose={() => setLightboxIndex(null)}
        />,
        document.body,
      )}
      {uploadModalOpen && typeof document !== "undefined" && createPortal(
        <ImageUploadModal
          title="Add photos"
          onImage={(file) => addImages([file])}
          onClose={() => setUploadModalOpen(false)}
          multi
        />,
        document.body,
      )}
      {pendingDeleteIndex !== null && (
        <ConfirmDeleteModal
          title="Are you sure you want to delete this image?"
          subtitle={versionLabel}
          onCancel={() => setPendingDeleteIndex(null)}
          onConfirm={() => {
            const idx = pendingDeleteIndex;
            setPendingDeleteIndex(null);
            void doRemoveSavedImage(idx);
          }}
        />
      )}
    </div>
  );
}

// Cancel / OK / Don't show for 24 hours — the same modal shape used for
// restock-row deletes. Reuses the existing DELETE_CONFIRM_SKIP_KEY so opting
// out of confirmations in one place opts out everywhere.
function shouldSkipDeleteConfirm(): boolean {
  if (typeof window === "undefined") return false;
  const skipUntil = Number(window.localStorage.getItem(DELETE_CONFIRM_SKIP_KEY) ?? 0);
  return skipUntil > Date.now();
}
function skipDeleteConfirmForDay(): void {
  if (typeof window === "undefined") return;
  window.localStorage.setItem(DELETE_CONFIRM_SKIP_KEY, String(Date.now() + 24 * 60 * 60 * 1000));
}

function ConfirmDeleteModal({
  title,
  subtitle,
  onConfirm,
  onCancel,
}: {
  title: string;
  subtitle?: string;
  onConfirm: () => void;
  onCancel: () => void;
}) {
  if (typeof document === "undefined") return null;
  // The shared deleteConfirm style sits at z-index 1000, but the sample /
  // vision drawers sit at 1200 — without an override the confirm modal
  // would be hidden behind the drawer it was opened from.
  return createPortal(
    <div style={{ ...s.deleteConfirm, zIndex: 1400 }} onClick={onCancel}>
      <div style={s.deleteConfirmCard} onClick={(e) => e.stopPropagation()}>
        <div style={s.deleteConfirmTitle}>{title}</div>
        {subtitle ? <div style={s.deleteConfirmText}>{subtitle}</div> : null}
        <div style={s.deleteConfirmActions}>
          <button type="button" style={s.deleteConfirmButton} onClick={onCancel}>Cancel</button>
          <button type="button" style={{ ...s.deleteConfirmButton, ...s.deleteConfirmDanger }} onClick={onConfirm}>OK</button>
          <button type="button" style={s.deleteConfirmButton} onClick={() => { skipDeleteConfirmForDay(); onConfirm(); }}>
            Don’t show for 24 hours
          </button>
        </div>
      </div>
    </div>,
    document.body,
  );
}

function ImageUploadModal({
  title,
  onImage,
  onClose,
  multi,
}: {
  title: string;
  onImage: (file: File) => void;
  onClose: () => void;
  multi?: boolean;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragOver, setDragOver] = useState(false);

  useEffect(() => {
    const handler = (e: KeyboardEvent) => { if (e.key === "Escape") onClose(); };
    document.addEventListener("keydown", handler);
    return () => document.removeEventListener("keydown", handler);
  }, [onClose]);

  const handleFiles = (files: FileList | null) => {
    if (!files) return;
    if (multi) {
      Array.from(files).forEach((f) => onImage(f));
    } else {
      onImage(files[0]);
    }
    onClose();
  };

  return (
    <div style={s.imageUploadModalBackdrop} onClick={onClose}>
      <div style={s.imageUploadModal} onClick={(e) => e.stopPropagation()}>
        <div style={s.imageUploadModalHeader}>
          <span style={{ fontWeight: 700, fontSize: 15 }}>{title}</span>
          <button type="button" onClick={onClose} style={{ background: "none", border: "none", fontSize: 20, cursor: "pointer", color: "#6b7280", lineHeight: 1 }}>×</button>
        </div>
        <div
          style={{ ...s.imageUploadDropZone, ...(dragOver ? s.imageUploadDropZoneActive : {}) }}
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFiles(e.dataTransfer.files); }}
          onClick={() => inputRef.current?.click()}
        >
          <div style={s.imageUploadDropIcon}>📷</div>
          <div style={s.imageUploadDropText}>Drop {multi ? "images" : "an image"} here</div>
          <div style={s.imageUploadDropSubtext}>or click to browse</div>
        </div>
        <input
          ref={inputRef}
          type="file"
          accept="image/*"
          multiple={!!multi}
          style={{ display: "none" }}
          onChange={(e) => handleFiles(e.target.files)}
        />
      </div>
    </div>
  );
}

function ImageLightbox({
  images,
  initialIndex,
  onClose,
}: {
  images: string[];
  initialIndex: number;
  onClose: () => void;
}) {
  const [index, setIndex] = useState(initialIndex);

  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if (e.key === "Escape") onClose();
      if (e.key === "ArrowLeft") setIndex((i) => Math.max(0, i - 1));
      if (e.key === "ArrowRight") setIndex((i) => Math.min(images.length - 1, i + 1));
    };
    document.addEventListener("keydown", handler);
    return () => document.removeEventListener("keydown", handler);
  }, [onClose, images.length]);

  return (
    <div style={s.lightboxBackdrop} onClick={onClose}>
      <button type="button" style={s.lightboxClose} onClick={onClose} aria-label="Close">×</button>
      <button
        type="button"
        style={{ ...s.lightboxArrow, ...s.lightboxArrowLeft, opacity: index === 0 ? 0.25 : 1 }}
        onClick={(e) => { e.stopPropagation(); setIndex((i) => Math.max(0, i - 1)); }}
        aria-label="Previous"
        disabled={index === 0}
      >‹</button>
      <img
        src={images[index]}
        alt={`${index + 1} of ${images.length}`}
        style={s.lightboxImage}
        onClick={(e) => e.stopPropagation()}
      />
      <button
        type="button"
        style={{ ...s.lightboxArrow, ...s.lightboxArrowRight, opacity: index === images.length - 1 ? 0.25 : 1 }}
        onClick={(e) => { e.stopPropagation(); setIndex((i) => Math.min(images.length - 1, i + 1)); }}
        aria-label="Next"
        disabled={index === images.length - 1}
      >›</button>
      <div style={s.lightboxCounter}>{index + 1} / {images.length}</div>
    </div>
  );
}

// ─── Vision Board ─────────────────────────────────────────────────────────────
// V2 components live in app/portal-vision-board.tsx. The loader feeds them
// slim data (active board's items, no fields/notes/images) and the action
// handles the vb_* intents.

// Generate a tiny thumbnail for the card grid (~5–15 KB). Re-uses the same
// canvas pipeline as the full-image compressor but with smaller targets.
async function generateThumbnail(input: string, maxDim = 240, quality = 0.62): Promise<string | null> {
  if (typeof input !== "string" || !input.startsWith("data:image/")) return null;
  try {
    const img = await new Promise<HTMLImageElement>((resolve, reject) => {
      const i = new Image();
      i.onload = () => resolve(i);
      i.onerror = () => reject(new Error("decode failed"));
      i.src = input;
    });
    const longEdge = Math.max(img.width, img.height);
    const scale = longEdge > maxDim ? maxDim / longEdge : 1;
    const w = Math.max(1, Math.round(img.width * scale));
    const h = Math.max(1, Math.round(img.height * scale));
    const canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext("2d");
    if (!ctx) return null;
    ctx.drawImage(img, 0, 0, w, h);
    return canvas.toDataURL("image/jpeg", quality);
  } catch {
    return null;
  }
}

// Compress an existing data URL (e.g. legacy stored base64). Aggressive defaults
// suited to a personal-use board: max 800px on the long edge, JPEG q=0.75.
// Typical photos drop from several MB to ~50-100KB.
// Hard cap: every uploaded image (after compression) must be at most 500 KB
// in binary form. Below the cap means the cards / drawer images load fast
// and the DB stores small JSONB payloads.
const IMAGE_TARGET_BYTES = 500 * 1024;

// Estimate the decoded binary size of a data URL.
function dataUrlBinaryBytes(url: string): number {
  const idx = url.indexOf(",");
  if (idx < 0) return url.length;
  const b64 = url.slice(idx + 1);
  let padding = 0;
  if (b64.endsWith("==")) padding = 2;
  else if (b64.endsWith("=")) padding = 1;
  return Math.max(0, Math.floor((b64.length * 3) / 4) - padding);
}

function renderJpegToDataUrl(img: HTMLImageElement, maxDim: number, quality: number): string | null {
  const longEdge = Math.max(img.width, img.height);
  const scale = longEdge > maxDim ? maxDim / longEdge : 1;
  const w = Math.max(1, Math.round(img.width * scale));
  const h = Math.max(1, Math.round(img.height * scale));
  const canvas = document.createElement("canvas");
  canvas.width = w;
  canvas.height = h;
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;
  ctx.drawImage(img, 0, 0, w, h);
  return canvas.toDataURL("image/jpeg", quality);
}

// Encode `img` as JPEG, then keep tightening (lower quality, then smaller
// dimensions) until the result fits under IMAGE_TARGET_BYTES — or we hit
// the floor (q 0.35 / 400px) and return whatever's smallest.
async function compressBelowTarget(img: HTMLImageElement, startMaxDim: number, startQuality: number): Promise<string | null> {
  let dim = startMaxDim;
  let q = startQuality;
  let attempt = renderJpegToDataUrl(img, dim, q);
  if (!attempt) return null;
  let attempts = 0;
  while (dataUrlBinaryBytes(attempt) > IMAGE_TARGET_BYTES && attempts < 10) {
    if (q > 0.4) q = Math.max(0.4, q - 0.1);
    else if (dim > 400) dim = Math.max(400, dim - 150);
    else break;
    const next = renderJpegToDataUrl(img, dim, q);
    if (!next) break;
    attempt = next;
    attempts++;
  }
  return attempt;
}

async function compressDataUrl(input: string, maxDim = 800, quality = 0.75): Promise<string> {
  if (typeof input !== "string" || !input.startsWith("data:image/")) return input;
  // Already small enough — leave it alone.
  if (dataUrlBinaryBytes(input) <= IMAGE_TARGET_BYTES * 0.7) return input;
  try {
    const img = await new Promise<HTMLImageElement>((resolve, reject) => {
      const i = new Image();
      i.onload = () => resolve(i);
      i.onerror = () => reject(new Error("decode failed"));
      i.src = input;
    });
    const out = await compressBelowTarget(img, maxDim, quality);
    if (!out) return input;
    return out.length < input.length ? out : input;
  } catch {
    return input;
  }
}

// POST a form to the current portal route's action and parse its JSON response.
// Used by background workflows (image compression, drawer uploads). Always
// passes noRevalidate=1 so the call doesn't trigger a full route loader re-run
// on top of the action.
//
// Returns { ok: false } only on a real HTTP failure. A 200 with an empty or
// `null` JSON body (which is what most of the route's action handlers return)
// is reported as { ok: true, data: null }. Callers should branch on `ok`,
// not on `data`, otherwise success-with-null-body is mistaken for failure.
type PortalActionResult<T> = { ok: boolean; data: T | null };
async function postPortalAction<T = unknown>(data: Record<string, string>): Promise<PortalActionResult<T>> {
  const fd = new FormData();
  for (const [k, v] of Object.entries(data)) fd.set(k, v);
  if (!fd.has("noRevalidate")) fd.set("noRevalidate", "1");
  let res: Response;
  try {
    res = await fetch(window.location.pathname + window.location.search, {
      method: "POST",
      body: fd,
      headers: { Accept: "application/json" },
    });
  } catch {
    return { ok: false, data: null };
  }
  if (!res.ok) return { ok: false, data: null };
  try {
    const body = await res.json();
    return { ok: true, data: body as T | null };
  } catch {
    return { ok: true, data: null };
  }
}

// Silently re-compress legacy uncompressed images in the background after page
// mount. Tracks done IDs in localStorage so each item is processed at most once
// per browser. Skips items whose images are already small.
// Bump these to v2 so any browser that already ran the v1 backfill (which
// only re-compressed images and didn't generate the new dedicated thumbnail
// column) re-runs and populates thumbnails for every item.
const COMPRESSED_IDS_VISION_KEY = "portal-compressed-vision-items-v2";
const COMPRESSED_IDS_SAMPLE_ITER_KEY = "portal-compressed-sample-iters-v2";

function loadCompressedIds(key: string): Set<number> {
  try {
    const raw = typeof window !== "undefined" ? window.localStorage.getItem(key) : null;
    if (!raw) return new Set();
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? new Set(arr.filter((n: unknown): n is number => typeof n === "number")) : new Set();
  } catch { return new Set(); }
}
function saveCompressedIds(key: string, ids: Set<number>): void {
  try {
    if (typeof window !== "undefined") window.localStorage.setItem(key, JSON.stringify(Array.from(ids)));
  } catch { /* quota or private mode — fine to lose */ }
}

function useAutoCompressSampleIterations(samples: SampleType[]): void {
  useEffect(() => {
    if (typeof window === "undefined") return;
    type Task = { sampleId: number; iterationId: number };
    const candidates: Task[] = [];
    for (const s of samples) {
      for (const it of s.iterations) {
        const count = it.imageCount ?? (Array.isArray(it.images) ? (it.images as string[]).length : 0);
        // Skip iterations the loader has already told us are processed.
        if (count > 0 && !it.hasThumbnail) candidates.push({ sampleId: s.id, iterationId: it.id });
      }
    }
    if (candidates.length === 0) return;
    let cancelled = false;
    const done = loadCompressedIds(COMPRESSED_IDS_SAMPLE_ITER_KEY);
    const sampleCache = new Map<number, { iterations?: Array<{ id: number; images?: unknown; thumbnail?: string | null }> } | null>();

    (async () => {
      for (const task of candidates) {
        if (cancelled) return;
        if (done.has(task.iterationId)) continue;
        try {
          let sample = sampleCache.get(task.sampleId);
          if (sample === undefined) {
            const r = await postPortalAction<{ sample: { iterations?: Array<{ id: number; images?: unknown; thumbnail?: string | null }> } | null }>({
              intent: "get_sample_full",
              sampleId: String(task.sampleId),
            });
            sample = r.data?.sample ?? null;
            sampleCache.set(task.sampleId, sample);
          }
          const iter = sample?.iterations?.find((it) => it.id === task.iterationId);
          const images = Array.isArray(iter?.images) ? (iter!.images as string[]) : [];
          const existingThumb = iter?.thumbnail ?? null;
          const totalBefore = images.reduce((sum, s) => sum + (typeof s === "string" ? s.length : 0), 0);
          const imagesAlreadySmall = totalBefore < 150_000 * Math.max(1, images.length);

          let thumbToSend: string | null = null;
          if (!existingThumb && images.length > 0) {
            thumbToSend = await generateThumbnail(images[0]);
          }
          let compressedToSend: string[] | null = null;
          if (!imagesAlreadySmall && images.length > 0) {
            const compressed = await Promise.all(images.map((img) => compressDataUrl(img)));
            const totalAfter = compressed.reduce((sum, s) => sum + s.length, 0);
            if (totalAfter < totalBefore * 0.95) compressedToSend = compressed;
          }

          if (thumbToSend || compressedToSend) {
            const payload: Record<string, string> = {
              intent: "update_sample_iteration",
              iterationId: String(task.iterationId),
            };
            if (compressedToSend) {
              payload.imagesReplace = JSON.stringify(compressedToSend);
              if (!thumbToSend && compressedToSend.length > 0) {
                thumbToSend = await generateThumbnail(compressedToSend[0]);
              }
            }
            if (thumbToSend) payload.thumbnail = thumbToSend;
            await postPortalAction(payload);
          }
          done.add(task.iterationId);
          saveCompressedIds(COMPRESSED_IDS_SAMPLE_ITER_KEY, done);
        } catch { /* keep going */ }
      }
    })();

    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [samples.length]);
}

// Compress an image File client-side: scale down to max 800px on the long edge
// and re-encode as JPEG q=0.75. Typical photos drop from several MB to ~50-100KB.
// Returns a data URL. Falls back to the original file if anything fails.
async function compressImageToDataUrl(file: File, maxDim = 800, quality = 0.75): Promise<string> {
  if (!file.type.startsWith("image/")) {
    return await new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(String(r.result));
      r.onerror = () => reject(r.error);
      r.readAsDataURL(file);
    });
  }
  const dataUrl = await new Promise<string>((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result));
    r.onerror = () => reject(r.error);
    r.readAsDataURL(file);
  });
  // Already small enough — skip the canvas round-trip entirely.
  if (file.size <= IMAGE_TARGET_BYTES * 0.7) return dataUrl;
  try {
    const img = await new Promise<HTMLImageElement>((resolve, reject) => {
      const i = new Image();
      i.onload = () => resolve(i);
      i.onerror = () => reject(new Error("image decode failed"));
      i.src = dataUrl;
    });
    const out = await compressBelowTarget(img, maxDim, quality);
    return out ?? dataUrl;
  } catch {
    return dataUrl;
  }
}


// ─── Collections ─────────────────────────────────────────────────────────────

type CollectionListItem = {
  id: number;
  name: string;
  sortOrder: number;
  hasThumbnail: boolean;
  rowCount: number;
  createdAt: Date | string;
  updatedAt: Date | string;
};

type CollectionColumnDef = {
  id: string;
  label: string;
  type?: "text" | "number" | "date";
  width?: number;
};

type CollectionFullType = {
  id: number;
  name: string;
  sortOrder: number;
  thumbnail: string | null;
  columns: unknown;
  rows: unknown;
  createdAt: Date | string;
  updatedAt: Date | string;
};

// Default columns for a fresh collection — modelled after the spreadsheet
// layout the user shared. Each column is just a string field for v1; image /
// dropdown / date-picker cell types can be added per-column later.
const DEFAULT_COLLECTION_COLUMNS: CollectionColumnDef[] = [
  { id: "release", label: "Release", width: 90 },
  { id: "modelPicture", label: "Model PICTURE", width: 120 },
  { id: "fabric", label: "FABRIC", width: 120 },
  { id: "name", label: "Name", width: 160 },
  { id: "notes", label: "Notes", width: 140 },
  { id: "sku", label: "SKU", width: 80 },
  { id: "xs", label: "XS", type: "number", width: 50 },
  { id: "s", label: "S", type: "number", width: 50 },
  { id: "m", label: "M", type: "number", width: 50 },
  { id: "l", label: "L", type: "number", width: 50 },
  { id: "xl", label: "XL", type: "number", width: 50 },
  { id: "xxl", label: "2XL", type: "number", width: 50 },
  { id: "xxxl", label: "3XL", type: "number", width: 50 },
  { id: "sm", label: "S/M", type: "number", width: 50 },
  { id: "ml", label: "M/L", type: "number", width: 50 },
  { id: "lxl", label: "L/XL", type: "number", width: 50 },
  { id: "totalOrdered", label: "TOTAL Ordered", type: "number", width: 100 },
  { id: "status", label: "STATUS", width: 110 },
  { id: "sample", label: "Sample", width: 110 },
  { id: "sampleReceived", label: "Sample RECEIVED", type: "date", width: 120 },
  { id: "price", label: "Price", type: "number", width: 70 },
  { id: "cost", label: "Cost", type: "number", width: 70 },
  { id: "eta", label: "ETA", type: "date", width: 90 },
  { id: "maniPicsTaken", label: "mani Pics Taken", width: 130 },
  { id: "loadingNotes", label: "Loading Notes", width: 140 },
  { id: "duplicateFrom", label: "DUPLICATE FROM", width: 140 },
  { id: "modelHeightSize", label: "Model height and size", width: 130 },
  { id: "createdBy", label: "Created by", width: 100 },
  { id: "link", label: "Link", width: 130 },
  { id: "title", label: "Title", width: 120 },
  { id: "description", label: "Description", width: 160 },
  { id: "categories", label: "Categories", width: 110 },
  { id: "productType", label: "Product type", width: 110 },
  { id: "collectionsField", label: "Collections", width: 110 },
  { id: "tags", label: "Tags", width: 100 },
  { id: "variants", label: "Variants", width: 100 },
  { id: "hsCode", label: "HS Code", width: 90 },
  { id: "countryOfOrigin", label: "Country of Origin", width: 130 },
  { id: "compareAtPrice", label: "Compare at price", type: "number", width: 110 },
  { id: "complProducts", label: "Compl. products", width: 130 },
  { id: "colour", label: "Colour", width: 90 },
  { id: "sizeGuide", label: "Size guide", width: 110 },
  { id: "seoTitleDesc", label: "SEO (title, desc)", width: 130 },
  { id: "activation", label: "Activation", width: 100 },
  { id: "schedules", label: "Schedules", width: 100 },
  { id: "reviews", label: "Reviews", width: 90 },
  { id: "swatches", label: "Swatches", width: 100 },
];

function normalizeCollectionColumns(value: unknown): CollectionColumnDef[] {
  if (!Array.isArray(value) || value.length === 0) return DEFAULT_COLLECTION_COLUMNS;
  return (value as Array<Record<string, unknown>>).map((c, i) => ({
    id: typeof c?.id === "string" ? c.id : `col_${i}`,
    label: typeof c?.label === "string" ? c.label : `Column ${i + 1}`,
    type: c?.type === "number" || c?.type === "date" ? c.type : "text",
    width: typeof c?.width === "number" ? c.width : undefined,
  }));
}
function normalizeCollectionRows(value: unknown): Record<string, string>[] {
  if (!Array.isArray(value)) return [];
  return (value as Array<Record<string, unknown>>).map((row) => {
    const out: Record<string, string> = {};
    for (const [k, v] of Object.entries(row ?? {})) out[k] = v == null ? "" : String(v);
    return out;
  });
}

function CollectionsPanel({ collections: initialCollections }: { collections: CollectionListItem[] }) {
  const fetcher = useFetcher();
  const [searchParams, setSearchParams] = useSearchParams();
  const collectionIdParam = searchParams.get("collectionId");
  const selectedId = collectionIdParam ? Number(collectionIdParam) : null;
  const [collections, setCollections] = useState<CollectionListItem[]>(initialCollections);
  const [dragId, setDragId] = useState<number | null>(null);
  const [dragOverId, setDragOverId] = useState<number | null>(null);
  const deletedRef = useRef<Set<number>>(new Set());
  const [addOpen, setAddOpen] = useState(false);
  const [addName, setAddName] = useState("");

  useEffect(() => {
    setCollections(initialCollections.filter((c) => !deletedRef.current.has(c.id)));
  }, [initialCollections]);

  const openCollection = (id: number) => {
    const next = new URLSearchParams(searchParams);
    next.set("collectionId", String(id));
    setSearchParams(next, { replace: false });
  };
  const closeCollection = () => {
    const next = new URLSearchParams(searchParams);
    next.delete("collectionId");
    setSearchParams(next, { replace: false });
  };
  const selectedCollection = selectedId ? collections.find((c) => c.id === selectedId) ?? null : null;

  const submitAdd = () => {
    const name = addName.trim() || "Untitled collection";
    fetcher.submit({ intent: "add_collection", name }, { method: "post" });
    setAddOpen(false);
    setAddName("");
  };

  const handleDelete = (id: number) => {
    deletedRef.current.add(id);
    setCollections((prev) => prev.filter((c) => c.id !== id));
    if (selectedId === id) closeCollection();
    fetcher.submit({ intent: "delete_collection", collectionId: String(id) }, { method: "post" });
  };

  const handleRename = (id: number, name: string) => {
    setCollections((prev) => prev.map((c) => c.id === id ? { ...c, name } : c));
    fetcher.submit({ intent: "rename_collection", collectionId: String(id), name }, { method: "post" });
  };

  const reorder = (targetId: number) => {
    if (!dragId || dragId === targetId) return;
    const fromIdx = collections.findIndex((c) => c.id === dragId);
    const toIdx = collections.findIndex((c) => c.id === targetId);
    if (fromIdx < 0 || toIdx < 0) return;
    const next = [...collections];
    const [moved] = next.splice(fromIdx, 1);
    next.splice(toIdx, 0, moved);
    setCollections(next);
    fetcher.submit({ intent: "reorder_collections", collectionIds: JSON.stringify(next.map((c) => c.id)) }, { method: "post" });
  };

  // When a collection is selected via the URL, render the spreadsheet page
  // instead of the tile grid — full inline view, not an overlay.
  if (selectedCollection) {
    return (
      <CollectionSpreadsheetPage
        listItem={selectedCollection}
        onBack={closeCollection}
        onLocalNameChange={(name) => handleRename(selectedCollection.id, name)}
      />
    );
  }

  return (
    <div style={s.productInfoPage}>
      <div style={s.productInfoToolbar}>
        <div style={s.productInfoToolbarLeft}>
          <div>
            <h2 style={s.productInfoHeading}>Collections</h2>
            <div style={s.productInfoMeta}>{collections.length} collection{collections.length !== 1 ? "s" : ""}</div>
          </div>
        </div>
        <div style={s.productInfoActions}>
          <button type="button" style={s.primaryActionButton} onClick={() => { setAddName(""); setAddOpen(true); }}>
            Add Collection
          </button>
        </div>
      </div>

      <div style={{ ...s.productInfoList, gridTemplateColumns: "repeat(5, minmax(0, 1fr))" }}>
        {collections.map((c) => (
          <CollectionCard
            key={c.id}
            collection={c}
            isDragging={dragId === c.id}
            isDragOver={dragOverId === c.id && dragId !== c.id}
            onOpen={() => openCollection(c.id)}
            onRename={(name) => handleRename(c.id, name)}
            onDelete={() => handleDelete(c.id)}
            onDragStart={(e) => { setDragId(c.id); e.dataTransfer.effectAllowed = "move"; }}
            onDragOver={(e) => { if (!dragId) return; e.preventDefault(); setDragOverId(c.id); }}
            onDragLeave={() => setDragOverId((cur) => cur === c.id ? null : cur)}
            onDrop={(e) => { e.preventDefault(); reorder(c.id); setDragId(null); setDragOverId(null); }}
            onDragEnd={() => { setDragId(null); setDragOverId(null); }}
          />
        ))}
        {collections.length === 0 && (
          <div style={{ gridColumn: "1 / -1", padding: "48px 0", textAlign: "center", color: "#9ca3af", fontSize: 14 }}>
            No collections yet. Click Add Collection to create your first one.
          </div>
        )}
      </div>

      {addOpen && typeof document !== "undefined" && createPortal(
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 1300, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div style={{ background: "#fff", borderRadius: 10, padding: 22, minWidth: 360, boxShadow: "0 20px 50px rgba(0,0,0,0.25)" }} onClick={(e) => e.stopPropagation()}>
            <h3 style={{ margin: "0 0 12px", fontSize: 16 }}>New Collection</h3>
            <input
              autoFocus
              value={addName}
              onChange={(e) => setAddName(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter") submitAdd(); if (e.key === "Escape") setAddOpen(false); }}
              placeholder="Collection name"
              style={{ width: "100%", border: "1px solid #d1d5db", borderRadius: 6, padding: "8px 10px", fontSize: 14, boxSizing: "border-box" }}
            />
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 14 }}>
              <button type="button" onClick={() => setAddOpen(false)} style={{ background: "#f3f4f6", border: "none", borderRadius: 6, padding: "8px 14px", fontSize: 13, cursor: "pointer" }}>Cancel</button>
              <button type="button" onClick={submitAdd} style={{ background: "#111827", color: "#fff", border: "none", borderRadius: 6, padding: "8px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>Create</button>
            </div>
          </div>
        </div>,
        document.body,
      )}
    </div>
  );
}

function CollectionCard({
  collection,
  isDragging,
  isDragOver,
  onOpen,
  onRename,
  onDelete,
  onDragStart,
  onDragOver,
  onDragLeave,
  onDrop,
  onDragEnd,
}: {
  collection: CollectionListItem;
  isDragging: boolean;
  isDragOver: boolean;
  onOpen: () => void;
  onRename: (name: string) => void;
  onDelete: () => void;
  onDragStart: (e: React.DragEvent<HTMLDivElement>) => void;
  onDragOver: (e: React.DragEvent<HTMLDivElement>) => void;
  onDragLeave: () => void;
  onDrop: (e: React.DragEvent<HTMLDivElement>) => void;
  onDragEnd: () => void;
}) {
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [hover, setHover] = useState(false);
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(collection.name);
  useEffect(() => { setDraft(collection.name); }, [collection.name]);

  return (
    <div
      style={{ ...s.productStyleCard, ...(isDragging ? s.productStyleCardDragging : {}), ...(isDragOver ? s.productStyleCardDropTarget : {}), cursor: "pointer" }}
      onClick={() => { if (!editing) onOpen(); }}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      onDragOver={onDragOver}
      onDragLeave={onDragLeave}
      onDrop={onDrop}
    >
      <span
        draggable
        style={{ ...s.productStyleDragHandle, pointerEvents: "auto", cursor: "grab", opacity: hover ? 0.6 : 0.25, transition: "opacity 0.15s" }}
        title="Drag to reorder"
        onDragStart={onDragStart}
        onDragEnd={onDragEnd}
        onClick={(e) => e.stopPropagation()}
      >::</span>
      {(hover || confirmDelete) && (
        <button
          type="button"
          title="Delete collection"
          onClick={(e) => { e.stopPropagation(); setConfirmDelete(true); }}
          style={{ position: "absolute", top: 8, right: 8, width: 26, height: 26, borderRadius: "50%", border: "none", background: "#ef4444", color: "#fff", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16, fontWeight: 700, lineHeight: 1, padding: 0, zIndex: 2, boxShadow: "0 2px 6px rgba(0,0,0,0.18)" }}
        >×</button>
      )}
      {confirmDelete && (
        <div style={{ position: "absolute", inset: 0, background: "rgba(0,0,0,0.6)", zIndex: 10, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 10, borderRadius: 8 }} onClick={(e) => e.stopPropagation()}>
          <div style={{ color: "#fff", fontSize: 13, fontWeight: 600 }}>Delete this collection?</div>
          <div style={{ display: "flex", gap: 8 }}>
            <button type="button" onClick={(e) => { e.stopPropagation(); onDelete(); }} style={{ padding: "7px 16px", borderRadius: 7, border: "none", background: "#ef4444", color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>Delete</button>
            <button type="button" onClick={(e) => { e.stopPropagation(); setConfirmDelete(false); }} style={{ padding: "7px 16px", borderRadius: 7, border: "1px solid rgba(255,255,255,0.3)", background: "transparent", color: "#fff", fontSize: 13, cursor: "pointer" }}>Cancel</button>
          </div>
        </div>
      )}
      <div style={{ ...s.productStyleImageWrap, aspectRatio: "1.3 / 1.8" }}>
        {collection.hasThumbnail ? (
          <img
            src={`/portal/thumbnail/collection/${collection.id}?v=${new Date(collection.updatedAt).getTime()}`}
            alt={collection.name}
            style={s.productStyleImage}
            loading="lazy"
            decoding="async"
          />
        ) : (
          <div style={s.productStyleImageEmpty}>No image yet</div>
        )}
      </div>
      <div style={{ ...s.productStyleCardBody, paddingBottom: 14 }}>
        {editing ? (
          <input
            autoFocus
            value={draft}
            onChange={(e) => setDraft(e.target.value)}
            onClick={(e) => e.stopPropagation()}
            onBlur={() => { setEditing(false); const t = draft.trim(); if (t && t !== collection.name) onRename(t); else setDraft(collection.name); }}
            onKeyDown={(e) => { if (e.key === "Enter") (e.target as HTMLInputElement).blur(); if (e.key === "Escape") { setDraft(collection.name); setEditing(false); } }}
            style={{ ...s.productStyleTitle, border: "1px solid #d1d5db", borderRadius: 4, padding: "2px 6px", textAlign: "center", width: "90%" }}
          />
        ) : (
          <span
            style={s.productStyleTitle}
            onClick={(e) => { e.stopPropagation(); setEditing(true); }}
            title="Click to rename"
          >{collection.name || "Untitled"}</span>
        )}
        <span style={{ fontSize: 11, color: "#9ca3af", marginTop: 2 }}>
          {collection.rowCount} row{collection.rowCount !== 1 ? "s" : ""}
        </span>
      </div>
    </div>
  );
}

function CollectionSpreadsheetPage({
  listItem,
  onBack,
  onLocalNameChange,
}: {
  listItem: CollectionListItem;
  onBack: () => void;
  onLocalNameChange: (name: string) => void;
}) {
  const fetcher = useFetcher();
  const loadFetcher = useFetcher<{ collection: CollectionFullType | null }>();
  const [columns, setColumns] = useState<CollectionColumnDef[]>(DEFAULT_COLLECTION_COLUMNS);
  const [rows, setRows] = useState<Record<string, string>[]>([]);
  const [loaded, setLoaded] = useState(false);
  const [nameDraft, setNameDraft] = useState(listItem.name);
  const [editingName, setEditingName] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    loadFetcher.submit(
      { intent: "get_collection_full", collectionId: String(listItem.id) },
      { method: "post" },
    );
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [listItem.id]);

  useEffect(() => {
    const col = loadFetcher.data?.collection;
    if (col && col.id === listItem.id) {
      setColumns(normalizeCollectionColumns(col.columns));
      setRows(normalizeCollectionRows(col.rows));
      setLoaded(true);
    }
  }, [loadFetcher.data, listItem.id]);

  const persistRows = (next: Record<string, string>[]) => {
    fetcher.submit(
      { intent: "update_collection", collectionId: String(listItem.id), rows: JSON.stringify(next) },
      { method: "post" },
    );
  };

  const updateCell = (rowIdx: number, colId: string, value: string) => {
    setRows((prev) => {
      const next = prev.map((r, i) => i === rowIdx ? { ...r, [colId]: value } : r);
      persistRows(next);
      return next;
    });
  };

  const addRow = () => {
    setRows((prev) => {
      const next = [...prev, {} as Record<string, string>];
      persistRows(next);
      return next;
    });
  };

  const removeRow = (idx: number) => {
    if (!window.confirm("Delete this row?")) return;
    setRows((prev) => {
      const next = prev.filter((_, i) => i !== idx);
      persistRows(next);
      return next;
    });
  };

  const saveName = () => {
    setEditingName(false);
    const t = nameDraft.trim();
    if (t && t !== listItem.name) {
      onLocalNameChange(t);
      fetcher.submit({ intent: "rename_collection", collectionId: String(listItem.id), name: t }, { method: "post" });
    } else {
      setNameDraft(listItem.name);
    }
  };

  const handleCoverUpload = async (file: File) => {
    const dataUrl = await compressImageToDataUrl(file);
    const thumb = await generateThumbnail(dataUrl);
    fetcher.submit(
      { intent: "update_collection", collectionId: String(listItem.id), thumbnail: thumb || dataUrl },
      { method: "post" },
    );
  };

  const noopResize = () => {};

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", minHeight: 0, gap: 14 }}>
      <div style={{ ...s.productInfoToolbar, flexShrink: 0 }}>
        <div style={s.productInfoToolbarLeft}>
          <button
            type="button"
            onClick={onBack}
            style={{ background: "transparent", border: "1px solid #d1d5db", color: "#374151", borderRadius: 6, padding: "6px 12px", fontSize: 13, fontWeight: 600, cursor: "pointer" }}
          >← Collections</button>
          <div>
            {editingName ? (
              <input
                autoFocus
                value={nameDraft}
                onChange={(e) => setNameDraft(e.target.value)}
                onBlur={saveName}
                onKeyDown={(e) => { if (e.key === "Enter") (e.target as HTMLInputElement).blur(); if (e.key === "Escape") { setNameDraft(listItem.name); setEditingName(false); } }}
                style={{ ...s.productInfoHeading, border: "1px solid #d1d5db", borderRadius: 6, padding: "2px 8px" }}
              />
            ) : (
              <h2 style={{ ...s.productInfoHeading, cursor: "pointer", margin: 0 }} onClick={() => setEditingName(true)} title="Click to rename">
                {listItem.name || "Untitled"}
              </h2>
            )}
            <div style={s.productInfoMeta}>{rows.length} row{rows.length !== 1 ? "s" : ""}</div>
          </div>
        </div>
        <div style={s.productInfoActions}>
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            style={{ background: "transparent", border: "1px solid #d1d5db", color: "#374151", borderRadius: 6, padding: "6px 12px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}
          >Set cover image</button>
          <input
            ref={fileInputRef}
            type="file"
            accept="image/*"
            style={{ display: "none" }}
            onChange={(e) => {
              const file = e.target.files?.[0];
              if (file) void handleCoverUpload(file);
              e.target.value = "";
            }}
          />
          <button type="button" onClick={addRow} style={s.primaryActionButton}>+ Add row</button>
        </div>
      </div>

      <div className="portal-table-scroll" style={{ ...s.tableWrap, flex: 1, maxHeight: "none", minHeight: 0 }}>
        {!loaded ? (
          <div style={{ padding: 40, textAlign: "center", color: "#94a3b8", fontSize: 13 }}>Loading…</div>
        ) : (
          <table style={s.table}>
            <colgroup>
              <col style={{ width: 48 }} />
              {columns.map((col) => (
                <col key={col.id} style={{ width: col.width ?? 110 }} />
              ))}
            </colgroup>
            <thead>
              <tr style={s.headerRow}>
                <th style={{ ...s.th, ...s.rowNumberHeader }}>#</th>
                {columns.map((col) => (
                  <Th
                    key={col.id}
                    headerKey={`collection:${col.id}`}
                    columnId={col.id}
                    onResizeStart={noopResize}
                  >
                    {col.label}
                  </Th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((row, rIdx) => (
                <tr key={rIdx} style={s.row}>
                  <RowNumberCell
                    rowNumber={rIdx + 1}
                    actions={[
                      { label: "Delete row", danger: true, onClick: () => removeRow(rIdx) },
                    ]}
                  />
                  {columns.map((col, colIdx) => (
                    <Td key={col.id} rowIndex={rIdx} colIndex={colIdx}>
                      <CollectionCell
                        value={row[col.id] ?? ""}
                        type={col.type ?? "text"}
                        columnId={col.id}
                        onCommit={(v) => updateCell(rIdx, col.id, v)}
                      />
                    </Td>
                  ))}
                </tr>
              ))}
              {rows.length === 0 && (
                <tr style={s.row}>
                  <td colSpan={columns.length + 1} style={{ ...s.td, padding: 32, textAlign: "center", color: "#9ca3af", fontSize: 13 }}>
                    No rows yet. Click + Add row above to add one.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

function CollectionCell({
  value,
  type,
  columnId,
  onCommit,
}: {
  value: string;
  type: "text" | "number" | "date";
  columnId: string;
  onCommit: (next: string) => void;
}) {
  const isImageColumn = columnId === "modelPicture" || columnId === "fabric";
  if (isImageColumn) {
    return <CollectionImageCell value={value} onCommit={onCommit} />;
  }
  const [draft, setDraft] = useState(value);
  useEffect(() => { setDraft(value); }, [value]);
  const inputType = type === "number" ? "number" : type === "date" ? "date" : "text";
  return (
    <input
      type={inputType}
      value={draft}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={() => { if (draft !== value) onCommit(draft); }}
      onKeyDown={(e) => { if (e.key === "Enter") (e.target as HTMLInputElement).blur(); }}
      style={{ width: "100%", border: "none", outline: "none", padding: "6px 8px", fontSize: 12, fontFamily: "inherit", background: "transparent", boxSizing: "border-box" }}
    />
  );
}

function CollectionImageCell({
  value,
  onCommit,
}: {
  value: string;
  onCommit: (next: string) => void;
}) {
  const [hover, setHover] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const [busy, setBusy] = useState(false);
  const trimmed = value.trim();
  const hasImage = isFabricImageValue(trimmed);

  const handleFile = async (file: File | null | undefined) => {
    if (!file || !file.type.startsWith("image/")) return;
    setBusy(true);
    try {
      const dataUrl = await compressImageToDataUrl(file);
      onCommit(dataUrl);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div style={s.fabricImageEditCell}>
      <div
        tabIndex={0}
        style={{
          ...s.fabricImageDrop,
          ...(dragOver ? { borderColor: "#2563eb", background: "#eff6ff" } : {}),
        }}
        onMouseEnter={() => setHover(true)}
        onMouseLeave={() => setHover(false)}
        onFocus={() => setHover(true)}
        onBlur={() => setHover(false)}
        onPaste={(event) => {
          const file = Array.from(event.clipboardData.files).find((item) => item.type.startsWith("image/"));
          if (file) void handleFile(file);
        }}
        onDragOver={(event) => {
          event.preventDefault();
          event.stopPropagation();
          if (!dragOver) setDragOver(true);
        }}
        onDragEnter={(event) => {
          event.preventDefault();
          event.stopPropagation();
          setDragOver(true);
        }}
        onDragLeave={(event) => {
          event.preventDefault();
          event.stopPropagation();
          setDragOver(false);
        }}
        onDrop={(event) => {
          event.preventDefault();
          event.stopPropagation();
          setDragOver(false);
          const file = Array.from(event.dataTransfer.files).find((item) => item.type.startsWith("image/"));
          if (file) void handleFile(file);
        }}
        title="Paste, drop, or click to upload image"
      >
        {hasImage
          ? <img src={trimmed} alt="" style={s.fabricSheetImage} />
          : <span>{busy ? "Uploading…" : "Paste, drop or upload"}</span>}
        {hasImage && (
          <button
            type="button"
            style={{ ...s.imageDeleteOverlay, ...(hover ? s.imageDeleteOverlayVisible : {}) }}
            onClick={(event) => {
              event.preventDefault();
              event.stopPropagation();
              onCommit("");
            }}
          >
            Delete
          </button>
        )}
        <input
          type="file"
          accept="image/*"
          style={s.hiddenFileInput}
          onChange={(event) => {
            void handleFile(event.currentTarget.files?.[0] ?? null);
            event.currentTarget.value = "";
          }}
        />
      </div>
    </div>
  );
}

// ─── Fabric Sheets ───────────────────────────────────────────────────────────

function ProductInformationPanel({
  productInfo,
  selectedCategoryId,
  search,
  updateParams,
}: {
  productInfo: ProductInfo;
  selectedCategoryId: string;
  search: string;
  updateParams: (updates: Record<string, string>) => void;
}) {
  const fetcher = useFetcher();
  const [showHidden, setShowHidden] = useState(false);
  const [styleChoice, setStyleChoice] = useState<ProductInfoStyle | null>(null);
  const [detailStyle, setDetailStyle] = useState<ProductInfoStyle | null>(null);
  const [detailDraft, setDetailDraft] = useState<Record<string, string>>({});
  const [dragStyleId, setDragStyleId] = useState<string | null>(null);
  const [dragOverStyleId, setDragOverStyleId] = useState<string | null>(null);
  const selectedCategory = productInfo.categories.find((category) => category.id === selectedCategoryId)
    ?? productInfo.categories[0]
    ?? null;
  const normalizedSearch = search.trim().toLowerCase();
  const visibleStyles = selectedCategory
    ? selectedCategory.styles.filter((style) => (showHidden || !style.hidden) && (!normalizedSearch || style.name.toLowerCase().includes(normalizedSearch)))
    : [];
  const hiddenStyleCount = selectedCategory?.styles.filter((style) => style.hidden).length ?? 0;
  const isSubmitting = fetcher.state !== "idle";
  const gridColumns = productInfo.gridColumns === 3 ? 3 : 4;

  const submitProductInfo = (fields: Record<string, string>) => {
    fetcher.submit(fields, { method: "post" });
  };

  const addStyle = () => {
    if (!selectedCategory) return;
    const name = window.prompt("Style name");
    if (!name?.trim()) return;
    submitProductInfo({ intent: "add_product_style", categoryId: selectedCategory.id, name: name.trim() });
  };

  const deleteStyle = (style: ProductInfoStyle) => {
    if (!selectedCategory) return;
    submitProductInfo({ intent: "delete_product_style", categoryId: selectedCategory.id, styleId: style.id });
  };

  const hideStyle = (style: ProductInfoStyle) => {
    if (!selectedCategory) return;
    submitProductInfo({ intent: "hide_product_style", categoryId: selectedCategory.id, styleId: style.id });
  };

  const unhideStyle = (style: ProductInfoStyle) => {
    if (!selectedCategory) return;
    submitProductInfo({ intent: "unhide_product_style", categoryId: selectedCategory.id, styleId: style.id });
  };

  const updateStyleImage = (style: ProductInfoStyle, imageUrl: string) => {
    if (!selectedCategory) return;
    submitProductInfo({
      intent: "update_product_style_image",
      categoryId: selectedCategory.id,
      styleId: style.id,
      imageUrl,
    });
  };

  const replaceStyleImage = (style: ProductInfoStyle, file: File | null) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => updateStyleImage(style, String(reader.result ?? ""));
    reader.readAsDataURL(file);
  };

  const updateGridColumns = (nextColumns: 3 | 4) => {
    submitProductInfo({ intent: "update_product_info_grid", gridColumns: String(nextColumns) });
  };

  const numberToDraft = (value?: number) => (
    typeof value === "number" && Number.isFinite(value) ? String(value) : ""
  );

  const openStyleDetails = (style: ProductInfoStyle) => {
    setDetailStyle(style);
    setDetailDraft({
      averageMeters: numberToDraft(style.averageMeters),
      averageTrimMeters: numberToDraft(style.averageTrimMeters),
      zipButtonType: style.zipButtonType ?? "",
      stitchingCost: numberToDraft(style.stitchingCost),
      zipButtonsCost: numberToDraft(style.zipButtonsCost),
      sheetCount: numberToDraft(style.sheetCount),
      costingNotes: style.costingNotes ?? "",
    });
  };

  const updateDetailDraft = (field: string, value: string) => {
    setDetailDraft((current) => ({ ...current, [field]: value }));
  };

  const saveStyleDetails = () => {
    if (!selectedCategory || !detailStyle) return;
    submitProductInfo({
      intent: "update_product_style_details",
      categoryId: selectedCategory.id,
      styleId: detailStyle.id,
      ...detailDraft,
    });
    setDetailStyle(null);
  };

  const reorderStyles = (targetStyleId: string) => {
    if (!selectedCategory || !dragStyleId || dragStyleId === targetStyleId || normalizedSearch) return;
    const currentIds = selectedCategory.styles.map((style) => style.id);
    const fromIndex = currentIds.indexOf(dragStyleId);
    const toIndex = currentIds.indexOf(targetStyleId);
    if (fromIndex < 0 || toIndex < 0) return;
    const nextIds = [...currentIds];
    const [movedId] = nextIds.splice(fromIndex, 1);
    nextIds.splice(toIndex, 0, movedId);
    submitProductInfo({
      intent: "reorder_product_styles",
      categoryId: selectedCategory.id,
      styleIds: JSON.stringify(nextIds),
    });
  };

  return (
    <div style={s.productInfoPage}>
      <div style={s.productInfoToolbar}>
        <div style={s.productInfoToolbarLeft}>
          <label style={s.productInfoSelectLabel}>
            Category
            <select
              value={selectedCategory?.id ?? ""}
              onChange={(event) => updateParams({ category: event.currentTarget.value })}
              style={s.productInfoSelect}
            >
              {productInfo.categories.map((category) => (
                <option key={category.id} value={category.id}>{category.name}</option>
              ))}
            </select>
          </label>
          <div>
            <h2 style={s.productInfoHeading}>{selectedCategory?.name ?? "Product Styles"}</h2>
            <div style={s.productInfoMeta}>
              {visibleStyles.length} of {selectedCategory?.styles.length ?? 0} styles
            </div>
          </div>
        </div>
        <div style={s.productInfoActions}>
          <div style={s.productInfoSegmented} aria-label="Style cards per row">
            {[3, 4].map((count) => (
              <button
                key={count}
                type="button"
                style={gridColumns === count ? { ...s.productInfoSegmentButton, ...s.productInfoSegmentButtonActive } : s.productInfoSegmentButton}
                onClick={() => updateGridColumns(count as 3 | 4)}
                disabled={isSubmitting}
              >
                {count}
              </button>
            ))}
          </div>
          <button type="button" style={s.primaryActionButton} onClick={addStyle} disabled={isSubmitting || !selectedCategory}>
            Add Style
          </button>
        </div>
      </div>

      <div style={{ ...s.productInfoList, gridTemplateColumns: `repeat(${gridColumns}, minmax(0, 1fr))` }}>
        {visibleStyles.map((style) => (
          <div
            key={style.id}
            draggable={!normalizedSearch}
            style={{
              ...s.productStyleCard,
              ...(dragStyleId === style.id ? s.productStyleCardDragging : {}),
              ...(dragOverStyleId === style.id && dragStyleId !== style.id ? s.productStyleCardDropTarget : {}),
              cursor: normalizedSearch ? "default" : "grab",
            }}
            onDragStart={(event) => {
              if (normalizedSearch) {
                event.preventDefault();
                return;
              }
              setDragStyleId(style.id);
              event.dataTransfer.effectAllowed = "move";
              event.dataTransfer.setData("text/plain", style.id);
            }}
            onDragOver={(event) => {
              if (!dragStyleId || normalizedSearch) return;
              event.preventDefault();
              setDragOverStyleId(style.id);
            }}
            onDragLeave={() => setDragOverStyleId((current) => current === style.id ? null : current)}
            onDrop={(event) => {
              event.preventDefault();
              reorderStyles(style.id);
              setDragStyleId(null);
              setDragOverStyleId(null);
            }}
            onDragEnd={() => {
              setDragStyleId(null);
              setDragOverStyleId(null);
            }}
          >
            <span style={s.productStyleDragHandle} title="Drag to reorder">::</span>
            <div style={s.productStyleImageWrap}>
              <button
                type="button"
                style={s.productStyleImageButton}
                onClick={() => openStyleDetails(style)}
              >
                {style.imageUrl ? (
                  <img src={style.imageUrl} alt={style.name} style={s.productStyleImage} loading="lazy" />
                ) : (
                  <div style={s.productStyleImageEmpty}>No image</div>
                )}
              </button>
            </div>
            <div style={s.productStyleCardBody}>
              <span style={s.productStyleTitle}>{style.name}</span>
              <span style={s.productStyleMeta}>
                {style.hidden ? "Hidden" : style.averageMeters ? `${style.averageMeters}m avg fabric` : "Click image for details"}
              </span>
            </div>
            <div style={s.productStyleCardActions}>
              <label style={s.secondaryButton}>
                Replace image
                <input
                  type="file"
                  accept="image/*"
                  style={{ display: "none" }}
                  onChange={(event) => {
                    replaceStyleImage(style, event.currentTarget.files?.[0] ?? null);
                    event.currentTarget.value = "";
                  }}
                />
              </label>
              <button
                type="button"
                style={s.secondaryButton}
                onClick={() => updateStyleImage(style, "")}
                disabled={!style.imageUrl || isSubmitting}
              >
                Remove image
              </button>
              <button
                type="button"
                style={s.removeUserButton}
                onClick={() => setStyleChoice(style)}
                title={`Remove ${style.name}`}
              >
                Remove style
              </button>
            </div>
          </div>
        ))}
        {selectedCategory && !visibleStyles.length && (
          <div style={s.productInfoEmpty}>
            {normalizedSearch ? "No styles match this search." : "No styles in this category yet."}
          </div>
        )}
        {!selectedCategory && (
          <div style={s.productInfoEmpty}>Add a category to start building product styles.</div>
        )}
      </div>
      <div style={s.productInfoFooterActions}>
        <button
          type="button"
          style={s.secondaryButton}
          onClick={() => setShowHidden((current) => !current)}
          disabled={!hiddenStyleCount}
        >
          {showHidden ? "Hide hidden styles" : `Show hidden styles${hiddenStyleCount ? ` (${hiddenStyleCount})` : ""}`}
        </button>
      </div>
      {detailStyle && selectedCategory && (
        <div style={s.productInfoModalBackdrop}>
          <div style={s.productInfoDetailsModal}>
            <div style={s.productInfoDetailsHeader}>
              <div>
                <h3 style={s.productInfoModalTitle}>{detailStyle.name}</h3>
                <p style={s.productInfoModalText}>
                  Edit the averaged production details for this style.
                </p>
              </div>
              {detailStyle.imageUrl && <img src={detailStyle.imageUrl} alt="" style={s.productInfoDetailsThumb} />}
            </div>
            <div style={s.productInfoDetailsGrid}>
              <label style={s.productInfoDetailsField}>
                Fabric meters average
                <input
                  type="text"
                  inputMode="decimal"
                  value={detailDraft.averageMeters ?? ""}
                  onChange={(event) => updateDetailDraft("averageMeters", event.currentTarget.value)}
                  style={s.productInfoDetailsInput}
                />
              </label>
              <label style={s.productInfoDetailsField}>
                Lining/trim meters
                <input
                  type="text"
                  inputMode="decimal"
                  value={detailDraft.averageTrimMeters ?? ""}
                  onChange={(event) => updateDetailDraft("averageTrimMeters", event.currentTarget.value)}
                  style={s.productInfoDetailsInput}
                />
              </label>
              <label style={s.productInfoDetailsField}>
                Zip/button size/type
                <input
                  type="text"
                  value={detailDraft.zipButtonType ?? ""}
                  onChange={(event) => updateDetailDraft("zipButtonType", event.currentTarget.value)}
                  style={s.productInfoDetailsInput}
                />
              </label>
              <label style={s.productInfoDetailsField}>
                Stitching cost
                <input
                  type="text"
                  inputMode="decimal"
                  value={detailDraft.stitchingCost ?? ""}
                  onChange={(event) => updateDetailDraft("stitchingCost", event.currentTarget.value)}
                  style={s.productInfoDetailsInput}
                />
              </label>
              <label style={s.productInfoDetailsField}>
                Zip/buttons cost
                <input
                  type="text"
                  inputMode="decimal"
                  value={detailDraft.zipButtonsCost ?? ""}
                  onChange={(event) => updateDetailDraft("zipButtonsCost", event.currentTarget.value)}
                  style={s.productInfoDetailsInput}
                />
              </label>
              <label style={{ ...s.productInfoDetailsField, gridColumn: "1 / -1" }}>
                Notes
                <textarea
                  rows={3}
                  value={detailDraft.costingNotes ?? ""}
                  onChange={(event) => updateDetailDraft("costingNotes", event.currentTarget.value)}
                  style={s.productInfoDetailsTextarea}
                />
              </label>
            </div>
            <div style={s.productInfoModalActions}>
              <button type="button" style={s.primaryActionButton} onClick={saveStyleDetails}>
                Save details
              </button>
              <button type="button" style={s.secondaryButton} onClick={() => setDetailStyle(null)}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}
      {styleChoice && selectedCategory && (
        <div style={s.productInfoModalBackdrop}>
          <div style={s.productInfoModal}>
            <h3 style={s.productInfoModalTitle}>Remove this style?</h3>
            <p style={s.productInfoModalText}>
              Are you sure you want to remove "{styleChoice.name}"? You can hide it instead if you may use it again later.
            </p>
            <div style={s.productInfoModalActions}>
              {styleChoice.hidden ? (
                <button type="button" style={s.secondaryButton} onClick={() => { unhideStyle(styleChoice); setStyleChoice(null); }}>
                  Unhide
                </button>
              ) : (
                <button type="button" style={s.secondaryButton} onClick={() => { hideStyle(styleChoice); setStyleChoice(null); }}>
                  Hide
                </button>
              )}
              <button type="button" style={s.removeUserButton} onClick={() => { deleteStyle(styleChoice); setStyleChoice(null); }}>
                Remove
              </button>
              <button type="button" style={s.secondaryButton} onClick={() => setStyleChoice(null)}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

const UNIFIED_FABRIC_COLUMNS = [
  { key: "supplier", label: "Supplier", header: "Supplier", defaultWidth: 130 },
  { key: "fabricType", label: "Fabric Type", header: "Fabric Type", defaultWidth: 130 },
  { key: "fabricImage", label: "Fabric", header: "Fabric", defaultWidth: 120 },
  { key: "name", label: "Name", header: "Name", defaultWidth: 150 },
  { key: "collection", label: "Collection", header: "Collection", defaultWidth: 120 },
  { key: "costPerMeter", label: "Cost per Meter", header: "Cost per Meter", defaultWidth: 100 },
  { key: "cutPieces", label: "Cut Pieces", header: "Cut Pieces", defaultWidth: 110 },
  { key: "receivedDate", label: "Received / Date", header: "Received / Date", defaultWidth: 150 },
  { key: "products", label: "Products", header: "Products", defaultWidth: 160 },
  { key: "inStock", label: "In Stock", header: "Meters in Stock", defaultWidth: 90 },
  { key: "onOrder", label: "On Order", header: "Quantity Ordered", defaultWidth: 90 },
  { key: "orderDate", label: "Order Date", header: "Order Date", defaultWidth: 110 },
] as const;

const FABRIC_COLUMN_DEFAULT_WIDTHS: Record<string, number> = Object.fromEntries(
  UNIFIED_FABRIC_COLUMNS.map((column) => [column.key, column.defaultWidth]),
);

type UnifiedFabricKey = typeof UNIFIED_FABRIC_COLUMNS[number]["key"];
type UnifiedFabricCell = {
  gid: string;
  sourceRowIndex: number;
  colIndex: number;
  header: string;
  value: string;
  originalValue: string;
};
type UnifiedFabricRowEntry = {
  primarySheet: FabricSheetData;
  primaryRowIndex: number;
  cells: Record<UnifiedFabricKey, UnifiedFabricCell | null>;
};

function unifyFabricRow(sheet: FabricSheetData, displayRowIndex: number): UnifiedFabricRowEntry {
  const sourceRowIndex = sheet.rowKeys?.[displayRowIndex] ?? displayRowIndex;
  const row = sheet.rows[displayRowIndex] ?? [];
  const originalRow = sheet.originalRows?.[sourceRowIndex] ?? row;
  const find = (predicate: (header: string) => boolean) =>
    sheet.headers.findIndex((header) => predicate(header.trim().toLowerCase()));
  const supplierIdx = find((h) => /^supplier$/.test(h));
  const fabricTypeIdx = find((h) => /^fabric\s*type$/.test(h) || /^type$/.test(h));
  const imageIdx = find((h) => /^fabric$/.test(h) || /^picture$/.test(h));
  const nameIdx = find((h) => /^name$/.test(h));
  const collectionIdx = find((h) => /^collection$/.test(h));
  const costIdx = find((h) => /cost\s*per\s*meter|price\s*per\s*meter|^price$/.test(h));
  const cutPiecesIdx = find((h) => /^cut\s*pieces?$/.test(h));
  const receivedIdx = find((h) => /received/.test(h));
  const orderDateIdx = find((h) => /^order\s*date$/.test(h));
  const productsIdx = find((h) => /^products?$/.test(h));
  const inStockIdx = find((h) => /meters?\s*in\s*stock|meters?\s*available|^meters?$/.test(h));
  const onOrderIdx = find((h) => /quantity\s*ordered/.test(h));
  const make = (idx: number, header: string): UnifiedFabricCell | null => idx < 0 ? null : {
    gid: sheet.gid,
    sourceRowIndex,
    colIndex: idx,
    header,
    value: row[idx] ?? "",
    originalValue: originalRow[idx] ?? "",
  };
  return {
    primarySheet: sheet,
    primaryRowIndex: sourceRowIndex,
    cells: {
      supplier: make(supplierIdx, "Supplier"),
      fabricType: fabricTypeIdx < 0 ? null : {
        gid: sheet.gid,
        sourceRowIndex,
        colIndex: fabricTypeIdx,
        header: "Fabric Type",
        value: canonicalizeFabricType(row[fabricTypeIdx] ?? ""),
        originalValue: canonicalizeFabricType(originalRow[fabricTypeIdx] ?? ""),
      },
      fabricImage: make(imageIdx, "Fabric"),
      name: make(nameIdx, "Name"),
      collection: make(collectionIdx, "Collection"),
      costPerMeter: make(costIdx, "Cost per Meter"),
      cutPieces: make(cutPiecesIdx, "Cut Pieces"),
      receivedDate: make(receivedIdx, "Received / Date"),
      products: make(productsIdx, "Products"),
      inStock: make(inStockIdx, "Meters in Stock"),
      onOrder: make(onOrderIdx, "Quantity Ordered"),
      orderDate: make(orderDateIdx, "Order Date"),
    },
  };
}

function mergeFabricRowEntries(rows: UnifiedFabricRowEntry[]): UnifiedFabricRowEntry[] {
  const byKey = new Map<string, UnifiedFabricRowEntry>();
  const merged: UnifiedFabricRowEntry[] = [];
  for (const entry of rows) {
    const fabricType = canonicalizeFabricType(entry.cells.fabricType?.value.trim() ?? "").toLowerCase();
    const name = (entry.cells.name?.value ?? "").trim().toLowerCase();
    if (!fabricType || !name) {
      merged.push(entry);
      continue;
    }
    const key = `${fabricType}|${name}`;
    const existing = byKey.get(key);
    if (!existing) {
      byKey.set(key, entry);
      merged.push(entry);
      continue;
    }
    const existingIsOnOrder = existing.primarySheet.gid === COMBINED_FABRIC_ON_ORDER_GID;
    const newIsOnOrder = entry.primarySheet.gid === COMBINED_FABRIC_ON_ORDER_GID;
    if (existingIsOnOrder === newIsOnOrder) {
      // Both are stock, or both are on-order — keep as separate rows
      merged.push(entry);
      continue;
    }
    const stockEntry = existingIsOnOrder ? entry : existing;
    const orderEntry = existingIsOnOrder ? existing : entry;
    const mergedCells: Record<UnifiedFabricKey, UnifiedFabricCell | null> = { ...stockEntry.cells };
    mergedCells.onOrder = orderEntry.cells.onOrder ?? mergedCells.onOrder;
    const mergedEntry: UnifiedFabricRowEntry = {
      primarySheet: stockEntry.primarySheet,
      primaryRowIndex: stockEntry.primaryRowIndex,
      cells: mergedCells,
    };
    byKey.set(key, mergedEntry);
    const index = merged.indexOf(existing);
    if (index >= 0) merged[index] = mergedEntry;
  }
  return merged;
}

function parseFabricNumberCell(value: string | undefined) {
  if (!value) return 0;
  const n = parseFloat(String(value).replace(/[, ]/g, ""));
  return Number.isFinite(n) ? n : 0;
}

function CombinedFabricStockPanel({
  sheets,
  fabricSettings,
  productInfo,
  users,
  rowHeights,
  inrToAudRate,
  nameSearch,
}: {
  sheets: FabricSheetData[];
  fabricSettings: FabricSettings;
  productInfo: ProductInfo;
  users: PortalUser[];
  rowHeights: Record<string, number>;
  inrToAudRate: number | null;
  nameSearch: string;
}) {
  const fetcher = useFetcher();
  const [fabricTypeFilter, setFabricTypeFilter] = useState("");
  const [pendingDeletes, setPendingDeletes] = useState<Set<string>>(new Set());
  useEffect(() => setPendingDeletes(new Set()), [sheets]);

  // Bulk image compaction disabled — was contending with the page's image
  // loads. Reintroduce behind an explicit user-triggered button if needed.
  const markRowDeleted = useCallback((gid: string, rowIndex: number) => {
    setPendingDeletes((prev) => {
      const next = new Set(prev);
      next.add(`${gid}:${rowIndex}`);
      return next;
    });
  }, []);
  const allRows = useMemo(() => {
    const entries: UnifiedFabricRowEntry[] = [];
    for (const sheet of sheets) {
      if (sheet.error) continue;
      for (let i = 0; i < sheet.rows.length; i++) {
        entries.push(unifyFabricRow(sheet, i));
      }
    }
    const merged = mergeFabricRowEntries(entries);
    return merged.filter((entry) => !pendingDeletes.has(`${entry.primarySheet.gid}:${entry.primaryRowIndex}`));
  }, [sheets, pendingDeletes]);

  const fabricTypeChoices = useMemo(() => {
    const rowLabels = new Set<string>();
    for (const entry of allRows) {
      const raw = entry.cells.fabricType ? entry.cells.fabricType.value.trim() : entry.primarySheet.name.trim();
      if (raw) rowLabels.add(canonicalizeFabricType(raw));
    }
    const labels = new Set<string>();
    for (const option of fabricSettings.fabricTypeOptions) {
      const label = canonicalizeFabricType(option.label);
      if (label && rowLabels.has(label)) labels.add(label);
    }
    return [...labels].sort((a, b) => a.localeCompare(b));
  }, [allRows, fabricSettings.fabricTypeOptions]);

  const filteredRows = useMemo(() => {
    const search = nameSearch.trim().toLowerCase();
    const typeFilter = canonicalizeFabricType(fabricTypeFilter).toLowerCase();
    return allRows.filter((entry) => {
      if (typeFilter) {
        const raw = entry.cells.fabricType ? entry.cells.fabricType.value.trim() : entry.primarySheet.name.trim();
        if (canonicalizeFabricType(raw).toLowerCase() !== typeFilter) return false;
      }
      if (search) {
        const name = (entry.cells.name?.value ?? "").toLowerCase();
        if (!name.includes(search)) return false;
      }
      return true;
    });
  }, [allRows, nameSearch, fabricTypeFilter]);

  const orderedColumns = useMemo(() => {
    const map = new Map(UNIFIED_FABRIC_COLUMNS.map((column) => [column.key, column]));
    const ordered: typeof UNIFIED_FABRIC_COLUMNS[number][] = [];
    for (const key of fabricSettings.combinedColumnOrder) {
      const column = map.get(key as UnifiedFabricKey);
      if (column) {
        ordered.push(column);
        map.delete(key as UnifiedFabricKey);
      }
    }
    for (const column of UNIFIED_FABRIC_COLUMNS) {
      if (map.has(column.key)) ordered.push(column);
    }
    return ordered;
  }, [fabricSettings.combinedColumnOrder]);

  const [localColumns, setLocalColumns] = useState(orderedColumns);
  useEffect(() => setLocalColumns(orderedColumns), [orderedColumns]);
  const [dragKey, setDragKey] = useState<UnifiedFabricKey | null>(null);
  const saveColumnOrder = (next: typeof UNIFIED_FABRIC_COLUMNS[number][]) => {
    submitPortalCell(
      fetcher,
      {
        intent: "update_fabric_settings",
        value: JSON.stringify({ ...fabricSettings, combinedColumnOrder: next.map((column) => column.key) }),
      },
      {
        label: "Undo column order",
        fields: { intent: "update_fabric_settings", value: JSON.stringify(fabricSettings) },
      },
    );
  };

  const [columnWidths, setColumnWidths] = useState<Record<string, number>>(fabricSettings.combinedColumnWidths);
  useEffect(() => setColumnWidths(fabricSettings.combinedColumnWidths), [fabricSettings.combinedColumnWidths]);
  const widthFor = (key: string) => columnWidths[key] ?? FABRIC_COLUMN_DEFAULT_WIDTHS[key] ?? 140;
  const startColumnResize = (columnKey: UnifiedFabricKey, event: React.MouseEvent<HTMLSpanElement>) => {
    event.preventDefault();
    event.stopPropagation();
    const startX = event.clientX;
    const startWidth = widthFor(columnKey);
    const previousCursor = document.body.style.cursor;
    const previousUserSelect = document.body.style.userSelect;
    document.body.style.cursor = "col-resize";
    document.body.style.userSelect = "none";
    let nextWidth = startWidth;
    const handleMove = (moveEvent: MouseEvent) => {
      nextWidth = Math.max(60, startWidth + moveEvent.clientX - startX);
      setColumnWidths((current) => ({ ...current, [columnKey]: nextWidth }));
    };
    const handleUp = () => {
      document.body.style.cursor = previousCursor;
      document.body.style.userSelect = previousUserSelect;
      document.removeEventListener("mousemove", handleMove);
      document.removeEventListener("mouseup", handleUp);
      if (nextWidth === startWidth) return;
      const nextWidths = { ...columnWidths, [columnKey]: nextWidth };
      setColumnWidths(nextWidths);
      submitPortalCell(
        fetcher,
        {
          intent: "update_fabric_settings",
          value: JSON.stringify({ ...fabricSettings, combinedColumnWidths: nextWidths }),
          // Skip loader revalidation so our optimistic widths aren't reset
          // (matches how packing/restock persist column widths).
          noRevalidate: "1",
        },
        { label: "Undo column resize", fields: { intent: "update_fabric_settings", value: JSON.stringify(fabricSettings) } },
      );
    };
    document.addEventListener("mousemove", handleMove);
    document.addEventListener("mouseup", handleUp);
  };
  // With table-layout: fixed the browser only honours the <col> widths when
  // the table itself has an explicit width — so resizing has no effect unless
  // we set it to the sum of the column widths (matches packing/restock).
  const fabricTableWidth = 48 + localColumns.reduce((sum, column) => sum + widthFor(column.key), 0);

  const totalInStock = filteredRows.reduce((sum, entry) => sum + parseFabricNumberCell(entry.cells.inStock?.value), 0);
  const totalOnOrder = filteredRows.reduce((sum, entry) => sum + parseFabricNumberCell(entry.cells.onOrder?.value), 0);
  const totalCostInr = filteredRows.reduce((sum, entry) => {
    if (!entry.cells.inStock) return sum;
    return sum + parseFabricNumberCell(entry.cells.inStock.value) * parseFabricNumberCell(entry.cells.costPerMeter?.value);
  }, 0);
  const totalCostAud = inrToAudRate && totalCostInr ? totalCostInr * inrToAudRate : null;

  return (
    <div style={s.fabricPage}>
      <div style={s.fabricToolbar}>
        <div style={s.fabricToolbarLeft}>
          <label style={s.filterLabel}>
            Fabric type
            <select
              value={fabricTypeFilter}
              onChange={(event) => setFabricTypeFilter(event.currentTarget.value)}
              style={s.productTypeFilter}
            >
              <option value="">All fabric types</option>
              {fabricTypeChoices.map((type) => (
                <option key={type} value={type}>{type}</option>
              ))}
            </select>
          </label>
          {fabricTypeFilter && (
            <button type="button" style={s.secondaryButton} onClick={() => setFabricTypeFilter("")}>
              Clear filter
            </button>
          )}
        </div>
        <div style={s.fabricToolbarMeta}>
          <span><strong>{filteredRows.length}</strong> rows</span>
          <span><strong>{formatFabricNumber(totalInStock)}</strong> in stock</span>
          <span><strong>{formatFabricNumber(totalOnOrder)}</strong> on order</span>
          {totalCostInr > 0 && <span><strong>{formatCurrency(totalCostInr)}</strong> value</span>}
          {totalCostAud != null && <span><strong>{formatAudCurrency(totalCostAud)}</strong> AUD</span>}
        </div>
      </div>
      <div style={s.fabricTableShell}>
        <div style={s.fabricTableWrap}>
          <table style={{ ...s.fabricTable, width: fabricTableWidth, minWidth: "100%" }} onKeyDown={handleTableGridKeyDown}>
            <colgroup>
              <col style={{ width: 48 }} />
              {localColumns.map((column) => (
                <col key={column.key} style={{ width: widthFor(column.key) }} />
              ))}
            </colgroup>
            <thead>
              <tr>
                <th style={{ ...s.fabricTh, ...s.rowNumberHeader }}>#</th>
                {localColumns.map((column) => (
                  <th
                    key={column.key}
                    onDragOver={(event) => {
                      if (!dragKey || dragKey === column.key) return;
                      event.preventDefault();
                      const from = localColumns.findIndex((item) => item.key === dragKey);
                      const to = localColumns.findIndex((item) => item.key === column.key);
                      if (from < 0 || to < 0) return;
                      const next = [...localColumns];
                      const [moved] = next.splice(from, 1);
                      next.splice(to, 0, moved);
                      setLocalColumns(next);
                    }}
                    style={{
                      ...s.fabricTh,
                      ...(dragKey === column.key ? { opacity: 0.55 } : {}),
                    }}
                  >
                    <span
                      draggable
                      onDragStart={(event) => {
                        setDragKey(column.key);
                        event.dataTransfer.effectAllowed = "move";
                      }}
                      onDragEnd={() => {
                        setDragKey(null);
                        if (localColumns.map((c) => c.key).join("|") !== orderedColumns.map((c) => c.key).join("|")) {
                          saveColumnOrder(localColumns);
                        }
                      }}
                      title="Drag to reorder column"
                      style={{ cursor: "grab", marginRight: 6, color: "#94a3b8", userSelect: "none" }}
                    >
                      ⠿
                    </span>
                    {column.label}
                    <span
                      role="separator"
                      aria-orientation="vertical"
                      aria-label={`Resize ${column.label} column`}
                      onMouseDown={(event) => startColumnResize(column.key, event)}
                      onMouseEnter={(event) => { event.currentTarget.style.background = "#2563eb"; }}
                      onMouseLeave={(event) => { event.currentTarget.style.background = "transparent"; }}
                      style={{
                        position: "absolute",
                        top: 0,
                        right: 0,
                        width: 12,
                        height: "100%",
                        cursor: "col-resize",
                        zIndex: 30,
                        background: "transparent",
                        touchAction: "none",
                      }}
                    />
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filteredRows.map((entry, displayIdx) => (
                <CombinedFabricRow
                  key={`${entry.primarySheet.gid}:${entry.primaryRowIndex}`}
                  entry={entry}
                  displayIndex={displayIdx}
                  columns={localColumns}
                  fetcher={fetcher}
                  fabricSettings={fabricSettings}
                  productInfo={productInfo}
                  users={users}
                  rowHeights={rowHeights}
                  sheets={sheets}
                  onMarkDeleted={markRowDeleted}
                />
              ))}
              {!filteredRows.length && (
                <tr>
                  <td colSpan={localColumns.length + 1} style={{ ...s.fabricTd, padding: 24, textAlign: "center" }}>
                    No fabric rows match the current filters.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function CombinedFabricRow({
  entry,
  displayIndex,
  columns,
  fetcher,
  fabricSettings,
  productInfo,
  users,
  rowHeights,
  sheets,
  onMarkDeleted,
}: {
  entry: UnifiedFabricRowEntry;
  displayIndex: number;
  columns: typeof UNIFIED_FABRIC_COLUMNS[number][];
  fetcher: ReturnType<typeof useFetcher>;
  fabricSettings: FabricSettings;
  productInfo: ProductInfo;
  users: PortalUser[];
  rowHeights: Record<string, number>;
  sheets: FabricSheetData[];
  onMarkDeleted: (gid: string, rowIndex: number) => void;
}) {
  const primaryGid = entry.primarySheet.gid;
  const primaryRowIndex = entry.primaryRowIndex;
  const rowHeightKey = `fabric:${primaryGid}:${primaryRowIndex}`;
  const fabricImageUrl = entry.cells.fabricImage?.value ?? "";
  const fabricName = entry.cells.name?.value ?? "";
  const moveTargets = sheets.filter((item) => item.gid !== primaryGid && !isHiddenFabricSheet(item.name));
  return (
    <tr style={{ ...s.row, ...(rowHeights[rowHeightKey] ? { height: rowHeights[rowHeightKey] } : {}) }}>
      <RowNumberCell
        rowNumber={displayIndex + 1}
        actions={[
          { label: "Add row", onClick: () => submitPortalCell(fetcher, { intent: "add_fabric_row", gid: primaryGid }) },
          { label: "Duplicate row", onClick: () => submitPortalCell(fetcher, { intent: "duplicate_fabric_row", gid: primaryGid, rowIndex: primaryRowIndex }) },
          {
            label: "Move to fabric type",
            options: moveTargets.map((item) => ({ label: item.name, value: item.gid })),
            onSelect: (targetGid) => submitPortalCell(fetcher, { intent: "move_fabric_row", gid: primaryGid, rowIndex: primaryRowIndex, targetGid }),
          },
          { label: "Delete row", danger: true, onClick: () => {
              if (!window.confirm("Delete this fabric row?")) return;
              onMarkDeleted(primaryGid, primaryRowIndex);
              submitPortalCell(fetcher, { intent: "delete_fabric_row", gid: primaryGid, rowIndex: primaryRowIndex });
            } },
        ]}
        heightKey={rowHeightKey}
      />
      {columns.map((column, colDisplayIdx) => {
        const cell = entry.cells[column.key];
        return (
          <FabricTd key={column.key} rowIndex={displayIndex} colIndex={colDisplayIdx}>
            {cell ? (
              <FabricCell
                gid={cell.gid}
                rowIndex={cell.sourceRowIndex}
                colIndex={cell.colIndex}
                value={cell.value}
                originalValue={cell.originalValue}
                fabricImageUrl={String(fabricImageUrl)}
                fabricName={fabricName}
                header={cell.header}
                fetcher={fetcher}
                fabricSettings={fabricSettings}
                productInfo={productInfo}
                users={users}
              />
            ) : null}
          </FabricTd>
        );
      })}
    </tr>
  );
}

function FabricHeaderCell({
  gid,
  columnId,
  index,
  label,
}: {
  gid: string;
  columnId?: string;
  index: number;
  label: string;
}) {
  const fetcher = useFetcher();
  const [menu, setMenu] = useState<{ x: number; y: number } | null>(null);
  const [width, setWidth] = useState<number | null>(null);
  const startResize = (event: React.MouseEvent<HTMLSpanElement>) => {
    event.preventDefault();
    const th = event.currentTarget.closest("th");
    const startX = event.clientX;
    const startWidth = th?.getBoundingClientRect().width ?? 140;
    const previousCursor = document.body.style.cursor;
    const previousUserSelect = document.body.style.userSelect;
    document.body.style.cursor = "col-resize";
    document.body.style.userSelect = "none";
    const handleMove = (moveEvent: MouseEvent) => {
      setWidth(Math.max(MIN_COLUMN_WIDTH, startWidth + moveEvent.clientX - startX));
    };
    const handleUp = () => {
      document.body.style.cursor = previousCursor;
      document.body.style.userSelect = previousUserSelect;
      document.removeEventListener("mousemove", handleMove);
      document.removeEventListener("mouseup", handleUp);
    };
    document.addEventListener("mousemove", handleMove);
    document.addEventListener("mouseup", handleUp);
  };

  return (
    <th
      style={{ ...s.fabricTh, ...(width ? { width, minWidth: width } : {}) }}
      onContextMenu={(event) => {
        event.preventDefault();
        setMenu({ x: event.clientX, y: event.clientY });
      }}
    >
      {isLockedFabricCalculationHeader(label)
        ? <span style={s.thContent} title="Locked because this column is used in fabric totals">{label}</span>
        : <FabricEditableHeaderLabel headerKey={`fabric:${gid}:${index}`} value={label} />}
      <span role="separator" aria-orientation="vertical" aria-label={`Resize ${label} column`} onMouseDown={startResize} style={s.resizeHandle} />
      {menu && typeof document !== "undefined" && createPortal(
        <div style={{ ...s.contextMenu, left: menu.x, top: menu.y }} onMouseDown={(event) => event.stopPropagation()}>
          <button
            type="button"
            style={s.contextMenuButton}
            onClick={() => {
              setMenu(null);
              const nextLabel = window.prompt("New column name?");
              if (!nextLabel?.trim()) return;
              const nextColumnId = `custom_${Date.now()}`;
              submitPortalCell(
                fetcher,
                { intent: "add_table_column", table: "fabric", gid, columnId: nextColumnId, label: nextLabel.trim() },
                { label: "Undo add column", fields: { intent: "remove_table_column", table: "fabric", gid, columnId: nextColumnId } },
              );
            }}
          >
            Add column
          </button>
          <button
            type="button"
            disabled={!columnId?.startsWith("custom_")}
            style={{ ...s.contextMenuButton, ...(!columnId?.startsWith("custom_") ? s.contextMenuDisabled : {}), ...(columnId?.startsWith("custom_") ? s.contextMenuDanger : {}) }}
            onClick={() => {
              if (!columnId?.startsWith("custom_")) return;
              setMenu(null);
              submitPortalCell(
                fetcher,
                { intent: "remove_table_column", table: "fabric", gid, columnId },
                { label: "Undo remove column", fields: { intent: "add_table_column", table: "fabric", gid, columnId, label } },
              );
            }}
          >
            Remove column
          </button>
        </div>,
        document.body,
      )}
    </th>
  );
}

function FabricEditableHeaderLabel({ headerKey, value }: { headerKey: string; value: string }) {
  const fetcher = useFetcher();
  const [draft, setDraft] = useState(value);
  useEffect(() => setDraft(value), [value]);
  const save = (nextValue: string) => {
    const trimmed = nextValue.trim();
    if (!trimmed || trimmed === value) return;
    submitPortalCell(
      fetcher,
      { intent: "update_table_header", key: headerKey, value: trimmed },
      { label: "Undo heading text", fields: { intent: "update_table_header", key: headerKey, value } },
    );
  };
  return (
    <input
      value={draft}
      onChange={(event) => setDraft(event.currentTarget.value)}
      onBlur={(event) => save(event.currentTarget.value)}
      onKeyDown={(event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          event.currentTarget.blur();
        }
      }}
      style={s.headerEditInput}
      title="Edit fabric heading"
    />
  );
}

function FabricTd({
  children,
  rowIndex,
  colIndex,
}: {
  children: React.ReactNode;
  rowIndex: number;
  colIndex: number;
}) {
  return (
    <td
      data-grid-row={rowIndex}
      data-grid-col={colIndex}
      tabIndex={0}
      onFocus={(event) => {
        if (event.target !== event.currentTarget) return;
        const focusTarget = event.currentTarget.querySelector<HTMLElement>(FOCUSABLE_CELL_SELECTOR);
        if (!focusTarget) return;
        window.setTimeout(() => {
          focusTarget.focus();
          if (focusTarget instanceof HTMLInputElement || focusTarget instanceof HTMLTextAreaElement) {
            focusTarget.select();
          }
        }, 0);
      }}
      style={s.fabricTd}
    >
      {children}
    </td>
  );
}

function isFabricImageValue(value: string) {
  const trimmed = value.trim();
  return /^data:image\//i.test(trimmed)
    || /^blob:/i.test(trimmed)
    || /^(?:https?:\/\/|\/).+\.(png|jpe?g|webp|gif|avif)(\?.*)?$/i.test(trimmed);
}

// Browser-side shrink: a 4000x3000 phone photo (~5 MB) becomes ~800x600 (~120 KB)
// before it ever hits the network. Keeps uploads under the server's size cap,
// makes the stored blob smaller, and makes image cells render faster.
async function resizeImageForFabricUpload(file: File): Promise<File> {
  if (!file.type.startsWith("image/")) return file;
  // Tiny files: skip the work.
  if (file.size < 200 * 1024) return file;
  const dataUrl: string = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = () => reject(reader.error ?? new Error("read failed"));
    reader.readAsDataURL(file);
  });
  const img = new Image();
  await new Promise<void>((resolve, reject) => {
    img.onload = () => resolve();
    img.onerror = () => reject(new Error("decode failed"));
    img.src = dataUrl;
  });
  const maxEdge = 800;
  const longest = Math.max(img.naturalWidth, img.naturalHeight);
  if (!longest) return file;
  const ratio = Math.min(1, maxEdge / longest);
  const width = Math.max(1, Math.round(img.naturalWidth * ratio));
  const height = Math.max(1, Math.round(img.naturalHeight * ratio));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");
  if (!ctx) return file;
  ctx.drawImage(img, 0, 0, width, height);
  const blob: Blob | null = await new Promise((resolve) => {
    canvas.toBlob((result) => resolve(result), "image/webp", 0.82);
  });
  if (!blob || blob.size >= file.size) return file;
  return new File([blob], file.name.replace(/\.[a-z]+$/i, "") + ".webp", {
    type: "image/webp",
    lastModified: Date.now(),
  });
}

function isNumericFabricCell(header: string, value: string) {
  const normalizedHeader = header.trim().toLowerCase();
  if (/cost|price|meter|quantity|received|^k$|^l$|additional/.test(normalizedHeader)) return true;
  const trimmed = value.trim();
  return Boolean(trimmed) && /^[-₹$]?\d[\d,]*(?:\.\d+)?$/.test(trimmed);
}

function FabricCell({
  gid,
  rowIndex,
  colIndex,
  value,
  originalValue,
  fabricImageUrl,
  fabricName,
  header,
  fetcher,
  fabricSettings,
  productInfo,
  users,
}: {
  gid: string;
  rowIndex: number;
  colIndex: number;
  value: string;
  originalValue: string;
  fabricImageUrl: string;
  fabricName: string;
  header: string;
  fetcher: ReturnType<typeof useFetcher>;
  fabricSettings: FabricSettings;
  productInfo: ProductInfo;
  users: PortalUser[];
}) {
  const revalidator = useRevalidator();
  const [draft, setDraft] = useState(value);
  const [imageHover, setImageHover] = useState(false);
  useEffect(() => setDraft(value), [value]);
  const trimmed = draft.trim();
  const imageValue = isFabricImageValue(trimmed);
  const normalizedHeader = header.trim().toLowerCase();
  const imageColumn = /picture|image/i.test(header) || normalizedHeader === "fabric" || imageValue || isFabricImageValue(originalValue);
  const chipKind = /^supplier$/i.test(header)
    ? "supplierOptions"
    : /fabric\s*type/i.test(header)
      ? "fabricTypeOptions"
      : null;
  const chipOptions = chipKind === "supplierOptions"
    ? fabricSettings.supplierOptions
    : chipKind === "fabricTypeOptions"
      ? fabricSettings.fabricTypeOptions
      : null;
  const chipOption = chipOptions?.find((option) => option.label === draft || option.value === slugForOption(draft));
  const centerValue = isNumericFabricCell(header, draft);
  const save = (nextValue: string) => {
    if (nextValue === value) return;
    submitPortalCell(
      fetcher,
      {
        intent: "update_fabric_cell",
        gid,
        rowIndex,
        colIndex,
        value: nextValue,
      },
      { label: "Undo fabric cell", fields: { intent: "update_fabric_cell", gid, rowIndex, colIndex, value } },
    );
  };
  const uploadImage = async (file: File | null) => {
    if (!file || !file.type.startsWith("image/")) return;
    // Instant preview via an object URL while the resize + encode run in the background.
    const previewUrl = URL.createObjectURL(file);
    setDraft(previewUrl);
    let processed = file;
    try {
      processed = await resizeImageForFabricUpload(file);
    } catch {
      processed = file;
    }
    if (processed.size > 5 * 1024 * 1024) {
      window.alert(`That image is ${(processed.size / 1024 / 1024).toFixed(1)} MB — over the 5 MB limit. Try a smaller one.`);
      return;
    }
    let dataUrl = "";
    try {
      dataUrl = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result ?? ""));
        reader.onerror = () => reject(reader.error ?? new Error("read failed"));
        reader.readAsDataURL(processed);
      });
    } catch {
      window.alert("Couldn't read the image file.");
      return;
    }
    pushPortalUndo({ label: "Undo fabric image", fields: { intent: "update_fabric_cell", gid, rowIndex, colIndex, value } });
    submitPortalCell(fetcher, {
      intent: "update_fabric_cell",
      gid,
      rowIndex,
      colIndex,
      value: dataUrl,
    });
  };

  if (/^products?$/i.test(normalizedHeader)) {
    return (
      <FabricProductsCell
        value={draft}
        originalValue={originalValue}
        fabricImageUrl={fabricImageUrl}
        fabricName={fabricName}
        productInfo={productInfo}
        onDraftChange={setDraft}
        onSave={save}
      />
    );
  }

  if (/^notes?$/i.test(normalizedHeader)) {
    return (
      <FabricMentionCell
        value={draft}
        originalValue={originalValue}
        users={users}
        onDraftChange={setDraft}
        onSave={save}
      />
    );
  }

  if (imageColumn) {
    return (
      <FabricImageEditCell
        value={trimmed}
        imageValue={imageValue}
        originalValue={originalValue}
        draft={draft}
        setDraft={setDraft}
        save={save}
        uploadImage={uploadImage}
        imageHover={imageHover}
        setImageHover={setImageHover}
      />
    );
  }

  if (chipOptions) {
    const options = chipOptions.some((option) => option.label === draft || option.value === slugForOption(draft))
      ? chipOptions
      : draft
        ? [...chipOptions, { value: slugForOption(draft) || draft, label: draft, bg: "#f3f4f6", color: "#374151" }]
        : chipOptions;
    return (
      <FabricChipDropdown
        value={draft}
        option={chipOption}
        options={options}
        chipKind={chipKind}
        fabricSettings={fabricSettings}
        onChange={(nextValue) => {
          setDraft(nextValue);
          save(nextValue);
        }}
      />
    );
  }

  return (
    <>
      <AutoGrowTextarea
        value={draft}
        onChange={setDraft}
        onCommit={save}
        style={{
          ...s.fabricCellTextarea,
          ...(centerValue ? s.fabricNumberCellInput : {}),
        }}
        placeholder="—"
      />
      {originalValue && originalValue !== draft && (
        <div style={s.fabricCellActions}>
          <button
            type="button"
            style={s.fabricMiniButton}
            onClick={() => {
              setDraft(originalValue);
              save(originalValue);
            }}
          >
            Restore
          </button>
        </div>
      )}
    </>
  );
}

function ImageUploadDialog({
  open,
  onClose,
  onFile,
  hasCurrentImage,
  onRemove,
}: {
  open: boolean;
  onClose: () => void;
  onFile: (file: File) => void;
  hasCurrentImage: boolean;
  onRemove: () => void;
}) {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [dragOver, setDragOver] = useState(false);
  const [pasteMenu, setPasteMenu] = useState<{ x: number; y: number } | null>(null);
  useEffect(() => {
    if (!open) return;
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") onClose();
    };
    const onPaste = (event: ClipboardEvent) => {
      const file = Array.from(event.clipboardData?.files ?? []).find((item) => item.type.startsWith("image/"));
      if (file) {
        event.preventDefault();
        onFile(file);
      }
    };
    document.addEventListener("keydown", onKey);
    document.addEventListener("paste", onPaste);
    return () => {
      document.removeEventListener("keydown", onKey);
      document.removeEventListener("paste", onPaste);
    };
  }, [open, onClose, onFile]);
  const pasteFromClipboard = async () => {
    setPasteMenu(null);
    if (!navigator.clipboard || typeof navigator.clipboard.read !== "function") {
      window.alert("This browser doesn't allow reading the clipboard. Use Cmd/Ctrl+V instead.");
      return;
    }
    try {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        const imageType = item.types.find((type) => type.startsWith("image/"));
        if (!imageType) continue;
        const blob = await item.getType(imageType);
        onFile(new File([blob], "pasted-image", { type: imageType }));
        return;
      }
      window.alert("No image found in the clipboard.");
    } catch (error) {
      window.alert(`Paste failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  };
  if (!open || typeof document === "undefined") return null;
  return createPortal(
    <div
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(15, 23, 42, 0.55)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 9999,
      }}
      onClick={onClose}
    >
      <div
        onClick={(event) => event.stopPropagation()}
        onDragOver={(event) => {
          event.preventDefault();
          event.dataTransfer.dropEffect = "copy";
        }}
        onDragEnter={() => setDragOver(true)}
        onDragLeave={(event) => {
          if (event.currentTarget.contains(event.relatedTarget as Node | null)) return;
          setDragOver(false);
        }}
        onDrop={(event) => {
          event.preventDefault();
          setDragOver(false);
          const file = Array.from(event.dataTransfer.files).find((item) => item.type.startsWith("image/"));
          if (file) onFile(file);
        }}
        style={{
          background: "#fff",
          borderRadius: 10,
          padding: 24,
          width: 460,
          maxWidth: "92vw",
          boxShadow: "0 25px 50px -12px rgba(0, 0, 0, 0.3)",
          display: "flex",
          flexDirection: "column",
          gap: 16,
        }}
      >
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <strong style={{ fontSize: 16 }}>Add image</strong>
          <button
            type="button"
            onClick={onClose}
            style={{ border: "none", background: "transparent", fontSize: 22, cursor: "pointer", lineHeight: 1, color: "#64748b" }}
            aria-label="Close"
          >
            ×
          </button>
        </div>
        <div
          onContextMenu={(event) => {
            event.preventDefault();
            event.stopPropagation();
            setPasteMenu({ x: event.clientX, y: event.clientY });
          }}
          style={{
            border: dragOver ? "3px dashed #2563eb" : "3px dashed #cbd5e1",
            background: dragOver ? "#dbeafe" : "#f8fafc",
            borderRadius: 8,
            padding: 36,
            textAlign: "center",
            color: "#475569",
            fontSize: 14,
            fontWeight: 600,
          }}
        >
          <div style={{ marginBottom: 6 }}>Drop an image here</div>
          <div style={{ fontSize: 13, color: "#94a3b8" }}>or paste with Cmd/Ctrl+V — or right-click for paste</div>
        </div>
        <button
          type="button"
          style={{
            background: "#2563eb",
            color: "#fff",
            border: "none",
            padding: "10px 16px",
            borderRadius: 6,
            fontWeight: 700,
            fontSize: 14,
            cursor: "pointer",
          }}
          onClick={() => fileInputRef.current?.click()}
        >
          Choose file…
        </button>
        {hasCurrentImage && (
          <button
            type="button"
            style={{
              background: "transparent",
              color: "#991b1b",
              border: "1px solid #fecaca",
              padding: "8px 16px",
              borderRadius: 6,
              fontWeight: 700,
              fontSize: 13,
              cursor: "pointer",
            }}
            onClick={() => {
              onRemove();
              onClose();
            }}
          >
            Remove current image
          </button>
        )}
        <input
          ref={fileInputRef}
          type="file"
          accept="image/*"
          style={{ display: "none" }}
          onChange={(event) => {
            const file = event.currentTarget.files?.[0];
            event.currentTarget.value = "";
            if (file) onFile(file);
          }}
        />
      </div>
      {pasteMenu && (
        <div
          onClick={(event) => event.stopPropagation()}
          style={{
            position: "fixed",
            left: pasteMenu.x,
            top: pasteMenu.y,
            background: "#fff",
            border: "1px solid #cbd5e1",
            borderRadius: 6,
            boxShadow: "0 10px 20px rgba(15,23,42,0.18)",
            padding: 4,
            zIndex: 10000,
            minWidth: 180,
          }}
          onMouseLeave={() => setPasteMenu(null)}
        >
          <button
            type="button"
            style={{
              display: "block",
              width: "100%",
              padding: "8px 12px",
              background: "transparent",
              border: "none",
              textAlign: "left",
              cursor: "pointer",
              fontSize: 13,
              fontWeight: 600,
              color: "#1f2937",
            }}
            onClick={() => void pasteFromClipboard()}
          >
            Paste from clipboard
          </button>
        </div>
      )}
    </div>,
    document.body,
  );
}

function FabricImageEditCell({
  value,
  imageValue,
  originalValue,
  draft,
  setDraft,
  save,
  uploadImage,
  imageHover,
  setImageHover,
}: {
  value: string;
  imageValue: boolean;
  originalValue: string;
  draft: string;
  setDraft: (next: string) => void;
  save: (next: string) => void;
  uploadImage: (file: File | null) => void;
  imageHover: boolean;
  setImageHover: (hover: boolean) => void;
}) {
  const [dialogOpen, setDialogOpen] = useState(false);
  return (
    <div style={s.fabricImageEditCell}>
      <div
        style={s.fabricImageDrop}
        onMouseEnter={() => setImageHover(true)}
        onMouseLeave={() => setImageHover(false)}
        onContextMenu={(event) => {
          event.preventDefault();
          setDialogOpen(true);
        }}
        title="Right-click to upload, drop, or paste"
      >
        {imageValue ? <img src={value} alt="" style={s.fabricSheetImage} /> : <span>Right-click to add image</span>}
        {imageValue && (
          <button
            type="button"
            style={{ ...s.imageDeleteOverlay, ...(imageHover ? s.imageDeleteOverlayVisible : {}) }}
            onClick={(event) => {
              event.preventDefault();
              event.stopPropagation();
              setDraft("");
              save("");
            }}
          >
            Delete
          </button>
        )}
      </div>
      {originalValue && originalValue !== draft && (
        <button
          type="button"
          style={s.fabricMiniButton}
          onClick={() => {
            setDraft(originalValue);
            save(originalValue);
          }}
        >
          Restore
        </button>
      )}
      <ImageUploadDialog
        open={dialogOpen}
        onClose={() => setDialogOpen(false)}
        onFile={(file) => {
          setDialogOpen(false);
          void uploadImage(file);
        }}
        hasCurrentImage={imageValue}
        onRemove={() => {
          setDraft("");
          save("");
        }}
      />
    </div>
  );
}

function AutoGrowTextarea({
  value,
  onChange,
  onCommit,
  style,
  placeholder,
}: {
  value: string;
  onChange: (next: string) => void;
  onCommit: (next: string) => void;
  style?: React.CSSProperties;
  placeholder?: string;
}) {
  const ref = useRef<HTMLTextAreaElement>(null);
  const resize = () => {
    const el = ref.current;
    if (!el) return;
    el.style.height = "auto";
    el.style.height = `${Math.max(el.scrollHeight, 22)}px`;
  };
  useEffect(() => { resize(); }, [value]);
  useEffect(() => {
    const handle = () => resize();
    window.addEventListener("resize", handle);
    return () => window.removeEventListener("resize", handle);
  }, []);
  return (
    <textarea
      ref={ref}
      value={value}
      onChange={(event) => onChange(event.currentTarget.value)}
      onBlur={(event) => onCommit(event.currentTarget.value)}
      onKeyDown={(event) => {
        if (event.key === "Enter" && !event.shiftKey) {
          event.preventDefault();
          onCommit(event.currentTarget.value);
        }
      }}
      rows={1}
      style={{
        ...style,
        boxSizing: "border-box",
        whiteSpace: "pre-wrap",
        wordBreak: "break-word",
        overflow: "hidden",
        resize: "none",
      }}
      placeholder={placeholder}
    />
  );
}

function FabricChipDropdown({
  value,
  option,
  options,
  chipKind,
  fabricSettings,
  onChange,
}: {
  value: string;
  option?: RestockOption;
  options: RestockOption[];
  chipKind: "supplierOptions" | "fabricTypeOptions";
  fabricSettings: FabricSettings;
  onChange: (value: string) => void;
}) {
  const settingsFetcher = useFetcher();
  const buttonRef = useRef<HTMLButtonElement | null>(null);
  const [open, setOpen] = useState(false);
  const [rect, setRect] = useState<DOMRect | null>(null);
  const [editingChip, setEditingChip] = useState<RestockOption | null>(null);
  const [addingChip, setAddingChip] = useState(false);
  const [editLabel, setEditLabel] = useState("");
  const [editBg, setEditBg] = useState("#f3f4f6");
  const [editColor, setEditColor] = useState("#374151");
  const selectedOption = option ?? options.find((item) => item.label === value || item.value === slugForOption(value));
  const startEdit = (item?: RestockOption) => {
    setEditingChip(item ?? null);
    setAddingChip(!item);
    setEditLabel(item?.label ?? "");
    setEditBg(item?.bg ?? "#f3f4f6");
    setEditColor(item?.color ?? "#374151");
  };
  const stopEdit = () => {
    setEditingChip(null);
    setAddingChip(false);
    setEditLabel("");
  };
  const updateRect = () => {
    if (buttonRef.current) setRect(buttonRef.current.getBoundingClientRect());
  };
  useEffect(() => {
    if (!open) return;
    updateRect();
    const close = (event: MouseEvent) => {
      const target = event.target as Node;
      if (buttonRef.current?.contains(target)) return;
      const menu = document.querySelector("[data-fabric-chip-menu='true']");
      if (menu?.contains(target)) return;
      setOpen(false);
    };
    document.addEventListener("mousedown", close);
    window.addEventListener("resize", updateRect);
    window.addEventListener("scroll", updateRect, true);
    return () => {
      document.removeEventListener("mousedown", close);
      window.removeEventListener("resize", updateRect);
      window.removeEventListener("scroll", updateRect, true);
    };
  }, [open]);

  const updateChipOption = (label: string, patch: Partial<RestockOption>) => {
    if (!label.trim()) return;
    const nextOptions = [...options];
    const existingIndex = nextOptions.findIndex((item) => item.label.toLowerCase() === label.trim().toLowerCase());
    const existing = existingIndex >= 0
      ? nextOptions[existingIndex]
      : { value: slugForOption(label) || label, label, bg: "#f3f4f6", color: "#374151" };
    const nextLabel = String(patch.label ?? existing.label ?? label).trim();
    const nextOption = { ...existing, ...patch, value: existing.value || slugForOption(nextLabel) || nextLabel, label: nextLabel };
    if (existingIndex >= 0) nextOptions[existingIndex] = nextOption;
    else nextOptions.push(nextOption);
    submitPortalCell(
      settingsFetcher,
      {
        intent: "update_fabric_settings",
        value: JSON.stringify({ ...fabricSettings, [chipKind]: nextOptions }),
      },
      { label: "Undo fabric chip", fields: { intent: "update_fabric_settings", value: JSON.stringify(fabricSettings) } },
    );
    if (value === label) onChange(nextOption.label);
  };

  const saveEdit = () => {
    const nextLabel = editLabel.trim();
    if (!nextLabel) return;
    updateChipOption(editingChip?.label ?? nextLabel, { label: nextLabel, bg: editBg, color: editColor });
    if (addingChip) onChange(nextLabel);
    stopEdit();
  };

  const dropdown = open && rect && typeof document !== "undefined" ? createPortal(
    <div
      data-fabric-chip-menu="true"
      style={{
        ...s.fabricChipMenu,
        top: rect.bottom + 6,
        left: rect.left,
        minWidth: Math.max(rect.width, 240),
      }}
    >
      <button
        type="button"
        style={s.fabricChipMenuOption}
        onClick={() => {
          onChange("");
          setOpen(false);
        }}
      >
        <span>—</span>
      </button>
      {options.map((item) => {
        const selected = item.label === value || item.value === slugForOption(value);
        return (
          <div key={item.value} style={s.fabricChipMenuItem}>
            <button
              type="button"
              style={s.fabricChipEditButton}
              onClick={() => startEdit(item)}
            >
              Edit
            </button>
            <button
              type="button"
              style={s.fabricChipMenuOption}
              onClick={() => {
                onChange(item.label);
                setOpen(false);
              }}
            >
              <span style={s.fabricChipCheck}>{selected ? "✓" : ""}</span>
              <span style={{ ...s.fabricChipMenuPill, background: item.bg, color: item.color }}>{item.label}</span>
            </button>
          </div>
        );
      })}
      {(editingChip || addingChip) && (
        <div style={s.fabricChipEditor}>
          <input
            value={editLabel}
            onChange={(event) => setEditLabel(event.currentTarget.value)}
            style={s.fabricChipEditInput}
            placeholder="Chip text"
            autoFocus
          />
          <label style={s.fabricChipMenuToolLabel}>
            Chip
            <input type="color" value={editBg} style={s.fabricChipColor} onChange={(event) => setEditBg(event.currentTarget.value)} />
          </label>
          <label style={s.fabricChipMenuToolLabel}>
            Text
            <input type="color" value={editColor} style={s.fabricChipColor} onChange={(event) => setEditColor(event.currentTarget.value)} />
          </label>
          <div style={s.fabricChipEditActions}>
            <button type="button" style={s.fabricMiniButton} onClick={stopEdit}>Cancel</button>
            <button type="button" style={s.fabricMiniButton} onClick={saveEdit}>Save</button>
          </div>
        </div>
      )}
      <div style={s.fabricChipMenuTools}>
        <button
          type="button"
          style={s.fabricMiniButton}
          onClick={() => {
            startEdit();
          }}
        >
          Add chip
        </button>
      </div>
    </div>,
    document.body,
  ) : null;

  return (
    <div style={s.fabricChipCell}>
      <button
        ref={buttonRef}
        type="button"
        style={{
          ...s.fabricChipSelect,
          background: selectedOption?.bg ?? "#f3f4f6",
          color: selectedOption?.color ?? "#374151",
        }}
        onClick={() => {
          updateRect();
          setOpen((current) => !current);
        }}
      >
        <span style={s.fabricChipButtonText}>{value || "—"}</span>
        <span style={s.fabricChipChevron}>⌄</span>
      </button>
      {dropdown}
    </div>
  );
}

type FabricStyleUsage = {
  styleId: string;
  styleName: string;
  meters: string;
};

function productInfoStyleSearchOptions(productInfo: ProductInfo) {
  return productInfo.categories.flatMap((category) =>
    category.styles
      .filter((style) => !style.hidden)
      .map((style) => ({
        id: style.id,
        name: style.name,
        categoryName: category.name,
        averageMeters: style.averageMeters,
      })),
  );
}

function parseFabricStyleUsage(value: string): FabricStyleUsage[] {
  const trimmed = value.trim();
  if (!trimmed) return [];
  try {
    const parsed = JSON.parse(trimmed) as { styles?: Partial<FabricStyleUsage>[] };
    if (Array.isArray(parsed.styles)) {
      return parsed.styles
        .map((item) => ({
          styleId: String(item.styleId ?? slugForOption(String(item.styleName ?? ""))),
          styleName: String(item.styleName ?? "").trim(),
          meters: String(item.meters ?? "").trim(),
        }))
        .filter((item) => item.styleName);
    }
  } catch {
    // Older rows were saved as plain product names. Keep them editable.
  }
  return trimmed
    .split(/\n|,/)
    .map((item) => item.trim())
    .filter(Boolean)
    .map((styleName) => ({ styleId: slugForOption(styleName), styleName, meters: "" }));
}

function serializeFabricStyleUsage(items: FabricStyleUsage[]) {
  const styles = items
    .slice()
    .sort(compareFabricStyleUsage)
    .map((item) => ({
      styleId: item.styleId || slugForOption(item.styleName),
      styleName: item.styleName.trim(),
      meters: item.meters.trim(),
    }))
    .filter((item) => item.styleName);
  return styles.length ? JSON.stringify({ styles }) : "";
}

function fabricStyleSortGroup(styleName: string) {
  const normalized = styleName.toLowerCase();
  if (/\bjacket(s)?\b/.test(normalized)) return 0;
  if (/\bdress(es)?\b/.test(normalized)) return 1;
  if (/\btop(s)?\b/.test(normalized)) return 2;
  if (/\bskirt(s)?\b/.test(normalized)) return 3;
  if (/\bpant(s)?\b|\bpants\b/.test(normalized)) return 4;
  return 5;
}

function compareFabricStyleUsage(a: FabricStyleUsage, b: FabricStyleUsage) {
  const groupDiff = fabricStyleSortGroup(a.styleName) - fabricStyleSortGroup(b.styleName);
  if (groupDiff !== 0) return groupDiff;
  return a.styleName.localeCompare(b.styleName, undefined, { sensitivity: "base" });
}

function FabricProductsCell({
  value,
  originalValue,
  fabricImageUrl,
  fabricName,
  productInfo,
  onDraftChange,
  onSave,
}: {
  value: string;
  originalValue: string;
  fabricImageUrl: string;
  fabricName: string;
  productInfo: ProductInfo;
  onDraftChange: (value: string) => void;
  onSave: (value: string) => void;
}) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [items, setItems] = useState<FabricStyleUsage[]>(() => parseFabricStyleUsage(value));
  useEffect(() => setItems(parseFabricStyleUsage(value)), [value]);
  const options = productInfoStyleSearchOptions(productInfo);
  const normalizedQuery = query.trim().toLowerCase();
  const searchResults = normalizedQuery
    ? options
        .filter((option) =>
          `${option.name} ${option.categoryName}`.toLowerCase().includes(normalizedQuery) &&
          !items.some((item) => item.styleId === option.id),
        )
        .slice(0, 8)
    : [];
  const previewItems = parseFabricStyleUsage(value);
  const sortedPreviewItems = previewItems.slice().sort(compareFabricStyleUsage);
  const sortedItems = items.slice().sort(compareFabricStyleUsage);
  const hasChanges = originalValue && originalValue !== value;
  const save = () => {
    const nextValue = serializeFabricStyleUsage(items);
    onDraftChange(nextValue);
    onSave(nextValue);
    setOpen(false);
  };
  const restore = () => {
    const nextItems = parseFabricStyleUsage(originalValue);
    setItems(nextItems);
    onDraftChange(originalValue);
    onSave(originalValue);
  };

  return (
    <div style={s.fabricProductsCell}>
      <button
        type="button"
        style={sortedPreviewItems.length ? s.fabricProductsButton : s.fabricProductsEmptyButton}
        onClick={() => {
          setItems(parseFabricStyleUsage(value));
          setQuery("");
          setOpen(true);
        }}
      >
        {sortedPreviewItems.length ? (
          <>
            {sortedPreviewItems.slice(0, 5).map((item) => (
              <span key={`${item.styleId}-${item.styleName}`} style={s.fabricProductChip}>
                {item.styleName}{item.meters ? ` ${item.meters}m` : ""}
              </span>
            ))}
            {sortedPreviewItems.length > 5 && <span style={s.fabricProductMore}>Show more</span>}
          </>
        ) : "Add style usage"}
      </button>
      {open && typeof document !== "undefined" && createPortal(
        <div style={s.productInfoModalBackdrop}>
          <div style={s.fabricStyleUsageModal}>
            <div style={s.fabricStyleUsageLayout}>
              <div style={s.fabricStyleUsageImagePane}>
                <div style={s.fabricStyleUsageImageFrame}>
                  {isFabricImageValue(fabricImageUrl) ? (
                    <img src={fabricImageUrl} alt="" style={s.fabricStyleUsageImage} />
                  ) : (
                    <div style={s.fabricStyleUsageImageEmpty}>No image</div>
                  )}
                </div>
                {fabricName.trim() && <h1 style={s.fabricStyleUsagePrintName}>{fabricName.trim()}</h1>}
              </div>
              <div style={s.fabricStyleUsageContent}>
                <div style={s.fabricStyleUsageHeader}>
                  <div>
                    <h2 style={s.productInfoModalTitle}>Styles for this fabric</h2>
                    <p style={s.productInfoModalText}>Search Product Information styles and enter meters for this fabric.</p>
                  </div>
                  <button type="button" style={s.secondaryButton} onClick={() => setOpen(false)}>Close</button>
                </div>
                <div style={s.fabricStyleSearchWrap}>
                  <input
                    type="search"
                    value={query}
                    onChange={(event) => setQuery(event.currentTarget.value)}
                    autoFocus
                    placeholder="Search product style"
                    style={s.fabricStyleSearchInput}
                  />
                  {searchResults.length > 0 && (
                    <div style={s.fabricStyleSearchResults}>
                      {searchResults.map((option) => (
                        <button
                          key={option.id}
                          type="button"
                          style={s.fabricStyleSearchResult}
                          onClick={() => {
                            setItems((current) => [
                              ...current,
                              {
                                styleId: option.id,
                                styleName: option.name,
                                meters: option.averageMeters ? String(option.averageMeters) : "",
                              },
                            ]);
                            setQuery("");
                          }}
                        >
                          <span style={s.fabricStyleSearchName}>{option.name}</span>
                          <span style={s.fabricStyleSearchCategory}>{option.categoryName}</span>
                        </button>
                      ))}
                    </div>
                  )}
                </div>
                <div style={s.fabricStyleUsageTableWrap}>
                  <table style={s.fabricStyleUsageTable}>
                    <thead>
                      <tr>
                        <th style={s.fabricStyleUsageTh}>Style</th>
                        <th style={s.fabricStyleUsageTh}>Meters per garment</th>
                        <th style={s.fabricStyleUsageTh}> </th>
                      </tr>
                    </thead>
                    <tbody>
                      {sortedItems.map((item) => {
                        const itemIndex = items.findIndex((currentItem) => currentItem === item);
                        return (
                        <tr key={`${item.styleId}-${itemIndex}`}>
                          <td style={s.fabricStyleUsageTd}>{item.styleName}</td>
                          <td style={s.fabricStyleUsageTd}>
                            <input
                              type="number"
                              min="0"
                              step="0.01"
                              className="no-number-spinner"
                              value={item.meters}
                              onChange={(event) => {
                                const meters = event.currentTarget.value;
                                setItems((current) => current.map((currentItem, currentIndex) => currentIndex === itemIndex ? { ...currentItem, meters } : currentItem));
                              }}
                              style={s.fabricStyleUsageInput}
                            />
                          </td>
                          <td style={{ ...s.fabricStyleUsageTd, textAlign: "right" }}>
                            <button
                              type="button"
                              style={s.removeUserButton}
                              onClick={() => setItems((current) => current.filter((_, currentIndex) => currentIndex !== itemIndex))}
                            >
                              Remove
                            </button>
                          </td>
                        </tr>
                      );})}
                      {!items.length && (
                        <tr>
                          <td colSpan={3} style={s.fabricStyleUsageEmpty}>No styles added yet.</td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
                <div style={s.fabricProductsActions}>
                  {hasChanges && <button type="button" style={s.fabricMiniButton} onClick={restore}>Restore</button>}
                  <button type="button" style={s.secondaryButton} onClick={() => setOpen(false)}>Cancel</button>
                  <button type="button" style={s.primaryActionButton} onClick={save}>Save</button>
                </div>
              </div>
            </div>
          </div>
        </div>,
        document.body,
      )}
      {hasChanges && (
        <div style={s.fabricCellActions}>
          <button type="button" style={s.fabricMiniButton} onClick={restore}>Restore</button>
        </div>
      )}
    </div>
  );
}

function FabricMentionCell({
  value,
  originalValue,
  users,
  onDraftChange,
  onSave,
}: {
  value: string;
  originalValue: string;
  users: PortalUser[];
  onDraftChange: (value: string) => void;
  onSave: (value: string) => void;
}) {
  const [focused, setFocused] = useState(false);
  const tagQuery = currentTagQuery(value);
  const suggestions = tagQuery == null
    ? []
    : users
        .filter((user) => user.active)
        .filter((user) => user.name.toLowerCase().includes(tagQuery))
        .slice(0, 5);

  return (
    <div style={s.fabricNoteCell}>
      <div style={s.noteTagWrap}>
        <textarea
          value={value}
          onChange={(event) => onDraftChange(event.currentTarget.value)}
          onFocus={() => setFocused(true)}
          onBlur={(event) => {
            window.setTimeout(() => setFocused(false), 120);
            onSave(event.currentTarget.value);
          }}
          rows={3}
          style={s.fabricCellTextarea}
          placeholder="Add note... use @name"
        />
        {focused && suggestions.length > 0 && (
          <div style={s.tagSuggestions}>
            {suggestions.map((user) => (
              <button
                key={user.id}
                type="button"
                style={s.tagSuggestionButton}
                onMouseDown={(event) => {
                  event.preventDefault();
                  onDraftChange(insertStaffTag(value, user.name));
                }}
              >
                @{user.name}
              </button>
            ))}
          </div>
        )}
      </div>
      {originalValue && originalValue !== value && (
        <div style={s.fabricCellActions}>
          <button
            type="button"
            style={s.fabricMiniButton}
            onClick={() => {
              onDraftChange(originalValue);
              onSave(originalValue);
            }}
          >
            Restore
          </button>
        </div>
      )}
    </div>
  );
}

function FabricRowActions({
  gid,
  rowIndex,
  sheets,
  fetcher,
}: {
  gid: string;
  rowIndex: number;
  sheets: FabricSheetData[];
  fetcher: ReturnType<typeof useFetcher>;
}) {
  const moveTargets = sheets.filter((sheet) => sheet.gid !== gid);
  const [targetGid, setTargetGid] = useState(moveTargets[0]?.gid ?? "");
  useEffect(() => {
    if (!moveTargets.some((sheet) => sheet.gid === targetGid)) {
      setTargetGid(moveTargets[0]?.gid ?? "");
    }
  }, [moveTargets, targetGid]);

  return (
    <div style={s.fabricRowActions}>
      <button
        type="button"
        style={s.removeUserButton}
        onClick={() => submitPortalCell(fetcher, {
          intent: "delete_fabric_row",
          gid,
          rowIndex,
        })}
      >
        Delete row
      </button>
      <select
        value={targetGid}
        onChange={(event) => setTargetGid(event.currentTarget.value)}
        style={s.fabricMoveSelect}
      >
        {moveTargets.map((sheet) => (
          <option key={sheet.gid} value={sheet.gid}>{sheet.name}</option>
        ))}
      </select>
      <button
        type="button"
        disabled={!targetGid}
        style={s.smallButton}
        onClick={() => submitPortalCell(fetcher, {
          intent: "move_fabric_row",
          gid,
          rowIndex,
          targetGid,
        })}
      >
        Move row
      </button>
    </div>
  );
}

function formatFabricNumber(value: number) {
  return Number(value).toLocaleString("en-AU", { maximumFractionDigits: 1 });
}

function formatCurrency(value: number) {
  return Number(value).toLocaleString("en-AU", {
    style: "currency",
    currency: "INR",
    notation: "compact",
    maximumFractionDigits: 0,
  });
}

function formatAudCurrency(value: number) {
  return Number(value).toLocaleString("en-AU", {
    style: "currency",
    currency: "AUD",
    notation: "compact",
    maximumFractionDigits: 1,
  });
}

// ─── Nav Order Section ────────────────────────────────────────────────────────

function NavOrderSection({ canManageUsers, settingsFetcher }: { canManageUsers: boolean; settingsFetcher: ReturnType<typeof useFetcher> }) {
  const { navOrder } = useLoaderData<typeof import("./portal._index").loader>();
  const [order, setOrder] = useState<NavItemId[]>(navOrder);
  const dragId = useRef<NavItemId | null>(null);

  const saveOrder = () => {
    const formData = new FormData();
    formData.set("intent", "update_nav_order");
    formData.set("value", JSON.stringify(order));
    settingsFetcher.submit(formData, { method: "post" });
  };

  return (
    <section style={s.settingsCard}>
      <div style={s.settingsHeader}>
        <div>
          <h2 style={s.settingsTitle}>Menu Order</h2>
          <p style={s.settingsHint}>Drag items to reorder the sidebar navigation.</p>
        </div>
        <button type="button" disabled={!canManageUsers} style={s.loginButton} onClick={saveOrder}>
          Save order
        </button>
      </div>
      {!canManageUsers && (
        <div style={s.settingsWarning}>Only an admin user can change menu order.</div>
      )}
      <div style={{ display: "flex", flexDirection: "column", gap: 6, marginTop: 12 }}>
        {order.map((id) => {
          const item = ALL_NAV_ITEMS.find((n) => n.id === id);
          if (!item) return null;
          return (
            <div
              key={id}
              draggable={canManageUsers}
              onDragStart={() => { dragId.current = id; }}
              onDragOver={(e) => {
                e.preventDefault();
                if (dragId.current === null || dragId.current === id) return;
                const from = order.indexOf(dragId.current);
                const to = order.indexOf(id);
                if (from === -1 || to === -1) return;
                const next = [...order];
                next.splice(from, 1);
                next.splice(to, 0, dragId.current);
                setOrder(next);
              }}
              onDragEnd={() => { dragId.current = null; }}
              style={{
                display: "flex", alignItems: "center", gap: 10,
                padding: "10px 14px", borderRadius: 8,
                background: "#f8fafc", border: "1px solid #e5e7eb",
                cursor: canManageUsers ? "grab" : "default",
                userSelect: "none",
              }}
            >
              <span style={{ color: "#9ca3af", fontSize: 18, lineHeight: 1 }}>⠿</span>
              <span style={{ fontWeight: 600, fontSize: 13, color: "#111827" }}>{item.label}</span>
            </div>
          );
        })}
      </div>
    </section>
  );
}

const ROLE_LABELS: Record<PortalUserRole, string> = { superadmin: "Super Admin", admin: "Admin", user: "User" };

function AddUserForm({ currentUser, onAdd }: { currentUser: PortalUser | null; onAdd: (f: Record<string, string>) => void }) {
  const [name, setName] = useState("");
  const [password, setPassword] = useState("");
  const [role, setRole] = useState<PortalUserRole>("user");
  const allowedRoles: PortalUserRole[] = currentUser?.role === "superadmin" ? ["superadmin", "admin", "user"] : ["admin", "user"];
  return (
    <div style={{ background: "#f8fafc", border: "1px dashed #cbd5e1", borderRadius: 8, padding: 16, display: "flex", flexDirection: "column" as const, gap: 10 }}>
      <strong style={{ fontSize: 13 }}>Add new user</strong>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
        <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Full name (used to log in)" style={s.addUserInput} />
        <input value={password} onChange={(e) => setPassword(e.target.value)} type="password" placeholder="Password" style={s.addUserInput} />
        <select value={role} onChange={(e) => setRole(e.target.value as PortalUserRole)} style={s.addUserInput}>
          {allowedRoles.map((r) => <option key={r} value={r}>{ROLE_LABELS[r]}</option>)}
        </select>
      </div>
      <button type="button" style={s.loginButton} onClick={() => { if (!name || !password) return; onAdd({ name, password, role }); setName(""); setPassword(""); setRole("user"); }}>
        Add user
      </button>
    </div>
  );
}

function UserEditForm({
  user, currentUser, navItems, onSave,
}: {
  user: PortalUser;
  currentUser: PortalUser | null;
  navItems: typeof ALL_NAV_ITEMS;
  onSave: (fields: Record<string, string>) => void;
}) {
  const [name, setName] = useState(user.name);
  const [password, setPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [role, setRole] = useState<PortalUserRole>(user.role);
  const [pageAccess, setPageAccess] = useState<Record<string, boolean>>(user.pageAccess ?? {});
  const [canLoadInventory, setCanLoadInventory] = useState(user.canLoadInventory);
  const allowedRoles: PortalUserRole[] = currentUser?.role === "superadmin" ? ["superadmin", "admin", "user"] : ["admin", "user"];
  const showPageAccess = role !== "superadmin";
  return (
    <div style={{ width: "100%", background: "#f8fafc", borderRadius: 8, padding: 14, display: "flex", flexDirection: "column" as const, gap: 12, marginTop: 4 }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
        <label style={{ fontSize: 12, fontWeight: 600, display: "flex", flexDirection: "column" as const, gap: 4 }}>
          Name <span style={{ fontWeight: 400, color: "#9ca3af" }}>(used to log in)</span>
          <input value={name} onChange={(e) => setName(e.target.value)} style={s.addUserInput} />
        </label>
        <label style={{ fontSize: 12, fontWeight: 600, display: "flex", flexDirection: "column" as const, gap: 4 }}>
          New password <span style={{ fontWeight: 400, color: "#9ca3af" }}>(leave blank to keep)</span>
          <div style={{ position: "relative" }}>
            <input value={password} onChange={(e) => setPassword(e.target.value)} type={showPassword ? "text" : "password"} placeholder="••••••••" style={{ ...s.addUserInput, paddingRight: 34 }} />
            <button type="button" onClick={() => setShowPassword((v) => !v)} style={{ position: "absolute", right: 8, top: "50%", transform: "translateY(-50%)", background: "none", border: "none", cursor: "pointer", padding: 2, color: "#9ca3af", display: "flex", alignItems: "center" }} title={showPassword ? "Hide password" : "Show password"}>
              {showPassword ? (
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94"/><path d="M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19"/><line x1="1" y1="1" x2="23" y2="23"/></svg>
              ) : (
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>
              )}
            </button>
          </div>
        </label>
        <label style={{ fontSize: 12, fontWeight: 600, display: "flex", flexDirection: "column" as const, gap: 4 }}>
          Role
          <select value={role} onChange={(e) => setRole(e.target.value as PortalUserRole)} style={s.addUserInput}>
            {allowedRoles.map((r) => <option key={r} value={r}>{ROLE_LABELS[r]}</option>)}
          </select>
        </label>
      </div>

      <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
        <input type="checkbox" checked={canLoadInventory} onChange={(e) => setCanLoadInventory(e.target.checked)} />
        Can load Shopify inventory (Packing Lists)
      </label>

      {showPageAccess && (
        <div>
          <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 8 }}>Page access</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(160px, 1fr))", gap: 6 }}>
            {navItems.map((item) => (
              <label key={item.id} style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, cursor: "pointer", padding: "4px 0" }}>
                <input
                  type="checkbox"
                  checked={Boolean(pageAccess[item.id])}
                  onChange={(e) => setPageAccess((p) => ({ ...p, [item.id]: e.target.checked }))}
                />
                {item.label}
              </label>
            ))}
          </div>
        </div>
      )}

      <button
        type="button"
        style={s.loginButton}
        onClick={() => {
          const fields: Record<string, string> = { name, role, canLoadInventory: canLoadInventory ? "on" : "off", pageAccess: JSON.stringify(pageAccess) };
          if (password) fields.password = password;
          onSave(fields);
        }}
      >
        Save changes
      </button>
    </div>
  );
}

function SettingsPanel({
  users,
  currentUser,
  restockSettings,
  universalSettings,
}: {
  users: PortalUser[];
  currentUser: PortalUser | null;
  restockSettings: RestockSettings;
  universalSettings: UniversalSettings;
}) {
  const settingsFetcher = useFetcher();
  const canManageUsers = currentUser?.role === "superadmin" || currentUser?.role === "admin";
  const [restockDraft, setRestockDraft] = useState<RestockSettings>(restockSettings);
  const [universalDraft, setUniversalDraft] = useState<UniversalSettings>(universalSettings);
  const [editUserId, setEditUserId] = useState<string | null>(null);
  const [universalJustSaved, setUniversalJustSaved] = useState(false);
  useEffect(() => {
    if (!universalJustSaved) return;
    const timer = window.setTimeout(() => setUniversalJustSaved(false), 3000);
    return () => window.clearTimeout(timer);
  }, [universalJustSaved]);
  const saveRestockSettings = () => submitPortalCell(
    settingsFetcher,
    {
      intent: "update_restock_settings",
      value: JSON.stringify(restockDraft),
    },
    { label: "Undo restock settings", fields: { intent: "update_restock_settings", value: JSON.stringify(restockSettings) } },
  );
  const saveUniversalSettings = () => submitPortalCell(
    settingsFetcher,
    {
      intent: "update_universal_settings",
      value: JSON.stringify(universalDraft),
    },
    { label: "Undo universal settings", fields: { intent: "update_universal_settings", value: JSON.stringify(universalSettings) } },
  );
  return (
    <div style={s.settingsPanel}>

      {/* ── Account ──────────────────────────────────────── */}
      <section style={s.settingsCard}>
        <div style={s.settingsHeader}>
          <div>
            <h2 style={s.settingsTitle}>Your account</h2>
            <p style={s.settingsHint}>Logged in as <strong>{currentUser?.name}</strong> ({currentUser?.role})</p>
          </div>
          {currentUser && (
            <form method="post">
              <input type="hidden" name="intent" value="portal_logout" />
              <button type="submit" style={s.secondaryButton}>Sign out</button>
            </form>
          )}
        </div>
      </section>

      {/* ── User Management ──────────────────────────────── */}
      {canManageUsers && (
        <section style={s.settingsCard}>
          <div style={s.settingsHeader}>
            <div>
              <h2 style={s.settingsTitle}>Users</h2>
              <p style={s.settingsHint}>Manage who can access the portal and what they can see.</p>
            </div>
          </div>

          <div style={s.userList}>
            {users.map((user) => {
              const isEditing = editUserId === user.id;
              const canEdit = currentUser?.role === "superadmin" || (currentUser?.role === "admin" && user.role !== "superadmin");
              const canDelete = canEdit && user.id !== currentUser?.id;
              return (
                <div key={user.id} style={{ ...s.userRow, flexWrap: "wrap" as const, gap: 8 }}>
                  <span style={s.activeUserBadge}>{initialsForName(user.name)}</span>
                  <span style={s.userName}>{user.name}</span>
                  <span style={{ ...s.adminPill, background: user.role === "superadmin" ? "#fef3c7" : user.role === "admin" ? "#dbeafe" : "#f1f5f9", color: user.role === "superadmin" ? "#92400e" : user.role === "admin" ? "#1e40af" : "#374151" }}>{user.role}</span>
                  <div style={{ marginLeft: "auto", display: "flex", gap: 6 }}>
                    {canEdit && (
                      <button type="button" style={s.secondaryButton} onClick={() => setEditUserId(isEditing ? null : user.id)}>
                        {isEditing ? "Cancel" : "Edit"}
                      </button>
                    )}
                    {canDelete && (
                      <button type="button" style={s.removeUserButton} onClick={() => {
                        if (window.confirm(`Remove ${user.name}?`)) settingsFetcher.submit({ intent: "remove_portal_user", userId: user.id }, { method: "post" });
                      }}>Remove</button>
                    )}
                  </div>

                  {isEditing && (
                    <UserEditForm
                      user={user}
                      currentUser={currentUser}
                      navItems={ALL_NAV_ITEMS}
                      onSave={(fields) => { settingsFetcher.submit({ intent: "update_portal_user", userId: user.id, ...fields }, { method: "post" }); setEditUserId(null); }}
                    />
                  )}
                </div>
              );
            })}
          </div>

          <div style={{ marginTop: 16 }}>
            <AddUserForm currentUser={currentUser} onAdd={(fields) => settingsFetcher.submit({ intent: "add_portal_user", ...fields }, { method: "post" })} />
          </div>
        </section>
      )}

      <section style={s.settingsCard}>
        <div style={s.settingsHeader}>
          <div>
            <h2 style={s.settingsTitle}>Universal Settings</h2>
            <p style={s.settingsHint}>Shared button, table text, and heading styling across restock and packing pages.</p>
          </div>
          <button
            type="button"
            disabled={!canManageUsers}
            style={s.loginButton}
            onClick={() => {
              saveUniversalSettings();
              setUniversalJustSaved(true);
            }}
          >
            {universalJustSaved ? "Settings saved" : "Save universal settings"}
          </button>
        </div>

        <div style={s.settingsSubCard}>
          <h3 style={s.settingsSubTitle}>Buttons</h3>
          <div style={s.settingsInlineFields}>
            <label style={s.settingsFieldLabel}>
              Button colour
              <ColorPickerInput value={universalDraft.primaryButtonBg} disabled={!canManageUsers}
                onChange={(hex) => setUniversalDraft((c) => ({ ...c, primaryButtonBg: hex }))} />
            </label>
            <label style={s.settingsFieldLabel}>
              Button text
              <ColorPickerInput value={universalDraft.primaryButtonColor} disabled={!canManageUsers}
                onChange={(hex) => setUniversalDraft((c) => ({ ...c, primaryButtonColor: hex }))} />
            </label>
            <span style={{ ...s.buttonPreview, background: universalDraft.primaryButtonBg, color: universalDraft.primaryButtonColor }}>
              Button
            </span>
          </div>
        </div>

        <div style={s.settingsSubCard}>
          <h3 style={s.settingsSubTitle}>Table text</h3>
          <div style={s.settingsInlineFields}>
            <label style={s.settingsFieldLabel}>
              Text size
              <input
                type="number"
                min={10}
                max={20}
                value={universalDraft.tableTextSize}
                disabled={!canManageUsers}
                onChange={(event) => {
                  const v = Number(event.currentTarget.value);
                  setUniversalDraft((current) => ({
                    ...current,
                    tableTextSize: v || current.tableTextSize,
                  }));
                }}
                style={s.settingsSmallInput}
              />
            </label>
            <label style={s.settingsFieldLabel}>
              Text colour
              <ColorPickerInput value={universalDraft.tableTextColor} disabled={!canManageUsers}
                onChange={(hex) => setUniversalDraft((c) => ({ ...c, tableTextColor: hex }))} />
            </label>
            <span style={{ ...s.qtyPreview, fontSize: universalDraft.tableTextSize, color: universalDraft.tableTextColor }}>
              Table text
            </span>
          </div>
        </div>

        <div style={s.settingsSubCard}>
          <h3 style={s.settingsSubTitle}>Quantity numbers</h3>
          <p style={s.settingsHint}>Size of the per-size qty cells and TOTAL on restock and packing tables.</p>
          <div style={s.settingsInlineFields}>
            <label style={s.settingsFieldLabel}>
              Font size
              <input
                type="number"
                min={9}
                max={32}
                value={universalDraft.inventoryFontSize}
                disabled={!canManageUsers}
                onChange={(event) => {
                  const v = Number(event.currentTarget.value);
                  setUniversalDraft((current) => ({
                    ...current,
                    inventoryFontSize: v || current.inventoryFontSize,
                  }));
                }}
                style={s.settingsSmallInput}
              />
            </label>
            <span style={{ ...s.qtyPreview, fontSize: universalDraft.inventoryFontSize, color: "#374151" }}>
              25
            </span>
          </div>
        </div>

        <div style={s.settingsSubCard}>
          <h3 style={s.settingsSubTitle}>Slide-out panel text (Samples / Vision Board)</h3>
          <div style={s.settingsInlineFields}>
            <label style={s.settingsFieldLabel}>
              Text size
              <input
                type="number"
                min={11}
                max={22}
                value={universalDraft.panelTextSize}
                disabled={!canManageUsers}
                onChange={(event) => {
                  const v = Number(event.target.value);
                  setUniversalDraft((current) => ({
                    ...current,
                    panelTextSize: v || current.panelTextSize,
                  }));
                }}
                style={s.settingsSmallInput}
              />
            </label>
            <span style={{ ...s.qtyPreview, fontSize: universalDraft.panelTextSize, color: "#111827" }}>
              Panel text
            </span>
          </div>
        </div>

        <div style={s.settingsSubCard}>
          <h3 style={s.settingsSubTitle}>Headings</h3>
          <div style={s.settingsInlineFields}>
            <label style={s.settingsFieldLabel}>
              Heading size
              <input
                type="number"
                min={14}
                max={34}
                value={universalDraft.headingTextSize}
                disabled={!canManageUsers}
                onChange={(event) => {
                  const v = Number(event.currentTarget.value);
                  setUniversalDraft((current) => ({
                    ...current,
                    headingTextSize: v || current.headingTextSize,
                  }));
                }}
                style={s.settingsSmallInput}
              />
            </label>
            <label style={s.settingsFieldLabel}>
              Heading colour
              <ColorPickerInput value={universalDraft.headingTextColor} disabled={!canManageUsers}
                onChange={(hex) => setUniversalDraft((c) => ({ ...c, headingTextColor: hex }))} />
            </label>
            <span style={{ ...s.headingPreview, fontSize: universalDraft.headingTextSize, color: universalDraft.headingTextColor }}>
              Heading
            </span>
          </div>
        </div>

        <div style={s.settingsSubCard}>
          <h3 style={s.settingsSubTitle}>Menu colours</h3>
          <div style={s.settingsInlineFields}>
            <label style={s.settingsFieldLabel}>
              Menu background
              <ColorPickerInput value={universalDraft.menuBg} disabled={!canManageUsers}
                onChange={(hex) => setUniversalDraft((c) => ({ ...c, menuBg: hex }))} />
            </label>
            <label style={s.settingsFieldLabel}>
              Menu text
              <ColorPickerInput value={universalDraft.menuTextColor} disabled={!canManageUsers}
                onChange={(hex) => setUniversalDraft((c) => ({ ...c, menuTextColor: hex }))} />
            </label>
            <span style={{ ...s.buttonPreview, background: universalDraft.menuBg, color: universalDraft.menuTextColor }}>
              Menu
            </span>
          </div>
        </div>

        <div style={s.settingsSubCard}>
          <h3 style={s.settingsSubTitle}>Page background</h3>
          <div style={s.settingsInlineFields}>
            <label style={s.settingsFieldLabel}>
              Page background colour
              <ColorPickerInput value={universalDraft.pageBg} disabled={!canManageUsers}
                onChange={(hex) => setUniversalDraft((c) => ({ ...c, pageBg: hex }))} />
            </label>
            <span style={{ ...s.buttonPreview, background: universalDraft.pageBg, color: "#111827", border: "1px solid #e5e7eb" }}>
              Page
            </span>
          </div>
        </div>

        <div style={s.settingsSubCard}>
          <h3 style={s.settingsSubTitle}>Logo</h3>
          <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
            {universalDraft.logoUrl ? (
              <img src={universalDraft.logoUrl} alt="Logo preview" style={{ maxHeight: 64, maxWidth: 160, borderRadius: 4, border: "1px solid #e5e7eb", objectFit: "contain", background: "#f9fafb", padding: 4 }} />
            ) : (
              <div style={{ width: 160, height: 64, borderRadius: 4, border: "1px dashed #d1d5db", display: "flex", alignItems: "center", justifyContent: "center", color: "#9ca3af", fontSize: 12 }}>
                No logo
              </div>
            )}
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              <label style={{ cursor: canManageUsers ? "pointer" : "default" }}>
                <input
                  type="file"
                  accept="image/*"
                  disabled={!canManageUsers}
                  style={{ display: "none" }}
                  onChange={(e) => {
                    const file = e.target.files?.[0];
                    if (!file) return;
                    const reader = new FileReader();
                    reader.onload = () => {
                      setUniversalDraft((c) => ({ ...c, logoUrl: reader.result as string }));
                    };
                    reader.readAsDataURL(file);
                    e.target.value = "";
                  }}
                />
                <span style={{ ...s.loginButton, pointerEvents: canManageUsers ? "auto" : "none", opacity: canManageUsers ? 1 : 0.5 }}>
                  {universalDraft.logoUrl ? "Replace image" : "Upload image"}
                </span>
              </label>
              {universalDraft.logoUrl && (
                <button
                  type="button"
                  disabled={!canManageUsers}
                  onClick={() => setUniversalDraft((c) => ({ ...c, logoUrl: "" }))}
                  style={{ background: "none", border: "none", color: "#d72c0d", cursor: "pointer", fontSize: 12, padding: 0, textAlign: "left" }}
                >
                  Remove logo
                </button>
              )}
            </div>
          </div>
        </div>
      </section>

      <NavOrderSection canManageUsers={canManageUsers} settingsFetcher={settingsFetcher} />

      <section style={s.settingsCard}>
        <div style={s.settingsHeader}>
          <div>
            <h2 style={s.settingsTitle}>Existing Product Restock Settings</h2>
            <p style={s.settingsHint}>Edit table number styling for the restock page.</p>
          </div>
          <button type="button" disabled={!canManageUsers} style={s.loginButton} onClick={saveRestockSettings}>
            Save restock settings
          </button>
        </div>

        {!canManageUsers && (
          <div style={s.settingsWarning}>Only an admin user can change restock settings.</div>
        )}

        <div style={s.settingsSubCard}>
          <h3 style={s.settingsSubTitle}>Quantity numbers</h3>
          <div style={s.settingsInlineFields}>
            <label style={s.settingsFieldLabel}>
              Font size
              <input
                type="number"
                min={10}
                max={32}
                value={restockDraft.quantityFontSize}
                disabled={!canManageUsers}
                onChange={(event) => {
                  const v = Number(event.currentTarget.value);
                  setRestockDraft((current) => ({
                    ...current,
                    quantityFontSize: v || current.quantityFontSize,
                  }));
                }}
                style={s.settingsSmallInput}
              />
            </label>
            <label style={s.settingsFieldLabel}>
              Font colour
              <ColorPickerInput value={restockDraft.quantityFontColor} disabled={!canManageUsers}
                onChange={(hex) => setRestockDraft((c) => ({ ...c, quantityFontColor: hex }))} />
            </label>
            <label style={s.settingsFieldLabel}>
              Inventory arrow colour
              <ColorPickerInput value={restockDraft.inventoryArrowColor} disabled={!canManageUsers}
                onChange={(hex) => setRestockDraft((c) => ({ ...c, inventoryArrowColor: hex }))} />
            </label>
            <span style={{ ...s.qtyPreview, fontSize: restockDraft.quantityFontSize, color: restockDraft.quantityFontColor }}>
              25
            </span>
            <span style={{ ...s.inventoryArrowPreview, color: restockDraft.inventoryArrowColor }}>▼</span>
          </div>
        </div>
      </section>

    </div>
  );
}

function ActivityLogPanel({
  activityLogs,
}: {
  activityLogs: { id: number; userName: string; action: string; entity: string; entityId: string | null; entityName: string | null; field: string | null; toValue: string | null; createdAt: Date | string }[];
}) {
  const [expandedDate, setExpandedDate] = useState<string | null>(null);
  const safeLog = Array.isArray(activityLogs) ? activityLogs : [];

  const grouped = safeLog.reduce<Record<string, typeof safeLog>>((acc, log) => {
    const dateKey = new Date(log.createdAt).toLocaleDateString("en-AU", {
      day: "2-digit", month: "short", year: "numeric",
    });
    if (!acc[dateKey]) acc[dateKey] = [];
    acc[dateKey].push(log);
    return acc;
  }, {});

  const dates = Object.keys(grouped);

  return (
    <section style={s.settingsCard}>
      <h2 style={s.settingsTitle}>Change Log</h2>
      <p style={s.settingsHint}>All changes made in the last 90 days. Click a date to see that day's activity.</p>

      {dates.length === 0 ? (
        <div style={{ marginTop: 16, color: "#6b7280", fontSize: 13 }}>No changes recorded yet.</div>
      ) : (
        <div style={{ marginTop: 14, display: "grid", gap: 8 }}>
          {dates.map((dateKey) => {
            const entries = grouped[dateKey];
            const isOpen = expandedDate === dateKey;
            return (
              <div key={dateKey} style={s.logDateBlock}>
                <button
                  type="button"
                  onClick={() => setExpandedDate(isOpen ? null : dateKey)}
                  style={s.logDateButton}
                >
                  <span>{dateKey}</span>
                  <span style={s.logDateCount}>{entries.length} change{entries.length !== 1 ? "s" : ""}</span>
                  <span style={{ marginLeft: "auto", fontSize: 12, color: "#6b7280" }}>{isOpen ? "▲" : "▼"}</span>
                </button>

                {isOpen && (
                  <div style={s.logEntries}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                      <thead>
                        <tr>
                          <th style={s.logTh}>Time</th>
                          <th style={s.logTh}>User</th>
                          <th style={s.logTh}>Action</th>
                          <th style={s.logTh}>Item</th>
                          <th style={s.logTh}>Field</th>
                          <th style={s.logTh}>New value</th>
                        </tr>
                      </thead>
                      <tbody>
                        {entries.map((log, idx) => (
                          <tr key={log.id} style={{ background: idx % 2 === 0 ? "#fff" : "#f8fafc" }}>
                            <td style={s.logTd}>
                              {new Date(log.createdAt).toLocaleTimeString("en-AU", { hour: "2-digit", minute: "2-digit" })}
                            </td>
                            <td style={{ ...s.logTd, fontWeight: 700 }}>{log.userName}</td>
                            <td style={s.logTd}>{log.action}</td>
                            <td style={s.logTd}>{log.entityName ?? log.entity}</td>
                            <td style={{ ...s.logTd, color: "#6b7280" }}>{log.field ?? "—"}</td>
                            <td style={{ ...s.logTd, color: "#111827" }}>{log.toValue ?? "—"}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </section>
  );
}

type PackingListWithLines = Awaited<ReturnType<typeof loader>>["packingLists"][number];

function PackingListsPanel({
  packingLists,
  selectedPackingList,
  savedPackingColumnWidths,
  tableHeaderLabels,
  customColumns,
  customCells,
  rowHeights,
  productSearch,
  packingSearchLineId,
  productResults,
  searchTitle,
  updateParams,
  canLoadInventory,
  canEditLockedQuantities,
  isAdmin,
  shopDomain,
}: {
  packingLists: PackingListWithLines[];
  selectedPackingList: PackingListWithLines | null;
  savedPackingColumnWidths: Record<string, number>;
  tableHeaderLabels: Record<string, string>;
  customColumns: TableCustomColumn[];
  customCells: Record<string, string>;
  rowHeights: Record<string, number>;
  productSearch: string;
  packingSearchLineId: number | null;
  productResults: ShopifySearchProduct[];
  searchTitle: string;
  updateParams: (updates: Record<string, string>) => void;
  canLoadInventory: boolean;
  canEditLockedQuantities: boolean;
  isAdmin: boolean;
  shopDomain: string | null;
}) {
  const fetcher = useFetcher();

  return (
    <div style={s.packingLayout}>
      {!selectedPackingList ? (
        <PackingListsOverview packingLists={packingLists} fetcher={fetcher} searchTitle={searchTitle} />
      ) : (
        <section style={s.packingDetail}>
          <PackingListDetail
            packingList={selectedPackingList}
            savedPackingColumnWidths={savedPackingColumnWidths}
            tableHeaderLabels={tableHeaderLabels}
            customColumns={customColumns}
            customCells={customCells}
            rowHeights={rowHeights}
            productSearch={productSearch}
            packingSearchLineId={packingSearchLineId}
            productResults={productResults}
            headerSearch={searchTitle}
            updateParams={updateParams}
            canLoadInventory={canLoadInventory}
            canEditLockedQuantities={canEditLockedQuantities}
            isAdmin={isAdmin}
            shopDomain={shopDomain}
          />
        </section>
      )}
    </div>
  );
}

function PackingListsOverview({
  packingLists,
  fetcher,
  searchTitle,
}: {
  packingLists: PackingListWithLines[];
  fetcher: ReturnType<typeof useFetcher>;
  searchTitle: string;
}) {
  const [searchParams] = useSearchParams();
  const [hoveredListId, setHoveredListId] = useState<number | null>(null);
  const [deleteWarningList, setDeleteWarningList] = useState<PackingListWithLines | null>(null);
  const showHidden = searchParams.get("showHidden") === "true";
  const visibleLists = packingLists.filter((list) => !list.hiddenAt);
  const hiddenLists = packingLists.filter((list) => list.hiddenAt);
  const baseRows = showHidden ? hiddenLists : visibleLists;
  const rows = searchTitle
    ? baseRows.filter((list) => {
        const q = searchTitle.toLowerCase();
        return (list.invoiceNumber ?? "").toLowerCase().includes(q)
          || (list.title ?? "").toLowerCase().includes(q)
          || `packing list #${list.id}`.includes(q);
      })
    : baseRows;
  const isImporting = fetcher.state !== "idle" && String(fetcher.formData?.get("intent") ?? "") === "import_supplier_packing_csv";

  return (
    <div style={s.packingOverview}>
      <section style={s.packingOverviewCreate}>
        <fetcher.Form method="post" style={s.packingCreateForm}>
          <input type="hidden" name="intent" value="create_packing_list" />
          <button type="submit" style={s.loginButton}>New packing list</button>
        </fetcher.Form>
        <fetcher.Form method="post" encType="multipart/form-data" style={s.packingImportForm}>
          <input type="hidden" name="intent" value="import_supplier_packing_csv" />
          <input type="file" name="packingCsv" accept=".csv,text/csv" required style={s.fileInput} />
          <button type="submit" style={isImporting ? { ...s.secondaryButton, ...s.busyButton } : s.secondaryButton} disabled={fetcher.state !== "idle"}>
            {isImporting ? "Importing..." : "Import supplier CSV"}
          </button>
        </fetcher.Form>
      </section>

      <section style={s.packingOverviewTableWrap}>
        <div style={s.packingOverviewBar}>
          <strong>{showHidden ? "Hidden packing lists" : "Packing lists"}</strong>
          <a href={`/portal?page=packing${showHidden ? "" : "&showHidden=true"}`} style={s.secondaryButton}>
            {showHidden ? "Show active lists" : `Show hidden lists (${hiddenLists.length})`}
          </a>
        </div>
        <table style={{ ...s.table, width: "100%" }}>
          <thead>
            <tr style={s.headerRow}>
              {["Invoice", "Boxes", "Total qty", "Estimated arrival", "Shipping", "Status", "Actions"].map((heading) => (
                <th key={heading} style={{ ...s.th, textAlign: heading === "Total qty" || heading === "Boxes" || heading === "Actions" ? "center" : "left" }}>
                  <span style={s.thContent}>{heading}</span>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.length ? rows.map((list) => {
              const isHovered = hoveredListId === list.id;
              const cellStyle = {
                ...s.td,
                ...(isHovered ? s.clickableOverviewCellHover : {}),
              };
              return (
                <tr
                  key={list.id}
                  style={s.clickableOverviewRow}
                  onClick={() => {
                    window.location.href = `/portal?page=packing&packingId=${list.id}`;
                  }}
                  onMouseEnter={() => setHoveredListId(list.id)}
                  onMouseLeave={() => setHoveredListId(null)}
                >
                  <td
                    style={cellStyle}
                    onContextMenu={(e) => { e.preventDefault(); e.stopPropagation(); document.dispatchEvent(new CustomEvent("show-cell-history", { detail: { x: e.clientX, y: e.clientY, entity: "Packing List", entityId: String(list.id), field: "Invoice number", entityName: list.invoiceNumber || `Packing list #${list.id}` } })); }}
                  >
                    <strong style={s.productName}>{list.invoiceNumber || `Packing list #${list.id}`}</strong>
                  </td>
                  <td style={{ ...cellStyle, textAlign: "center" }}><span style={s.total}>{packingListBoxCount(list)}</span></td>
                  <td style={{ ...cellStyle, textAlign: "center" }}><span style={s.total}>{packingListTotal(list)}</span></td>
                  <td
                    style={cellStyle}
                    onContextMenu={(e) => { e.preventDefault(); e.stopPropagation(); document.dispatchEvent(new CustomEvent("show-cell-history", { detail: { x: e.clientX, y: e.clientY, entity: "Packing List", entityId: String(list.id), field: "Estimated arrival", entityName: list.invoiceNumber || `Packing list #${list.id}` } })); }}
                  >{formatPortalDate(list.expectedLeaveFactoryDate ?? list.shipmentDate) || "—"}</td>
                  <td
                    style={cellStyle}
                    onContextMenu={(e) => { e.preventDefault(); e.stopPropagation(); document.dispatchEvent(new CustomEvent("show-cell-history", { detail: { x: e.clientX, y: e.clientY, entity: "Packing List", entityId: String(list.id), field: "Shipping method", entityName: list.invoiceNumber || `Packing list #${list.id}` } })); }}
                  >{list.shippingMethod ? list.shippingMethod.charAt(0).toUpperCase() + list.shippingMethod.slice(1) : "—"}</td>
                  <td
                    style={cellStyle}
                    onContextMenu={(e) => { e.preventDefault(); e.stopPropagation(); document.dispatchEvent(new CustomEvent("show-cell-history", { detail: { x: e.clientX, y: e.clientY, entity: "Packing List", entityId: String(list.id), field: "Status", entityName: list.invoiceNumber || `Packing list #${list.id}` } })); }}
                  >{labelForPackingStatus(list.status)}</td>
                  <td style={{ ...cellStyle, textAlign: "center" }}>
                    <div style={s.packingListActions} onClick={(event) => event.stopPropagation()}>
                      <fetcher.Form method="post">
                        <input type="hidden" name="intent" value="set_packing_list_hidden" />
                        <input type="hidden" name="packingId" value={list.id} />
                        <input type="hidden" name="hidden" value={showHidden ? "false" : "true"} />
                        <button type="submit" style={showHidden ? s.smallButton : s.hideListButton}>
                          {showHidden ? "Show" : "Hide"}
                        </button>
                      </fetcher.Form>
                      <button type="button" style={s.removeUserButton} onClick={() => setDeleteWarningList(list)}>
                        Delete
                      </button>
                    </div>
                  </td>
                </tr>
              );
            }) : (
              <tr style={s.row}>
                <td colSpan={6} style={{ ...s.td, textAlign: "center", padding: 40 }}>
                  {showHidden ? "No hidden packing lists." : "No packing lists yet."}
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </section>
      {deleteWarningList ? (
        <div style={s.deleteConfirm} onClick={() => setDeleteWarningList(null)}>
          <div style={s.deleteConfirmCard} onClick={(event) => event.stopPropagation()}>
            <div style={s.deleteConfirmTitle}>Delete packing list?</div>
            <div style={s.deleteConfirmText}>
              Are you sure you want to delete this list? You can also hide it from the list instead.
            </div>
            <div style={s.deleteConfirmActions}>
              <fetcher.Form method="post" onSubmit={() => setDeleteWarningList(null)}>
                <input type="hidden" name="intent" value="set_packing_list_hidden" />
                <input type="hidden" name="packingId" value={deleteWarningList.id} />
                <input type="hidden" name="hidden" value="true" />
                <button type="submit" style={s.hideListButton}>Hide</button>
              </fetcher.Form>
              <fetcher.Form method="post" onSubmit={() => setDeleteWarningList(null)}>
                <input type="hidden" name="intent" value="delete_packing_list" />
                <input type="hidden" name="packingId" value={deleteWarningList.id} />
                <button type="submit" style={{ ...s.deleteConfirmButton, ...s.deleteConfirmDanger }}>
                  Delete
                </button>
              </fetcher.Form>
              <button type="button" style={s.deleteConfirmButton} onClick={() => setDeleteWarningList(null)}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}

function PackingListDetail({
  packingList,
  savedPackingColumnWidths,
  tableHeaderLabels,
  customColumns,
  customCells,
  rowHeights,
  productSearch,
  packingSearchLineId,
  productResults,
  headerSearch = "",
  updateParams,
  canLoadInventory,
  canEditLockedQuantities,
  isAdmin,
  shopDomain,
}: {
  packingList: PackingListWithLines;
  savedPackingColumnWidths: Record<string, number>;
  tableHeaderLabels: Record<string, string>;
  customColumns: TableCustomColumn[];
  customCells: Record<string, string>;
  rowHeights: Record<string, number>;
  productSearch: string;
  packingSearchLineId: number | null;
  productResults: ShopifySearchProduct[];
  headerSearch?: string;
  updateParams: (updates: Record<string, string>) => void;
  canLoadInventory: boolean;
  canEditLockedQuantities: boolean;
  isAdmin: boolean;
  shopDomain: string | null;
}) {
  const fetcher = useFetcher();
  const loadInventoryFetcher = useFetcher();
  const columnWidthsFetcher = useFetcher();
  const [packingColumnWidths, setPackingColumnWidths] = useState<Record<string, number>>(savedPackingColumnWidths);
  const [skipWords, setSkipWords] = useState("");
  const [packingListSearch, setPackingListSearch] = useState("");
  const [statusValue, setStatusValue] = useState(packingList.status ?? "still_packing");
  useEffect(() => setStatusValue(packingList.status ?? "still_packing"), [packingList.status]);
  const [combineView, setCombineView] = useState(false);
  const bulkLoadedFetcher = useFetcher();
  const relinkFetcher = useFetcher<{ relinked: number; scanned: number; error?: string }>();
  const unlinkedLinesCount = packingList.lines.filter((line) => !line.productId && (line.productTitle ?? "").trim().length > 0).length;
  const [combinedSelectedCells, setCombinedSelectedCells] = useState<Set<string>>(new Set());
  const [combinedAnchorCell, setCombinedAnchorCell] = useState<string | null>(null);
  const [combinedCellMenu, setCombinedCellMenu] = useState<{ x: number; y: number; cells: { lineId: number; size: string }[] } | null>(null);
  useEffect(() => {
    if (!combineView) {
      setCombinedSelectedCells(new Set());
      setCombinedAnchorCell(null);
      setCombinedCellMenu(null);
    }
  }, [combineView]);
  useEffect(() => {
    if (!combinedCellMenu) return;
    const close = () => setCombinedCellMenu(null);
    document.addEventListener("mousedown", close);
    document.addEventListener("keydown", close);
    return () => {
      document.removeEventListener("mousedown", close);
      document.removeEventListener("keydown", close);
    };
  }, [combinedCellMenu]);
  // Quantities are editable only while "still_packing"; once it moves past that
  // they freeze (and status can't revert to still_packing) — unless an admin.
  const quantitiesLocked = (packingList.status ?? "still_packing") !== "still_packing" && !canEditLockedQuantities;
  const packingSizes = derivePackingSizes(packingList.lines);
  const packingColumns = [
    ...packingColumnsForSizes(packingSizes).filter((col) => col.id !== "shopify" || combineView),
    ...customColumns.map((column) => ({ id: column.id, label: column.label, width: 130, center: false })),
  ];
  const packingWidthFor = (columnId: string) => packingColumnWidths[columnId] ?? defaultPackingColumnWidth(columnId);
  const packingTableWidth = packingColumns.reduce((sum, column) => sum + packingWidthFor(column.id), 48);
  const normalizedPackingListSearch = (packingListSearch || headerSearch).trim().toLowerCase();
  const visiblePackingLines = normalizedPackingListSearch
    ? packingList.lines.filter((line) => packingLineMatchesSearch(line, normalizedPackingListSearch))
    : packingList.lines;
  const exportPackingList = () => {
    const headers = [
      "Box",
      "Name",
      "SKU",
      ...packingSizes,
      "Total",
      "Price rupees",
      "Value rupees",
      "Weight",
    ];
    const rows = packingList.lines.map((line) => {
      const qtys = normalizeQtys(line.qtys);
      const total = packingTotal(qtys);
      const price = line.priceRupees ?? 0;
      return [
        line.boxNumber ?? "",
        line.productTitle,
        line.sku ?? "",
        ...packingSizes.map((size) => qtys[size] || ""),
        total || "",
        line.priceRupees ?? "",
        total && price ? Math.round(total * price) : "",
        line.weight ?? "",
      ];
    });
    const totalQty = packingList.lines.reduce((acc, line) => acc + packingTotal(normalizeQtys(line.qtys)), 0);
    const totalValue = packingList.lines.reduce((acc, line) => {
      const qty = packingTotal(normalizeQtys(line.qtys));
      const price = line.priceRupees ?? 0;
      return acc + (qty && price ? Math.round(qty * price) : 0);
    }, 0);
    const totalWeight = packingList.lines.reduce((acc, line) => acc + (line.weight ?? 0), 0);
    const totalsRow = [
      "TOTAL",
      "",
      "",
      ...packingSizes.map(() => ""),
      totalQty || "",
      "",
      totalValue || "",
      totalWeight ? `${totalWeight.toFixed(2)}kg` : "",
    ];
    const csv = [headers, ...rows, [], totalsRow].map((row) => (row as (string | number)[]).map(csvCell).join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const fileName = `${packingList.invoiceNumber || packingList.title || `packing-list-${packingList.id}`}`
      .replace(/[^a-z0-9-]+/gi, "-")
      .replace(/^-+|-+$/g, "")
      .toLowerCase() || `packing-list-${packingList.id}`;
    link.href = url;
    link.download = `${fileName}.csv`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  };

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
      submitPortalCell(
        columnWidthsFetcher,
        { intent: "update_packing_column_widths", value: JSON.stringify(nextColumnWidths) },
        { label: "Undo column width", fields: { intent: "update_packing_column_widths", value: JSON.stringify(packingColumnWidths) } },
      );
    };

    document.addEventListener("mousemove", handleMove);
    document.addEventListener("mouseup", handleUp);
  };

  return (
    <div style={s.packingDetailInner}>
      <div style={s.packingTop}>
        {/* Row 1: navigation + core fields */}
        <div style={s.packingTopRow}>
          <a href="/portal?page=packing" style={s.secondaryButton}>Back</a>
          <button type="button" style={s.secondaryButton} onClick={exportPackingList}>Export packing list</button>
          <label style={s.packingToolbarLabel}>
            <span>Invoice number</span>
            <input
              defaultValue={packingList.invoiceNumber ?? ""}
              onBlur={(event) => submitPortalCell(
                fetcher,
                {
                  intent: "update_packing_list",
                  packingId: packingList.id,
                  field: "invoiceNumber",
                  value: event.currentTarget.value,
                },
                { label: "Undo invoice number", fields: { intent: "update_packing_list", packingId: packingList.id, field: "invoiceNumber", value: packingList.invoiceNumber ?? "" } },
              )}
              placeholder="Invoice number"
              style={{ ...s.packingInput, ...s.invoiceInput }}
            />
          </label>
          <label style={s.packingToolbarLabel}>
            <span>Estimated arrival</span>
            <input
              key={String(packingList.expectedLeaveFactoryDate ?? packingList.shipmentDate ?? "")}
              type="date"
              defaultValue={(() => {
                const d = packingList.expectedLeaveFactoryDate ?? packingList.shipmentDate;
                if (!d) return "";
                const dt = new Date(d);
                return isNaN(dt.getTime()) ? "" : dt.toISOString().slice(0, 10);
              })()}
              onBlur={(event) => submitPortalCell(
                fetcher,
                {
                  intent: "update_packing_list",
                  packingId: packingList.id,
                  field: "expectedLeaveFactoryDate",
                  value: event.currentTarget.value,
                },
                { label: "Undo estimated arrival", fields: { intent: "update_packing_list", packingId: packingList.id, field: "expectedLeaveFactoryDate", value: packingList.expectedLeaveFactoryDate ? new Date(packingList.expectedLeaveFactoryDate).toISOString().slice(0, 10) : "" } },
              )}
              style={{ ...s.packingInput, width: 150 }}
            />
          </label>
          <label style={s.packingToolbarLabel}>
            <span>Status</span>
            <select
              value={statusValue}
              onChange={(event) => {
                const next = event.currentTarget.value;
                setStatusValue(next);
                submitPortalCell(
                  fetcher,
                  {
                    intent: "update_packing_list",
                    packingId: packingList.id,
                    field: "status",
                    value: next,
                  },
                  { label: "Undo packing status", fields: { intent: "update_packing_list", packingId: packingList.id, field: "status", value: packingList.status ?? "still_packing" } },
                );
              }}
              style={{ ...s.packingInput, width: 160 }}
            >
              {PACKING_STATUS_OPTIONS.map((opt) => {
                // Once the list has moved past "still_packing", non-admins can't
                // set it back (which would re-open quantity editing).
                const lockStillPacking = opt.value === "still_packing"
                  && (packingList.status ?? "still_packing") !== "still_packing"
                  && !canEditLockedQuantities;
                return (
                  <option key={opt.value} value={opt.value} disabled={lockStillPacking}>{opt.label}</option>
                );
              })}
            </select>
          </label>
          <label style={s.packingToolbarLabel}>
            <span>Shipping method</span>
            <select
              key={packingList.shippingMethod ?? ""}
              defaultValue={packingList.shippingMethod ?? ""}
              onChange={(event) => submitPortalCell(
                fetcher,
                {
                  intent: "update_packing_list",
                  packingId: packingList.id,
                  field: "shippingMethod",
                  value: event.currentTarget.value,
                },
                { label: "Undo shipping method", fields: { intent: "update_packing_list", packingId: packingList.id, field: "shippingMethod", value: packingList.shippingMethod ?? "" } },
              )}
              style={{ ...s.packingInput, width: 120 }}
            >
              <option value="">—</option>
              <option value="sea">Sea</option>
              <option value="air">Air</option>
            </select>
          </label>
        </div>
        {/* Row 2: load inventory (left) + total quantity (right) */}
        <div style={s.packingBottomRow}>
          {canLoadInventory && !packingList.masterInventoryLoadedAt ? (
            <loadInventoryFetcher.Form
              method="post"
              style={s.loadInventoryForm}
              onSubmit={(event) => {
                const ok = window.confirm("Add these packing list quantities to current Shopify stock? You can only do this once — afterwards, use the per-product Load button in combined view.");
                if (!ok) event.preventDefault();
              }}
            >
              <input type="hidden" name="intent" value="load_packing_inventory" />
              <input type="hidden" name="packingId" value={packingList.id} />
              <label style={s.packingToolbarLabel}>
                <span>Skip words</span>
                <input
                  name="skipWords"
                  value={skipWords}
                  onChange={(event) => setSkipWords(event.currentTarget.value)}
                  placeholder="acacia, sample, fabric"
                  style={{ ...s.packingInput, ...s.skipWordsInput }}
                />
              </label>
              <button type="submit" style={{ ...s.loginButton, ...s.loadInventoryButton }} disabled={loadInventoryFetcher.state !== "idle"}>
                {loadInventoryFetcher.state === "idle" ? "Load inventory on Shopify" : "Loading..."}
              </button>
            </loadInventoryFetcher.Form>
          ) : <div />}
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            {unlinkedLinesCount > 0 && (
              <button
                type="button"
                disabled={relinkFetcher.state !== "idle"}
                onClick={() => {
                  if (!window.confirm(`Re-check Shopify for ${unlinkedLinesCount} unlinked product${unlinkedLinesCount === 1 ? "" : "s"}? Quantities won't change — this only links lines to Shopify products so admins can load inventory later.`)) return;
                  submitPortalCell(relinkFetcher, {
                    intent: "relink_packing_lines_to_shopify",
                    packingId: packingList.id,
                  });
                }}
                style={{
                  padding: "8px 14px",
                  borderRadius: 6,
                  border: "1px solid #cbd5e1",
                  background: "#fff",
                  color: "#1f2937",
                  fontWeight: 700,
                  fontSize: 13,
                  cursor: relinkFetcher.state !== "idle" ? "wait" : "pointer",
                }}
                title="Search Shopify by product title for lines that aren't linked yet, and link any exact matches."
              >
                {relinkFetcher.state !== "idle"
                  ? "Re-checking…"
                  : relinkFetcher.data
                    ? `Re-check Shopify (${unlinkedLinesCount} left, ${relinkFetcher.data.relinked} linked)`
                    : `Re-check Shopify (${unlinkedLinesCount} unlinked)`}
              </button>
            )}
            {isAdmin && (
              <button
                type="button"
                onClick={() => setCombineView((current) => !current)}
                style={{
                  padding: "8px 14px",
                  borderRadius: 6,
                  border: "1px solid #cbd5e1",
                  background: combineView ? "#2563eb" : "#fff",
                  color: combineView ? "#fff" : "#1f2937",
                  fontWeight: 700,
                  fontSize: 13,
                  cursor: "pointer",
                }}
                title="Combine rows with the same product into a single row"
              >
                {combineView ? "✓ Combined by product" : "Combine by product"}
              </button>
            )}
            <div style={s.packingTotalPill}>
              Total quantity <strong>{packingListTotal(packingList)}</strong>
            </div>
          </div>
        </div>
      </div>

      <div style={s.packingSearchBar}>
        <label style={s.packingToolbarLabel}>
          <span>Search packing list</span>
          <input
            value={packingListSearch}
            onChange={(event) => setPackingListSearch(event.currentTarget.value)}
            placeholder="Search title, SKU, box, quantity..."
            style={{ ...s.packingInput, ...s.packingSearchInput }}
          />
        </label>
        {packingListSearch ? (
          <button type="button" style={s.smallButton} onClick={() => setPackingListSearch("")}>Clear</button>
        ) : null}
        <span style={s.searchCount}>
          {visiblePackingLines.length} of {packingList.lines.length} rows
        </span>
      </div>

      <div className="portal-table-scroll" style={s.packingTableWrap}>
        <table style={{ ...s.table, width: packingTableWidth, minWidth: "100%" }} onKeyDown={handleTableGridKeyDown}>
          <colgroup>
            <col style={{ width: 48 }} />
            {packingColumns.map((column) => (
              <col key={column.id} style={{ width: packingWidthFor(column.id) }} />
            ))}
          </colgroup>
          <thead>
            <tr style={s.headerRow}>
              <th style={{ ...s.th, ...s.rowNumberHeader }}>#</th>
              {(() => {
                const packingFrozenLeft = [
                  48,
                  48 + packingWidthFor("box"),
                  48 + packingWidthFor("box") + packingWidthFor("picture"),
                  48 + packingWidthFor("box") + packingWidthFor("picture") + packingWidthFor("fabric"),
                ];
                return packingColumns.map((column, colIdx) => (
                  <Th
                    key={column.id}
                    center={column.center}
                    headerKey={`packing:${column.id}`}
                    columnId={column.id}
                    onResizeStart={(event) => startPackingResize(column.id, event)}
                    stickyLeft={colIdx < 4 ? packingFrozenLeft[colIdx] : undefined}
                    isLastFrozen={colIdx === 3}
                  >
                    {headerLabel(tableHeaderLabels, `packing:${column.id}`, column.label)}
                  </Th>
                ));
              })()}
            </tr>
          </thead>
          <tbody>
            {combineView ? (() => {
              const combinedRows = buildCombinedPackingRows(visiblePackingLines);
              const rowKeyToIndex = new Map<string, number>();
              const rowKeyToLineIds = new Map<string, number[]>();
              combinedRows.forEach((row, idx) => {
                rowKeyToIndex.set(row.key, idx);
                rowKeyToLineIds.set(row.key, row.lineIds);
              });
              const cellKey = (rowKey: string, size: string) => `${rowKey}|${size}`;
              const parseCellKey = (key: string): { rowKey: string; size: string } => {
                const sepIdx = key.lastIndexOf("|");
                return { rowKey: key.slice(0, sepIdx), size: key.slice(sepIdx + 1) };
              };
              const expandToLineCells = (selectionKeys: string[]): { lineId: number; size: string }[] => {
                const out: { lineId: number; size: string }[] = [];
                for (const key of selectionKeys) {
                  const { rowKey, size } = parseCellKey(key);
                  const lineIds = rowKeyToLineIds.get(rowKey) ?? [];
                  for (const lineId of lineIds) out.push({ lineId, size });
                }
                return out;
              };
              const buildRectSelection = (anchor: string, target: string): Set<string> => {
                const a = parseCellKey(anchor);
                const b = parseCellKey(target);
                const ar = rowKeyToIndex.get(a.rowKey);
                const br = rowKeyToIndex.get(b.rowKey);
                if (ar === undefined || br === undefined) return new Set([target]);
                const ac = packingSizes.indexOf(a.size);
                const bc = packingSizes.indexOf(b.size);
                if (ac < 0 || bc < 0) return new Set([target]);
                const r1 = Math.min(ar, br); const r2 = Math.max(ar, br);
                const c1 = Math.min(ac, bc); const c2 = Math.max(ac, bc);
                const next = new Set<string>();
                for (let r = r1; r <= r2; r += 1) {
                  const row = combinedRows[r];
                  for (let c = c1; c <= c2; c += 1) {
                    const size = packingSizes[c];
                    if ((row.qtys[size] ?? 0) > 0) next.add(cellKey(row.key, size));
                  }
                }
                return next;
              };
              const handleCellClick = (rowKey: string, size: string, event: React.MouseEvent) => {
                const key = cellKey(rowKey, size);
                if (event.shiftKey && combinedAnchorCell) {
                  setCombinedSelectedCells(buildRectSelection(combinedAnchorCell, key));
                } else if (event.metaKey || event.ctrlKey) {
                  const next = new Set(combinedSelectedCells);
                  if (next.has(key)) next.delete(key); else next.add(key);
                  setCombinedSelectedCells(next);
                  setCombinedAnchorCell(key);
                } else {
                  setCombinedSelectedCells(new Set([key]));
                  setCombinedAnchorCell(key);
                }
              };
              const handleCellContextMenu = (rowKey: string, size: string, event: React.MouseEvent) => {
                event.preventDefault();
                const key = cellKey(rowKey, size);
                let cells = combinedSelectedCells;
                if (!cells.has(key)) {
                  cells = new Set([key]);
                  setCombinedSelectedCells(cells);
                  setCombinedAnchorCell(key);
                }
                setCombinedCellMenu({
                  x: event.clientX,
                  y: event.clientY,
                  cells: expandToLineCells(Array.from(cells)),
                });
              };
              return combinedRows.map((row, rowIndex) => (
                <PackingCombinedRow
                  key={row.key}
                  row={row}
                  rowIndex={rowIndex}
                  customColumns={customColumns}
                  isAdmin={isAdmin}
                  shopDomain={shopDomain}
                  packingId={packingList.id}
                  selectedCells={combinedSelectedCells}
                  onCellClick={handleCellClick}
                  onCellContextMenu={handleCellContextMenu}
                  sizes={packingSizes}
                />
              ));
            })() : (visiblePackingLines.length ? visiblePackingLines.map((line, rowIndex) => {
              const packingFrozenOffsets = [
                48,
                48 + packingWidthFor("box"),
                48 + packingWidthFor("box") + packingWidthFor("picture"),
                48 + packingWidthFor("box") + packingWidthFor("picture") + packingWidthFor("fabric"),
              ];
              return (
              <PackingListLineRow
                key={line.id}
                line={line}
                rowIndex={rowIndex}
                updateParams={updateParams}
                customColumns={customColumns}
                customCells={customCells}
                rowHeights={rowHeights}
                frozenOffsets={packingFrozenOffsets}
                quantitiesLocked={quantitiesLocked}
                isAdmin={isAdmin}
                shopDomain={shopDomain}
                showShopifyColumn={combineView}
                sizes={packingSizes}
              />
              );
            }) : (
              <tr style={s.row}>
                <td colSpan={packingColumns.length + 1} style={{ ...s.td, textAlign: "center", padding: 40 }}>
                  No packing list rows match this search.
                </td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            {(() => {
              const totalQty = visiblePackingLines.reduce((sum, line) => sum + packingTotal(normalizeQtys(line.qtys)), 0);
              const totalValue = visiblePackingLines.reduce((sum, line) => {
                const qty = packingTotal(normalizeQtys(line.qtys));
                return sum + qty * (line.priceRupees ?? 0);
              }, 0);
              const packingFrozenOffsets = [
                48,
                48 + packingWidthFor("box"),
                48 + packingWidthFor("box") + packingWidthFor("picture"),
                48 + packingWidthFor("box") + packingWidthFor("picture") + packingWidthFor("fabric"),
              ];
              return (
                <tr style={{ ...s.row, background: "#eef2f7" }}>
                  <td style={{ ...s.rowNumberCell, background: "#e2e8f0" }}>—</td>
                  {packingColumns.map((col, colIdx) => {
                    const isLastFrozen = colIdx === 3;
                    const frozenStyle: React.CSSProperties = colIdx < 4 ? {
                      position: "sticky",
                      left: packingFrozenOffsets[colIdx],
                      zIndex: 40,
                      ...(isLastFrozen ? { boxShadow: "4px 0 6px -2px rgba(0,0,0,0.1)" } : {}),
                    } : {};
                    let content: React.ReactNode = null;
                    if (col.id === "name") content = <span style={{ fontWeight: 700, color: "#374151", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em" }}>Totals</span>;
                    if (col.id === "total") content = <span style={s.total}>{totalQty || ""}</span>;
                    if (col.id === "value") content = <span style={s.total}>{totalValue ? Math.round(totalValue) : ""}</span>;
                    return (
                      <td key={col.id} style={{ ...s.td, background: "#eef2f7", fontWeight: 700, textAlign: col.center ? "center" : "left", ...frozenStyle }}>{content}</td>
                    );
                  })}
                </tr>
              );
            })()}
          </tfoot>
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
      {combinedCellMenu && typeof document !== "undefined" && createPortal(
        <div
          style={{ ...s.contextMenu, left: combinedCellMenu.x, top: combinedCellMenu.y }}
          onMouseDown={(event) => event.stopPropagation()}
        >
          <div style={{ padding: "6px 10px", fontSize: 11, color: "#64748b", fontWeight: 600, borderBottom: "1px solid #e2e8f0" }}>
            {combinedCellMenu.cells.length} cell{combinedCellMenu.cells.length === 1 ? "" : "s"} selected
          </div>
          <button
            type="button"
            style={s.contextMenuButton}
            onClick={() => {
              const cells = combinedCellMenu.cells;
              setCombinedCellMenu(null);
              submitPortalCell(bulkLoadedFetcher, {
                intent: "bulk_set_packing_qty_manual_loaded",
                packingId: packingList.id,
                action: "mark",
                cells: JSON.stringify(cells),
              });
            }}
          >
            Mark as loaded
          </button>
          <button
            type="button"
            style={s.contextMenuButton}
            onClick={() => {
              const cells = combinedCellMenu.cells;
              setCombinedCellMenu(null);
              submitPortalCell(bulkLoadedFetcher, {
                intent: "bulk_set_packing_qty_manual_loaded",
                packingId: packingList.id,
                action: "unmark",
                cells: JSON.stringify(cells),
              });
            }}
          >
            Unmark loaded
          </button>
        </div>,
        document.body,
      )}
    </div>
  );
}

function PackingListLineRow({
  line,
  rowIndex,
  updateParams,
  customColumns,
  customCells,
  rowHeights,
  frozenOffsets,
  quantitiesLocked,
  isAdmin,
  shopDomain,
  showShopifyColumn,
  sizes,
}: {
  line: PackingListWithLines["lines"][number];
  rowIndex: number;
  updateParams: (updates: Record<string, string>) => void;
  customColumns: TableCustomColumn[];
  customCells: Record<string, string>;
  rowHeights: Record<string, number>;
  frozenOffsets?: number[];
  quantitiesLocked: boolean;
  isAdmin: boolean;
  shopDomain: string | null;
  showShopifyColumn: boolean;
  sizes: string[];
}) {
  const fetcher = useFetcher();
  const qtys = normalizeQtys(line.qtys);
  const shopifyLoadedQtys = normalizeQtys(line.shopifyLoadedQtys);
  const manuallyLoadedQtys = normalizeQtys(line.manuallyLoadedQtys);
  // A size is considered "loaded" if either the real Shopify push or the
  // manual-mark matches the current packed quantity.
  const isLoadedForSize = (size: string) => qtys[size] > 0 && (
    shopifyLoadedQtys[size] === qtys[size] || manuallyLoadedQtys[size] === qtys[size]
  );
  const total = packingTotal(qtys);
  const price = line.priceRupees ?? 0;
  const value = total * price;

  const rowHeightKey = `packing:${line.id}`;

  return (
    <tr style={{ ...s.row, ...(rowHeights[rowHeightKey] ? { height: rowHeights[rowHeightKey] } : {}) }}>
      <RowNumberCell rowNumber={rowIndex + 1} actions={[
        { label: "Add row", onClick: () => submitPortalCell(fetcher, { intent: "add_custom_packing_line", packingId: line.packingListId }) },
        { label: "Duplicate row", onClick: () => submitPortalCell(fetcher, { intent: "duplicate_packing_line", lineId: line.id }) },
        { label: "Move up", onClick: () => submitPortalCell(fetcher, { intent: "move_packing_line", lineId: line.id, direction: "up" }) },
        { label: "Move down", onClick: () => submitPortalCell(fetcher, { intent: "move_packing_line", lineId: line.id, direction: "down" }) },
        { label: "Delete row", danger: true, onClick: () => { if (window.confirm("Delete this packing row?")) submitPortalCell(fetcher, { intent: "delete_packing_line", lineId: line.id }); } },
      ]} heightKey={rowHeightKey} />
      <PackingTd rowIndex={rowIndex} colIndex={0} stickyLeft={frozenOffsets?.[0]}><PackingTextInput lineId={line.id} field="boxNumber" value={line.boxNumber ?? ""} /></PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={1} center stickyLeft={frozenOffsets?.[1]}><PackingImageCell lineId={line.id} field="productImageUrl" value={line.productImageUrl ?? ""} /></PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={2} center stickyLeft={frozenOffsets?.[2]}><PackingImageCell lineId={line.id} field="fabricImageData" value={line.fabricImageData ?? ""} /></PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={3} overflowVisible stickyLeft={frozenOffsets?.[3]} isLastFrozen>
        <PackingProductNameCell
          line={line}
          updateParams={updateParams}
        />
      </PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={4}><PackingSkuCell lineId={line.id} value={line.sku ?? ""} /></PackingTd>
      {sizes.map((size, sizeIndex) => (
        <PackingTd
          key={size}
          rowIndex={rowIndex}
          colIndex={5 + sizeIndex}
          center
          style={{
            ...(isLoadedForSize(size) ? s.loadedInventoryCell : {}),
          }}
          onContextMenu={(e) => { e.preventDefault(); document.dispatchEvent(new CustomEvent("show-cell-history", { detail: { x: e.clientX, y: e.clientY, entity: "Packing List Line", entityId: String(line.id), field: `Qty (${size})`, entityName: line.productTitle } })); }}
        >
          <input
            type="text"
            inputMode="numeric"
            readOnly={quantitiesLocked}
            defaultValue={qtys[size] || ""}
            onChange={(event) => { if (!quantitiesLocked) event.currentTarget.value = event.currentTarget.value.replace(/\D/g, ""); }}
            onBlur={quantitiesLocked ? undefined : (event) => submitPortalCell(
              fetcher,
              {
                intent: "update_packing_qty",
                lineId: line.id,
                size,
                value: event.currentTarget.value,
              },
              { label: "Undo packing quantity", fields: { intent: "update_packing_qty", lineId: line.id, size, value: qtys[size] ?? 0 } },
            )}
            title={quantitiesLocked
              ? "Quantities are locked once the list moves past Still packing"
              : undefined}
            style={{
              ...s.qtyInput,
              ...(isLoadedForSize(size) ? { color: "#fff" } : {}),
              ...(quantitiesLocked ? { cursor: "not-allowed", background: "#f1f5f9" } : {}),
            }}
          />
        </PackingTd>
      ))}
      <PackingTd rowIndex={rowIndex} colIndex={5 + sizes.length} center><span style={s.total}>{total}</span></PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={6 + sizes.length} center><PackingTextInput lineId={line.id} field="priceRupees" value={line.priceRupees?.toString() ?? ""} center /></PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={7 + sizes.length} center><span style={s.total}>{value ? Math.round(value) : ""}</span></PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={8 + sizes.length} center><PackingTextInput lineId={line.id} field="weight" value={line.weight?.toString() ?? ""} center /></PackingTd>
      {showShopifyColumn && (
        <PackingTd rowIndex={rowIndex} colIndex={9 + sizes.length} center>
          {isAdmin && line.productId && shopDomain ? (
            <a
              href={`https://${shopDomain}/admin/products/${line.productId.replace(/^gid:\/\/shopify\/Product\//, "")}`}
              target="_blank"
              rel="noopener noreferrer"
              title="Open product in Shopify admin"
              style={{
                display: "inline-flex",
                alignItems: "center",
                justifyContent: "center",
                padding: "4px 10px",
                background: "#ecfeff",
                border: "1px solid #67e8f9",
                color: "#0e7490",
                borderRadius: 4,
                fontSize: 12,
                fontWeight: 700,
                textDecoration: "none",
              }}
            >
              Open ↗
            </a>
          ) : null}
        </PackingTd>
      )}
      {customColumns.map((column, customIndex) => (
        <PackingTd key={column.id} rowIndex={rowIndex} colIndex={10 + sizes.length + customIndex}>
          <TableCustomCell cellKey={`packing:${line.id}:${column.id}`} value={customCells[`packing:${line.id}:${column.id}`] ?? ""} />
        </PackingTd>
      ))}
    </tr>
  );
}

type CombinedPackingRow = {
  key: string;
  productId: string | null;
  isCustom: boolean;
  productTitle: string;
  productImageUrl: string | null;
  fabricImageData: string | null;
  sku: string | null;
  boxNumbers: string[];
  lineIds: number[];
  qtys: Record<string, number>;
  shopifyLoadedQtys: Record<string, number>;
  manuallyLoadedQtys: Record<string, number>;
  totalPrice: number;
  totalWeight: number;
};

function buildCombinedPackingRows(lines: PackingListWithLines["lines"]): CombinedPackingRow[] {
  const byKey = new Map<string, CombinedPackingRow>();
  const order: string[] = [];
  // Convention: only the first row of a new box has the box number written in;
  // subsequent rows in the same box leave it blank. Track the "current box" as
  // we iterate so inherited rows still attribute to the right box.
  let currentBox = "";
  for (const line of lines) {
    const explicitBox = (line.boxNumber ?? "").trim();
    if (explicitBox) currentBox = explicitBox;
    const effectiveBox = currentBox;
    const key = line.productId
      ? `pid:${line.productId}`
      : `title:${(line.productTitle ?? "").trim().toLowerCase() || `line:${line.id}`}`;
    let entry = byKey.get(key);
    if (!entry) {
      entry = {
        key,
        productId: line.productId,
        isCustom: Boolean(line.isCustom),
        productTitle: line.productTitle ?? "",
        productImageUrl: line.productImageUrl ?? null,
        fabricImageData: line.fabricImageData ?? null,
        sku: line.sku ?? null,
        boxNumbers: [],
        lineIds: [],
        qtys: {},
        shopifyLoadedQtys: {},
        manuallyLoadedQtys: {},
        totalPrice: 0,
        totalWeight: 0,
      };
      byKey.set(key, entry);
      order.push(key);
    }
    if (effectiveBox && !entry.boxNumbers.includes(effectiveBox)) {
      entry.boxNumbers.push(effectiveBox);
    }
    entry.lineIds.push(line.id);
    const lineQtys = normalizeQtys(line.qtys);
    const lineLoaded = normalizeQtys(line.shopifyLoadedQtys);
    const lineManual = normalizeQtys(line.manuallyLoadedQtys);
    for (const [size, qty] of Object.entries(lineQtys)) {
      entry.qtys[size] = (entry.qtys[size] ?? 0) + qty;
    }
    for (const [size, qty] of Object.entries(lineLoaded)) {
      entry.shopifyLoadedQtys[size] = (entry.shopifyLoadedQtys[size] ?? 0) + qty;
    }
    for (const [size, qty] of Object.entries(lineManual)) {
      entry.manuallyLoadedQtys[size] = (entry.manuallyLoadedQtys[size] ?? 0) + qty;
    }
    if (typeof line.priceRupees === "number") entry.totalPrice += line.priceRupees;
    if (typeof line.weight === "number") entry.totalWeight += line.weight;
  }
  // Sort box numbers numerically when possible so 1,2,10 reads correctly.
  for (const entry of byKey.values()) {
    entry.boxNumbers.sort((a, b) => {
      const an = Number(a); const bn = Number(b);
      if (Number.isFinite(an) && Number.isFinite(bn)) return an - bn;
      return a.localeCompare(b);
    });
  }
  return order.map((key) => byKey.get(key)!).filter(Boolean);
}

function PackingCombinedRow({
  row,
  rowIndex,
  customColumns,
  isAdmin,
  shopDomain,
  packingId,
  selectedCells,
  onCellClick,
  onCellContextMenu,
  sizes,
}: {
  row: CombinedPackingRow;
  rowIndex: number;
  customColumns: TableCustomColumn[];
  isAdmin: boolean;
  shopDomain: string | null;
  packingId: number;
  selectedCells: Set<string>;
  onCellClick: (rowKey: string, size: string, event: React.MouseEvent) => void;
  onCellContextMenu: (rowKey: string, size: string, event: React.MouseEvent) => void;
  sizes: string[];
}) {
  const fetcher = useFetcher();
  const isLoadedForSize = (size: string) => {
    const want = row.qtys[size] ?? 0;
    if (want <= 0) return false;
    const got = (row.shopifyLoadedQtys[size] ?? 0) + (row.manuallyLoadedQtys[size] ?? 0);
    return got >= want;
  };
  const allSizesLoaded = Object.entries(row.qtys).every(([size, qty]) => qty <= 0 || isLoadedForSize(size));
  const canLoadThisRow = isAdmin && Boolean(row.productId) && !row.isCustom && !allSizesLoaded;
  const adminUrl = isAdmin && row.productId && shopDomain
    ? `https://${shopDomain}/admin/products/${row.productId.replace(/^gid:\/\/shopify\/Product\//, "")}`
    : null;
  const total = packingTotal(row.qtys);
  const value = total * (row.totalPrice / Math.max(1, row.boxNumbers.length || 1));
  const fabricImageSrc = row.fabricImageData || "";
  return (
    <tr style={{ ...s.row, background: "#fafbfc" }}>
      <RowNumberCell rowNumber={rowIndex + 1} actions={[]} />
      <PackingTd rowIndex={rowIndex} colIndex={0} center>
        <span style={{ color: "#6b7280", fontSize: 12, fontWeight: 600 }}>
          {row.boxNumbers.length ? row.boxNumbers.join(", ") : "—"}
        </span>
      </PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={1} center>
        {row.productImageUrl ? <img src={row.productImageUrl} alt="" style={{ width: 90, height: 110, objectFit: "cover", borderRadius: 4 }} /> : null}
      </PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={2} center>
        {fabricImageSrc ? <img src={fabricImageSrc} alt="" style={{ width: 90, height: 110, objectFit: "cover", borderRadius: 4 }} /> : null}
      </PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={3}>
        <div style={{ position: "relative", padding: "8px 10px", minHeight: 96 }}>
          <div style={{
            fontWeight: 700,
            color: "#111827",
            fontSize: "var(--portal-table-font-size, 13px)",
            lineHeight: 1.3,
          }}>
            {row.productTitle}
          </div>
          {canLoadThisRow && (
            <button
              type="button"
              onClick={() => {
                if (!window.confirm(`Load inventory for "${row.productTitle}" to Shopify?`)) return;
                submitPortalCell(
                  fetcher,
                  {
                    intent: "load_packing_inventory_for_product",
                    packingId,
                    productId: row.productId ?? "",
                  },
                );
              }}
              disabled={fetcher.state !== "idle"}
              style={{
                position: "absolute",
                bottom: 6,
                left: 6,
                padding: "4px 8px",
                background: "#0f766e",
                color: "#fff",
                border: "none",
                borderRadius: 4,
                fontSize: 11,
                fontWeight: 700,
                cursor: "pointer",
              }}
            >
              {fetcher.state !== "idle" ? "Loading…" : "Load inventory"}
            </button>
          )}
          {!canLoadThisRow && row.productId && !row.isCustom && allSizesLoaded && (
            <span style={{ position: "absolute", bottom: 6, left: 6, color: "#0f766e", fontSize: 11, fontWeight: 700 }}>
              ✓ Loaded
            </span>
          )}
        </div>
      </PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={4}>
        <span style={{ display: "block", padding: "8px 10px", color: "#374151", fontSize: 12, whiteSpace: "pre-wrap" }}>{row.sku ?? ""}</span>
      </PackingTd>
      {sizes.map((size, sizeIndex) => {
        const hasQty = (row.qtys[size] ?? 0) > 0;
        const canSelect = isAdmin && hasQty;
        const cellKey = `${row.key}|${size}`;
        const isSelected = canSelect ? selectedCells.has(cellKey) : false;
        return (
          <PackingTd
            key={size}
            rowIndex={rowIndex}
            colIndex={5 + sizeIndex}
            center
            style={{
              ...(isLoadedForSize(size) ? s.loadedInventoryCell : {}),
              ...(canSelect ? { cursor: "pointer" } : {}),
              ...(isSelected ? { outline: "2px solid #2563eb", outlineOffset: -2, position: "relative", zIndex: 1 } : {}),
            }}
            onClick={canSelect ? (event) => {
              onCellClick(row.key, size, event);
            } : undefined}
            onContextMenu={canSelect ? (event) => {
              onCellContextMenu(row.key, size, event);
            } : undefined}
          >
            <span
              style={{ display: "block", padding: "8px 0", fontWeight: 700, color: isLoadedForSize(size) ? "#fff" : "#111827" }}
              title={canSelect ? "Click to select. Shift+click to extend, Cmd/Ctrl+click to toggle. Right-click selection to mark loaded." : undefined}
            >
              {row.qtys[size] || ""}
            </span>
          </PackingTd>
        );
      })}
      <PackingTd rowIndex={rowIndex} colIndex={5 + sizes.length} center><span style={s.total}>{total}</span></PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={6 + sizes.length} center><span style={s.total}>{row.totalPrice ? Math.round(row.totalPrice) : ""}</span></PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={7 + sizes.length} center><span style={s.total}>{value ? Math.round(value) : ""}</span></PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={8 + sizes.length} center><span style={s.total}>{row.totalWeight ? Math.round(row.totalWeight) : ""}</span></PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={9 + sizes.length} center>
        {adminUrl ? (
          <a
            href={adminUrl}
            target="_blank"
            rel="noopener noreferrer"
            title="Open product in Shopify admin"
            style={{
              display: "inline-flex",
              alignItems: "center",
              justifyContent: "center",
              padding: "4px 10px",
              background: "#ecfeff",
              border: "1px solid #67e8f9",
              color: "#0e7490",
              borderRadius: 4,
              fontSize: 12,
              fontWeight: 700,
              textDecoration: "none",
            }}
          >
            Open ↗
          </a>
        ) : null}
      </PackingTd>
      {customColumns.map((column, customIndex) => (
        <PackingTd key={column.id} rowIndex={rowIndex} colIndex={10 + sizes.length + customIndex} />
      ))}
    </tr>
  );
}

function PackingProductNameCell({
  line,
  updateParams,
}: {
  line: PackingListWithLines["lines"][number];
  updateParams: (updates: Record<string, string>) => void;
}) {
  const fetcher = useFetcher();
  const searchFetcher = useFetcher<{ products: ShopifySearchProduct[] }>();
  const inputRef = useRef<HTMLInputElement | null>(null);
  const displayValue = line.isCustom && line.productTitle === "Custom item" ? "" : line.productTitle;
  const [value, setValue] = useState(displayValue);
  const [isFocused, setIsFocused] = useState(false);
  const [isProductSelected, setIsProductSelected] = useState(false);
  const [isChangingProduct, setIsChangingProduct] = useState(false);
  const [dropdownRect, setDropdownRect] = useState<DOMRect | null>(null);
  const [canPortalDropdown, setCanPortalDropdown] = useState(false);
  const [lastSearchedQuery, setLastSearchedQuery] = useState("");
  const hasLinkedProduct = Boolean(line.productId);
  const canSearch = !isProductSelected && (!hasLinkedProduct || isChangingProduct) && isFocused;
  const shouldShowResults = canSearch && value.trim().length >= 2;
  const searchResults: ShopifySearchProduct[] = searchFetcher.data?.products ?? [];
  const isSearching = searchFetcher.state !== "idle" || value.trim() !== lastSearchedQuery;
  const dropdownHeight = isSearching || !searchResults.length
    ? 48
    : Math.min(320, searchResults.length * 62 + 12);
  const updateDropdownRect = () => {
    if (!inputRef.current) return;
    setDropdownRect(inputRef.current.getBoundingClientRect());
  };

  useEffect(() => {
    setValue(displayValue);
    setIsProductSelected(false);
    setIsChangingProduct(false);
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
    const trimmed = value.trim();
    if (trimmed.length < 2) return;
    const timer = window.setTimeout(() => {
      setLastSearchedQuery(trimmed);
      searchFetcher.load(`/api/packing-search?q=${encodeURIComponent(trimmed)}`);
    }, 450);
    return () => window.clearTimeout(timer);
  }, [value, canSearch]);

  const applyProduct = (product: ShopifySearchProduct) => {
    setValue(product.title);
    setIsFocused(false);
    setIsProductSelected(true);
    setIsChangingProduct(false);
    setDropdownRect(null);
    inputRef.current?.blur();
    submitPortalCell(
      fetcher,
      {
        intent: "apply_product_to_packing_line",
        lineId: line.id,
        product: JSON.stringify(product),
      },
      {
        label: "Undo product selection",
        fields: {
          intent: "update_packing_line",
          lineId: line.id,
          field: "productTitle",
          value: displayValue,
        },
      },
    );
    updateParams({ productSearch: "", packingSearchLineId: "" });
  };

  if (hasLinkedProduct && !isChangingProduct) {
    return (
      <div style={s.linkedProductCell}>
        <span style={s.linkedProductTitle}>{displayValue || "Linked product"}</span>
        <button
          type="button"
          style={s.changeProductButton}
          onClick={() => {
            setIsChangingProduct(true);
            window.setTimeout(() => inputRef.current?.focus(), 0);
          }}
        >
          Change
        </button>
      </div>
    );
  }

  return (
    <div style={s.productCellSearch}>
      <input
        ref={inputRef}
        type="search"
        value={value}
        onFocus={() => {
          setIsFocused(true);
          setIsProductSelected(false);
          if (hasLinkedProduct) setIsChangingProduct(true);
        }}
        onChange={(event) => setValue(event.currentTarget.value)}
        onBlur={(event) => {
          if (isProductSelected) return;
          setIsFocused(false);
          if (hasLinkedProduct) {
            setValue(displayValue);
            setIsChangingProduct(false);
            return;
          }
          submitPortalCell(
            fetcher,
            {
              intent: "update_packing_line",
              lineId: line.id,
              field: "productTitle",
              value: event.currentTarget.value,
            },
            { label: "Undo product title", fields: { intent: "update_packing_line", lineId: line.id, field: "productTitle", value: displayValue } },
          );
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
          {isSearching ? (
            <div style={s.productCellResultEmpty}>Searching...</div>
          ) : searchResults.length ? searchResults.map((product) => (
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
    onBlur: (event: React.FocusEvent<HTMLInputElement | HTMLTextAreaElement>) => submitPortalCell(
      fetcher,
      {
        intent: "update_packing_line",
        lineId,
        field,
        value: event.currentTarget.value,
      },
      { label: "Undo packing cell", fields: { intent: "update_packing_line", lineId, field, value } },
    ),
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
      onBlur={(event) => submitPortalCell(
        fetcher,
        {
          intent: "update_packing_line",
          lineId,
          field: "sku",
          value: event.currentTarget.value,
        },
        { label: "Undo SKU", fields: { intent: "update_packing_line", lineId, field: "sku", value } },
      )}
      style={s.packingSkuTextarea}
    />
  );
}

function PackingImageCell({ lineId, field, value }: { lineId: number; field: "productImageUrl" | "fabricImageData"; value: string }) {
  const fetcher = useFetcher();
  const [imageHover, setImageHover] = useState(false);
  const [dialogOpen, setDialogOpen] = useState(false);
  const [localValue, setLocalValue] = useState(value);
  useEffect(() => setLocalValue(value), [value]);

  const saveFile = async (file: File) => {
    const previewUrl = URL.createObjectURL(file);
    setLocalValue(previewUrl);
    let processed = file;
    try {
      processed = await resizeImageForFabricUpload(file);
    } catch {
      processed = file;
    }
    if (processed.size > 5 * 1024 * 1024) {
      window.alert(`That image is ${(processed.size / 1024 / 1024).toFixed(1)} MB — over the 5 MB limit. Try a smaller one.`);
      return;
    }
    let dataUrl = "";
    try {
      dataUrl = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result ?? ""));
        reader.onerror = () => reject(reader.error ?? new Error("read failed"));
        reader.readAsDataURL(processed);
      });
    } catch {
      window.alert("Couldn't read the image file.");
      return;
    }
    submitPortalCell(
      fetcher,
      {
        intent: "update_packing_line",
        lineId,
        field,
        value: dataUrl,
      },
      { label: "Undo packing image", fields: { intent: "update_packing_line", lineId, field, value } },
    );
  };
  const clearImage = () => {
    setLocalValue("");
    submitPortalCell(
      fetcher,
      {
        intent: "update_packing_line",
        lineId,
        field,
        value: "",
      },
      { label: "Undo image delete", fields: { intent: "update_packing_line", lineId, field, value } },
    );
  };

  return (
    <div style={s.packingImageCell}>
      <div
        style={s.fabricImageDrop}
        onMouseEnter={() => setImageHover(true)}
        onMouseLeave={() => setImageHover(false)}
        onContextMenu={(event) => {
          event.preventDefault();
          setDialogOpen(true);
        }}
        title="Right-click to upload, drop, or paste"
      >
        {localValue ? <img src={localValue} alt="" style={s.fabricThumb} /> : <span>Right-click to add image</span>}
        {localValue && (
          <button
            type="button"
            style={{ ...s.imageDeleteOverlay, ...(imageHover ? s.imageDeleteOverlayVisible : {}) }}
            onClick={(event) => {
              event.preventDefault();
              event.stopPropagation();
              clearImage();
            }}
          >
            Delete
          </button>
        )}
      </div>
      <ImageUploadDialog
        open={dialogOpen}
        onClose={() => setDialogOpen(false)}
        onFile={(file) => {
          setDialogOpen(false);
          void saveFile(file);
        }}
        hasCurrentImage={Boolean(localValue)}
        onRemove={clearImage}
      />
    </div>
  );
}

// ─── Category Dropdown ───────────────────────────────────────────────────────

function CategoryDropdown({
  productGroups,
  showNewCategoryInput,
  newCategoryInput,
  onNewCategoryInput,
  onShowNewCategory,
  onSelect,
  onClose,
}: {
  productGroups: string[];
  showNewCategoryInput: boolean;
  newCategoryInput: string;
  onNewCategoryInput: (value: string) => void;
  onShowNewCategory: () => void;
  onSelect: (category: string) => void;
  onClose: () => void;
}) {
  const ref = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    const handleClick = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) onClose();
    };
    const handleKey = (e: KeyboardEvent) => { if (e.key === "Escape") onClose(); };
    document.addEventListener("mousedown", handleClick);
    document.addEventListener("keydown", handleKey);
    return () => {
      document.removeEventListener("mousedown", handleClick);
      document.removeEventListener("keydown", handleKey);
    };
  }, [onClose]);

  return (
    <div
      ref={ref}
      style={{
        position: "absolute",
        top: "calc(100% + 6px)",
        right: 0,
        zIndex: 9999,
        background: "#fff",
        border: "1px solid #d1d5db",
        borderRadius: 8,
        boxShadow: "0 4px 16px rgba(0,0,0,0.13)",
        minWidth: 220,
        overflow: "hidden",
      }}
    >
      <div style={{ padding: "8px 0" }}>
        <div style={{ padding: "4px 12px 6px", fontSize: 11, fontWeight: 700, color: "#9ca3af", textTransform: "uppercase", letterSpacing: 1 }}>
          Choose category
        </div>
        {productGroups.map((group) => (
          <button
            key={group}
            type="button"
            onClick={() => onSelect(group)}
            style={{
              display: "block",
              width: "100%",
              textAlign: "left",
              padding: "8px 14px",
              background: "none",
              border: "none",
              cursor: "pointer",
              fontSize: 13,
              color: "#111827",
            }}
            onMouseEnter={(e) => { (e.currentTarget as HTMLButtonElement).style.background = "#f3f4f6"; }}
            onMouseLeave={(e) => { (e.currentTarget as HTMLButtonElement).style.background = "none"; }}
          >
            {group}
          </button>
        ))}
        <div style={{ borderTop: "1px solid #e5e7eb", marginTop: 4 }}>
          {showNewCategoryInput ? (
            <div style={{ display: "flex", gap: 6, padding: "8px 10px" }}>
              <input
                autoFocus
                type="text"
                value={newCategoryInput}
                onChange={(e) => onNewCategoryInput(e.currentTarget.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter" && newCategoryInput.trim()) onSelect(newCategoryInput.trim());
                }}
                placeholder="New category name"
                style={{
                  flex: 1,
                  border: "1px solid #d1d5db",
                  borderRadius: 4,
                  padding: "5px 8px",
                  fontSize: 13,
                  outline: "none",
                }}
              />
              <button
                type="button"
                onClick={() => { if (newCategoryInput.trim()) onSelect(newCategoryInput.trim()); }}
                style={{
                  background: "#008060",
                  color: "#fff",
                  border: "none",
                  borderRadius: 4,
                  padding: "5px 10px",
                  fontSize: 12,
                  fontWeight: 700,
                  cursor: "pointer",
                }}
              >
                Add
              </button>
            </div>
          ) : (
            <button
              type="button"
              onClick={onShowNewCategory}
              style={{
                display: "block",
                width: "100%",
                textAlign: "left",
                padding: "8px 14px",
                background: "none",
                border: "none",
                cursor: "pointer",
                fontSize: 13,
                color: "#008060",
                fontWeight: 700,
              }}
              onMouseEnter={(e) => { (e.currentTarget as HTMLButtonElement).style.background = "#f0fdf4"; }}
              onMouseLeave={(e) => { (e.currentTarget as HTMLButtonElement).style.background = "none"; }}
            >
              + Add new category
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── Rows ────────────────────────────────────────────────────────────────────

function AddRestockOrderRow({
  rowIndex,
  sizes,
  initialProductGroup,
  productSearch,
  productResults,
  updateParams,
  restockSettings,
  onSaved,
}: {
  rowIndex: number;
  sizes: string[];
  initialProductGroup?: string;
  productSearch: string;
  productResults: ShopifySearchProduct[];
  updateParams: (updates: Record<string, string>) => void;
  restockSettings: RestockSettings;
  onSaved?: () => void;
}) {
  const fetcher = useFetcher();
  const rowRef = useRef<HTMLTableRowElement | null>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);
  const [searchValue, setSearchValue] = useState("");
  const [selectedProduct, setSelectedProduct] = useState<ShopifySearchProduct | null>(null);
  const [qtys, setQtys] = useState<Record<string, string>>({});
  const [status, setStatus] = useState(restockSettings.statusOptions[0]?.value ?? "on_order");
  const [priority, setPriority] = useState("");
  const [notes, setNotes] = useState("");
  const [eta, setEta] = useState("");
  const [focused, setFocused] = useState(false);
  const [dropdownRect, setDropdownRect] = useState<DOMRect | null>(null);
  const [canPortal, setCanPortal] = useState(false);

  const totalQty = Object.values(qtys).reduce((sum, v) => sum + (Number(v) || 0), 0);
  const totalCol = 5 + sizes.length;
  const statusCol = totalCol + 1;
  const notesCol = totalCol + 2;
  const priorityCol = totalCol + 3;
  const etaCol = totalCol + 4;

  const shouldShowResults = !selectedProduct && focused && searchValue.trim().length >= 2;
  const dropdownHeight = searchValue.trim() !== productSearch || !productResults.length
    ? 48 : Math.min(320, productResults.length * 62 + 12);

  const stateRef = useRef({ selectedProduct, qtys, status, priority, notes, eta });
  useEffect(() => { stateRef.current = { selectedProduct, qtys, status, priority, notes, eta }; });

  useEffect(() => { setCanPortal(true); }, []);

  const updateDropdownRect = () => {
    if (inputRef.current) setDropdownRect(inputRef.current.getBoundingClientRect());
  };
  useEffect(() => {
    if (!shouldShowResults) { setDropdownRect(null); return; }
    updateDropdownRect();
    window.addEventListener("resize", updateDropdownRect);
    window.addEventListener("scroll", updateDropdownRect, true);
    return () => {
      window.removeEventListener("resize", updateDropdownRect);
      window.removeEventListener("scroll", updateDropdownRect, true);
    };
  }, [shouldShowResults, searchValue]);

  useEffect(() => {
    if (!focused || selectedProduct) return;
    const timer = window.setTimeout(() => {
      const t = searchValue.trim();
      updateParams({ restockProductSearch: t.length >= 2 ? t : "" });
    }, 350);
    return () => window.clearTimeout(timer);
  }, [focused, selectedProduct, searchValue]);

  // Auto-save when clicking outside the row
  useEffect(() => {
    const handleMouseDown = (e: MouseEvent) => {
      const target = e.target as HTMLElement;
      if (rowRef.current?.contains(target)) return;
      if (target.closest("[data-add-row-portal]")) return;
      const { selectedProduct: sp, qtys: q, status: st, priority: pr, notes: no, eta: et } = stateRef.current;
      const total = Object.values(q).reduce((s, v) => s + (Number(v) || 0), 0);
      if (sp && total > 0) {
        const fd = new FormData();
        fd.set("intent", "create_restock_order_from_portal");
        fd.set("product", JSON.stringify(sp));
        fd.set("qtys", JSON.stringify(q));
        fd.set("productType", initialProductGroup ?? "");
        fd.set("supplierStatus", st);
        fd.set("priority", pr);
        fd.set("notes", no);
        fd.set("eta", et);
        fetcher.submit(fd, { method: "post" });
        setSelectedProduct(null);
        setSearchValue("");
        setQtys({});
        setNotes(""); setEta(""); setPriority("");
        setStatus(restockSettings.statusOptions[0]?.value ?? "on_order");
        updateParams({ restockProductSearch: "" });
        onSaved?.();
      } else {
        updateParams({ restockProductSearch: "" });
      }
    };
    document.addEventListener("mousedown", handleMouseDown);
    return () => document.removeEventListener("mousedown", handleMouseDown);
  }, []);

  const selectProduct = (product: ShopifySearchProduct) => {
    setSelectedProduct(product);
    setSearchValue(product.title);
    setFocused(false);
    setDropdownRect(null);
    // qtys keyed by size-only variant titles (already filtered by searchShopifyProducts)
    setQtys(Object.fromEntries(product.variants.map((v) => [v.title, ""])));
  };

  const clearProduct = () => {
    setSelectedProduct(null);
    setSearchValue("");
    setQtys({});
    updateParams({ restockProductSearch: "" });
    window.setTimeout(() => inputRef.current?.focus(), 0);
  };

  return (
    <tr ref={rowRef} style={{ ...s.row, ...s.addOrderRow }}>
      <RowNumberCell rowNumber={rowIndex} actions={[
        { label: "Blank add row", onClick: () => inputRef.current?.focus() },
      ]} />
      <Td rowIndex={0} colIndex={0}>
        {initialProductGroup && <span style={{ fontSize: 11, color: "#6b7280", fontStyle: "italic" }}>{initialProductGroup}</span>}
      </Td>
      <Td rowIndex={0} colIndex={1} center>{null}</Td>
      <Td rowIndex={0} colIndex={2} center>
        <div style={s.imageCell}>
          {selectedProduct?.imageUrl ? <img src={selectedProduct.imageUrl} alt="" style={s.thumb} /> : <div style={s.noImg}>—</div>}
        </div>
      </Td>
      <Td rowIndex={0} colIndex={3} overflowVisible>
        <div style={s.productCellSearch}>
          {selectedProduct ? (
            <div style={s.selectedRestockProduct}>
              <span style={s.productName}>{selectedProduct.title}</span>
              <button type="button" style={s.clearProductButton} onClick={clearProduct}>×</button>
            </div>
          ) : (
            <input
              ref={inputRef}
              type="search"
              value={searchValue}
              onFocus={() => setFocused(true)}
              onChange={(e) => setSearchValue(e.currentTarget.value)}
              onBlur={() => window.setTimeout(() => setFocused(false), 140)}
              placeholder="Search product…"
              style={s.restockSearchInput}
            />
          )}
          {shouldShowResults && canPortal && dropdownRect && createPortal(
            <div data-add-row-portal style={{ ...s.productCellResults, top: dropdownRect.bottom + 8, left: dropdownRect.left, width: Math.max(dropdownRect.width, 460), height: dropdownHeight }}>
              {searchValue.trim() !== productSearch
                ? <div style={s.productCellResultEmpty}>Searching…</div>
                : productResults.length
                  ? productResults.map((p) => (
                    <button key={p.id} type="button" style={s.productCellResult}
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => selectProduct(p)}
                    >
                      {p.imageUrl ? <img src={p.imageUrl} alt="" style={s.productCellResultImage} /> : <span style={s.productCellNoImage}>—</span>}
                      <span style={s.productCellResultText}>
                        <strong>{p.title}</strong>
                        <span>{p.skus.slice(0, 3).join(", ") || "No SKU"}</span>
                      </span>
                    </button>
                  ))
                  : <div style={s.productCellResultEmpty}>No products found.</div>
              }
            </div>,
            document.body,
          )}
        </div>
      </Td>
      <Td rowIndex={0} colIndex={4}>
        <span style={s.sku}>{selectedProduct?.skus?.join("\n") || "—"}</span>
      </Td>
      {sizes.map((size, sizeIndex) => (
        <Td key={size} rowIndex={0} colIndex={5 + sizeIndex} center>
          {selectedProduct && size in qtys ? (
            <input
              type="text"
              inputMode="numeric"
              pattern="[0-9]*"
              value={qtys[size] ?? ""}
              onChange={(e) => { const v = e.currentTarget.value.replace(/\D/g, ""); setQtys((c) => ({ ...c, [size]: v })); }}
              style={{ ...s.qtyInput, ...(Number(qtys[size]) > 0 ? s.qtyInputActive : s.qtyInputZero) }}
            />
          ) : (
            <span style={{ color: "#d1d5db" }}>—</span>
          )}
        </Td>
      ))}
      <Td rowIndex={0} colIndex={totalCol} center><span style={s.total}>{totalQty || 0}</span></Td>
      <Td rowIndex={0} colIndex={statusCol}>
        <select value={status} onChange={(e) => setStatus(e.currentTarget.value)} style={{ ...s.select, background: restockSettings.statusOptions.find((o) => o.value === status)?.bg ?? "#f3f4f6", color: restockSettings.statusOptions.find((o) => o.value === status)?.color ?? "#374151" }}>
          {restockSettings.statusOptions.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
        </select>
      </Td>
      <Td rowIndex={0} colIndex={notesCol} overflowVisible>
        <textarea value={notes} onChange={(e) => setNotes(e.currentTarget.value)} rows={2} placeholder="Notes" style={s.textarea} />
      </Td>
      <Td rowIndex={0} colIndex={priorityCol}>
        <select value={priority} onChange={(e) => setPriority(e.currentTarget.value)} style={s.select}>
          <option value="">— Priority —</option>
          {restockSettings.priorityOptions.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
        </select>
      </Td>
      <Td rowIndex={0} colIndex={etaCol}>
        <input value={eta} onChange={(e) => setEta(e.currentTarget.value)} style={s.dateInput} placeholder="dd/mm/yy" />
      </Td>
    </tr>
  );
}

function OrderRow({
  order,
  rowIndex,
  sizes,
  users,
  restockSettings,
  customColumns,
  customCells,
  rowHeights,
  frozenOffsets,
  fabricStockIndex,
}: {
  order: Order;
  rowIndex: number;
  sizes: string[];
  users: PortalUser[];
  restockSettings: RestockSettings;
  customColumns: TableCustomColumn[];
  customCells: Record<string, string>;
  rowHeights: Record<string, number>;
  frozenOffsets?: number[];
  fabricStockIndex: FabricStockEntry[];
}) {
  const fetcher = useFetcher();
  const [inventoryOpen, setInventoryOpen] = useState(false);
  const [deleteConfirmOpen, setDeleteConfirmOpen] = useState(false);
  const qtyBySize = order.lines.reduce<Record<string, number>>((acc, line) => {
    acc[line.variantTitle] = (acc[line.variantTitle] ?? 0) + line.qtyOrdered;
    return acc;
  }, {});
  const inventoryBySize = order.lines.reduce<Record<string, number | null>>((acc, line) => {
    const availableInventory = "availableInventory" in line ? line.availableInventory : null;
    if (Number.isFinite(Number(availableInventory))) {
      acc[line.variantTitle] = Number(availableInventory);
    } else if (!(line.variantTitle in acc)) {
      acc[line.variantTitle] = null;
    }
    return acc;
  }, {});
  const allSkus = order.lines.map((l) => l.sku).filter(Boolean).join("\n");
  const etaValue = formatPortalDate(order.eta);
  const orderDate = formatPortalDate(order.createdAt);
  const inventoryTotal = sizes.reduce((sum, size) => sum + (inventoryBySize[size] ?? 0), 0);
  const totalCol = 5 + sizes.length;
  const statusCol = totalCol + 1;
  const notesCol = totalCol + 2;
  const priorityCol = totalCol + 3;
  const etaCol = totalCol + 4;
  const fabricStockCol = totalCol + 5;
  const fabricMatches = findFabricStockMatches(order.productTitle, fabricStockIndex);
  const rowHeightKey = `restock:${order.id}`;
  const shouldSkipDeleteConfirm = () => {
    if (typeof window === "undefined") return false;
    const skipUntil = Number(window.localStorage.getItem(DELETE_CONFIRM_SKIP_KEY) ?? 0);
    return skipUntil > Date.now();
  };
  const deleteOrder = () => submitPortalCell(fetcher, { intent: "delete_order", orderId: order.id });
  const requestDeleteOrder = () => {
    if (shouldSkipDeleteConfirm()) {
      deleteOrder();
      return;
    }
    setDeleteConfirmOpen(true);
  };
  const skipDeleteConfirmForDay = () => {
    window.localStorage.setItem(DELETE_CONFIRM_SKIP_KEY, String(Date.now() + 24 * 60 * 60 * 1000));
  };

  return (
    <>
      <tr id={`order-${order.id}`} style={{ ...s.row, ...(rowHeights[rowHeightKey] ? { height: rowHeights[rowHeightKey] } : {}) }}>
        <RowNumberCell rowNumber={rowIndex} actions={[
          { label: "Duplicate row", onClick: () => submitPortalCell(fetcher, { intent: "duplicate_order", orderId: order.id }) },
          { label: "Delete row", danger: true, onClick: requestDeleteOrder },
        ]} heightKey={rowHeightKey} />
        {/* Factory notes */}
        <Td rowIndex={rowIndex} colIndex={0} overflowVisible historyEntity="Restock Order" historyEntityId={String(order.id)} historyField="Factory notes" historyEntityName={order.productTitle} stickyLeft={frozenOffsets?.[0]}><NotesCell orderId={order.id} field="factory_notes" value={order.factoryNotes ?? ""} users={users} /></Td>

        {/* Order date */}
        <Td rowIndex={rowIndex} colIndex={1} center stickyLeft={frozenOffsets?.[1]}><span style={s.dateText}>{orderDate}</span></Td>

        {/* Picture */}
        <Td rowIndex={rowIndex} colIndex={2} center historyEntity="Restock Order" historyEntityId={String(order.id)} historyField="Product image" historyEntityName={order.productTitle} stickyLeft={frozenOffsets?.[2]}>
          <div style={s.imageCell}>
            {order.productImageUrl
              ? <img src={order.productImageUrl} alt="" style={s.thumb} />
              : <div style={s.noImg}>—</div>}
          </div>
        </Td>

        {/* Name */}
        <Td rowIndex={rowIndex} colIndex={3} historyEntity="Restock Order" historyEntityId={String(order.id)} historyField="Product name" historyEntityName={order.productTitle} stickyLeft={frozenOffsets?.[3]} isLastFrozen><span style={s.productName}>{order.productTitle}</span></Td>

        {/* SKU */}
        <Td rowIndex={rowIndex} colIndex={4} overflowVisible historyEntity="Restock Order" historyEntityId={String(order.id)} historyField="SKU" historyEntityName={order.productTitle}>
          <div style={s.skuCellWithToggle}>
            <span style={s.sku}>{allSkus || "—"}</span>
            <button
              type="button"
              onClick={() => setInventoryOpen((current) => !current)}
              style={{ ...s.inventoryToggle, color: restockSettings.inventoryArrowColor }}
              aria-label={inventoryOpen ? "Hide Shopify inventory" : "Show Shopify inventory"}
              title={inventoryOpen ? "Hide Shopify inventory" : "Show Shopify inventory"}
            >
              {inventoryOpen ? "▲" : "▼"}
            </button>
          </div>
        </Td>

        {/* Size columns */}
        {sizes.map((sz, sizeIndex) => (
          <Td key={sz} rowIndex={rowIndex} colIndex={5 + sizeIndex} center historyEntity="Restock Order" historyEntityId={String(order.id)} historyField={`Qty (${sz})`} historyEntityName={order.productTitle}>
            <QtyCell orderId={order.id} size={sz} value={qtyBySize[sz] ?? 0} restockSettings={restockSettings} />
          </Td>
        ))}

        {/* Total */}
        <Td rowIndex={rowIndex} colIndex={totalCol} center><span style={s.total}>{order.totalQty}</span></Td>

        {/* Status */}
        <Td rowIndex={rowIndex} colIndex={statusCol} historyEntity="Restock Order" historyEntityId={String(order.id)} historyField="Status" historyEntityName={order.productTitle}><StatusCell orderId={order.id} value={order.supplierStatus} restockSettings={restockSettings} /></Td>

        {/* Notes (from order) */}
        <Td rowIndex={rowIndex} colIndex={notesCol} overflowVisible historyEntity="Restock Order" historyEntityId={String(order.id)} historyField="Notes" historyEntityName={order.productTitle}><NotesCell orderId={order.id} field="notes" value={order.notes ?? ""} users={users} /></Td>

        {/* Priority */}
        <Td rowIndex={rowIndex} colIndex={priorityCol} historyEntity="Restock Order" historyEntityId={String(order.id)} historyField="Priority" historyEntityName={order.productTitle}><PriorityCell orderId={order.id} value={order.priority ?? ""} restockSettings={restockSettings} /></Td>

        {/* ETA */}
        <Td rowIndex={rowIndex} colIndex={etaCol} historyEntity="Restock Order" historyEntityId={String(order.id)} historyField="ETA" historyEntityName={order.productTitle}><EtaCell orderId={order.id} value={etaValue} /></Td>

        {/* Fabric in stock — looked up from the fabric name in the product title */}
        <Td rowIndex={rowIndex} colIndex={fabricStockCol} center>
          {fabricMatches.length === 0 ? (
            <span style={{ color: "#111827", fontWeight: 600, fontSize: 12 }}>Fabric not found</span>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 2, alignItems: "center" }}>
              {fabricMatches.filter((m) => m.kind === "stock").map((match, idx) => (
                <span key={`s-${idx}`} title={`${match.name} — ${match.sheetName}`} style={{ fontSize: 12, color: "#111827" }}>
                  {match.meters === 0 ? (
                    <span style={{ color: "#dc2626", fontWeight: 600 }}>Out of stock</span>
                  ) : (
                    <>
                      <span style={{ fontWeight: 600 }}>{Math.round(match.meters).toLocaleString()}</span>
                      <span style={{ marginLeft: 3, fontWeight: 500 }}>m</span>
                    </>
                  )}
                  <span style={{ marginLeft: 4, fontWeight: 500 }}>({match.sheetName})</span>
                </span>
              ))}
              {fabricMatches.filter((m) => m.kind === "order").map((match, idx) => (
                <span key={`o-${idx}`} title={`${match.name} — ${match.sheetName}`} style={{ fontSize: 12, color: "#0369a1" }}>
                  <span style={{ fontWeight: 600 }}>{Math.round(match.meters).toLocaleString()}</span>
                  <span style={{ marginLeft: 3, fontWeight: 500 }}>m on order</span>
                  <span style={{ marginLeft: 4, fontWeight: 500 }}>({match.sheetName})</span>
                </span>
              ))}
            </div>
          )}
        </Td>

        {customColumns.map((column, customIndex) => (
          <Td key={column.id} rowIndex={rowIndex} colIndex={fabricStockCol + 1 + customIndex}>
            <TableCustomCell cellKey={`restock:${order.id}:${column.id}`} value={customCells[`restock:${order.id}:${column.id}`] ?? ""} />
          </Td>
        ))}

      </tr>
      {deleteConfirmOpen && (
        <tr>
          <td colSpan={fabricStockCol + customColumns.length + 2}>
            <div style={s.deleteConfirm} onClick={() => setDeleteConfirmOpen(false)}>
              <div style={s.deleteConfirmCard} onClick={(event) => event.stopPropagation()}>
                <div style={s.deleteConfirmTitle}>Delete this restock row?</div>
                <div style={s.deleteConfirmText}>{order.productTitle}</div>
                <div style={s.deleteConfirmActions}>
                  <button type="button" style={s.deleteConfirmButton} onClick={() => setDeleteConfirmOpen(false)}>
                    Cancel
                  </button>
                  <button
                    type="button"
                    style={{ ...s.deleteConfirmButton, ...s.deleteConfirmDanger }}
                    onClick={() => {
                      deleteOrder();
                      setDeleteConfirmOpen(false);
                    }}
                  >
                    OK
                  </button>
                  <button
                    type="button"
                    style={s.deleteConfirmButton}
                    onClick={() => {
                      skipDeleteConfirmForDay();
                      deleteOrder();
                      setDeleteConfirmOpen(false);
                    }}
                  >
                    Don’t show for 24 hours
                  </button>
                </div>
              </div>
            </div>
          </td>
        </tr>
      )}
      {inventoryOpen && (
        <tr style={s.inventoryRow}>
          <td style={s.inventoryBlankCell} />
          <td style={s.inventoryBlankCell} />
          <td style={s.inventoryBlankCell} />
          <td style={s.inventoryBlankCell} />
          <td style={s.inventoryBlankCell} />
          <td style={s.inventoryLabelCell}>Shopify</td>
          {sizes.map((size) => (
            <td key={size} style={{ ...s.td, ...s.inventoryQtyCell }}>
              {inventoryBySize[size] == null ? "—" : inventoryBySize[size]}
            </td>
          ))}
          <td style={{ ...s.td, ...s.inventoryQtyCell }}><span style={s.total}>{inventoryTotal}</span></td>
          <td style={{ ...s.td, ...s.inventoryStatusCell }}>Available</td>
          <td style={s.inventoryBlankCell} />
          <td style={s.inventoryBlankCell} />
          <td style={s.inventoryBlankCell} />
        </tr>
      )}
    </>
  );
}

// ─── Editable cells ───────────────────────────────────────────────────────────

const PORTAL_UNDO_STACK_KEY = "production-portal-undo-stack-v1";
const MAX_PORTAL_UNDO_ENTRIES = 80;

function readPortalUndoStack(): PortalUndoEntry[] {
  if (typeof window === "undefined") return [];
  try {
    const parsed = JSON.parse(window.localStorage.getItem(PORTAL_UNDO_STACK_KEY) ?? "[]");
    if (!Array.isArray(parsed)) return [];
    return parsed.filter((entry): entry is PortalUndoEntry => (
      entry
      && typeof entry === "object"
      && typeof (entry as PortalUndoEntry).label === "string"
      && (entry as PortalUndoEntry).fields
      && typeof (entry as PortalUndoEntry).fields === "object"
    ));
  } catch {
    return [];
  }
}

function pushPortalUndo(entry?: PortalUndoEntry | null) {
  if (!entry || typeof window === "undefined") return;
  const stack = readPortalUndoStack();
  stack.push(entry);
  window.localStorage.setItem(PORTAL_UNDO_STACK_KEY, JSON.stringify(stack.slice(-MAX_PORTAL_UNDO_ENTRIES)));
}

function submitLastPortalUndo(fetcher: ReturnType<typeof useFetcher>) {
  if (typeof window === "undefined") return false;
  const stack = readPortalUndoStack();
  const entry = stack.pop();
  if (!entry) return false;
  window.localStorage.setItem(PORTAL_UNDO_STACK_KEY, JSON.stringify(stack));
  submitPortalCell(fetcher, entry.fields, null);
  return true;
}

function submitPortalCell(
  fetcher: ReturnType<typeof useFetcher>,
  fields: Record<string, string | number>,
  undo?: PortalUndoEntry | null,
) {
  pushPortalUndo(undo);
  const formData = new FormData();
  for (const [key, value] of Object.entries(fields)) {
    formData.set(key, String(value));
  }
  fetcher.submit(formData, { method: "post" });
}

function TableCustomCell({ cellKey, value }: { cellKey: string; value: string }) {
  const fetcher = useFetcher();
  const [draft, setDraft] = useState(value);
  useEffect(() => setDraft(value), [value]);
  const save = (nextValue: string) => {
    if (nextValue === value) return;
    submitPortalCell(
      fetcher,
      { intent: "update_table_custom_cell", key: cellKey, value: nextValue },
      { label: "Undo custom cell", fields: { intent: "update_table_custom_cell", key: cellKey, value } },
    );
  };
  return (
    <textarea
      value={draft}
      onChange={(event) => setDraft(event.currentTarget.value)}
      onBlur={(event) => save(event.currentTarget.value)}
      rows={3}
      style={s.customCellTextarea}
      placeholder="Add..."
    />
  );
}

function RestockOptionChipDropdown({
  orderId,
  value,
  options,
  optionKind,
  restockSettings,
  updateIntent,
  undoLabel,
  emptyLabel,
}: {
  orderId: number;
  value: string;
  options: RestockOption[];
  optionKind: "statusOptions" | "priorityOptions";
  restockSettings: RestockSettings;
  updateIntent: "update_status" | "update_priority";
  undoLabel: string;
  emptyLabel?: string;
}) {
  const cellFetcher = useFetcher();
  const settingsFetcher = useFetcher();
  const buttonRef = useRef<HTMLButtonElement | null>(null);
  const [open, setOpen] = useState(false);
  const [rect, setRect] = useState<DOMRect | null>(null);
  const [editingChip, setEditingChip] = useState<RestockOption | null>(null);
  const [addingChip, setAddingChip] = useState(false);
  const [editLabel, setEditLabel] = useState("");
  const [editBg, setEditBg] = useState("#f3f4f6");
  const [editColor, setEditColor] = useState("#374151");
  const current = cellFetcher.formData ? String(cellFetcher.formData.get("value")) : value;
  const option = options.find((item) => item.value === current);

  const updateRect = () => {
    if (buttonRef.current) setRect(buttonRef.current.getBoundingClientRect());
  };
  const startEdit = (item?: RestockOption) => {
    setEditingChip(item ?? null);
    setAddingChip(!item);
    setEditLabel(item?.label ?? "");
    setEditBg(item?.bg ?? "#f3f4f6");
    setEditColor(item?.color ?? "#374151");
  };
  const stopEdit = () => {
    setEditingChip(null);
    setAddingChip(false);
    setEditLabel("");
  };
  useEffect(() => {
    if (!open) return;
    updateRect();
    const close = (event: MouseEvent) => {
      const target = event.target as Node;
      if (buttonRef.current?.contains(target)) return;
      const menu = document.querySelector(`[data-restock-chip-menu="${orderId}-${optionKind}"]`);
      if (menu?.contains(target)) return;
      setOpen(false);
    };
    document.addEventListener("mousedown", close);
    window.addEventListener("resize", updateRect);
    window.addEventListener("scroll", updateRect, true);
    return () => {
      document.removeEventListener("mousedown", close);
      window.removeEventListener("resize", updateRect);
      window.removeEventListener("scroll", updateRect, true);
    };
  }, [open, optionKind, orderId]);

  const selectValue = (nextValue: string) => {
    submitPortalCell(
      cellFetcher,
      { intent: updateIntent, orderId, value: nextValue },
      { label: undoLabel, fields: { intent: updateIntent, orderId, value } },
    );
  };

  const saveSettingsOptions = (nextOptions: RestockOption[]) => {
    submitPortalCell(
      settingsFetcher,
      {
        intent: "update_restock_settings",
        value: JSON.stringify({ ...restockSettings, [optionKind]: nextOptions }),
      },
      { label: "Undo restock chip", fields: { intent: "update_restock_settings", value: JSON.stringify(restockSettings) } },
    );
  };

  const saveEdit = () => {
    const nextLabel = editLabel.trim();
    if (!nextLabel) return;
    if (editingChip) {
      const nextOptions = options.map((item) => (
        item.value === editingChip.value ? { ...item, label: nextLabel, bg: editBg, color: editColor } : item
      ));
      saveSettingsOptions(nextOptions);
    } else {
      const valueSeed = slugForOption(nextLabel) || `chip_${Date.now()}`;
      const taken = new Set(options.map((item) => item.value));
      const nextValue = taken.has(valueSeed) ? `${valueSeed}_${Date.now()}` : valueSeed;
      saveSettingsOptions([...options, { value: nextValue, label: nextLabel, bg: editBg, color: editColor }]);
      selectValue(nextValue);
      setOpen(false);
    }
    stopEdit();
  };

  const dropdown = open && rect && typeof document !== "undefined" ? createPortal(
    <div
      data-restock-chip-menu={`${orderId}-${optionKind}`}
      style={{
        ...s.fabricChipMenu,
        top: rect.bottom + 6,
        left: rect.left,
        minWidth: Math.max(rect.width, 250),
      }}
    >
      {emptyLabel && (
        <button
          type="button"
          style={s.fabricChipMenuOption}
          onClick={() => {
            selectValue("");
            setOpen(false);
          }}
        >
          <span style={s.fabricChipCheck}>{current ? "" : "✓"}</span>
          <span>{emptyLabel}</span>
        </button>
      )}
      {options.map((item) => {
        const selected = item.value === current;
        return (
          <div key={item.value} style={s.fabricChipMenuItem}>
            <button type="button" style={s.fabricChipEditButton} onClick={() => startEdit(item)}>Edit</button>
            <button
              type="button"
              style={s.fabricChipMenuOption}
              onClick={() => {
                selectValue(item.value);
                setOpen(false);
              }}
            >
              <span style={s.fabricChipCheck}>{selected ? "✓" : ""}</span>
              <span style={{ ...s.fabricChipMenuPill, background: item.bg, color: item.color }}>{item.label}</span>
            </button>
          </div>
        );
      })}
      {(editingChip || addingChip) && (
        <div style={s.fabricChipEditor}>
          <input
            value={editLabel}
            onChange={(event) => setEditLabel(event.currentTarget.value)}
            style={s.fabricChipEditInput}
            placeholder="Chip text"
            autoFocus
          />
          <label style={s.fabricChipMenuToolLabel}>
            Chip
            <input type="color" value={editBg} style={s.fabricChipColor} onChange={(event) => setEditBg(event.currentTarget.value)} />
          </label>
          <label style={s.fabricChipMenuToolLabel}>
            Text
            <input type="color" value={editColor} style={s.fabricChipColor} onChange={(event) => setEditColor(event.currentTarget.value)} />
          </label>
          <div style={s.fabricChipEditActions}>
            <button type="button" style={s.fabricMiniButton} onClick={stopEdit}>Cancel</button>
            <button type="button" style={s.fabricMiniButton} onClick={saveEdit}>Save</button>
          </div>
        </div>
      )}
      <div style={s.fabricChipMenuTools}>
        <button type="button" style={s.fabricMiniButton} onClick={() => startEdit()}>Add chip</button>
      </div>
    </div>,
    document.body,
  ) : null;

  return (
    <div style={s.restockChipCell}>
      <button
        ref={buttonRef}
        type="button"
        style={{
          ...s.fabricChipSelect,
          background: option?.bg ?? "#f3f4f6",
          color: option?.color ?? "#374151",
        }}
        onClick={() => {
          updateRect();
          setOpen((currentOpen) => !currentOpen);
        }}
      >
        <span style={s.fabricChipButtonText}>{option?.label ?? emptyLabel ?? "—"}</span>
        <span style={s.fabricChipChevron}>⌄</span>
      </button>
      {dropdown}
    </div>
  );
}

function StatusCell({ orderId, value, restockSettings }: { orderId: number; value: string; restockSettings: RestockSettings }) {
  return (
    <RestockOptionChipDropdown
      orderId={orderId}
      value={value}
      options={restockSettings.statusOptions}
      optionKind="statusOptions"
      restockSettings={restockSettings}
      updateIntent="update_status"
      undoLabel="Undo status"
    />
  );
}

function PriorityCell({ orderId, value, restockSettings }: { orderId: number; value: string; restockSettings: RestockSettings }) {
  return (
    <RestockOptionChipDropdown
      orderId={orderId}
      value={value}
      options={restockSettings.priorityOptions}
      optionKind="priorityOptions"
      restockSettings={restockSettings}
      updateIntent="update_priority"
      undoLabel="Undo priority"
      emptyLabel="— Priority —"
    />
  );
}

function currentTagQuery(value: string) {
  const match = value.match(/(^|\s)@([a-z0-9._-]*)$/i);
  return match ? match[2].toLowerCase() : null;
}

function insertStaffTag(value: string, name: string) {
  const tag = `@${name.trim().split(/\s+/)[0]}`;
  if (/(^|\s)@[a-z0-9._-]*$/i.test(value)) {
    return value.replace(/(^|\s)@[a-z0-9._-]*$/i, (match, prefix) => `${prefix}${tag} `);
  }
  return `${value}${value.endsWith(" ") || !value ? "" : " "}${tag} `;
}

function NotesCell({
  orderId,
  field,
  value,
  users,
}: {
  orderId: number;
  field: string;
  value: string;
  users: PortalUser[];
}) {
  const fetcher = useFetcher();
  const [text, setText] = useState(value);
  const [focused, setFocused] = useState(false);
  const tagQuery = currentTagQuery(text);
  const suggestions = tagQuery == null
    ? []
    : users
        .filter((user) => user.active)
        .filter((user) => user.name.toLowerCase().includes(tagQuery))
        .slice(0, 5);

  return (
    <div style={s.noteTagWrap}>
      <textarea
        value={text}
        onChange={(e) => setText(e.currentTarget.value)}
        onFocus={() => setFocused(true)}
        onBlur={(e) => {
          window.setTimeout(() => setFocused(false), 120);
          submitPortalCell(
            fetcher,
            {
              intent: `update_${field}`,
              orderId,
              value: e.currentTarget.value,
            },
            { label: "Undo note", fields: { intent: `update_${field}`, orderId, value } },
          );
        }}
        rows={3}
        style={s.restockNoteTextarea}
        placeholder="Add note... use @name"
      />
      {focused && suggestions.length > 0 && (
        <div style={s.tagSuggestions}>
          {suggestions.map((user) => (
            <button
              key={user.id}
              type="button"
              style={s.tagSuggestionButton}
              onMouseDown={(event) => {
                event.preventDefault();
                setText((current) => insertStaffTag(current, user.name));
              }}
            >
              @{user.name}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function EtaCell({ orderId, value }: { orderId: number; value: string }) {
  const fetcher = useFetcher();
  return (
    <input
      type="text"
      defaultValue={value}
      onBlur={(e) => submitPortalCell(
        fetcher,
        {
          intent: "update_eta",
          orderId,
          value: e.currentTarget.value,
        },
        { label: "Undo ETA", fields: { intent: "update_eta", orderId, value } },
      )}
      style={s.dateInput}
      placeholder="dd/mm/yy"
    />
  );
}

function QtyCell({ orderId, size, value, restockSettings }: { orderId: number; size: string; value: number; restockSettings: RestockSettings }) {
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
      onBlur={(e) => submitPortalCell(
        fetcher,
        {
          intent: "update_qty",
          orderId,
          size,
          value: e.currentTarget.value,
        },
        { label: "Undo quantity", fields: { intent: "update_qty", orderId, size, value } },
      )}
      style={{
        ...s.qtyInput,
        ...(numericCurrent > 0 ? s.qtyInputActive : s.qtyInputZero),
        ...(numericCurrent > 0 ? { color: restockSettings.quantityFontColor } : {}),
      }}
    />
  );
}

function OrderActionsCell({ orderId }: { orderId: number }) {
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
    <div style={s.rowActions}>
      <fetcher.Form method="post">
        <input type="hidden" name="intent" value="duplicate_order" />
        <input type="hidden" name="orderId" value={orderId} />
        <button type="submit" style={s.smallButton}>Duplicate</button>
      </fetcher.Form>
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
    </div>
  );
}

// ─── Table helpers ────────────────────────────────────────────────────────────

function Th({
  children,
  center,
  headerKey,
  columnId,
  onResizeStart,
  stickyLeft,
  isLastFrozen,
}: {
  children: React.ReactNode;
  center?: boolean;
  headerKey?: string;
  columnId?: string;
  onResizeStart: (event: React.MouseEvent<HTMLSpanElement>) => void;
  stickyLeft?: number;
  isLastFrozen?: boolean;
}) {
  const frozenStyle: React.CSSProperties = stickyLeft !== undefined ? {
    left: stickyLeft,
    zIndex: 55,
    ...(isLastFrozen ? { boxShadow: "4px 0 6px -2px rgba(0,0,0,0.1)" } : {}),
  } : {};
  return (
    <th
      style={{ ...s.th, textAlign: center ? "center" : "left", ...frozenStyle }}
    >
      {headerKey && typeof children === "string"
        ? <EditableHeaderLabel headerKey={headerKey} value={children} />
        : <span style={s.thContent}>{children}</span>}
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

function EditableHeaderLabel({ headerKey, value }: { headerKey: string; value: string }) {
  return <span style={s.thContent} data-header-key={headerKey}>{value}</span>;
}
function Td({
  children,
  center,
  rowIndex,
  colIndex,
  overflowVisible,
  historyEntity,
  historyEntityId,
  historyField,
  historyEntityName,
  stickyLeft,
  isLastFrozen,
}: {
  children: React.ReactNode;
  center?: boolean;
  rowIndex: number;
  colIndex: number;
  overflowVisible?: boolean;
  historyEntity?: string;
  historyEntityId?: string;
  historyField?: string;
  historyEntityName?: string;
  stickyLeft?: number;
  isLastFrozen?: boolean;
}) {
  const frozenStyle: React.CSSProperties = stickyLeft !== undefined ? {
    position: "sticky",
    left: stickyLeft,
    zIndex: 40,
    ...(isLastFrozen ? { boxShadow: "4px 0 6px -2px rgba(0,0,0,0.1)" } : {}),
  } : {};
  return (
    <td
      data-grid-row={rowIndex}
      data-grid-col={colIndex}
      data-history-entity={historyEntity}
      data-history-entity-id={historyEntityId}
      data-history-field={historyField}
      data-history-entity-name={historyEntityName}
      tabIndex={0}
      onFocus={(event) => {
        if (event.target !== event.currentTarget) return;
        const focusTarget = event.currentTarget.querySelector<HTMLElement>(FOCUSABLE_CELL_SELECTOR);
        if (!focusTarget) return;
        window.setTimeout(() => {
          focusTarget.focus();
          if (focusTarget instanceof HTMLInputElement || focusTarget instanceof HTMLTextAreaElement) {
            focusTarget.select();
          }
        }, 0);
      }}
      style={{ ...s.td, textAlign: center ? "center" : "left", ...(overflowVisible ? { overflow: "visible" } : {}), ...frozenStyle }}
    >
      {children}
    </td>
  );
}

function PackingTd({
  children,
  center,
  rowIndex,
  colIndex,
  overflowVisible,
  style,
  onContextMenu,
  onClick,
  stickyLeft,
  isLastFrozen,
}: {
  children: React.ReactNode;
  center?: boolean;
  rowIndex: number;
  colIndex: number;
  overflowVisible?: boolean;
  style?: React.CSSProperties;
  onContextMenu?: (e: React.MouseEvent<HTMLTableCellElement>) => void;
  onClick?: (e: React.MouseEvent<HTMLTableCellElement>) => void;
  stickyLeft?: number;
  isLastFrozen?: boolean;
}) {
  const frozenStyle: React.CSSProperties = stickyLeft !== undefined ? {
    position: "sticky",
    left: stickyLeft,
    zIndex: 40,
    ...(isLastFrozen ? { boxShadow: "4px 0 6px -2px rgba(0,0,0,0.1)" } : {}),
  } : {};
  return (
    <td
      data-grid-row={rowIndex}
      data-grid-col={colIndex}
      tabIndex={0}
      onContextMenu={onContextMenu}
      onClick={onClick}
      onFocus={(event) => {
        if (event.target !== event.currentTarget) return;
        const focusTarget = event.currentTarget.querySelector<HTMLElement>(FOCUSABLE_CELL_SELECTOR);
        if (!focusTarget) return;
        window.setTimeout(() => {
          focusTarget.focus();
          if (focusTarget instanceof HTMLInputElement || focusTarget instanceof HTMLTextAreaElement) {
            focusTarget.select();
          }
        }, 0);
      }}
      style={{
        ...s.td,
        textAlign: center ? "center" : "left",
        ...(overflowVisible ? { overflow: "visible" } : {}),
        ...frozenStyle,
        ...style,
      }}
    >
      {children}
    </td>
  );
}

// ─── Styles ───────────────────────────────────────────────────────────────────

const s: Record<string, React.CSSProperties> = {
  appShell: {
    minHeight: "100vh",
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
    display: "flex",
    alignItems: "flex-start",
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
    position: "sticky",
    top: 0,
    height: "100vh",
    overflowY: "auto",
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
    color: "inherit",
    textDecoration: "none",
    borderRadius: 8,
    padding: "10px 12px",
    fontSize: 13,
    fontWeight: 700,
  },
  navSubItem: {
    display: "block",
    color: "inherit",
    textDecoration: "none",
    borderRadius: 8,
    padding: "8px 12px 8px 24px",
    fontSize: 12,
    fontWeight: 700,
  },
  navItemActive: { background: "rgba(255,255,255,0.15)", color: "#fff" },
  iconNavItem: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: 40,
    height: 40,
    color: "inherit",
    textDecoration: "none",
    borderRadius: 10,
    fontSize: 13,
    fontWeight: 800,
    border: "1px solid rgba(148,163,184,0.18)",
    background: "rgba(15,23,42,0.35)",
  },
  iconNavItemActive: { background: "rgba(255,255,255,0.2)", color: "#fff", borderColor: "rgba(255,255,255,0.4)" },
  settingsLink: { marginTop: "auto" },
  count: { fontSize: 13, color: "#6b7280" },
  main: { flex: 1, minWidth: 0, padding: "24px 16px", minHeight: "100vh" },
  pageHeader: {
    display: "flex",
    alignItems: "flex-start",
    justifyContent: "space-between",
    gap: 16,
    marginBottom: 10,
  },
  pageTitle: { margin: 0, fontSize: "var(--portal-heading-font-size)", color: "var(--portal-heading-text-color)", lineHeight: 1.2 },
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
  addProductButton: {
    border: "1px solid #008060",
    borderRadius: 6,
    padding: "7px 14px",
    background: "#008060",
    color: "#fff",
    fontSize: 13,
    fontWeight: 700,
    cursor: "pointer",
    whiteSpace: "nowrap" as const,
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
    background: "var(--portal-primary-button-bg)",
    color: "var(--portal-primary-button-color)",
    fontSize: 12,
    fontWeight: 800,
    letterSpacing: "0.04em",
  },
  activeUserEmpty: { color: "#6b7280", fontSize: 12, fontWeight: 700 },
  activeUserTooltip: {
    position: "fixed",
    transform: "translateX(-50%)",
    zIndex: 1300,
    padding: "8px 14px",
    background: "#111827",
    color: "#fff",
    fontSize: 15,
    fontWeight: 700,
    borderRadius: 8,
    boxShadow: "0 18px 42px rgba(15,23,42,0.24)",
    pointerEvents: "none",
    whiteSpace: "nowrap",
  },
  messagesWrap: { position: "relative" },
  messagesButton: {
    position: "relative",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    border: "1px solid #d1d5db",
    borderRadius: 8,
    padding: 7,
    background: "#fff",
    color: "#374151",
    width: 34,
    height: 34,
    cursor: "pointer",
  },
  messagesButtonActive: {
    position: "relative",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    border: "1px solid #f97316",
    borderRadius: 8,
    padding: 7,
    background: "#fff7ed",
    color: "#9a3412",
    width: 34,
    height: 34,
    cursor: "pointer",
  },
  messageCount: {
    position: "absolute",
    top: -6,
    right: -6,
    minWidth: 17,
    height: 17,
    padding: "0 4px",
    borderRadius: 999,
    background: "#ef4444",
    color: "#fff",
    fontSize: 10,
    fontWeight: 900,
    lineHeight: "17px",
    textAlign: "center",
    boxShadow: "0 0 0 2px #fff",
  },
  messagesPopover: {
    position: "absolute",
    top: 38,
    right: 0,
    width: 360,
    maxHeight: 420,
    overflow: "auto",
    background: "#fff",
    border: "1px solid #cbd5e1",
    borderRadius: 12,
    boxShadow: "0 18px 40px rgba(15,23,42,0.24)",
    zIndex: 200,
    padding: 8,
  },
  messagesHeader: { padding: "8px 10px", fontSize: 13, fontWeight: 900, color: "#111827" },
  messageItem: {
    display: "grid",
    gridTemplateColumns: "1fr auto",
    gap: 8,
    alignItems: "start",
    padding: 8,
    borderTop: "1px solid #e5e7eb",
  },
  messageLink: {
    display: "grid",
    gap: 3,
    color: "#111827",
    textDecoration: "none",
    fontSize: 12,
    fontWeight: 700,
  },
  messageBody: {
    color: "#6b7280",
    fontWeight: 600,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  messageReadButton: {
    border: "1px solid #d1d5db",
    borderRadius: 6,
    padding: "5px 7px",
    background: "#fff",
    color: "#374151",
    fontSize: 11,
    fontWeight: 800,
    cursor: "pointer",
  },
  messageEmpty: { padding: 14, color: "#6b7280", fontSize: 12, fontWeight: 700 },
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
  loginLabel: {
    display: "flex",
    flexDirection: "column" as const,
    gap: 5,
    fontSize: 13,
    fontWeight: 600,
    color: "#374151",
  },
  loginInput: {
    border: "1px solid #b6c0cc",
    borderRadius: 8,
    padding: "10px 12px",
    fontSize: 15,
    fontFamily: "inherit",
    background: "#fff",
    outline: "none",
  },
  loginButton: {
    border: "1px solid var(--portal-primary-button-bg)",
    borderRadius: 8,
    padding: "10px 14px",
    background: "var(--portal-primary-button-bg)",
    color: "var(--portal-primary-button-color)",
    fontSize: 14,
    fontWeight: 800,
    cursor: "pointer",
  },
  settingsPanel: { display: "grid", gap: 16, maxWidth: 1140 },
  settingsCard: {
    background: "#fff",
    border: "1px solid #cbd5e1",
    borderRadius: 14,
    padding: 18,
    boxShadow: "0 1px 2px rgba(15,23,42,0.08)",
  },
  settingsHeader: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 12 },
  settingsSaveGroup: { display: "flex", alignItems: "center", gap: 10 },
  settingsSavedText: { color: "#166534", fontSize: 13, fontWeight: 900, whiteSpace: "nowrap" },
  settingsTitle: { margin: 0, fontSize: 18, color: "var(--portal-heading-text-color)" },
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
  settingsSubCard: {
    display: "grid",
    gap: 10,
    padding: 12,
    border: "1px solid #e5e7eb",
    borderRadius: 10,
    background: "#f8fafc",
    marginTop: 14,
  },
  settingsSubTitle: { margin: 0, fontSize: 14, color: "#111827" },
  settingsInlineFields: { display: "flex", alignItems: "end", flexWrap: "wrap", gap: 12 },
  settingsFieldLabel: { display: "grid", gap: 5, fontSize: 12, fontWeight: 800, color: "#4b5563" },
  settingsSmallInput: {
    border: "1px solid #cbd5e1",
    borderRadius: 7,
    padding: "8px 9px",
    width: 90,
    fontSize: 13,
    fontWeight: 700,
  },
  qtyPreview: {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    minWidth: 50,
    minHeight: 34,
    borderRadius: 8,
    border: "1px solid #cbd5e1",
    background: "#fff",
    fontWeight: 900,
  },
  buttonPreview: {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    minWidth: 82,
    minHeight: 34,
    borderRadius: 8,
    fontSize: 13,
    fontWeight: 900,
    boxShadow: "inset 0 1px 0 rgba(255,255,255,0.18)",
  },
  headingPreview: {
    display: "inline-flex",
    alignItems: "center",
    minHeight: 34,
    fontWeight: 900,
    lineHeight: 1.1,
  },
  inventoryArrowPreview: {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    width: 34,
    height: 34,
    borderRadius: 8,
    border: "1px solid #cbd5e1",
    background: "#fff",
    fontSize: 14,
    fontWeight: 900,
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
  busyButton: {
    opacity: 0.7,
    cursor: "wait",
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
  inventoryAccessCheckbox: {
    display: "flex",
    alignItems: "center",
    gap: 6,
    color: "#374151",
    fontSize: 12,
    fontWeight: 800,
    whiteSpace: "nowrap",
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
    alignItems: "center",
    flexWrap: "wrap",
    justifyContent: "flex-end",
    gap: 10,
    background: "#fff",
    border: "1px solid #cbd5e1",
    borderRadius: 12,
    padding: 12,
  },
  packingOverviewTableWrap: {
    overflow: "auto",
    background: "#fff",
    border: "1px solid #cbd5e1",
    boxShadow: "0 1px 2px rgba(15,23,42,0.08)",
  },
  packingOverviewBar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 10,
    padding: 12,
    color: "#111827",
    borderBottom: "1px solid #cbd5e1",
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
  packingSearchBar: {
    display: "flex",
    alignItems: "flex-end",
    flexWrap: "wrap",
    gap: 10,
    background: "#fff",
    border: "1px solid #cbd5e1",
    borderRadius: 12,
    padding: "10px 12px",
  },
  packingSearchInput: { minWidth: 280 },
  searchCount: {
    alignSelf: "center",
    color: "#6b7280",
    fontSize: 13,
    fontWeight: 800,
  },
  packingCreateForm: { display: "flex", alignItems: "center", justifyContent: "flex-end" },
  packingImportForm: { display: "flex", alignItems: "center", justifyContent: "flex-end", gap: 8, flexWrap: "wrap" },
  fileInput: {
    border: "1px solid #b6c0cc",
    borderRadius: 8,
    padding: "7px 9px",
    background: "#fff",
    color: "#374151",
    fontSize: 13,
    fontWeight: 700,
  },
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
    display: "grid",
    gap: 10,
    background: "#fff",
    border: "1px solid #cbd5e1",
    borderRadius: 12,
    padding: "12px 14px",
  },
  packingTopRow: {
    display: "flex",
    alignItems: "center",
    flexWrap: "wrap",
    gap: 10,
  },
  packingBottomRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    flexWrap: "wrap",
    gap: 10,
  },
  packingTopLeft: {
    display: "flex",
    alignItems: "center",
    flexWrap: "wrap",
    gap: 12,
    minWidth: 0,
  },
  packingMeta: { display: "flex", alignItems: "center", flexWrap: "wrap", gap: 10 },
  packingToolbarLabel: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    color: "#374151",
    fontSize: 13,
    fontWeight: 800,
    whiteSpace: "nowrap",
  },
  packingBackButton: {},
  invoiceInput: { width: 240 },
  skipWordsInput: { width: 250 },
  packingTotalPill: {
    border: "1px solid #d1d5db",
    borderRadius: 8,
    padding: "8px 10px",
    background: "#fff",
    color: "#374151",
    fontSize: 13,
    fontWeight: 800,
    display: "inline-flex",
    alignItems: "center",
    gap: 5,
    whiteSpace: "nowrap",
  },
  packingActions: { display: "flex", gap: 8 },
  loadInventoryForm: {
    display: "flex",
    alignItems: "center",
    flexWrap: "wrap",
    gap: 10,
  },
  loadInventoryButton: {
    minHeight: 42,
    whiteSpace: "nowrap",
  },
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
    overflowX: "scroll",
    overflowY: "visible",
    scrollbarGutter: "stable",
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
  fabricThumb: { width: "100%", height: "100%", objectFit: "cover", borderRadius: 3 },
  fabricImageDrop: {
    position: "relative",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: 128,
    height: 172,
    minHeight: 172,
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
    zIndex: 1,
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
  linkedProductCell: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 10,
    width: "100%",
  },
  linkedProductTitle: {
    minWidth: 0,
    color: "#111827",
    fontSize: 13,
    fontWeight: 800,
    lineHeight: 1.25,
    wordBreak: "break-word" as const,
    whiteSpace: "normal" as const,
    overflowWrap: "anywhere" as const,
  },
  changeProductButton: {
    flex: "0 0 auto",
    border: "1px solid #d1d5db",
    borderRadius: 999,
    padding: "3px 7px",
    background: "#fff",
    color: "#4b5563",
    fontSize: 11,
    fontWeight: 800,
    cursor: "pointer",
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
  hideListButton: {
    border: "1px solid #fed7aa",
    borderRadius: 6,
    padding: "5px 7px",
    background: "#fff7ed",
    color: "#9a3412",
    fontSize: 11,
    fontWeight: 800,
    cursor: "pointer",
  },
  packingListActions: {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    gap: 8,
    flexWrap: "wrap",
  },
  clickableOverviewRow: {
    cursor: "pointer",
    transition: "box-shadow 120ms ease",
  },
  clickableOverviewCellHover: {
    background: "#eaf3ff",
    boxShadow: "inset 0 0 0 9999px rgba(37,99,235,0.05)",
  },
  empty: { background: "#fff", borderRadius: 12, padding: 40, textAlign: "center", color: "#6b7280" },
  primaryActionButton: {
    border: "1px solid var(--portal-primary-button-bg, #111827)",
    borderRadius: 8,
    padding: "9px 12px",
    background: "var(--portal-primary-button-bg, #111827)",
    color: "var(--portal-primary-button-color, #ffffff)",
    fontSize: 13,
    fontWeight: 900,
    cursor: "pointer",
  },
  productInfoPage: {
    display: "grid",
    gap: 14,
  },
  productInfoToolbar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 12,
    padding: "14px 16px",
    background: "#fff",
    border: "1px solid #dbe3ee",
    borderRadius: 10,
  },
  productInfoToolbarLeft: {
    display: "flex",
    alignItems: "center",
    gap: 12,
    flexWrap: "wrap",
  },
  productInfoActions: {
    display: "flex",
    alignItems: "center",
    justifyContent: "flex-end",
    gap: 8,
    flexWrap: "wrap",
  },
  productInfoSegmented: {
    display: "inline-flex",
    alignItems: "center",
    overflow: "hidden",
    border: "1px solid #cbd5e1",
    borderRadius: 8,
    background: "#fff",
  },
  productInfoSegmentButton: {
    minWidth: 38,
    border: 0,
    borderRight: "1px solid #cbd5e1",
    padding: "9px 11px",
    background: "#fff",
    color: "#475569",
    fontSize: 13,
    fontWeight: 900,
    cursor: "pointer",
  },
  productInfoSegmentButtonActive: {
    background: "var(--portal-primary-button-bg, #111827)",
    color: "var(--portal-primary-button-color, #ffffff)",
  },
  productInfoSelectLabel: {
    display: "grid",
    gap: 5,
    color: "#4b5563",
    fontSize: 12,
    fontWeight: 900,
  },
  productInfoSelect: {
    minWidth: 240,
    border: "1px solid #cbd5e1",
    borderRadius: 8,
    padding: "9px 10px",
    background: "#fff",
    color: "#111827",
    fontSize: 13,
    fontWeight: 800,
  },
  productInfoHeading: {
    margin: 0,
    color: "var(--portal-heading-text-color, #111827)",
    fontSize: 18,
    fontWeight: 900,
    lineHeight: 1.2,
  },
  productInfoMeta: {
    marginTop: 3,
    color: "#64748b",
    fontSize: 12,
    fontWeight: 800,
  },
  productInfoGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))",
    gap: 12,
  },
  productInfoList: {
    display: "grid",
    gap: 12,
    background: "#fff",
    border: "1px solid #dbe3ee",
    borderRadius: 10,
    padding: 10,
  },
  productStyleCard: {
    position: "relative",
    minWidth: 0,
    display: "grid",
    gridTemplateRows: "auto 1fr auto",
    overflow: "hidden",
    border: "1px solid #dbe3ee",
    borderRadius: 8,
    background: "#fff",
    boxShadow: "0 1px 2px rgba(15,23,42,0.06)",
  },
  productStyleCardDragging: {
    opacity: 0.56,
    boxShadow: "0 10px 24px rgba(15,23,42,0.16)",
  },
  productStyleCardDropTarget: {
    outline: "3px solid var(--portal-primary-button-bg, #111827)",
    outlineOffset: 2,
  },
  productStyleDragHandle: {
    position: "absolute",
    top: 8,
    left: 8,
    zIndex: 2,
    display: "grid",
    placeItems: "center",
    width: 28,
    height: 28,
    border: "1px solid rgba(15,23,42,0.22)",
    borderRadius: 7,
    background: "rgba(255,255,255,0.88)",
    color: "#475569",
    fontSize: 14,
    fontWeight: 900,
    letterSpacing: 1,
    boxShadow: "0 2px 8px rgba(15,23,42,0.16)",
    pointerEvents: "none",
  },
  productStyleImageWrap: {
    aspectRatio: "256 / 361",
    background: "#eef2f7",
  },
  productStyleImageButton: {
    width: "100%",
    height: "100%",
    display: "block",
    border: 0,
    padding: 0,
    background: "transparent",
    cursor: "pointer",
  },
  productStyleImage: {
    width: "100%",
    height: "100%",
    display: "block",
    objectFit: "cover",
  },
  productStyleImageEmpty: {
    width: "100%",
    height: "100%",
    display: "grid",
    placeItems: "center",
    color: "#64748b",
    fontSize: 13,
    fontWeight: 900,
  },
  productStyleCardBody: {
    display: "grid",
    gap: 4,
    padding: "12px 12px 4px",
    textAlign: "center",
    justifyItems: "center",
  },
  productStyleCardActions: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: 6,
    flexWrap: "wrap",
    padding: 12,
    textAlign: "center",
  },
  productStyleRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 12,
    minHeight: 58,
    padding: "10px 12px",
    border: "1px solid #e5e7eb",
    borderRadius: 8,
    background: "#f8fafc",
  },
  productInfoFooterActions: {
    display: "flex",
    justifyContent: "flex-end",
  },
  productInfoModalBackdrop: {
    position: "fixed",
    inset: 0,
    zIndex: 99998,
    display: "grid",
    placeItems: "center",
    padding: 20,
    background: "rgba(15,23,42,0.36)",
  },
  productInfoModal: {
    width: "min(440px, 100%)",
    display: "grid",
    gap: 12,
    border: "1px solid #cbd5e1",
    borderRadius: 12,
    padding: 18,
    background: "#fff",
    boxShadow: "0 24px 60px rgba(15,23,42,0.28)",
  },
  productInfoDetailsModal: {
    width: "min(760px, 100%)",
    maxHeight: "calc(100vh - 40px)",
    overflow: "auto",
    display: "grid",
    gap: 14,
    border: "1px solid #cbd5e1",
    borderRadius: 12,
    padding: 18,
    background: "#fff",
    boxShadow: "0 24px 60px rgba(15,23,42,0.28)",
  },
  productInfoDetailsHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 14,
  },
  productInfoDetailsThumb: {
    width: 72,
    aspectRatio: "256 / 361",
    objectFit: "cover",
    borderRadius: 7,
    border: "1px solid #e5e7eb",
    flex: "0 0 auto",
  },
  productInfoDetailsGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
    gap: 10,
  },
  productInfoDetailsField: {
    display: "grid",
    gap: 5,
    color: "#4b5563",
    fontSize: 12,
    fontWeight: 900,
  },
  productInfoDetailsInput: {
    width: "100%",
    border: "1px solid #cbd5e1",
    borderRadius: 8,
    padding: "9px 10px",
    color: "#111827",
    fontSize: 13,
    fontWeight: 800,
    boxSizing: "border-box",
  },
  productInfoDetailsTextarea: {
    width: "100%",
    border: "1px solid #cbd5e1",
    borderRadius: 8,
    padding: "9px 10px",
    color: "#111827",
    fontSize: 13,
    fontWeight: 800,
    boxSizing: "border-box",
    resize: "vertical",
    fontFamily: "inherit",
  },
  productInfoModalTitle: {
    margin: 0,
    color: "#111827",
    fontSize: 18,
    fontWeight: 900,
  },
  productInfoModalText: {
    margin: 0,
    color: "#4b5563",
    fontSize: 13,
    fontWeight: 700,
    lineHeight: 1.45,
  },
  productInfoModalActions: {
    display: "flex",
    justifyContent: "flex-end",
    gap: 8,
    flexWrap: "wrap",
  },
  productStyleRowMain: {
    minWidth: 0,
    display: "grid",
    gap: 3,
  },
  productCategoryTile: {
    position: "relative",
    minHeight: 132,
    border: "1px solid #dbe3ee",
    borderRadius: 10,
    background: "#fff",
    boxShadow: "0 1px 2px rgba(15,23,42,0.06)",
  },
  productTileMainButton: {
    width: "100%",
    height: "100%",
    minHeight: 132,
    display: "grid",
    alignContent: "start",
    gap: 10,
    border: 0,
    borderRadius: 10,
    padding: "18px 16px 48px",
    background: "transparent",
    textAlign: "left",
    cursor: "pointer",
  },
  productCategoryTitle: {
    color: "var(--portal-heading-text-color, #111827)",
    fontSize: 16,
    fontWeight: 900,
    lineHeight: 1.25,
    paddingRight: 48,
  },
  productCategoryMeta: {
    color: "#64748b",
    fontSize: 12,
    fontWeight: 800,
  },
  productStyleTile: {
    position: "relative",
    minHeight: 116,
    display: "grid",
    alignContent: "start",
    gap: 8,
    border: "1px solid #dbe3ee",
    borderRadius: 10,
    background: "#fff",
    padding: "18px 16px 48px",
    boxShadow: "0 1px 2px rgba(15,23,42,0.06)",
  },
  productStyleTitle: {
    color: "var(--portal-heading-text-color, #111827)",
    fontSize: 16,
    fontWeight: 900,
    lineHeight: 1.25,
  },
  productStyleMeta: {
    color: "#64748b",
    fontSize: 12,
    fontWeight: 800,
  },
  productTileRemove: {
    position: "absolute",
    right: 10,
    bottom: 10,
    border: "1px solid #fecaca",
    borderRadius: 7,
    padding: "5px 8px",
    background: "#fff5f5",
    color: "#991b1b",
    fontSize: 11,
    fontWeight: 900,
    cursor: "pointer",
  },
  productInfoEmpty: {
    gridColumn: "1 / -1",
    minHeight: 120,
    display: "grid",
    placeItems: "center",
    border: "1px dashed #cbd5e1",
    borderRadius: 10,
    background: "#fff",
    color: "#64748b",
    fontSize: 13,
    fontWeight: 800,
  },
  fabricPage: { display: "flex", flexDirection: "column", gap: 14 },
  fabricTotalsBar: {
    display: "flex",
    alignItems: "stretch",
    gap: 10,
    flexWrap: "wrap",
    padding: "12px",
    background: "#fff",
    border: "1px solid #dbe3ee",
    borderRadius: 10,
  },
  fabricTileActions: {
    display: "flex",
    justifyContent: "flex-end",
    marginBottom: 12,
  },
  fabricTotalsItem: {
    minWidth: 170,
    display: "grid",
    gap: 2,
    padding: "10px 12px",
    border: "1px solid #e5e7eb",
    borderRadius: 8,
    background: "#f8fafc",
    color: "var(--portal-table-text-color, #374151)",
    fontSize: "var(--portal-table-font-size, 13px)",
    fontWeight: 800,
  },
  fabricTotalsHelp: {
    position: "relative",
    width: 24,
    height: 24,
    borderRadius: "50%",
    border: "1px solid #cbd5e1",
    background: "#f8fafc",
    color: "#374151",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: 13,
    fontWeight: 900,
    cursor: "help",
    alignSelf: "center",
    flex: "0 0 auto",
  },
  fabricTotalsHelpBubble: {
    position: "absolute",
    top: 30,
    left: 0,
    zIndex: 50,
    width: 290,
    padding: "10px 12px",
    borderRadius: 8,
    border: "1px solid #cbd5e1",
    background: "#111827",
    color: "#fff",
    fontSize: 12,
    lineHeight: 1.35,
    fontWeight: 700,
    boxShadow: "0 12px 28px rgba(15,23,42,0.22)",
  },
  fabricIntro: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 12,
    padding: "16px 18px",
    background: "#fff",
    border: "1px solid #dbe3ee",
    borderRadius: 10,
  },
  fabricIntroTitle: { margin: 0, fontSize: 18, fontWeight: 800, color: "#111827" },
  fabricIntroText: { margin: "4px 0 0", color: "#6b7280", fontSize: 13, fontWeight: 600 },
  fabricGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))",
    gap: 12,
  },
  fabricCard: {
    position: "relative",
    minHeight: 112,
    border: "1px solid #dbe3ee",
    borderRadius: 10,
    background: "#fff",
    padding: 16,
    textAlign: "left",
    cursor: "pointer",
    boxShadow: "0 1px 2px rgba(15,23,42,0.06)",
  },
  fabricCardDragging: { opacity: 0.62, boxShadow: "0 10px 26px rgba(15,23,42,0.16)" },
  fabricCardHandle: {
    position: "absolute",
    top: 10,
    right: 12,
    color: "#94a3b8",
    fontSize: 10,
    fontWeight: 800,
    textTransform: "uppercase",
    letterSpacing: "0.04em",
  },
  fabricCardTitle: { display: "block", color: "var(--portal-heading-text-color, #111827)", fontSize: 15, fontWeight: 800, lineHeight: 1.25, paddingRight: 38 },
  fabricCardMeta: { display: "block", marginTop: 12, color: "#6b7280", fontSize: 12, fontWeight: 700 },
  fabricCardQuantity: { display: "block", marginTop: 4, color: "#008060", fontSize: 13, fontWeight: 800 },
  fabricCardCost: { display: "block", marginTop: 4, color: "var(--portal-table-text-color, #374151)", fontSize: 13, fontWeight: 800 },
  fabricToolbar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 12,
    padding: "10px 12px",
    background: "#fff",
    border: "1px solid #dbe3ee",
    borderRadius: 10,
  },
  fabricToolbarMeta: { display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", color: "#6b7280", fontSize: 13, fontWeight: 700 },
  fabricToolbarLeft: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    flexWrap: "wrap",
  },
  fabricSearchLabel: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    color: "#6b7280",
    fontSize: 12,
    fontWeight: 800,
  },
  fabricSearchInput: {
    width: 220,
    border: "1px solid #cbd5e1",
    borderRadius: 8,
    padding: "8px 10px",
    background: "#fff",
    color: "#111827",
    fontSize: 13,
    fontWeight: 700,
    outline: "none",
  },
  fabricTableShell: {
    display: "grid",
    gap: 10,
  },
  fabricTableWrap: {
    maxHeight: "calc(100vh - 170px)",
    overflow: "auto",
    background: "#fff",
    border: "1px solid #cbd5e1",
    boxShadow: "0 1px 2px rgba(15,23,42,0.08)",
  },
  fabricTable: { borderCollapse: "separate", borderSpacing: 0, minWidth: 960, fontSize: 13, tableLayout: "fixed" },
  fabricTh: {
    position: "sticky",
    top: 0,
    zIndex: 10,
    padding: "9px 10px",
    textAlign: "left",
    background: "#eef2f7",
    border: "1px solid #cbd5e1",
    color: "var(--portal-heading-text-color, #111827)",
    fontSize: 11,
    fontWeight: 800,
    textTransform: "uppercase",
    letterSpacing: "0.04em",
    whiteSpace: "nowrap",
  },
  fabricTd: {
    padding: 0,
    borderRight: "1px solid #e5e7eb",
    borderBottom: "1px solid #e5e7eb",
    verticalAlign: "top",
    color: "var(--portal-table-text-color, #1f2937)",
    fontWeight: 600,
    minWidth: 60,
    fontSize: "var(--portal-table-font-size, 13px)",
  },
  fabricCellInput: {
    width: "100%",
    minHeight: 96,
    border: "none",
    outline: "none",
    padding: "9px 10px",
    background: "transparent",
    color: "var(--portal-table-text-color, #1f2937)",
    font: "inherit",
    fontWeight: 700,
  },
  fabricCellTextarea: {
    width: "100%",
    minHeight: 96,
    border: "none",
    outline: "none",
    resize: "none",
    padding: "9px 10px",
    background: "transparent",
    color: "var(--portal-table-text-color, #1f2937)",
    font: "inherit",
    fontWeight: 700,
    lineHeight: 1.35,
  },
  fabricChipCell: {
    minHeight: 96,
    display: "grid",
    alignContent: "start",
    gap: 7,
    padding: 8,
  },
  restockChipCell: {
    minHeight: 44,
    display: "grid",
    alignContent: "center",
    padding: 8,
  },
  fabricChipColor: {
    width: 28,
    height: 28,
    padding: 0,
    border: "1px solid #cbd5e1",
    borderRadius: 6,
    background: "#fff",
    cursor: "pointer",
  },
  fabricNumberCellInput: {
    textAlign: "center",
  },
  fabricImageEditCell: {
    display: "grid",
    gap: 6,
    justifyItems: "center",
    padding: "8px 10px",
  },
  fabricTableFooter: {
    display: "flex",
    justifyContent: "flex-end",
    padding: "6px 0 4px",
  },
  fabricRowActions: {
    minWidth: 170,
    display: "grid",
    gap: 7,
    padding: 8,
  },
  fabricMoveSelect: {
    width: "100%",
    border: "1px solid #cbd5e1",
    borderRadius: 6,
    padding: "5px 7px",
    background: "#fff",
    color: "#374151",
    fontSize: 11,
    fontWeight: 800,
  },
  fabricChipSelect: {
    width: "100%",
    minHeight: 34,
    margin: 0,
    border: "1px solid rgba(15,23,42,0.12)",
    borderRadius: 8,
    padding: "6px 8px 6px 10px",
    outline: "none",
    fontSize: "var(--portal-table-font-size, 13px)",
    fontWeight: 900,
    cursor: "pointer",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 8,
    textAlign: "left",
  },
  fabricChipButtonText: {
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  fabricChipChevron: {
    flex: "0 0 auto",
    fontSize: 16,
    lineHeight: 1,
    fontWeight: 900,
    marginLeft: "auto",
    minWidth: 20,
    textAlign: "right",
  },
  fabricChipMenu: {
    position: "fixed",
    zIndex: 2147483647,
    maxHeight: 420,
    overflow: "auto",
    display: "grid",
    gap: 3,
    padding: 8,
    background: "rgba(255,255,255,0.96)",
    border: "1px solid #cbd5e1",
    borderRadius: 10,
    boxShadow: "0 18px 42px rgba(15,23,42,0.24)",
    backdropFilter: "blur(8px)",
  },
  fabricChipMenuOption: {
    width: "100%",
    display: "grid",
    gridTemplateColumns: "18px minmax(0, 1fr)",
    alignItems: "center",
    gap: 6,
    border: 0,
    borderRadius: 7,
    padding: "6px 8px",
    background: "transparent",
    color: "#111827",
    fontSize: 13,
    fontWeight: 800,
    textAlign: "left",
    cursor: "pointer",
  },
  fabricChipMenuItem: {
    display: "grid",
    gridTemplateColumns: "44px minmax(0, 1fr)",
    alignItems: "center",
    gap: 6,
  },
  fabricChipEditButton: {
    border: "1px solid #cbd5e1",
    borderRadius: 6,
    padding: "4px 6px",
    background: "#fff",
    color: "#475569",
    fontSize: 10,
    fontWeight: 900,
    cursor: "pointer",
    lineHeight: 1.1,
  },
  fabricChipCheck: {
    color: "#111827",
    fontSize: 13,
    fontWeight: 900,
    textAlign: "center",
  },
  fabricChipMenuPill: {
    minWidth: 0,
    borderRadius: 7,
    padding: "4px 9px",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  fabricChipEditor: {
    marginTop: 6,
    padding: 8,
    border: "1px solid #dbe3ef",
    borderRadius: 8,
    background: "#f8fafc",
    display: "grid",
    gridTemplateColumns: "minmax(120px, 1fr) auto auto",
    alignItems: "center",
    gap: 8,
  },
  fabricChipEditInput: {
    minWidth: 0,
    border: "1px solid #cbd5e1",
    borderRadius: 6,
    padding: "6px 8px",
    color: "#111827",
    fontSize: 12,
    fontWeight: 800,
    outline: "none",
  },
  fabricChipEditActions: {
    gridColumn: "1 / -1",
    display: "flex",
    justifyContent: "flex-end",
    gap: 6,
  },
  fabricChipMenuTools: {
    marginTop: 6,
    paddingTop: 8,
    borderTop: "1px solid #e5e7eb",
    display: "flex",
    alignItems: "center",
    flexWrap: "wrap",
    gap: 8,
  },
  fabricChipMenuToolLabel: {
    display: "inline-flex",
    alignItems: "center",
    gap: 5,
    color: "#475569",
    fontSize: 11,
    fontWeight: 900,
  },
  fabricCellActions: {
    display: "flex",
    gap: 6,
    flexWrap: "wrap",
    padding: "0 10px 8px",
  },
  fabricProductsCell: {
    position: "relative",
    minWidth: 190,
    minHeight: 96,
    padding: 10,
  },
  fabricProductsButton: {
    width: "100%",
    minHeight: 76,
    display: "flex",
    alignContent: "flex-start",
    alignItems: "flex-start",
    gap: 6,
    flexWrap: "wrap",
    border: "1px solid transparent",
    borderRadius: 6,
    padding: 0,
    background: "transparent",
    color: "var(--portal-table-text-color, #1f2937)",
    font: "inherit",
    fontWeight: 800,
    textAlign: "left",
    cursor: "pointer",
  },
  fabricProductsEmptyButton: {
    width: "100%",
    minHeight: 76,
    border: "1px dashed #cbd5e1",
    borderRadius: 6,
    background: "#f8fafc",
    color: "#64748b",
    font: "inherit",
    fontWeight: 900,
    cursor: "pointer",
  },
  fabricProductChip: {
    display: "inline-flex",
    maxWidth: "100%",
    borderRadius: 999,
    padding: "4px 8px",
    background: "#eef2ff",
    color: "#3730a3",
    fontSize: 11,
    fontWeight: 900,
    lineHeight: 1.2,
  },
  fabricProductMore: {
    display: "inline-flex",
    borderRadius: 999,
    padding: "4px 8px",
    background: "#e2e8f0",
    color: "#334155",
    fontSize: 11,
    fontWeight: 900,
  },
  fabricProductsPopover: {
    position: "absolute",
    left: 10,
    top: 74,
    zIndex: 220,
    width: 310,
    display: "grid",
    gap: 10,
    padding: 12,
    background: "#fff",
    border: "1px solid #cbd5e1",
    borderRadius: 8,
    boxShadow: "0 18px 38px rgba(15,23,42,0.22)",
  },
  fabricProductsLabel: {
    display: "grid",
    gap: 6,
    color: "#334155",
    fontSize: 12,
    fontWeight: 900,
  },
  fabricProductsTextarea: {
    width: "100%",
    border: "1px solid #cbd5e1",
    borderRadius: 6,
    padding: "8px 9px",
    color: "#1f2937",
    font: "inherit",
    fontSize: 12,
    fontWeight: 700,
    resize: "none",
    outline: "none",
  },
  fabricProductsActions: {
    display: "flex",
    justifyContent: "flex-end",
    gap: 8,
  },
  fabricStyleUsageModal: {
    width: "min(860px, 100%)",
    maxHeight: "calc(100vh - 40px)",
    overflow: "auto",
    border: "1px solid #cbd5e1",
    borderRadius: 12,
    padding: 18,
    background: "#fff",
    boxShadow: "0 24px 60px rgba(15,23,42,0.28)",
  },
  fabricStyleUsageLayout: {
    display: "grid",
    gridTemplateColumns: "30% minmax(0, 1fr)",
    gap: 16,
    alignItems: "start",
  },
  fabricStyleUsageImagePane: {
    position: "sticky",
    top: 0,
    width: "100%",
    display: "grid",
    gap: 12,
  },
  fabricStyleUsageImageFrame: {
    width: "100%",
    aspectRatio: "1 / 1.25",
    overflow: "hidden",
    border: "1px solid #dbe3ef",
    borderRadius: 10,
    background: "#f8fafc",
  },
  fabricStyleUsageImage: {
    width: "100%",
    height: "100%",
    objectFit: "cover",
    display: "block",
  },
  fabricStyleUsageImageEmpty: {
    height: "100%",
    display: "grid",
    placeItems: "center",
    color: "#94a3b8",
    fontSize: 13,
    fontWeight: 900,
  },
  fabricStyleUsagePrintName: {
    margin: 0,
    color: "#111827",
    fontSize: 26,
    lineHeight: 1.08,
    fontWeight: 900,
    textAlign: "center",
  },
  fabricStyleUsageContent: {
    minWidth: 0,
    display: "grid",
    gap: 14,
  },
  fabricStyleUsageHeader: {
    display: "flex",
    alignItems: "flex-start",
    justifyContent: "space-between",
    gap: 12,
  },
  fabricStyleSearchWrap: {
    position: "relative",
    display: "grid",
    gap: 8,
  },
  fabricStyleSearchInput: {
    width: "100%",
    border: "1px solid #cbd5e1",
    borderRadius: 8,
    padding: "10px 12px",
    color: "#111827",
    font: "inherit",
    fontSize: 14,
    fontWeight: 800,
    outline: "none",
  },
  fabricStyleSearchResults: {
    display: "grid",
    gap: 6,
    maxHeight: 220,
    overflow: "auto",
    border: "1px solid #e2e8f0",
    borderRadius: 8,
    padding: 6,
    background: "#f8fafc",
  },
  fabricStyleSearchResult: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 12,
    border: "1px solid #e2e8f0",
    borderRadius: 7,
    padding: "8px 10px",
    background: "#fff",
    color: "#111827",
    textAlign: "left",
    cursor: "pointer",
  },
  fabricStyleSearchName: {
    fontSize: 13,
    fontWeight: 900,
  },
  fabricStyleSearchCategory: {
    color: "#64748b",
    fontSize: 12,
    fontWeight: 800,
    whiteSpace: "nowrap",
  },
  fabricStyleUsageTableWrap: {
    overflow: "auto",
    border: "1px solid #e2e8f0",
    borderRadius: 8,
  },
  fabricStyleUsageTable: {
    width: "100%",
    borderCollapse: "collapse",
    minWidth: 520,
  },
  fabricStyleUsageTh: {
    padding: "9px 10px",
    background: "#f8fafc",
    borderBottom: "1px solid #e2e8f0",
    color: "#475569",
    fontSize: 12,
    fontWeight: 900,
    textAlign: "left",
    textTransform: "uppercase",
  },
  fabricStyleUsageTd: {
    padding: "9px 10px",
    borderBottom: "1px solid #eef2f7",
    color: "#111827",
    fontSize: 13,
    fontWeight: 800,
    verticalAlign: "middle",
  },
  fabricStyleUsageInput: {
    width: 140,
    border: "1px solid #cbd5e1",
    borderRadius: 7,
    padding: "7px 9px",
    color: "#111827",
    font: "inherit",
    fontSize: 13,
    fontWeight: 800,
    outline: "none",
  },
  fabricStyleUsageEmpty: {
    padding: 22,
    color: "#64748b",
    fontSize: 13,
    fontWeight: 800,
    textAlign: "center",
  },
  fabricNoteCell: {
    minWidth: 180,
  },
  fabricMiniButton: {
    border: "1px solid #cbd5e1",
    borderRadius: 6,
    padding: "4px 7px",
    background: "#fff",
    color: "#374151",
    fontSize: 11,
    fontWeight: 800,
    cursor: "pointer",
  },
  fabricSheetImage: { width: "100%", height: "100%", objectFit: "cover", borderRadius: 6, border: "1px solid #e5e7eb" },
  imageDeleteOverlay: {
    position: "absolute",
    right: 8,
    top: 8,
    zIndex: 3,
    border: "1px solid rgba(255,255,255,0.7)",
    borderRadius: 999,
    padding: "4px 8px",
    background: "rgba(17,24,39,0.88)",
    color: "#fff",
    fontSize: 11,
    fontWeight: 900,
    cursor: "pointer",
    opacity: 0,
    transition: "opacity 120ms ease",
  },
  imageDeleteOverlayVisible: {
    opacity: 1,
  },
  packingImageCell: {
    display: "grid",
    gap: 6,
    justifyItems: "center",
    width: "100%",
  },
  fabricLink: { color: "#2563eb", fontWeight: 800, textDecoration: "none" },
  tableWrap: {
    maxHeight: "calc(100vh - 118px)",
    overflowX: "scroll" as const,
    overflowY: "auto" as const,
    scrollbarGutter: "stable" as const,
    background: "#fff",
    border: "1px solid #cbd5e1",
    boxShadow: "0 1px 2px rgba(15,23,42,0.08)",
  },
  table: {
    borderCollapse: "separate",
    borderSpacing: 0,
    fontSize: "var(--portal-table-font-size)",
    minWidth: 900,
    tableLayout: "fixed",
  },
  headerRow: { background: "#eef2f7" },
  th: {
    padding: "8px 10px",
    fontWeight: 700,
    fontSize: "calc(var(--portal-table-font-size) - 2px)",
    textTransform: "uppercase",
    letterSpacing: "0.05em",
    color: "var(--portal-heading-text-color)",
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
  headerEditInput: {
    width: "100%",
    border: 0,
    outline: "none",
    background: "transparent",
    color: "inherit",
    font: "inherit",
    fontWeight: "inherit",
    textTransform: "inherit",
    letterSpacing: "inherit",
    padding: 0,
    minWidth: 42,
  },
  rowNumberHeader: {
    width: 48,
    minWidth: 48,
    textAlign: "center",
    color: "#64748b",
    left: 0,
    zIndex: 57,
  },
  rowNumberCell: {
    width: 48,
    minWidth: 48,
    padding: "8px 6px",
    verticalAlign: "middle",
    textAlign: "center",
    color: "#64748b",
    background: "#f8fafc",
    border: "1px solid #d1d5db",
    fontSize: 12,
    fontWeight: 900,
    cursor: "context-menu",
    position: "sticky",
    left: 0,
    zIndex: 42,
  },
  rowResizeHandle: {
    position: "absolute",
    left: 0,
    right: 0,
    bottom: -4,
    height: 8,
    cursor: "row-resize",
    zIndex: 4,
  },
  contextMenu: {
    position: "fixed",
    zIndex: 1200,
    minWidth: 180,
    display: "grid",
    gap: 4,
    padding: 6,
    background: "#fff",
    border: "1px solid #cbd5e1",
    borderRadius: 8,
    boxShadow: "0 18px 42px rgba(15,23,42,0.24)",
  },
  contextMenuButton: {
    border: 0,
    borderRadius: 6,
    padding: "8px 10px",
    background: "#fff",
    color: "#111827",
    fontSize: 12,
    fontWeight: 800,
    textAlign: "left",
    cursor: "pointer",
  },
  contextMenuLabel: {
    display: "grid",
    gap: 5,
    padding: "8px 10px",
    color: "#111827",
    fontSize: 12,
    fontWeight: 800,
  },
  contextMenuSelect: {
    width: "100%",
    border: "1px solid #cbd5e1",
    borderRadius: 6,
    padding: "6px 8px",
    background: "#fff",
    color: "#111827",
    fontSize: 12,
    fontWeight: 700,
  },
  contextMenuDanger: { color: "#991b1b" },
  contextMenuDisabled: { color: "#94a3b8", cursor: "not-allowed" },
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
  customCellTextarea: {
    width: "100%",
    minHeight: 96,
    border: "none",
    outline: "none",
    resize: "vertical",
    background: "transparent",
    padding: "9px 10px",
    color: "var(--portal-table-text-color, #1f2937)",
    font: "inherit",
    fontWeight: 700,
    boxSizing: "border-box",
  },
  row: { background: "#fff" },
  addOrderRow: { background: "#f8fafc" },
  td: {
    padding: "8px 10px",
    verticalAlign: "middle",
    color: "var(--portal-table-text-color)",
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
  skuCellWithToggle: {
    position: "relative",
    minHeight: 72,
    display: "flex",
    alignItems: "center",
  },
  inventoryToggle: {
    position: "absolute",
    right: -9,
    bottom: -11,
    width: 16,
    height: 16,
    borderRadius: 3,
    border: 0,
    background: "transparent",
    fontSize: 9,
    fontWeight: 900,
    lineHeight: "16px",
    textAlign: "center",
    cursor: "pointer",
    zIndex: 35,
    padding: 0,
  },
  inventoryRow: { background: "#f8fafc" },
  inventoryBlankCell: {
    padding: "8px 10px",
    verticalAlign: "middle",
    border: "1px solid #d1d5db",
    background: "#f8fafc",
  },
  inventoryLabelCell: {
    padding: "8px 10px",
    verticalAlign: "middle",
    border: "1px solid #d1d5db",
    background: "#f8fafc",
    color: "#374151",
    fontSize: 13,
    fontWeight: 900,
  },
  inventoryQtyCell: {
    background: "#f8fafc",
    color: "#374151",
    fontSize: 13,
    fontWeight: 900,
    textAlign: "center",
  },
  inventoryStatusCell: {
    background: "#f8fafc",
    color: "#374151",
    fontSize: 13,
    fontWeight: 800,
  },
  addOrderHint: { color: "#6b7280", fontWeight: 800, fontSize: 12 },
  addOrderSelect: {
    border: "1px solid #b6c0cc",
    borderRadius: 6,
    padding: "6px 8px",
    width: "100%",
    background: "#fff",
    color: "#374151",
    fontSize: 12,
    fontWeight: 700,
  },
  restockSearchInput: {
    width: "100%",
    border: "1px solid #b6c0cc",
    borderRadius: 7,
    padding: "8px 10px",
    background: "#fff",
    color: "#111827",
    fontSize: 13,
    fontWeight: 700,
    outline: "none",
    boxSizing: "border-box",
  },
  selectedRestockProduct: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 8,
    color: "#111827",
    fontWeight: 800,
    lineHeight: 1.3,
  },
  clearProductButton: {
    border: 0,
    borderRadius: 999,
    width: 22,
    height: 22,
    background: "#6b7280",
    color: "#fff",
    fontSize: 15,
    fontWeight: 900,
    lineHeight: "20px",
    cursor: "pointer",
    flex: "0 0 auto",
  },
  addOrderButtonReady: {
    background: "#111827",
    borderColor: "#111827",
    color: "#fff",
  },
  qty: { fontWeight: 700, color: "#111827" },
  qtyZero: { color: "#d1d5db" },
  dateText: { color: "#374151", fontWeight: 600, whiteSpace: "nowrap" },
  total: { fontWeight: 700, fontSize: "var(--portal-inventory-font-size, 14px)", color: "#111827" },
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
  restockNoteTextarea: {
    width: "100%",
    minHeight: 96,
    border: "none",
    outline: "none",
    resize: "none",
    padding: "9px 10px",
    background: "transparent",
    color: "var(--portal-table-text-color, #1f2937)",
    font: "inherit",
    fontWeight: 700,
    lineHeight: 1.35,
    boxSizing: "border-box",
  },
  noteTagWrap: { position: "relative" },
  tagSuggestions: {
    position: "absolute",
    left: 0,
    top: "calc(100% + 4px)",
    minWidth: 180,
    display: "grid",
    gap: 4,
    padding: 6,
    background: "#fff",
    border: "1px solid #cbd5e1",
    borderRadius: 8,
    boxShadow: "0 14px 30px rgba(15,23,42,0.2)",
    zIndex: 180,
  },
  tagSuggestionButton: {
    border: 0,
    borderRadius: 6,
    padding: "7px 8px",
    background: "#f8fafc",
    color: "#111827",
    fontSize: 12,
    fontWeight: 800,
    textAlign: "left",
    cursor: "pointer",
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
    fontSize: "var(--portal-inventory-font-size, 13px)",
    fontWeight: 700,
    textAlign: "center",
    outline: "none",
    background: "transparent",
    boxSizing: "border-box",
  },
  loadedInventoryCell: {
    background: "#92AD9B",
    color: "#fff",
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
  logDateBlock: {
    border: "1px solid #e5e7eb",
    borderRadius: 10,
    overflow: "hidden",
  },
  logDateButton: {
    width: "100%",
    display: "flex",
    alignItems: "center",
    gap: 12,
    padding: "10px 14px",
    background: "#f1f5f9",
    border: "none",
    cursor: "pointer",
    fontSize: 13,
    fontWeight: 700,
    color: "#111827",
    textAlign: "left" as const,
  },
  logDateCount: {
    background: "#e2e8f0",
    borderRadius: 999,
    padding: "2px 9px",
    fontSize: 11,
    fontWeight: 800,
    color: "#475569",
  },
  logEntries: {
    overflowX: "auto" as const,
    background: "#fff",
  },
  logTh: {
    padding: "7px 10px",
    borderBottom: "2px solid #e5e7eb",
    textAlign: "left" as const,
    fontSize: 11,
    fontWeight: 800,
    color: "#6b7280",
    whiteSpace: "nowrap" as const,
    background: "#f8fafc",
  },
  logTd: {
    padding: "6px 10px",
    borderBottom: "1px solid #f1f5f9",
    fontSize: 12,
    color: "#374151",
    whiteSpace: "nowrap" as const,
  },
  samplePanel: {
    position: "fixed" as const,
    top: 0,
    right: 0,
    bottom: 0,
    width: 800,
    maxWidth: "95vw",
    background: "#fff",
    boxShadow: "-4px 0 32px rgba(0,0,0,0.18)",
    zIndex: 1200,
    display: "flex",
    flexDirection: "column" as const,
    overflow: "hidden",
  },
  samplePanelBackdrop: {
    position: "fixed" as const,
    inset: 0,
    background: "rgba(15,23,42,0.35)",
    zIndex: 1199,
  },
  samplePanelHeader: {
    display: "flex",
    alignItems: "flex-start",
    justifyContent: "space-between",
    gap: 12,
    padding: "20px 24px 16px",
    borderBottom: "1px solid #e5e7eb",
    flexShrink: 0,
  },
  samplePanelNameWrap: {
    display: "flex",
    flexDirection: "column" as const,
    gap: 2,
    flex: 1,
    minWidth: 0,
  },
  samplePanelName: {
    fontSize: 20,
    fontWeight: 700,
    color: "#111827",
    margin: 0,
    cursor: "pointer",
    borderRadius: 4,
    padding: "2px 4px",
    marginLeft: -4,
  },
  samplePanelNameInput: {
    fontSize: 20,
    fontWeight: 700,
    color: "#111827",
    border: "2px solid #2563eb",
    borderRadius: 6,
    padding: "2px 8px",
    outline: "none",
    width: "100%",
    background: "#fff",
  },
  samplePanelVersionCount: {
    fontSize: 12,
    color: "#9ca3af",
    paddingLeft: 4,
  },
  samplePanelClose: {
    flexShrink: 0,
    width: 32,
    height: 32,
    border: "none",
    background: "#f1f5f9",
    borderRadius: 8,
    cursor: "pointer",
    fontSize: 20,
    lineHeight: "32px",
    textAlign: "center" as const,
    color: "#6b7280",
    padding: 0,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },
  samplePanelTopActions: {
    padding: "14px 24px",
    borderBottom: "1px solid #f1f5f9",
    flexShrink: 0,
  },
  samplePanelIterations: {
    flex: 1,
    minHeight: 0,
    overflowY: "auto" as const,
    WebkitOverflowScrolling: "touch" as const,
    padding: "16px 24px 32px",
    display: "flex",
    flexDirection: "column" as const,
    gap: 16,
  },
  samplePanelEmpty: {
    color: "#9ca3af",
    fontSize: 14,
    textAlign: "center" as const,
    padding: "48px 0",
  },
  sampleIterationBlock: {
    border: "1px solid #e5e7eb",
    borderRadius: 12,
    overflow: "hidden",
    background: "#fafafa",
  },
  sampleIterationHeader: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    padding: "12px 16px",
    background: "#fff",
    borderBottom: "1px solid #f1f5f9",
  },
  sampleIterationVersion: {
    fontSize: 14,
    fontWeight: 700,
    color: "#111827",
    flex: 1,
  },
  sampleIterationDate: {
    fontSize: 12,
    color: "#9ca3af",
  },
  sampleIterationImages: {
    display: "flex",
    flexWrap: "wrap" as const,
    gap: 10,
    padding: "14px 16px",
  },
  sampleIterationImageWrap: {
    position: "relative" as const,
    width: 140,
    height: 186,
    borderRadius: 8,
    overflow: "hidden",
    background: "#f1f5f9",
    flexShrink: 0,
    border: "1px solid #e5e7eb",
  },
  sampleIterationImage: {
    width: "100%",
    height: "100%",
    objectFit: "cover" as const,
    display: "block",
  },
  sampleIterationImageRemove: {
    position: "absolute" as const,
    top: 4,
    right: 4,
    width: 22,
    height: 22,
    borderRadius: "50%",
    background: "rgba(0,0,0,0.55)",
    color: "#fff",
    border: "none",
    cursor: "pointer",
    fontSize: 14,
    lineHeight: "22px",
    textAlign: "center" as const,
    padding: 0,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 2,
  },
  sampleIterationAddImage: {
    width: 140,
    height: 186,
    borderRadius: 8,
    border: "2px dashed #cbd5e1",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
    color: "#94a3b8",
    fontSize: 13,
    fontWeight: 600,
    flexShrink: 0,
    background: "#f8fafc",
  },
  sampleIterationNotes: {
    width: "100%",
    border: "none",
    padding: "8px 16px 14px",
    fontSize: "var(--portal-panel-font-size, 13px)" as unknown as number,
    color: "#374151",
    resize: "vertical" as const,
    minHeight: 84,
    background: "#fafafa",
    fontFamily: "inherit",
    boxSizing: "border-box" as const,
    outline: "none",
  },
  sampleVersionBadge: {
    fontSize: 11,
    fontWeight: 700,
    background: "#e0e7ef",
    color: "#4b5563",
    borderRadius: 99,
    padding: "1px 7px",
  },
  sampleCardOverlay: {
    position: "absolute" as const,
    inset: 0,
    background: "rgba(0,0,0,0.52)",
    display: "flex",
    flexDirection: "column" as const,
    alignItems: "center",
    justifyContent: "center",
    gap: 8,
    borderRadius: 0,
  },
  sampleCardOverlayBtn: {
    background: "rgba(255,255,255,0.95)",
    color: "#111827",
    border: "none",
    borderRadius: 7,
    padding: "8px 0",
    fontSize: 13,
    fontWeight: 600,
    cursor: "pointer",
    width: 160,
    textAlign: "center" as const,
    display: "block",
  },
  sampleCardOverlayBtnDanger: {
    color: "#dc2626",
    background: "rgba(255,255,255,0.92)",
  },

  imageUploadModalBackdrop: {
    position: "fixed" as const,
    inset: 0,
    background: "rgba(0,0,0,0.55)",
    zIndex: 1400,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },
  imageUploadModal: {
    background: "#fff",
    borderRadius: 12,
    padding: "24px 28px",
    width: 420,
    maxWidth: "90vw",
    boxShadow: "0 8px 40px rgba(0,0,0,0.22)",
    display: "flex",
    flexDirection: "column" as const,
    gap: 16,
  },
  imageUploadModalHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
  },
  imageUploadDropZone: {
    border: "2px dashed #d1d5db",
    borderRadius: 10,
    padding: "36px 24px",
    display: "flex",
    flexDirection: "column" as const,
    alignItems: "center",
    gap: 8,
    cursor: "pointer",
    transition: "border-color 0.15s, background 0.15s",
    background: "#fafafa",
  },
  imageUploadDropZoneActive: {
    borderColor: "#2563eb",
    background: "#eff6ff",
  },
  imageUploadDropIcon: {
    fontSize: 36,
    lineHeight: 1,
  },
  imageUploadDropText: {
    fontWeight: 600,
    fontSize: 14,
    color: "#374151",
  },
  imageUploadDropSubtext: {
    fontSize: 12,
    color: "#9ca3af",
  },

  lightboxBackdrop: {
    position: "fixed" as const,
    inset: 0,
    background: "rgba(0,0,0,0.88)",
    zIndex: 1500,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },
  lightboxImage: {
    maxWidth: "88vw",
    maxHeight: "88vh",
    objectFit: "contain" as const,
    borderRadius: 4,
    pointerEvents: "none" as const,
    userSelect: "none" as const,
  },
  lightboxClose: {
    position: "absolute" as const,
    top: 18,
    right: 24,
    background: "none",
    border: "none",
    color: "#fff",
    fontSize: 36,
    lineHeight: 1,
    cursor: "pointer",
    zIndex: 1,
  },
  lightboxArrow: {
    position: "absolute" as const,
    top: "50%",
    transform: "translateY(-50%)",
    background: "rgba(255,255,255,0.12)",
    border: "none",
    color: "#fff",
    fontSize: 48,
    lineHeight: 1,
    cursor: "pointer",
    borderRadius: 6,
    padding: "4px 14px",
    zIndex: 1,
    transition: "background 0.15s",
  },
  lightboxArrowLeft: {
    left: 18,
  },
  lightboxArrowRight: {
    right: 18,
  },
  lightboxCounter: {
    position: "absolute" as const,
    bottom: 20,
    left: "50%",
    transform: "translateX(-50%)",
    color: "rgba(255,255,255,0.75)",
    fontSize: 13,
    fontWeight: 500,
    letterSpacing: "0.04em",
    pointerEvents: "none" as const,
  },
};

export function ErrorBoundary() {
  const error = useRouteError();
  const message = isRouteErrorResponse(error)
    ? error.data
    : error instanceof Error
    ? error.message
    : "An unexpected error occurred.";
  return (
    <div style={{ fontFamily: "Inter, sans-serif", padding: "3rem 2rem", maxWidth: 480, margin: "0 auto", color: "#374151" }}>
      <h2 style={{ fontSize: 20, fontWeight: 700, marginBottom: 8 }}>Something went wrong</h2>
      <p style={{ color: "#6b7280", marginBottom: 24, fontSize: 14 }}>{String(message)}</p>
      <button
        type="button"
        onClick={() => window.location.reload()}
        style={{ background: "#2563eb", color: "#fff", border: "none", borderRadius: 6, padding: "8px 18px", cursor: "pointer", fontSize: 14, marginRight: 12 }}
      >
        Reload page
      </button>
      <a href="/portal" style={{ color: "#2563eb", fontSize: 14 }}>← Back to portal</a>
    </div>
  );
}
