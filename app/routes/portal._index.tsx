import { useEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";
import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import { useFetcher, useLoaderData, useSearchParams } from "react-router";
import prisma from "../db.server";
import { fabricStockSheets, type FabricStockSheet } from "../fabric-stock-data";
import { syncOrderNoteMessages } from "../portal-messages.server";
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
  const messageOrderId = Number(url.searchParams.get("messageOrderId") ?? 0) || null;
  const packingId = Number(url.searchParams.get("packingId") ?? 0) || null;
  const productSearch = url.searchParams.get("productSearch") ?? "";
  const restockProductSearch = url.searchParams.get("restockProductSearch") ?? "";
  const packingSearchLineId = Number(url.searchParams.get("packingSearchLineId") ?? 0) || null;
  const sortBy = url.searchParams.get("sortBy") ?? "orderDateDesc";
  const selectedFabricTab = url.searchParams.get("fabricTab") ?? "";
  const [allOrders, columnWidthsSetting, packingColumnWidthsSetting, restockSettingsSetting, universalSettingsSetting, fabricSettingsSetting, loginRequiredSetting, usersSetting, activeUsersSetting, packingLists, navOrderSetting] = await Promise.all([
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
      where: { key: RESTOCK_SETTINGS_KEY },
      select: { value: true },
    }),
    prisma.portalSetting.findUnique({
      where: { key: UNIVERSAL_SETTINGS_KEY },
      select: { value: true },
    }),
    prisma.portalSetting.findUnique({
      where: { key: FABRIC_SETTINGS_KEY },
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
    prisma.portalSetting.findUnique({ where: { key: PORTAL_NAV_ORDER_KEY }, select: { value: true } }),
  ]);
  const ninetyDaysAgo = new Date(Date.now() - 90 * 24 * 60 * 60 * 1000);
  const activityLogs = await prisma.activityLog.findMany({
      where: { createdAt: { gte: ninetyDaysAgo } },
      orderBy: { createdAt: "desc" },
      take: 1000,
    }).catch(() => [] as { id: number; userName: string; action: string; entity: string; entityId: string | null; entityName: string | null; field: string | null; toValue: string | null; createdAt: Date }[]);
  const users = normalizePortalUsers(usersSetting?.value);
  const restockSettings = normalizeRestockSettings(restockSettingsSetting?.value);
  const universalSettings = normalizeUniversalSettings(universalSettingsSetting?.value);
  const fabricSettings = normalizeFabricSettings(fabricSettingsSetting?.value);
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
    .sort((a, b) => labelForOption(restockSettings.statusOptions, a).localeCompare(labelForOption(restockSettings.statusOptions, b)));
  const priorityFilters = Array.from(new Set(normalizedOrders.map((order) => order.priority).filter(Boolean) as string[]))
    .sort((a, b) => labelForOption(restockSettings.priorityOptions, a).localeCompare(labelForOption(restockSettings.priorityOptions, b)));
  const filteredOrders = normalizedOrders
    .filter((order) => !selectedProductGroup || order.productType === selectedProductGroup)
    .filter((order) => !selectedStatus || order.supplierStatus === selectedStatus)
    .filter((order) => !selectedPriority || order.priority === selectedPriority)
    .filter((order) => !searchTitle || order.productTitle.toLowerCase().includes(searchTitle.toLowerCase()))
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
  const selectedPackingList = packingId
    ? packingLists.find((list) => list.id === packingId) ?? null
    : null;
  const productResults = page === "packing" && selectedPackingList && productSearch.trim().length >= 2
    ? await searchShopifyProducts(productSearch)
    : [];
  const restockProductResults = page === "restock" && restockProductSearch.trim().length >= 2
    ? await searchShopifyProducts(restockProductSearch)
    : [];
  const messages = currentUser
    ? await prisma.portalMessage.findMany({
        where: { userId: currentUser.id, readAt: null },
        orderBy: { createdAt: "desc" },
        take: 25,
      })
    : [];
  const fabricCellOverridesSetting = page === "fabric"
    ? await prisma.portalSetting.findUnique({
        where: { key: FABRIC_CELL_OVERRIDES_KEY },
        select: { value: true },
      })
    : null;
  const fabricCustomRowsSetting = page === "fabric"
    ? await prisma.portalSetting.findUnique({
        where: { key: FABRIC_CUSTOM_ROWS_KEY },
        select: { value: true },
      })
    : null;
  const fabricDeletedRowsSetting = page === "fabric"
    ? await prisma.portalSetting.findUnique({
        where: { key: FABRIC_DELETED_ROWS_KEY },
        select: { value: true },
      })
    : null;
  const inrToAudRate = page === "fabric" ? await getInrToAudRate() : null;
  const fabricSheets = page === "fabric"
    ? getFabricSheets(
        fabricCellOverridesSetting?.value,
        fabricCustomRowsSetting?.value,
        fabricDeletedRowsSetting?.value,
        fabricSettings.tileOrder,
      )
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
    priorityFilters,
    sortBy,
    page,
    columnWidths: normalizeColumnWidths(columnWidthsSetting?.value),
    packingColumnWidths: normalizeColumnWidths(packingColumnWidthsSetting?.value),
    restockSettings,
    universalSettings,
    fabricSettings,
    packingLists,
    selectedPackingList,
    productSearch,
    restockProductSearch,
    packingSearchLineId,
    productResults,
    restockProductResults,
    loginRequired,
    users,
    currentUser,
    activeUsers,
    messages,
    messageOrderId,
    loginBlocked: loginRequired && users.length > 0 && !currentUser,
    activityLogs,
    navOrder,
    fabricSheets,
    selectedFabricTab,
    inrToAudRate,
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
    const line = await prisma.packingListLine.findUnique({ where: { id: lineId }, select: { qtys: true, shopifyLoadedQtys: true, productTitle: true } });
    if (!line || !size) return null;
    const qtys = normalizeQtys(line.qtys);
    const shopifyLoadedQtys = normalizeQtys(line.shopifyLoadedQtys);
    const previousQty = qtys[size] ?? 0;
    qtys[size] = value;
    delete shopifyLoadedQtys[size];
    await prisma.packingListLine.update({ where: { id: lineId }, data: { qtys, shopifyLoadedQtys } });
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

  if (intent === "load_packing_inventory") {
    const packingId = Number(form.get("packingId"));
    const skipWords = String(form.get("skipWords") ?? "")
      .split(",")
      .map((word) => word.trim().toLowerCase())
      .filter(Boolean);
    const packingList = await prisma.packingList.findUnique({
      where: { id: packingId },
      include: { lines: { orderBy: [{ boxNumber: "asc" }, { sortOrder: "asc" }, { id: "asc" }] } },
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

      const qtys = normalizeQtys(line.qtys);
      const loadedQtys = normalizeQtys(line.shopifyLoadedQtys);
      const variants = await getVariants(line.productId);
      const changes: ShopifyInventoryChange[] = [];
      const nextLoadedQtys: Record<string, number> = { ...loadedQtys };

      for (const [size, qty] of Object.entries(qtys)) {
        if (qty <= 0 || loadedQtys[size] === qty) continue;
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

  if (intent === "update_fabric_cell") {
    const gid = String(form.get("gid") ?? "");
    const rowIndex = Number(form.get("rowIndex"));
    const colIndex = Number(form.get("colIndex"));
    const value = String(form.get("value") ?? "");
    if (!gid || !Number.isInteger(rowIndex) || !Number.isInteger(colIndex) || rowIndex < 0 || colIndex < 0) {
      return null;
    }

    const setting = await prisma.portalSetting.findUnique({
      where: { key: FABRIC_CELL_OVERRIDES_KEY },
      select: { value: true },
    });
    const overrides = normalizeFabricCellOverrides(setting?.value);
    const key = fabricCellKey(gid, rowIndex, colIndex);
    overrides[key] = value;

    await prisma.portalSetting.upsert({
      where: { key: FABRIC_CELL_OVERRIDES_KEY },
      create: { key: FABRIC_CELL_OVERRIDES_KEY, value: overrides },
      update: { value: overrides },
    });
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
    const setting = await prisma.portalSetting.findUnique({
      where: { key: FABRIC_CELL_OVERRIDES_KEY },
      select: { value: true },
    });
    const overrides = normalizeFabricCellOverrides(setting?.value);
    overrides[fabricCellKey(gid, rowIndex, colIndex)] = dataUrl;

    await prisma.portalSetting.upsert({
      where: { key: FABRIC_CELL_OVERRIDES_KEY },
      create: { key: FABRIC_CELL_OVERRIDES_KEY, value: overrides },
      update: { value: overrides },
    });
    return null;
  }

  if (intent === "delete_fabric_row" || intent === "move_fabric_row") {
    const gid = String(form.get("gid") ?? "");
    const rowIndex = Number(form.get("rowIndex"));
    const targetGid = String(form.get("targetGid") ?? "");
    const sourceSheet = fabricStockSheets.find((item) => item.gid === gid);
    const targetSheet = fabricStockSheets.find((item) => item.gid === targetGid);
    if (!sourceSheet || !Number.isInteger(rowIndex) || rowIndex < 0) return null;
    if (intent === "move_fabric_row" && (!targetSheet || targetSheet.gid === sourceSheet.gid || isHiddenFabricSheet(targetSheet.name))) return null;

    const [customRowsSetting, deletedRowsSetting, overridesSetting] = await Promise.all([
      prisma.portalSetting.findUnique({ where: { key: FABRIC_CUSTOM_ROWS_KEY }, select: { value: true } }),
      prisma.portalSetting.findUnique({ where: { key: FABRIC_DELETED_ROWS_KEY }, select: { value: true } }),
      prisma.portalSetting.findUnique({ where: { key: FABRIC_CELL_OVERRIDES_KEY }, select: { value: true } }),
    ]);
    const customRows = normalizeFabricCustomRows(customRowsSetting?.value);
    const deletedRows = normalizeFabricDeletedRows(deletedRowsSetting?.value);
    const overrides = normalizeFabricCellOverrides(overridesSetting?.value);

    if (intent === "move_fabric_row" && targetSheet) {
      const originalRows = [...sourceSheet.rows, ...(customRows[sourceSheet.gid] ?? [])];
      const row = originalRows[rowIndex]?.map((value, colIndex) => overrides[fabricCellKey(sourceSheet.gid, rowIndex, colIndex)] ?? value);
      if (!row) return null;
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
      customRows[targetSheet.gid] = [...(customRows[targetSheet.gid] ?? []), movedRow];
    }

    if (rowIndex >= sourceSheet.rows.length) {
      const customIndex = rowIndex - sourceSheet.rows.length;
      customRows[sourceSheet.gid] = (customRows[sourceSheet.gid] ?? []).filter((_, index) => index !== customIndex);
    } else {
      deletedRows[fabricRowKey(sourceSheet.gid, rowIndex)] = true;
    }

    await Promise.all([
      prisma.portalSetting.upsert({
        where: { key: FABRIC_CUSTOM_ROWS_KEY },
        create: { key: FABRIC_CUSTOM_ROWS_KEY, value: customRows },
        update: { value: customRows },
      }),
      prisma.portalSetting.upsert({
        where: { key: FABRIC_DELETED_ROWS_KEY },
        create: { key: FABRIC_DELETED_ROWS_KEY, value: deletedRows },
        update: { value: deletedRows },
      }),
    ]);
    return null;
  }

  if (intent === "add_fabric_row") {
    const gid = String(form.get("gid") ?? "");
    const sheet = fabricStockSheets.find((item) => item.gid === gid);
    if (!sheet) return null;

    const setting = await prisma.portalSetting.findUnique({
      where: { key: FABRIC_CUSTOM_ROWS_KEY },
      select: { value: true },
    });
    const customRows = normalizeFabricCustomRows(setting?.value);
    customRows[gid] = [...(customRows[gid] ?? []), Array.from({ length: sheet.headers.length }, () => "")];

    await prisma.portalSetting.upsert({
      where: { key: FABRIC_CUSTOM_ROWS_KEY },
      create: { key: FABRIC_CUSTOM_ROWS_KEY, value: customRows },
      update: { value: customRows },
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
  { value: "loaded", label: "Inventory loaded" },
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

function labelForOption(options: RestockOption[], value: string) {
  return options.find((option) => option.value === value)?.label ?? value;
}

function labelForPackingStatus(value: string) {
  return PACKING_STATUS_OPTIONS.find((option) => option.value === value)?.label ?? value;
}

const COLUMN_WIDTHS_KEY = "supplier-portal-column-widths-v1";
const PACKING_COLUMN_WIDTHS_KEY = "supplier-portal-packing-column-widths-v1";
const RESTOCK_SETTINGS_KEY = "supplier-portal-restock-settings-v1";
const UNIVERSAL_SETTINGS_KEY = "production-portal-universal-settings-v1";
const PORTAL_NAV_ORDER_KEY = "production-portal-nav-order-v1";
const FABRIC_SETTINGS_KEY = "production-portal-fabric-settings-v1";
const FABRIC_CELL_OVERRIDES_KEY = "production-portal-fabric-cell-overrides-v1";
const FABRIC_CUSTOM_ROWS_KEY = "production-portal-fabric-custom-rows-v1";
const FABRIC_DELETED_ROWS_KEY = "production-portal-fabric-deleted-rows-v1";
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
] as const;
type NavItemId = typeof ALL_NAV_ITEMS[number]["id"];
const DEFAULT_NAV_ORDER: NavItemId[] = ["dashboard", "restock", "fabric", "packing", "pricelist", "productinfo", "samples", "newproduct"];
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
  delete: 104,
};

type ColumnDef = { id: string; label: string; center?: boolean };
type PortalUser = { id: string; name: string; admin: boolean; active: boolean };
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
};
type UniversalSettings = {
  primaryButtonBg: string;
  primaryButtonColor: string;
  tableTextSize: number;
  tableTextColor: string;
  headingTextSize: number;
  headingTextColor: string;
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

function buildFabricDefaults(pattern: RegExp): RestockOption[] {
  const labels = new Set<string>();
  for (const sheet of fabricStockSheets) {
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
    fabricTypeOptions: normalizeRestockOptions(settings.fabricTypeOptions, buildFabricDefaults(/fabric\s*type/i)),
    tileOrder: Array.isArray(settings.tileOrder)
      ? settings.tileOrder.map((item) => String(item)).filter(Boolean)
      : [],
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
    menuBg: normalizeHexColor(settings.menuBg, "#111827"),
    menuTextColor: normalizeHexColor(settings.menuTextColor, "#cbd5e1"),
    pageBg: normalizeHexColor(settings.pageBg, "#f3f4f6"),
    logoUrl: typeof settings.logoUrl === "string" ? settings.logoUrl : "",
  };
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

function isHiddenFabricSheet(name: string) {
  return HIDDEN_FABRIC_SHEET_NAMES.has(normalizeFabricName(name));
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
  const onOrderSheet = fabricStockSheets.find((sheet) => sheet.gid === "759049382");
  if (!onOrderSheet) return recovered;

  const brokenMoveKeys = new Set(
    (customRows[onOrderSheet.gid] ?? [])
      .map((row, index) => ({ row, rowIndex: onOrderSheet.rows.length + index }))
      .filter(({ row, rowIndex }) => isBrokenMovedOnOrderRow(onOrderSheet, row, rowIndex))
      .map(({ row }) => fabricMoveMatchKey(onOrderSheet, row))
      .filter(Boolean),
  );
  if (!brokenMoveKeys.size) return recovered;

  for (const sheet of fabricStockSheets) {
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

async function getInrToAudRate() {
  try {
    const response = await fetch("https://open.er-api.com/v6/latest/INR", {
      headers: { Accept: "application/json" },
    });
    if (!response.ok) return null;
    const data = await response.json() as { rates?: Record<string, number> };
    const rate = Number(data.rates?.AUD);
    return Number.isFinite(rate) && rate > 0 ? rate : null;
  } catch {
    return null;
  }
}

function getFabricSheets(overridesValue?: unknown, customRowsValue?: unknown, deletedRowsValue?: unknown, tileOrder: string[] = []): FabricSheetData[] {
  const overrides = normalizeFabricCellOverrides(overridesValue);
  const customRows = normalizeFabricCustomRows(customRowsValue);
  const deletedRows = normalizeFabricDeletedRows(deletedRowsValue);
  const recoveredRows = recoveredBrokenMoveRows(customRows, deletedRows);
  const sheets = fabricStockSheets.filter((sheet) => !isHiddenFabricSheet(sheet.name)).map((sheet) => {
    const originalRows = [...sheet.rows, ...(customRows[sheet.gid] ?? [])];
    const rowEntries = originalRows
      .map((row, rowIndex) => ({ row, rowIndex }))
      .filter(({ row, rowIndex }) => !isBrokenMovedOnOrderRow(sheet, row, rowIndex))
      .filter(({ rowIndex }) => {
        const rowKey = fabricRowKey(sheet.gid, rowIndex);
        return !deletedRows[rowKey] || recoveredRows.has(rowKey) || shouldRestoreLostPeacockRow(sheet.gid, rowIndex);
      });
    const rows = rowEntries.map(({ row, rowIndex }) => (
      row.map((value, colIndex) => overrides[fabricCellKey(sheet.gid, rowIndex, colIndex)] ?? value)
    ));
    return {
      ...sheet,
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
            variants(first: 100) {
              edges {
                node {
                  id
                  title
                  sku
                  inventoryQuantity
                  selectedOptions { name value }
                }
              }
            }
          }
        }
      }
    }
  `;

  const VALID_SIZES = new Set(["XS", "S", "M", "L", "XL", "2XL", "3XL", "S-M", "M-L", "L-XL"]);

  const mapProducts = (json: any) => (json.data?.products?.edges ?? []).map((edge: any) => {
    const rawVariants: any[] = edge.node.variants.edges.map((variantEdge: any) => variantEdge.node);

    const seen = new Set<string>();
    const sizeVariants = rawVariants
      .map((v: any) => {
        // Extract each option value and check if it's a valid size
        const sizeValue: string | null =
          (v.selectedOptions as { name: string; value: string }[] | undefined)
            ?.map((o) => o.value)
            .find((val) => VALID_SIZES.has(val)) ?? null;
        return { raw: v, sizeValue };
      })
      .filter(({ sizeValue }) => {
        if (!sizeValue || seen.has(sizeValue)) return false;
        seen.add(sizeValue);
        return true;
      })
      .map(({ raw: v, sizeValue }) => ({
        id: String(v.id ?? ""),
        title: sizeValue as string,
        sku: v.sku ? String(v.sku) : null,
        availableInventory: Number.isFinite(Number(v.inventoryQuantity)) ? Number(v.inventoryQuantity) : null,
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

type ShopifyVariantInfo = { id: string; title: string; sku: string | null; availableInventory: number | null };
type ShopifyInventoryVariantInfo = ShopifyVariantInfo & { inventoryItemId: string | null };
type ShopifyInventoryChange = { size: string; qty: number; inventoryItemId: string };

async function getShopifyProductVariants(shop: string, productId: string): Promise<ShopifyVariantInfo[]> {
  const session = await prisma.session.findFirst({
    where: { shop, accessToken: { not: "" } },
    orderBy: { isOnline: "asc" },
  });
  if (!session) return [];

  const graphqlQuery = `
    query ProductVariants($id: ID!) {
      product(id: $id) {
        variants(first: 100) {
          nodes { id title sku inventoryQuantity }
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
        availableInventory: Number.isFinite(Number(variant.inventoryQuantity)) ? Number(variant.inventoryQuantity) : null,
      }))
      .filter((variant: ShopifyVariantInfo) => variant.id && variant.title);

  try {
    const { admin } = await unauthenticated.admin(session.shop);
    const response = await admin.graphql(graphqlQuery, { variables: { id: productId } });
    return mapVariants(await response.json());
  } catch {
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
  }
}

function normalizeVariantSizeLabel(value: string) {
  return value.trim().toLowerCase().replace(/\s+/g, "").replace(/-/g, "/");
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
  const session = await prisma.session.findFirst({
    where: { shop, accessToken: { not: "" } },
    orderBy: { isOnline: "asc" },
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
            inventoryQuantity
            inventoryItem { id }
          }
        }
      }
    }
  `;
  const json = await shopifyGraphql<any>(session.shop, session.accessToken, graphqlQuery, { id: productId });

  return (json?.data?.product?.variants?.nodes ?? [])
    .map((variant: any) => ({
      id: String(variant.id ?? ""),
      title: String(variant.title ?? ""),
      sku: variant.sku ? String(variant.sku) : null,
      availableInventory: Number.isFinite(Number(variant.inventoryQuantity)) ? Number(variant.inventoryQuantity) : null,
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
    restockSettings,
    universalSettings,
    fabricSettings,
    packingLists,
    selectedPackingList,
    productSearch,
    restockProductSearch,
    packingSearchLineId,
    productResults,
    restockProductResults,
    loginRequired,
    users,
    currentUser,
    activeUsers,
    messages,
    messageOrderId,
    loginBlocked,
    activityLogs,
    navOrder,
    fabricSheets,
    selectedFabricTab,
    inrToAudRate,
  } = useLoaderData<typeof loader>();
  const [searchParams, setSearchParams] = useSearchParams();
  const columnWidthsFetcher = useFetcher();
  const [showAddRow, setShowAddRow] = useState(false);
  const [showCategoryDropdown, setShowCategoryDropdown] = useState(false);
  const [selectedAddCategory, setSelectedAddCategory] = useState<string | null>(null);
  const [newCategoryInput, setNewCategoryInput] = useState("");
  const [showNewCategoryInput, setShowNewCategoryInput] = useState(false);
  const [historyMenu, setHistoryMenu] = useState<{ x: number; y: number; entity: string; entityId: string; field: string; entityName: string } | null>(null);
  useEffect(() => {
    const handler = (e: Event) => {
      const detail = (e as CustomEvent).detail as { x: number; y: number; entity: string; entityId: string; field: string; entityName: string };
      setHistoryMenu(detail);
    };
    document.addEventListener("show-cell-history", handler);
    return () => document.removeEventListener("show-cell-history", handler);
  }, []);
  const [columnWidths, setColumnWidths] = useState<Record<string, number>>(savedColumnWidths);
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
    { id: "delete", label: "Actions", center: true },
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
  const activePageTitle = page === "dashboard" ? "Dashboard"
    : page === "fabric" ? "Fabric in stock"
    : page === "settings" ? "Settings"
    : page === "packing" ? "Packing Lists"
    : page === "pricelist" ? "Price List"
    : page === "productinfo" ? "Product Information"
    : page === "samples" ? "Samples"
    : page === "newproduct" ? "New Product Orders"
    : selectedProductGroup || "Existing Products Restock";
  const orderedNavItems = navOrder.map((id) => ALL_NAV_ITEMS.find((item) => item.id === id)).filter(Boolean) as typeof ALL_NAV_ITEMS[number][];

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
      } as React.CSSProperties}
    >
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
              {page === "restock" && (
                <>
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
                </>
              )}
              <MessagesMenu messages={messages} />
              <div style={s.activeUsers} title="Currently active">
                <span style={s.activeUsersLabel}>Active</span>
                {activeUsers.length ? activeUsers.map((user) => (
                  <span key={user.id} style={s.activeUserBadge} title={user.name}>{user.initials}</span>
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
                      <option key={status} value={status}>{labelForOption(restockSettings.statusOptions, status)}</option>
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
                <button
                  type="button"
                  onClick={() => {
                    setShowCategoryDropdown((open) => !open);
                    setShowNewCategoryInput(false);
                    setNewCategoryInput("");
                  }}
                  style={s.addProductButton}
                >
                  + Add product
                </button>
                {showCategoryDropdown && (
                  <CategoryDropdown
                    productGroups={productGroups}
                    showNewCategoryInput={showNewCategoryInput}
                    newCategoryInput={newCategoryInput}
                    onNewCategoryInput={setNewCategoryInput}
                    onShowNewCategory={() => setShowNewCategoryInput(true)}
                    onSelect={(category) => {
                      setSelectedAddCategory(category);
                      setShowCategoryDropdown(false);
                      setShowNewCategoryInput(false);
                      setNewCategoryInput("");
                      setShowAddRow(true);
                    }}
                    onClose={() => {
                      setShowCategoryDropdown(false);
                      setShowNewCategoryInput(false);
                      setNewCategoryInput("");
                    }}
                  />
                )}
              </div>
            )}
          </div>
        </header>

        {page === "settings" ? (
          <div style={{ display: "grid", gap: 16 }}>
            <SettingsPanel
              users={users}
              currentUser={currentUser}
              loginRequired={loginRequired}
              restockSettings={restockSettings}
              universalSettings={universalSettings}
              fabricSettings={fabricSettings}
            />
            <ActivityLogPanel activityLogs={activityLogs} />
          </div>
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
        ) : page === "fabric" ? (
          <FabricSheetsPanel
            sheets={fabricSheets}
            selectedGid={selectedFabricTab}
            fabricSettings={fabricSettings}
            inrToAudRate={inrToAudRate}
            onSelect={(gid) => updateParams({ fabricTab: gid })}
          />
        ) : page !== "restock" ? (
          <div style={s.empty}>{activePageTitle} will be set up here.</div>
        ) : (
          <div style={s.tableWrap}>
            <table style={{ ...s.table, width: tableWidth }} onKeyDown={handleTableGridKeyDown}>
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
                {showAddRow && (
                  <AddRestockOrderRow
                    sizes={sizes}
                    initialProductGroup={selectedAddCategory ?? ""}
                    productSearch={restockProductSearch}
                    productResults={restockProductResults}
                    updateParams={updateParams}
                    restockSettings={restockSettings}
                    onClose={() => {
                      setShowAddRow(false);
                      setSelectedAddCategory(null);
                    }}
                  />
                )}
                {orders.map((order, rowIndex) => (
                  <OrderRow key={order.id} order={order} rowIndex={rowIndex + 1} sizes={sizes} users={users} restockSettings={restockSettings} />
                ))}
                {orders.length === 0 && (
                  <tr style={s.row}>
                    <td colSpan={columns.length} style={{ ...s.td, textAlign: "center", color: "#6b7280", fontWeight: 700 }}>
                      {messageOrderId ? "That message is for an order that is no longer open." : "No open orders yet. Click \"+ Add product\" to add one."}
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

function PortalLogin({ users }: { users: PortalUser[] }) {
  return (
    <div style={s.loginShell}>
      <form method="post" style={s.loginCard}>
        <input type="hidden" name="intent" value="portal_login" />
        <h1 style={s.loginTitle}>Production Portal</h1>
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
                href={`/portal?messageOrderId=${message.orderId}#order-${message.orderId}`}
                style={s.messageLink}
              >
                <strong>{message.productTitle || `Order #${message.orderId}`}</strong>
                <span>{message.field === "factory_notes" ? "Factory notes" : "Notes"}</span>
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

// ─── Fabric Sheets ───────────────────────────────────────────────────────────

function FabricSheetsPanel({
  sheets,
  selectedGid,
  fabricSettings,
  inrToAudRate,
  onSelect,
}: {
  sheets: FabricSheetData[];
  selectedGid: string;
  fabricSettings: FabricSettings;
  inrToAudRate: number | null;
  onSelect: (gid: string) => void;
}) {
  const fetcher = useFetcher();
  const selectedSheet = sheets.find((sheet) => sheet.gid === selectedGid) ?? null;
  const [orderedSheets, setOrderedSheets] = useState(sheets);
  const [dragGid, setDragGid] = useState<string | null>(null);
  const [fabricNameSearch, setFabricNameSearch] = useState("");
  useEffect(() => setOrderedSheets(sheets), [sheets]);
  const saveTileOrder = (nextSheets: FabricSheetData[]) => {
    submitPortalCell(fetcher, {
      intent: "update_fabric_settings",
      value: JSON.stringify({ ...fabricSettings, tileOrder: nextSheets.map((sheet) => sheet.gid) }),
    });
  };
  const totalSheets = orderedSheets.filter((sheet) => !isFabricTotalsExcluded(sheet.name));
  const totalQuantity = totalSheets.reduce((sum, sheet) => sum + (sheet.totalQuantity ?? 0), 0);
  const totalCostInr = totalSheets.reduce((sum, sheet) => sum + (sheet.totalCost ?? 0), 0);
  const totalCostAud = inrToAudRate && totalCostInr ? totalCostInr * inrToAudRate : null;
  const [showTotalsHelp, setShowTotalsHelp] = useState(false);

  if (selectedSheet) {
    return (
      <div style={s.fabricPage}>
        <div style={s.fabricToolbar}>
          <div style={s.fabricToolbarLeft}>
            <button type="button" onClick={() => onSelect("")} style={s.secondaryButton}>
              Back to fabric types
            </button>
            <label style={s.fabricSearchLabel}>
              Search fabric name
              <input
                type="search"
                value={fabricNameSearch}
                onChange={(event) => setFabricNameSearch(event.currentTarget.value)}
                placeholder="Search by name"
                style={s.fabricSearchInput}
              />
            </label>
          </div>
          <div style={s.fabricToolbarMeta}>
            <strong>{selectedSheet.name}</strong>
            <span>{selectedSheet.rowCount} rows</span>
            {selectedSheet.totalQuantity != null && <span>{formatFabricNumber(selectedSheet.totalQuantity)} total</span>}
            {selectedSheet.totalCost != null && <span>{formatCurrency(selectedSheet.totalCost)} fabric cost</span>}
          </div>
        </div>
        <FabricSheetTable
          sheet={selectedSheet}
          sheets={sheets}
          fetcher={fetcher}
          fabricSettings={fabricSettings}
          nameSearch={fabricNameSearch}
        />
      </div>
    );
  }

  return (
    <div style={s.fabricPage}>
      <div style={s.fabricTotalsBar}>
        <span
          style={s.fabricTotalsHelp}
          onMouseEnter={() => setShowTotalsHelp(true)}
          onMouseLeave={() => setShowTotalsHelp(false)}
          onFocus={() => setShowTotalsHelp(true)}
          onBlur={() => setShowTotalsHelp(false)}
          tabIndex={0}
        >
          !
          {showTotalsHelp && (
            <span style={s.fabricTotalsHelpBubble}>
              Totals do not include On Order, Fabric on Order, New Fabric on Order, Random Fabric Bits and Bobs, or Fabric Samples under consideration.
            </span>
          )}
        </span>
        <span style={s.fabricTotalsItem}>
          <strong>{formatFabricNumber(totalQuantity)}</strong>
          <span>Total fabric</span>
        </span>
        <span style={s.fabricTotalsItem}>
          <strong>{formatCurrency(totalCostInr)}</strong>
          <span>Total value</span>
        </span>
        <span style={s.fabricTotalsItem}>
          <strong>{totalCostAud == null ? "—" : formatAudCurrency(totalCostAud)}</strong>
          <span>AUD value</span>
        </span>
      </div>
      <div style={s.fabricGrid}>
        {orderedSheets.map((sheet) => (
          <button
            key={sheet.gid}
            type="button"
            draggable
            onDragStart={() => setDragGid(sheet.gid)}
            onDragOver={(event) => {
              event.preventDefault();
              if (!dragGid || dragGid === sheet.gid) return;
              const from = orderedSheets.findIndex((item) => item.gid === dragGid);
              const to = orderedSheets.findIndex((item) => item.gid === sheet.gid);
              if (from < 0 || to < 0) return;
              const next = [...orderedSheets];
              const [moved] = next.splice(from, 1);
              next.splice(to, 0, moved);
              setOrderedSheets(next);
              saveTileOrder(next);
            }}
            onDragEnd={() => setDragGid(null)}
            onClick={() => onSelect(sheet.gid)}
            style={{ ...s.fabricCard, ...(dragGid === sheet.gid ? s.fabricCardDragging : {}) }}
          >
            <span style={s.fabricCardHandle}>Drag</span>
            <span style={s.fabricCardTitle}>{sheet.name}</span>
            <span style={s.fabricCardMeta}>
              {sheet.error ? "Unable to load" : `${sheet.rowCount} rows`}
            </span>
            {!sheet.error && sheet.totalQuantity != null && (
              <span style={s.fabricCardQuantity}>{formatFabricNumber(sheet.totalQuantity)} total</span>
            )}
            {!sheet.error && sheet.totalCost != null && (
              <span style={s.fabricCardCost}>{formatCurrency(sheet.totalCost)} fabric cost</span>
            )}
          </button>
        ))}
      </div>
    </div>
  );
}

function FabricSheetTable({
  sheet,
  sheets,
  fetcher,
  fabricSettings,
  nameSearch,
}: {
  sheet: FabricSheetData;
  sheets: FabricSheetData[];
  fetcher: ReturnType<typeof useFetcher>;
  fabricSettings: FabricSettings;
  nameSearch: string;
}) {
  if (sheet.error) {
    return <div style={s.empty}>Unable to load {sheet.name}: {sheet.error}</div>;
  }
  if (!sheet.rows.length) {
    return <div style={s.empty}>No rows found for {sheet.name}.</div>;
  }

  const nameIndex = sheet.headers.findIndex((header) => /^name$/i.test(header.trim()));
  const searchText = nameSearch.trim().toLowerCase();
  const displayRows = sheet.rows
    .map((row, rowIndex) => ({ row, rowIndex, sourceRowIndex: sheet.rowKeys?.[rowIndex] ?? rowIndex }))
    .filter(({ row }) => !searchText || (nameIndex >= 0 && String(row[nameIndex] ?? "").toLowerCase().includes(searchText)));

  return (
    <div style={s.fabricTableShell}>
      <div style={s.fabricTableWrap}>
        <table style={s.fabricTable} onKeyDown={handleTableGridKeyDown}>
          <thead>
            <tr>
              {sheet.headers.map((header, index) => (
                <th key={`${header}-${index}`} style={s.fabricTh}>{header}</th>
              ))}
              <th style={s.fabricTh}>Actions</th>
            </tr>
          </thead>
          <tbody>
            {displayRows.map(({ row, rowIndex, sourceRowIndex }, displayRowIndex) => (
              <tr key={rowIndex} style={s.row}>
                {sheet.headers.map((_, colIndex) => (
                  <FabricTd key={colIndex} rowIndex={displayRowIndex} colIndex={colIndex}>
                    <FabricCell
                      gid={sheet.gid}
                      rowIndex={sourceRowIndex}
                      colIndex={colIndex}
                      value={row[colIndex] ?? ""}
                      originalValue={sheet.originalRows?.[sourceRowIndex]?.[colIndex] ?? ""}
                      header={sheet.headers[colIndex] ?? ""}
                      fetcher={fetcher}
                      fabricSettings={fabricSettings}
                    />
                  </FabricTd>
                ))}
                <FabricTd rowIndex={displayRowIndex} colIndex={sheet.headers.length}>
                  <FabricRowActions
                    gid={sheet.gid}
                    rowIndex={sourceRowIndex}
                    sheets={sheets}
                    fetcher={fetcher}
                  />
                </FabricTd>
              </tr>
            ))}
            {!displayRows.length && (
              <tr>
                <td colSpan={sheet.headers.length + 1} style={{ ...s.fabricTd, padding: 24, textAlign: "center" }}>
                  No fabric names match this search.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      <div style={s.fabricTableFooter}>
        <button
          type="button"
          style={s.loginButton}
          onClick={() => submitPortalCell(fetcher, { intent: "add_fabric_row", gid: sheet.gid })}
        >
          Add row
        </button>
      </div>
    </div>
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
  return /^data:image\//i.test(value.trim()) || /^(?:https?:\/\/|\/).+\.(png|jpe?g|webp|gif|avif)(\?.*)?$/i.test(value.trim());
}

function isNumericFabricCell(header: string, value: string) {
  const normalizedHeader = header.trim().toLowerCase();
  if (/cost|price|meter|quantity|received|products|^k$|^l$|additional/.test(normalizedHeader)) return true;
  const trimmed = value.trim();
  return Boolean(trimmed) && /^[-₹$]?\d[\d,]*(?:\.\d+)?$/.test(trimmed);
}

function FabricCell({
  gid,
  rowIndex,
  colIndex,
  value,
  originalValue,
  header,
  fetcher,
  fabricSettings,
}: {
  gid: string;
  rowIndex: number;
  colIndex: number;
  value: string;
  originalValue: string;
  header: string;
  fetcher: ReturnType<typeof useFetcher>;
  fabricSettings: FabricSettings;
}) {
  const [draft, setDraft] = useState(value);
  const [imageHover, setImageHover] = useState(false);
  useEffect(() => setDraft(value), [value]);
  const trimmed = draft.trim();
  const imageValue = isFabricImageValue(trimmed);
  const normalizedHeader = header.trim().toLowerCase();
  const imageColumn = /picture|image/i.test(header) || normalizedHeader === "fabric" || imageValue || isFabricImageValue(originalValue);
  const chipOptions = /^supplier$/i.test(header)
    ? fabricSettings.supplierOptions
    : /fabric\s*type/i.test(header)
      ? fabricSettings.fabricTypeOptions
      : null;
  const chipOption = chipOptions?.find((option) => option.label === draft || option.value === slugForOption(draft));
  const multiline = draft.includes("\n") || draft.length > 34;
  const centerValue = isNumericFabricCell(header, draft);
  const save = (nextValue: string) => {
    if (nextValue === value) return;
    submitPortalCell(fetcher, {
      intent: "update_fabric_cell",
      gid,
      rowIndex,
      colIndex,
      value: nextValue,
    });
  };
  const uploadImage = (file: File | null) => {
    if (!file) return;
    const formData = new FormData();
    formData.set("intent", "upload_fabric_image");
    formData.set("gid", gid);
    formData.set("rowIndex", String(rowIndex));
    formData.set("colIndex", String(colIndex));
    formData.set("image", file);
    fetcher.submit(formData, { method: "post", encType: "multipart/form-data" });
  };

  if (imageColumn) {
    return (
      <div style={s.fabricImageEditCell}>
        <div
          tabIndex={0}
          style={s.fabricImageDrop}
          onMouseEnter={() => setImageHover(true)}
          onMouseLeave={() => setImageHover(false)}
          onFocus={() => setImageHover(true)}
          onBlur={() => setImageHover(false)}
          onPaste={(event) => {
            const file = Array.from(event.clipboardData.files).find((item) => item.type.startsWith("image/"));
            if (file) uploadImage(file);
          }}
          title="Paste or upload image"
        >
          {imageValue ? <img src={trimmed} alt="" style={s.fabricSheetImage} /> : <span>Paste image</span>}
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
          <input
            type="file"
            accept="image/*"
            style={s.hiddenFileInput}
            onChange={(event) => {
              uploadImage(event.currentTarget.files?.[0] ?? null);
              event.currentTarget.value = "";
            }}
          />
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
      </div>
    );
  }

  if (chipOptions) {
    const options = chipOptions.some((option) => option.label === draft || option.value === slugForOption(draft))
      ? chipOptions
      : draft
        ? [...chipOptions, { value: slugForOption(draft) || draft, label: draft, bg: "#f3f4f6", color: "#374151" }]
        : chipOptions;
    return (
      <select
        value={draft}
        onChange={(event) => {
          setDraft(event.currentTarget.value);
          save(event.currentTarget.value);
        }}
        style={{
          ...s.fabricChipSelect,
          background: chipOption?.bg ?? "#f3f4f6",
          color: chipOption?.color ?? "#374151",
        }}
      >
        <option value="">—</option>
        {options.map((option) => (
          <option key={option.value} value={option.label}>{option.label}</option>
        ))}
      </select>
    );
  }

  const common = {
    value: draft,
    onChange: (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => setDraft(event.currentTarget.value),
    onBlur: (event: React.FocusEvent<HTMLInputElement | HTMLTextAreaElement>) => save(event.currentTarget.value),
    style: {
      ...(multiline ? s.fabricCellTextarea : s.fabricCellInput),
      ...(centerValue ? s.fabricNumberCellInput : {}),
    },
    placeholder: "—",
  };

  return (
    <>
      {multiline ? <textarea rows={3} {...common} /> : <input type="text" {...common} />}
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

function SettingsPanel({
  users,
  currentUser,
  loginRequired,
  restockSettings,
  universalSettings,
  fabricSettings,
}: {
  users: PortalUser[];
  currentUser: PortalUser | null;
  loginRequired: boolean;
  restockSettings: RestockSettings;
  universalSettings: UniversalSettings;
  fabricSettings: FabricSettings;
}) {
  const settingsFetcher = useFetcher();
  const canManageUsers = !loginRequired || users.length === 0 || currentUser?.admin;
  const [restockDraft, setRestockDraft] = useState<RestockSettings>(restockSettings);
  const [universalDraft, setUniversalDraft] = useState<UniversalSettings>(universalSettings);
  const [fabricDraft, setFabricDraft] = useState<FabricSettings>(fabricSettings);
  const updateRestockOption = (kind: "statusOptions" | "priorityOptions", index: number, patch: Partial<RestockOption>) => {
    setRestockDraft((current) => ({
      ...current,
      [kind]: current[kind].map((option, optionIndex) => {
        if (optionIndex !== index) return option;
        const label = patch.label ?? option.label;
        return {
          ...option,
          ...patch,
          label,
          value: patch.label && option.value.startsWith("new_") ? slugForOption(patch.label) || option.value : option.value,
        };
      }),
    }));
  };
  const addRestockOption = (kind: "statusOptions" | "priorityOptions") => {
    setRestockDraft((current) => ({
      ...current,
      [kind]: [
        ...current[kind],
        {
          value: `new_${Date.now()}`,
          label: "New option",
          bg: kind === "statusOptions" ? "#f3f4f6" : "#111827",
          color: kind === "statusOptions" ? "#374151" : "#ffffff",
        },
      ],
    }));
  };
  const removeRestockOption = (kind: "statusOptions" | "priorityOptions", index: number) => {
    setRestockDraft((current) => ({
      ...current,
      [kind]: current[kind].filter((_, optionIndex) => optionIndex !== index),
    }));
  };
  const updateFabricOption = (kind: "supplierOptions" | "fabricTypeOptions", index: number, patch: Partial<RestockOption>) => {
    setFabricDraft((current) => ({
      ...current,
      [kind]: current[kind].map((option, optionIndex) => {
        if (optionIndex !== index) return option;
        const label = patch.label ?? option.label;
        return {
          ...option,
          ...patch,
          label,
          value: patch.label && option.value.startsWith("new_") ? slugForOption(patch.label) || option.value : option.value,
        };
      }),
    }));
  };
  const addFabricOption = (kind: "supplierOptions" | "fabricTypeOptions") => {
    setFabricDraft((current) => ({
      ...current,
      [kind]: [
        ...current[kind],
        {
          value: `new_${Date.now()}`,
          label: "New option",
          bg: "#f3f4f6",
          color: "#374151",
        },
      ],
    }));
  };
  const removeFabricOption = (kind: "supplierOptions" | "fabricTypeOptions", index: number) => {
    setFabricDraft((current) => ({
      ...current,
      [kind]: current[kind].filter((_, optionIndex) => optionIndex !== index),
    }));
  };
  const saveRestockSettings = () => submitPortalCell(settingsFetcher, {
    intent: "update_restock_settings",
    value: JSON.stringify(restockDraft),
  });
  const saveUniversalSettings = () => submitPortalCell(settingsFetcher, {
    intent: "update_universal_settings",
    value: JSON.stringify(universalDraft),
  });
  const saveFabricSettings = () => submitPortalCell(settingsFetcher, {
    intent: "update_fabric_settings",
    value: JSON.stringify(fabricDraft),
  });

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

      <section style={s.settingsCard}>
        <div style={s.settingsHeader}>
          <div>
            <h2 style={s.settingsTitle}>Universal Settings</h2>
            <p style={s.settingsHint}>Shared button, table text, and heading styling across restock and packing pages.</p>
          </div>
          <button type="button" disabled={!canManageUsers} style={s.loginButton} onClick={saveUniversalSettings}>
            Save universal settings
          </button>
        </div>

        {!canManageUsers && (
          <div style={s.settingsWarning}>Only an admin user can change universal settings.</div>
        )}

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
                onChange={(event) => setUniversalDraft((current) => ({
                  ...current,
                  tableTextSize: Number(event.currentTarget.value) || current.tableTextSize,
                }))}
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
                onChange={(event) => setUniversalDraft((current) => ({
                  ...current,
                  headingTextSize: Number(event.currentTarget.value) || current.headingTextSize,
                }))}
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
            <h2 style={s.settingsTitle}>Fabric in Stock Settings</h2>
            <p style={s.settingsHint}>Edit supplier and fabric type chips used on the fabric pages.</p>
          </div>
          <button type="button" disabled={!canManageUsers} style={s.loginButton} onClick={saveFabricSettings}>
            Save fabric settings
          </button>
        </div>

        {!canManageUsers && (
          <div style={s.settingsWarning}>Only an admin user can change fabric settings.</div>
        )}

        <div style={s.settingsGrid}>
          <RestockOptionsEditor
            title="Supplier chips"
            options={fabricDraft.supplierOptions}
            disabled={!canManageUsers}
            onChange={(index, patch) => updateFabricOption("supplierOptions", index, patch)}
            onAdd={() => addFabricOption("supplierOptions")}
            onRemove={(index) => removeFabricOption("supplierOptions", index)}
          />
          <RestockOptionsEditor
            title="Fabric type chips"
            options={fabricDraft.fabricTypeOptions}
            disabled={!canManageUsers}
            onChange={(index, patch) => updateFabricOption("fabricTypeOptions", index, patch)}
            onAdd={() => addFabricOption("fabricTypeOptions")}
            onRemove={(index) => removeFabricOption("fabricTypeOptions", index)}
          />
        </div>
      </section>

      <section style={s.settingsCard}>
        <div style={s.settingsHeader}>
          <div>
            <h2 style={s.settingsTitle}>Existing Product Restock Settings</h2>
            <p style={s.settingsHint}>Edit dropdown options and table number styling for the restock page.</p>
          </div>
          <button type="button" disabled={!canManageUsers} style={s.loginButton} onClick={saveRestockSettings}>
            Save restock settings
          </button>
        </div>

        {!canManageUsers && (
          <div style={s.settingsWarning}>Only an admin user can change restock settings.</div>
        )}

        <div style={s.settingsGrid}>
          <RestockOptionsEditor
            title="Status options"
            options={restockDraft.statusOptions}
            disabled={!canManageUsers}
            onChange={(index, patch) => updateRestockOption("statusOptions", index, patch)}
            onAdd={() => addRestockOption("statusOptions")}
            onRemove={(index) => removeRestockOption("statusOptions", index)}
          />
          <RestockOptionsEditor
            title="Priority options"
            options={restockDraft.priorityOptions}
            disabled={!canManageUsers}
            onChange={(index, patch) => updateRestockOption("priorityOptions", index, patch)}
            onAdd={() => addRestockOption("priorityOptions")}
            onRemove={(index) => removeRestockOption("priorityOptions", index)}
          />
        </div>

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
                onChange={(event) => setRestockDraft((current) => ({
                  ...current,
                  quantityFontSize: Number(event.currentTarget.value) || current.quantityFontSize,
                }))}
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

function RestockOptionsEditor({
  title,
  options,
  disabled,
  onChange,
  onAdd,
  onRemove,
}: {
  title: string;
  options: RestockOption[];
  disabled: boolean;
  onChange: (index: number, patch: Partial<RestockOption>) => void;
  onAdd: () => void;
  onRemove: (index: number) => void;
}) {
  return (
    <div style={s.settingsSubCard}>
      <div style={s.settingsSubHeader}>
        <h3 style={s.settingsSubTitle}>{title}</h3>
        <button type="button" disabled={disabled} style={s.smallButton} onClick={onAdd}>Add option</button>
      </div>
      <div style={s.optionRows}>
        {options.map((option, index) => (
          <div key={`${option.value}-${index}`} style={s.optionRow}>
            <input
              value={option.label}
              disabled={disabled}
              onChange={(event) => onChange(index, { label: event.currentTarget.value })}
              style={s.optionLabelInput}
            />
            <ColorPickerInput value={option.bg} disabled={disabled}
              onChange={(hex) => onChange(index, { bg: hex })} />
            <ColorPickerInput value={option.color} disabled={disabled}
              onChange={(hex) => onChange(index, { color: hex })} />
            <span style={{ ...s.optionPreview, background: option.bg, color: option.color }}>{option.label || "Option"}</span>
            <button type="button" disabled={disabled || options.length <= 1} style={s.removeUserButton} onClick={() => onRemove(index)}>
              Remove
            </button>
          </div>
        ))}
      </div>
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
  const [searchParams] = useSearchParams();
  const [hoveredListId, setHoveredListId] = useState<number | null>(null);
  const [deleteWarningList, setDeleteWarningList] = useState<PackingListWithLines | null>(null);
  const showHidden = searchParams.get("showHidden") === "true";
  const visibleLists = packingLists.filter((list) => !list.hiddenAt);
  const hiddenLists = packingLists.filter((list) => list.hiddenAt);
  const rows = showHidden ? hiddenLists : visibleLists;
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
              {["Invoice", "Total qty", "Estimated arrival", "Status", "Actions"].map((heading) => (
                <th key={heading} style={{ ...s.th, textAlign: heading === "Total qty" || heading === "Actions" ? "center" : "left" }}>
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
                  <td style={{ ...cellStyle, textAlign: "center" }}><span style={s.total}>{packingListTotal(list)}</span></td>
                  <td
                    style={cellStyle}
                    onContextMenu={(e) => { e.preventDefault(); e.stopPropagation(); document.dispatchEvent(new CustomEvent("show-cell-history", { detail: { x: e.clientX, y: e.clientY, entity: "Packing List", entityId: String(list.id), field: "Estimated arrival", entityName: list.invoiceNumber || `Packing list #${list.id}` } })); }}
                  >{formatPortalDate(list.expectedLeaveFactoryDate ?? list.shipmentDate) || "—"}</td>
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
                <td colSpan={5} style={{ ...s.td, textAlign: "center", padding: 40 }}>
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
  const loadInventoryFetcher = useFetcher();
  const columnWidthsFetcher = useFetcher();
  const [packingColumnWidths, setPackingColumnWidths] = useState<Record<string, number>>(savedPackingColumnWidths);
  const [skipWords, setSkipWords] = useState("");
  const [packingListSearch, setPackingListSearch] = useState("");
  const packingWidthFor = (columnId: string) => packingColumnWidths[columnId] ?? defaultPackingColumnWidth(columnId);
  const packingTableWidth = PACKING_COLUMNS.reduce((sum, column) => sum + packingWidthFor(column.id), 0);
  const normalizedPackingListSearch = packingListSearch.trim().toLowerCase();
  const visiblePackingLines = normalizedPackingListSearch
    ? packingList.lines.filter((line) => packingLineMatchesSearch(line, normalizedPackingListSearch))
    : packingList.lines;
  const exportPackingList = () => {
    const headers = [
      "Box",
      "Product image URL",
      "Fabric image",
      "Name",
      "SKU",
      ...PACKING_SIZES,
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
        line.productImageUrl ?? "",
        line.fabricImageData ? "Image added" : "",
        line.productTitle,
        line.sku ?? "",
        ...PACKING_SIZES.map((size) => qtys[size] || ""),
        total || "",
        line.priceRupees ?? "",
        total && price ? Math.round(total * price) : "",
        line.weight ?? "",
      ];
    });
    const csv = [headers, ...rows].map((row) => row.map(csvCell).join(",")).join("\n");
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
        {/* Row 1: navigation + core fields */}
        <div style={s.packingTopRow}>
          <a href="/portal?page=packing" style={s.secondaryButton}>Back</a>
          <button type="button" style={s.secondaryButton} onClick={exportPackingList}>Export packing list</button>
          <label style={s.packingToolbarLabel}>
            <span>Invoice number</span>
            <input
              defaultValue={packingList.invoiceNumber ?? ""}
              onBlur={(event) => submitPortalCell(fetcher, {
                intent: "update_packing_list",
                packingId: packingList.id,
                field: "invoiceNumber",
                value: event.currentTarget.value,
              })}
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
              onBlur={(event) => submitPortalCell(fetcher, {
                intent: "update_packing_list",
                packingId: packingList.id,
                field: "expectedLeaveFactoryDate",
                value: event.currentTarget.value,
              })}
              style={{ ...s.packingInput, width: 150 }}
            />
          </label>
          <label style={s.packingToolbarLabel}>
            <span>Status</span>
            <select
              key={packingList.status ?? "still_packing"}
              defaultValue={packingList.status ?? "still_packing"}
              onChange={(event) => submitPortalCell(fetcher, {
                intent: "update_packing_list",
                packingId: packingList.id,
                field: "status",
                value: event.currentTarget.value,
              })}
              style={{ ...s.packingInput, width: 160 }}
            >
              {PACKING_STATUS_OPTIONS.map((opt) => (
                <option key={opt.value} value={opt.value}>{opt.label}</option>
              ))}
            </select>
          </label>
        </div>
        {/* Row 2: load inventory (left) + total quantity (right) */}
        <div style={s.packingBottomRow}>
          <loadInventoryFetcher.Form
            method="post"
            style={s.loadInventoryForm}
            onSubmit={(event) => {
              const ok = window.confirm("Add these packing list quantities to current Shopify stock?");
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
          <div style={s.packingTotalPill}>
            Total quantity <strong>{packingListTotal(packingList)}</strong>
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

      <div style={s.packingTableWrap}>
        <table style={{ ...s.table, width: packingTableWidth, minWidth: "100%" }} onKeyDown={handleTableGridKeyDown}>
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
            {visiblePackingLines.length ? visiblePackingLines.map((line, rowIndex) => (
              <PackingListLineRow
                key={line.id}
                line={line}
                rowIndex={rowIndex}
                activeSearchLineId={packingSearchLineId}
                productSearch={productSearch}
                productResults={productResults}
                updateParams={updateParams}
              />
            )) : (
              <tr style={s.row}>
                <td colSpan={PACKING_COLUMNS.length} style={{ ...s.td, textAlign: "center", padding: 40 }}>
                  No packing list rows match this search.
                </td>
              </tr>
            )}
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
  rowIndex,
  activeSearchLineId,
  productSearch,
  productResults,
  updateParams,
}: {
  line: PackingListWithLines["lines"][number];
  rowIndex: number;
  activeSearchLineId: number | null;
  productSearch: string;
  productResults: ShopifySearchProduct[];
  updateParams: (updates: Record<string, string>) => void;
}) {
  const fetcher = useFetcher();
  const qtys = normalizeQtys(line.qtys);
  const shopifyLoadedQtys = normalizeQtys(line.shopifyLoadedQtys);
  const total = packingTotal(qtys);
  const price = line.priceRupees ?? 0;
  const value = total * price;

  return (
    <tr style={s.row}>
      <PackingTd rowIndex={rowIndex} colIndex={0}><PackingTextInput lineId={line.id} field="boxNumber" value={line.boxNumber ?? ""} /></PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={1} center><PackingImageCell lineId={line.id} field="productImageUrl" value={line.productImageUrl ?? ""} /></PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={2} center><PackingImageCell lineId={line.id} field="fabricImageData" value={line.fabricImageData ?? ""} /></PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={3} overflowVisible>
        <PackingProductNameCell
          line={line}
          isActiveSearch={activeSearchLineId === line.id}
          productSearch={productSearch}
          productResults={productResults}
          updateParams={updateParams}
        />
      </PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={4}><PackingSkuCell lineId={line.id} value={line.sku ?? ""} /></PackingTd>
      {PACKING_SIZES.map((size, sizeIndex) => (
        <PackingTd
          key={size}
          rowIndex={rowIndex}
          colIndex={5 + sizeIndex}
          center
          style={{
            ...(qtys[size] > 0 && shopifyLoadedQtys[size] === qtys[size] ? s.loadedInventoryCell : {}),
          }}
          onContextMenu={(e) => { e.preventDefault(); document.dispatchEvent(new CustomEvent("show-cell-history", { detail: { x: e.clientX, y: e.clientY, entity: "Packing List Line", entityId: String(line.id), field: `Qty (${size})`, entityName: line.productTitle } })); }}
        >
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
            style={{ ...s.qtyInput, ...(qtys[size] > 0 && shopifyLoadedQtys[size] === qtys[size] ? { color: "#fff" } : {}) }}
          />
        </PackingTd>
      ))}
      <PackingTd rowIndex={rowIndex} colIndex={15} center><span style={s.total}>{total}</span></PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={16} center><PackingTextInput lineId={line.id} field="priceRupees" value={line.priceRupees?.toString() ?? ""} center /></PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={17} center><span style={s.total}>{value ? Math.round(value) : ""}</span></PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={18} center><PackingTextInput lineId={line.id} field="weight" value={line.weight?.toString() ?? ""} center /></PackingTd>
      <PackingTd rowIndex={rowIndex} colIndex={19} center>
        <div style={s.rowActions}>
          <button type="button" style={s.smallButton} onClick={() => submitPortalCell(fetcher, { intent: "duplicate_packing_line", lineId: line.id })}>Duplicate</button>
          <button type="button" style={s.removeUserButton} onClick={() => submitPortalCell(fetcher, { intent: "delete_packing_line", lineId: line.id })}>Delete</button>
        </div>
      </PackingTd>
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
  const [isChangingProduct, setIsChangingProduct] = useState(false);
  const [dropdownRect, setDropdownRect] = useState<DOMRect | null>(null);
  const [canPortalDropdown, setCanPortalDropdown] = useState(false);
  const hasLinkedProduct = Boolean(line.productId);
  const canSearch = !isProductSelected && (!hasLinkedProduct || isChangingProduct) && (isFocused || isActiveSearch);
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
    setIsChangingProduct(false);
    setDropdownRect(null);
    inputRef.current?.blur();
    submitPortalCell(fetcher, {
      intent: "apply_product_to_packing_line",
      lineId: line.id,
      product: JSON.stringify(product),
    });
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
          if (value.trim().length >= 2) {
            updateParams({ productSearch: value.trim(), packingSearchLineId: String(line.id) });
          }
        }}
        onChange={(event) => setValue(event.currentTarget.value)}
        onBlur={(event) => {
          if (isProductSelected) return;
          setIsFocused(false);
          if (hasLinkedProduct) {
            setValue(displayValue);
            setIsChangingProduct(false);
            updateParams({ productSearch: "", packingSearchLineId: "" });
            return;
          }
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

function PackingImageCell({ lineId, field, value }: { lineId: number; field: "productImageUrl" | "fabricImageData"; value: string }) {
  const fetcher = useFetcher();
  const [imageHover, setImageHover] = useState(false);

  const saveFile = (file: File) => {
    const reader = new FileReader();
    reader.onload = () => submitPortalCell(fetcher, {
      intent: "update_packing_line",
      lineId,
      field,
      value: String(reader.result ?? ""),
    });
    reader.readAsDataURL(file);
  };
  const clearImage = () => submitPortalCell(fetcher, {
    intent: "update_packing_line",
    lineId,
    field,
    value: "",
  });

  return (
    <div style={s.packingImageCell}>
      <div
        tabIndex={0}
        style={s.fabricImageDrop}
        onMouseEnter={() => setImageHover(true)}
        onMouseLeave={() => setImageHover(false)}
        onFocus={() => setImageHover(true)}
        onBlur={() => setImageHover(false)}
        onPaste={(event) => {
          const file = Array.from(event.clipboardData.files).find((item) => item.type.startsWith("image/"));
          if (file) saveFile(file);
        }}
        title="Paste or upload image"
      >
        {value ? <img src={value} alt="" style={s.fabricThumb} /> : <span>Paste image</span>}
        {value && (
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
        <input
          type="file"
          accept="image/*"
          style={s.hiddenFileInput}
          onChange={(event) => {
            const file = event.currentTarget.files?.[0];
            if (file) saveFile(file);
            event.currentTarget.value = "";
          }}
        />
      </div>
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
  sizes,
  initialProductGroup,
  productSearch,
  productResults,
  updateParams,
  restockSettings,
  onClose,
}: {
  sizes: string[];
  initialProductGroup?: string;
  productSearch: string;
  productResults: ShopifySearchProduct[];
  updateParams: (updates: Record<string, string>) => void;
  restockSettings: RestockSettings;
  onClose?: () => void;
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
  const deleteCol = totalCol + 5;

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
      } else {
        onClose?.();
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
    <tr ref={rowRef} style={{ ...s.row, ...s.addOrderRow, outline: "2px solid #008060", outlineOffset: -1 }}>
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
              autoFocus
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
      <Td rowIndex={0} colIndex={deleteCol} center>
        {onClose && <button type="button" onClick={onClose} style={s.deleteButton}>✕</button>}
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
}: {
  order: Order;
  rowIndex: number;
  sizes: string[];
  users: PortalUser[];
  restockSettings: RestockSettings;
}) {
  const [inventoryOpen, setInventoryOpen] = useState(false);
  const qtyBySize = order.lines.reduce<Record<string, number>>((acc, line) => {
    acc[line.variantTitle] = (acc[line.variantTitle] ?? 0) + line.qtyOrdered;
    return acc;
  }, {});
  const inventoryBySize = order.lines.reduce<Record<string, number | null>>((acc, line) => {
    const availableInventory = "availableInventory" in line ? line.availableInventory : null;
    acc[line.variantTitle] = Number.isFinite(Number(availableInventory)) ? Number(availableInventory) : null;
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
  const deleteCol = totalCol + 5;

  return (
    <>
      <tr id={`order-${order.id}`} style={s.row}>
        {/* Factory notes */}
        <Td rowIndex={rowIndex} colIndex={0} overflowVisible historyEntity="Restock Order" historyEntityId={String(order.id)} historyField="Factory notes" historyEntityName={order.productTitle}><NotesCell orderId={order.id} field="factory_notes" value={order.factoryNotes ?? ""} users={users} /></Td>

        {/* Order date */}
        <Td rowIndex={rowIndex} colIndex={1} center><span style={s.dateText}>{orderDate}</span></Td>

        {/* Picture */}
        <Td rowIndex={rowIndex} colIndex={2} center historyEntity="Restock Order" historyEntityId={String(order.id)} historyField="Product image" historyEntityName={order.productTitle}>
          <div style={s.imageCell}>
            {order.productImageUrl
              ? <img src={order.productImageUrl} alt="" style={s.thumb} />
              : <div style={s.noImg}>—</div>}
          </div>
        </Td>

        {/* Name */}
        <Td rowIndex={rowIndex} colIndex={3} historyEntity="Restock Order" historyEntityId={String(order.id)} historyField="Product name" historyEntityName={order.productTitle}><span style={s.productName}>{order.productTitle}</span></Td>

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
        <Td rowIndex={rowIndex} colIndex={statusCol} historyEntity="Restock Order" historyEntityId={String(order.id)} historyField="Status" historyEntityName={order.productTitle}><StatusCell orderId={order.id} value={order.supplierStatus} options={restockSettings.statusOptions} /></Td>

        {/* Notes (from order) */}
        <Td rowIndex={rowIndex} colIndex={notesCol} overflowVisible historyEntity="Restock Order" historyEntityId={String(order.id)} historyField="Notes" historyEntityName={order.productTitle}><NotesCell orderId={order.id} field="notes" value={order.notes ?? ""} users={users} /></Td>

        {/* Priority */}
        <Td rowIndex={rowIndex} colIndex={priorityCol} historyEntity="Restock Order" historyEntityId={String(order.id)} historyField="Priority" historyEntityName={order.productTitle}><PriorityCell orderId={order.id} value={order.priority ?? ""} options={restockSettings.priorityOptions} /></Td>

        {/* ETA */}
        <Td rowIndex={rowIndex} colIndex={etaCol} historyEntity="Restock Order" historyEntityId={String(order.id)} historyField="ETA" historyEntityName={order.productTitle}><EtaCell orderId={order.id} value={etaValue} /></Td>

        {/* Actions */}
        <Td rowIndex={rowIndex} colIndex={deleteCol} center><OrderActionsCell orderId={order.id} /></Td>
      </tr>
      {inventoryOpen && (
        <tr style={s.inventoryRow}>
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
          <td style={s.inventoryBlankCell} />
        </tr>
      )}
    </>
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

function StatusCell({ orderId, value, options }: { orderId: number; value: string; options: RestockOption[] }) {
  const fetcher = useFetcher();
  const current = fetcher.formData ? String(fetcher.formData.get("value")) : value;
  const option = options.find((item) => item.value === current);

  return (
    <select
      value={current}
      onChange={(e) => submitPortalCell(fetcher, {
        intent: "update_status",
        orderId,
        value: e.currentTarget.value,
      })}
      style={{ ...s.select, background: option?.bg ?? "#f3f4f6", color: option?.color ?? "#374151" }}
    >
      {options.map((o) => (
        <option key={o.value} value={o.value}>{o.label}</option>
      ))}
    </select>
  );
}

function PriorityCell({ orderId, value, options }: { orderId: number; value: string; options: RestockOption[] }) {
  const fetcher = useFetcher();
  const current = fetcher.formData ? String(fetcher.formData.get("value")) : value;
  const opt = options.find((o) => o.value === current);

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
      {options.map((o) => (
        <option key={o.value} value={o.value}>{o.label}</option>
      ))}
    </select>
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
          submitPortalCell(fetcher, {
            intent: `update_${field}`,
            orderId,
            value: e.currentTarget.value,
          });
        }}
        rows={2}
        style={s.textarea}
        placeholder="Add note… use @name"
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
      onBlur={(e) => submitPortalCell(fetcher, {
        intent: "update_qty",
        orderId,
        size,
        value: e.currentTarget.value,
      })}
      style={{
        ...s.qtyInput,
        ...(numericCurrent > 0 ? s.qtyInputActive : s.qtyInputZero),
        fontSize: restockSettings.quantityFontSize,
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
  overflowVisible,
  historyEntity,
  historyEntityId,
  historyField,
  historyEntityName,
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
}) {
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
      style={{ ...s.td, textAlign: center ? "center" : "left", ...(overflowVisible ? { overflow: "visible" } : {}) }}
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
}: {
  children: React.ReactNode;
  center?: boolean;
  rowIndex: number;
  colIndex: number;
  overflowVisible?: boolean;
  style?: React.CSSProperties;
  onContextMenu?: (e: React.MouseEvent<HTMLTableCellElement>) => void;
}) {
  return (
    <td
      data-grid-row={rowIndex}
      data-grid-col={colIndex}
      tabIndex={0}
      onContextMenu={onContextMenu}
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
  loginSelect: {
    border: "1px solid #b6c0cc",
    borderRadius: 8,
    padding: "10px 12px",
    fontSize: 15,
    fontWeight: 700,
    background: "#fff",
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
  settingsGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: 14,
    marginTop: 16,
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
  settingsSubHeader: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 },
  settingsSubTitle: { margin: 0, fontSize: 14, color: "#111827" },
  optionRows: { display: "grid", gap: 8 },
  optionRow: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    overflow: "hidden",
  },
  optionLabelInput: {
    flex: 1,
    minWidth: 0,
    border: "1px solid #cbd5e1",
    borderRadius: 7,
    padding: "8px 9px",
    fontSize: 13,
    fontWeight: 700,
  },
  colorLabel: { display: "flex", alignItems: "center", gap: 5, fontSize: 11, fontWeight: 800, color: "#4b5563" },
  colorInput: { width: 36, height: 32, padding: 1, border: "1px solid #cbd5e1", borderRadius: 6, background: "#fff" },
  optionPreview: {
    borderRadius: 999,
    padding: "6px 9px",
    fontSize: 11,
    fontWeight: 900,
    whiteSpace: "nowrap",
  },
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
  fabricTable: { borderCollapse: "separate", borderSpacing: 0, minWidth: 960, width: "100%", fontSize: 13 },
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
    minWidth: 120,
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
    resize: "vertical",
    padding: "9px 10px",
    background: "transparent",
    color: "var(--portal-table-text-color, #1f2937)",
    font: "inherit",
    fontWeight: 700,
    lineHeight: 1.35,
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
    justifyContent: "flex-start",
    padding: "0 0 4px",
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
    width: "calc(100% - 20px)",
    minHeight: 34,
    margin: 10,
    border: "1px solid rgba(15,23,42,0.12)",
    borderRadius: 999,
    padding: "6px 28px 6px 10px",
    outline: "none",
    fontSize: "var(--portal-table-font-size, 13px)",
    fontWeight: 900,
    cursor: "pointer",
  },
  fabricCellActions: {
    display: "flex",
    gap: 6,
    flexWrap: "wrap",
    padding: "0 10px 8px",
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
    overflow: "auto",
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
    fontSize: 13,
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
};
