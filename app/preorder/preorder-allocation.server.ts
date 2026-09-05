import { Prisma } from "@prisma/client";
import prisma from "../db.server";
import { calculatePreorderCapacity, isPreorderEligibleStatus, type PreorderMarket } from "./preorder-rules.server";

// OrderLine.variantId is stored as a Shopify GID, but the order webhook sends a
// bare numeric variant_id. Match every form so the batch line is actually found;
// otherwise the reservation silently fails with "0 available" and the paid order
// never gets reserved.
function variantIdCandidates(value: string): string[] {
  const raw = String(value ?? "").trim();
  const numeric = raw.replace(/[^0-9]/g, "");
  return Array.from(new Set([
    raw,
    numeric,
    numeric ? `gid://shopify/ProductVariant/${numeric}` : "",
  ].filter(Boolean)));
}

export class PreorderCapacityError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "PreorderCapacityError";
  }
}

export type ReservePreorderLineInput = {
  shop: string;
  shopifyOrderId: string;
  shopifyOrderName?: string | null;
  shopifyLineItemId: string;
  productId?: string | null;
  variantId: string;
  variantTitle?: string | null;
  sku?: string | null;
  market: PreorderMarket;
  quantity: number;
  customerEmail?: string | null;
  preferredSupplierOrderId?: number | null;
};

function destinationForMarket(market: PreorderMarket) {
  return market === "USA" ? "send_to_usa" : "send_to_au";
}

function timestamp(value: Date | null | undefined) {
  return value?.getTime() ?? Number.MAX_SAFE_INTEGER;
}

function isSerializableConflict(error: unknown) {
  return error instanceof Prisma.PrismaClientKnownRequestError && error.code === "P2034";
}

export async function reservePreorderLine(input: ReservePreorderLineInput) {
  if (!input.shop.trim() || !input.shopifyOrderId.trim() || !input.shopifyLineItemId.trim() || !input.variantId.trim()) {
    throw new PreorderCapacityError("Missing preorder order or variant identifiers.");
  }
  if (!Number.isInteger(input.quantity) || input.quantity <= 0) {
    throw new PreorderCapacityError("Preorder quantity must be a positive whole number.");
  }
  if (input.preferredSupplierOrderId != null && (!Number.isInteger(input.preferredSupplierOrderId) || input.preferredSupplierOrderId <= 0)) {
    throw new PreorderCapacityError("Invalid preorder production batch reference.");
  }

  for (let attempt = 1; attempt <= 3; attempt += 1) {
    try {
      return await prisma.$transaction(async (tx) => {
        const existing = await tx.preorderReservation.findMany({
          where: {
            shop: input.shop,
            shopifyLineItemId: input.shopifyLineItemId,
            status: "reserved",
          },
          orderBy: { id: "asc" },
        });
        const existingQty = existing.reduce((sum, row) => sum + row.quantity, 0);
        if (existingQty === input.quantity) return existing;
        if (existingQty > 0) {
          throw new PreorderCapacityError(
            `This Shopify line already has ${existingQty} units reserved. Release or adjust that allocation before reserving ${input.quantity}.`,
          );
        }

        const enabledSettings = await tx.preorderBatchSetting.findMany({
          where: {
            shop: input.shop,
            enabled: true,
            ...(input.preferredSupplierOrderId ? { supplierOrderId: input.preferredSupplierOrderId } : {}),
          },
          select: {
            supplierOrderId: true,
            safetyBufferPercent: true,
            safetyBufferQty: true,
            shipDate: true,
          },
        });
        if (!enabledSettings.length) {
          if (input.preferredSupplierOrderId) {
            throw new PreorderCapacityError(`Preorder batch #${input.preferredSupplierOrderId} is no longer enabled for new reservations.`);
          }
          throw new PreorderCapacityError(`No active ${input.market} preorder production batches are available.`);
        }
        const settingByBatch = new Map(enabledSettings.map((setting) => [setting.supplierOrderId, setting]));

        const variantMatch = { in: variantIdCandidates(input.variantId) };
        const orders = await tx.supplierOrder.findMany({
          where: {
            id: { in: enabledSettings.map((setting) => setting.supplierOrderId) },
            shop: input.shop,
            status: "open",
            // Any production status qualifies (isPreorderEligibleStatus filters in
            // code below) — not just on_production — matching the storefront rule.
            destination: destinationForMarket(input.market),
            lines: { some: { variantId: variantMatch } },
          },
          select: {
            id: true,
            productId: true,
            productTitle: true,
            supplierStatus: true,
            eta: true,
            createdAt: true,
            lines: {
              where: { variantId: variantMatch },
              select: {
                variantId: true,
                variantTitle: true,
                sku: true,
                qtyOrdered: true,
                qtyReceived: true,
              },
            },
          },
        });

        const orderedBatches = orders
          .filter((order) => isPreorderEligibleStatus(order.supplierStatus))
          .filter((order) => order.lines.length > 0)
          .sort((a, b) => {
            if (input.preferredSupplierOrderId) return a.id - b.id;
            const aSetting = settingByBatch.get(a.id);
            const bSetting = settingByBatch.get(b.id);
            return (
              timestamp(aSetting?.shipDate ?? a.eta) - timestamp(bSetting?.shipDate ?? b.eta) ||
              a.createdAt.getTime() - b.createdAt.getTime() ||
              a.id - b.id
            );
          });

        if (!orderedBatches.length) {
          if (input.preferredSupplierOrderId) {
            throw new PreorderCapacityError(
              `Preorder batch #${input.preferredSupplierOrderId} is not eligible for this ${input.market} variant. The order has not been moved to a later batch automatically.`,
            );
          }
          throw new PreorderCapacityError(`No eligible ${input.market} production batch contains this variant.`);
        }

        const reservedGroups = await tx.preorderReservation.groupBy({
          by: ["supplierOrderId"],
          where: {
            supplierOrderId: { in: orderedBatches.map((order) => order.id) },
            variantId: variantMatch,
            status: "reserved",
          },
          _sum: { quantity: true },
        });
        const reservedByBatch = new Map(reservedGroups.map((row) => [row.supplierOrderId, row._sum.quantity ?? 0]));

        let remaining = input.quantity;
        const allocations: Array<{
          supplierOrderId: number;
          quantity: number;
          expectedShipDate: Date | null;
          productId: string;
          variantTitle: string | null;
          sku: string | null;
        }> = [];

        for (const order of orderedBatches) {
          if (remaining <= 0) break;
          const line = order.lines[0];
          const setting = settingByBatch.get(order.id);
          if (!line || !setting) continue;

          const incomingRemaining = Math.max(0, line.qtyOrdered - line.qtyReceived);
          const capacity = calculatePreorderCapacity({
            confirmedIncomingQty: incomingRemaining,
            reservedQty: reservedByBatch.get(order.id) ?? 0,
            safetyBufferPercent: setting.safetyBufferPercent,
            safetyBufferQty: setting.safetyBufferQty,
          });
          const take = Math.min(remaining, capacity.availableToPreorder);
          if (take <= 0) continue;

          allocations.push({
            supplierOrderId: order.id,
            quantity: take,
            expectedShipDate: setting.shipDate ?? order.eta ?? null,
            productId: order.productId,
            variantTitle: input.variantTitle ?? line.variantTitle ?? null,
            sku: input.sku ?? line.sku ?? null,
          });
          remaining -= take;
        }

        if (remaining > 0) {
          const available = input.quantity - remaining;
          if (input.preferredSupplierOrderId) {
            throw new PreorderCapacityError(
              `Preorder batch #${input.preferredSupplierOrderId} only has ${available} unit${available === 1 ? "" : "s"} available for this variant; ${input.quantity} requested. The order has not been moved to a later batch automatically.`,
            );
          }
          throw new PreorderCapacityError(
            `Only ${available} unit${available === 1 ? " is" : "s are"} available to preorder for this ${input.market} variant; ${input.quantity} requested.`,
          );
        }

        const created = [];
        for (const allocation of allocations) {
          created.push(await tx.preorderReservation.create({
            data: {
              shop: input.shop,
              shopifyOrderId: input.shopifyOrderId,
              shopifyOrderName: input.shopifyOrderName?.trim() || null,
              shopifyLineItemId: input.shopifyLineItemId,
              supplierOrderId: allocation.supplierOrderId,
              productId: input.productId ?? allocation.productId,
              variantId: input.variantId,
              variantTitle: allocation.variantTitle,
              sku: allocation.sku,
              market: input.market,
              quantity: allocation.quantity,
              status: "reserved",
              customerEmail: input.customerEmail?.trim() || null,
              expectedShipDate: allocation.expectedShipDate,
            },
          }));
        }
        return created;
      }, { isolationLevel: Prisma.TransactionIsolationLevel.Serializable });
    } catch (error) {
      if (isSerializableConflict(error) && attempt < 3) continue;
      throw error;
    }
  }

  throw new PreorderCapacityError("Could not reserve preorder capacity after retrying.");
}

export async function releasePreorderOrder(shop: string, shopifyOrderId: string) {
  const now = new Date();
  return prisma.preorderReservation.updateMany({
    where: { shop, shopifyOrderId, status: "reserved" },
    data: { status: "released", releasedAt: now },
  });
}

export async function releasePreorderLine(shop: string, shopifyLineItemId: string) {
  const now = new Date();
  return prisma.preorderReservation.updateMany({
    where: { shop, shopifyLineItemId, status: "reserved" },
    data: { status: "released", releasedAt: now },
  });
}

export async function fulfillPreorderOrder(shop: string, shopifyOrderId: string) {
  const now = new Date();
  return prisma.preorderReservation.updateMany({
    where: { shop, shopifyOrderId, status: "reserved" },
    data: { status: "fulfilled", fulfilledAt: now },
  });
}
