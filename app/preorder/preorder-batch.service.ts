import prisma from "../db.server";
import type { PortalMessageUser } from "../portal-messages.server";
import {
  canManagePreorders,
  canManagePreorderEta,
  canManageSafetyBuffer,
  type PreorderPermissionSettings,
} from "./preorder-permissions.server";
import { getPreorderEligibility } from "./preorder-rules.server";

export class PreorderPermissionError extends Error {
  constructor(message = "You do not have permission to manage preorder availability.") {
    super(message);
    this.name = "PreorderPermissionError";
  }
}

export class PreorderEligibilityError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "PreorderEligibilityError";
  }
}

export type SetPreorderEnabledInput = {
  supplierOrderId: number;
  enabled: boolean;
  actor: PortalMessageUser;
  permissions: PreorderPermissionSettings;
  pausedReason?: string | null;
  // Optional customer dispatch date to store alongside enabling. Lets the
  // per-row "Enable pre-order" control set the date and enable in one manage
  // action (a manager doesn't need the separate ship-date permission just to
  // turn a preorder on with its promised date).
  shipDate?: Date | null;
};

export type UpdatePreorderBatchSettingsInput = {
  supplierOrderId: number;
  actor: PortalMessageUser;
  permissions: PreorderPermissionSettings;
  shipDate?: Date | null;
  safetyBufferPercent?: number;
  safetyBufferQty?: number | null;
};

async function getOrderForPreorder(supplierOrderId: number) {
  const order = await prisma.supplierOrder.findUnique({
    where: { id: supplierOrderId },
    select: {
      id: true,
      shop: true,
      productTitle: true,
      supplierStatus: true,
      destination: true,
    },
  });
  if (!order) throw new PreorderEligibilityError("Production batch not found.");
  return order;
}

async function audit({
  actor,
  orderId,
  productTitle,
  action,
  field,
  toValue,
}: {
  actor: PortalMessageUser;
  orderId: number;
  productTitle: string;
  action: string;
  field: string;
  toValue: string | null;
}) {
  await prisma.activityLog.create({
    data: {
      userName: actor.name,
      action,
      entity: "supplier_order",
      entityId: String(orderId),
      entityName: productTitle,
      field,
      toValue,
    },
  }).catch((error) => {
    console.warn("[preorder audit] failed:", error);
  });
}

/**
 * Final server-side gate for activating/pause preorder availability on one
 * production batch. UI visibility is never treated as security.
 */
export async function setPreorderBatchEnabled(input: SetPreorderEnabledInput) {
  if (!canManagePreorders(input.actor, input.permissions)) {
    throw new PreorderPermissionError();
  }

  const order = await getOrderForPreorder(input.supplierOrderId);

  // Enabling must satisfy production/destination eligibility. Disabling is
  // always allowed so staff can immediately stop taking new reservations.
  if (input.enabled) {
    const eligibility = getPreorderEligibility({
      supplierStatus: order.supplierStatus,
      destination: order.destination,
      preorderEnabled: true,
    });
    if (!eligibility.eligible) {
      throw new PreorderEligibilityError(
        eligibility.reason === "not_on_production"
          ? "Set the batch to a production status (On Production, Ready or In Shipment) before enabling preorder."
          : "Batch must be assigned to the AUS or USA destination before preorder can be enabled.",
      );
    }
  }

  // Only touch shipDate when the caller explicitly passed one (undefined = leave
  // the existing date as-is; null = clear it). Enabling with a date is one atomic
  // manage action for the per-row control.
  const hasShipDate = Object.prototype.hasOwnProperty.call(input, "shipDate");
  const now = new Date();
  const setting = await prisma.preorderBatchSetting.upsert({
    where: { supplierOrderId: order.id },
    create: {
      supplierOrderId: order.id,
      shop: order.shop,
      enabled: input.enabled,
      pausedReason: input.enabled ? null : (input.pausedReason?.trim() || null),
      ...(hasShipDate ? { shipDate: input.shipDate ?? null } : {}),
      enabledByUserId: input.enabled ? input.actor.id : null,
      enabledByUserName: input.enabled ? input.actor.name : null,
      enabledAt: input.enabled ? now : null,
      updatedByUserId: input.actor.id,
      updatedByUserName: input.actor.name,
    },
    update: {
      enabled: input.enabled,
      pausedReason: input.enabled ? null : (input.pausedReason?.trim() || null),
      ...(hasShipDate ? { shipDate: input.shipDate ?? null } : {}),
      ...(input.enabled
        ? {
            enabledByUserId: input.actor.id,
            enabledByUserName: input.actor.name,
            enabledAt: now,
          }
        : {}),
      updatedByUserId: input.actor.id,
      updatedByUserName: input.actor.name,
    },
  });

  await audit({
    actor: input.actor,
    orderId: order.id,
    productTitle: order.productTitle,
    action: input.enabled ? "preorder_enabled" : "preorder_paused",
    field: "preorderEnabled",
    toValue: input.enabled ? "true" : "false",
  });

  return setting;
}

export async function updatePreorderBatchSettings(input: UpdatePreorderBatchSettingsInput) {
  const hasShipDateChange = Object.prototype.hasOwnProperty.call(input, "shipDate");
  const hasBufferPercentChange = Object.prototype.hasOwnProperty.call(input, "safetyBufferPercent");
  const hasBufferQtyChange = Object.prototype.hasOwnProperty.call(input, "safetyBufferQty");

  if (hasShipDateChange && !canManagePreorderEta(input.actor, input.permissions)) {
    throw new PreorderPermissionError("You do not have permission to change preorder ship dates.");
  }
  if ((hasBufferPercentChange || hasBufferQtyChange) && !canManageSafetyBuffer(input.actor, input.permissions)) {
    throw new PreorderPermissionError("You do not have permission to change preorder safety buffers.");
  }
  if (!hasShipDateChange && !hasBufferPercentChange && !hasBufferQtyChange) {
    throw new PreorderEligibilityError("No preorder settings were supplied.");
  }

  const order = await getOrderForPreorder(input.supplierOrderId);
  const updateData: {
    shipDate?: Date | null;
    safetyBufferPercent?: number;
    safetyBufferQty?: number | null;
    updatedByUserId: string;
    updatedByUserName: string;
  } = {
    updatedByUserId: input.actor.id,
    updatedByUserName: input.actor.name,
  };

  if (hasShipDateChange) updateData.shipDate = input.shipDate ?? null;
  if (hasBufferPercentChange) {
    const value = Number(input.safetyBufferPercent);
    if (!Number.isFinite(value) || value < 0 || value > 50) {
      throw new PreorderEligibilityError("Safety buffer percentage must be between 0 and 50.");
    }
    updateData.safetyBufferPercent = value;
  }
  if (hasBufferQtyChange) {
    const value = input.safetyBufferQty;
    if (value != null && (!Number.isInteger(value) || value < 0)) {
      throw new PreorderEligibilityError("Safety buffer quantity must be a whole number of zero or more.");
    }
    updateData.safetyBufferQty = value ?? null;
  }

  const setting = await prisma.preorderBatchSetting.upsert({
    where: { supplierOrderId: order.id },
    create: {
      supplierOrderId: order.id,
      shop: order.shop,
      enabled: false,
      ...updateData,
    },
    update: updateData,
  });

  if (hasShipDateChange) {
    await audit({
      actor: input.actor,
      orderId: order.id,
      productTitle: order.productTitle,
      action: "preorder_ship_date_updated",
      field: "shipDate",
      toValue: input.shipDate?.toISOString() ?? null,
    });
  }
  if (hasBufferPercentChange || hasBufferQtyChange) {
    await audit({
      actor: input.actor,
      orderId: order.id,
      productTitle: order.productTitle,
      action: "preorder_buffer_updated",
      field: "safetyBuffer",
      toValue: JSON.stringify({
        percent: setting.safetyBufferPercent,
        qty: setting.safetyBufferQty,
      }),
    });
  }

  return setting;
}
