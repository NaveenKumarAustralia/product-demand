import prisma from "../db.server";
import type { PortalMessageUser } from "../portal-messages.server";
import {
  canManagePreorders,
  type PreorderPermissionSettings,
} from "./preorder-permissions.server";
import { getPreorderEligibility } from "./preorder-rules.server";

export class PreorderPermissionError extends Error {
  constructor() {
    super("You do not have permission to manage preorder availability.");
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
};

/**
 * Final server-side gate for activating/pause preorder availability on one
 * production batch. UI visibility is never treated as security.
 */
export async function setPreorderBatchEnabled(input: SetPreorderEnabledInput) {
  if (!canManagePreorders(input.actor, input.permissions)) {
    throw new PreorderPermissionError();
  }

  const order = await prisma.supplierOrder.findUnique({
    where: { id: input.supplierOrderId },
    select: {
      id: true,
      shop: true,
      productTitle: true,
      supplierStatus: true,
      destination: true,
    },
  });
  if (!order) throw new PreorderEligibilityError("Production batch not found.");

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
          ? "Batch must be On Production before preorder can be enabled."
          : "Batch must be assigned to the AUS or USA destination before preorder can be enabled.",
      );
    }
  }

  const now = new Date();
  const setting = await prisma.preorderBatchSetting.upsert({
    where: { supplierOrderId: order.id },
    create: {
      supplierOrderId: order.id,
      shop: order.shop,
      enabled: input.enabled,
      pausedReason: input.enabled ? null : (input.pausedReason?.trim() || null),
      enabledByUserId: input.enabled ? input.actor.id : null,
      enabledByUserName: input.enabled ? input.actor.name : null,
      enabledAt: input.enabled ? now : null,
      updatedByUserId: input.actor.id,
      updatedByUserName: input.actor.name,
    },
    update: {
      enabled: input.enabled,
      pausedReason: input.enabled ? null : (input.pausedReason?.trim() || null),
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

  await prisma.activityLog.create({
    data: {
      userName: input.actor.name,
      action: input.enabled ? "preorder_enabled" : "preorder_paused",
      entity: "supplier_order",
      entityId: String(order.id),
      entityName: order.productTitle,
      field: "preorderEnabled",
      toValue: input.enabled ? "true" : "false",
    },
  }).catch((error) => {
    // Audit logging should not leave a successfully paused batch active if the
    // legacy activity table has an unrelated issue. Surface it operationally.
    console.warn("[preorder audit] failed:", error);
  });

  return setting;
}
