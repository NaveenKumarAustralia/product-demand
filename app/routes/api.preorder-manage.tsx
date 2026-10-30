import type { ActionFunctionArgs } from "react-router";
import prisma from "../db.server";
import { reconcilePreorderForProduct } from "../preorder/preorder-reconcile.server";
import {
  PreorderEligibilityError,
  PreorderPermissionError,
  setPreorderBatchEnabled,
  updatePreorderBatchSettings,
} from "../preorder/preorder-batch.service";
import { getPreorderPermissionContext } from "../preorder/preorder-permissions.server";
import { setPreorderLocationSettings } from "../preorder/preorder-locations.server";
import { setPreorderNotifyEnabled, setPreorderCombineWindowDays, setBackInStockThreshold } from "../preorder/preorder-storefront-settings.server";
import { combineExistingPreorderOrders } from "../preorder/preorder-release.server";
import {
  PreorderSellingPlanError,
  activatePreorderSellingPlan,
  deactivatePreorderSellingPlan,
  refreshPreorderSellingPlanDate,
} from "../preorder/preorder-selling-plan.service.server";
import {
  requirePreorderPortalUser,
  requireSameOrigin,
} from "../preorder/preorder-portal-auth.server";

function jsonError(message: string, status: number) {
  return Response.json({ ok: false, error: message }, { status });
}

function parseId(value: unknown) {
  const id = Number(value);
  if (!Number.isInteger(id) || id <= 0) return null;
  return id;
}

function parseOptionalDate(value: unknown): Date | null | undefined {
  if (value === undefined) return undefined;
  const text = String(value ?? "").trim();
  if (!text) return null;
  const date = new Date(`${text}T00:00:00.000Z`);
  if (Number.isNaN(date.getTime())) throw new PreorderEligibilityError("Invalid ship date.");
  return date;
}

export const action = async ({ request }: ActionFunctionArgs) => {
  if (request.method !== "POST") return jsonError("Method not allowed", 405);

  try {
    requireSameOrigin(request);
    const actor = await requirePreorderPortalUser(request);
    const { permissions } = await getPreorderPermissionContext();
    const payload = await request.json().catch(() => null) as Record<string, unknown> | null;
    if (!payload) return jsonError("Invalid request body.", 400);

    const operation = String(payload.operation ?? "");

    if (operation === "set-notify-enabled") {
      if (actor.admin !== true) return jsonError("Only a portal admin can change the notify-me block.", 403);
      const enabled = await setPreorderNotifyEnabled(payload.enabled === true, actor.name);
      return Response.json({ ok: true, notifyBlockEnabled: enabled });
    }

    if (operation === "set-combine-window") {
      if (actor.admin !== true) return jsonError("Only a portal admin can change the combine window.", 403);
      const days = await setPreorderCombineWindowDays(Number(payload.days), actor.name);
      return Response.json({ ok: true, combineWindowDays: days });
    }

    if (operation === "set-backinstock-threshold") {
      if (actor.admin !== true) return jsonError("Only a portal admin can change the back-in-stock threshold.", 403);
      const threshold = await setBackInStockThreshold(Number(payload.threshold), actor.name);
      return Response.json({ ok: true, backInStockThreshold: threshold });
    }

    if (operation === "combine-existing") {
      if (actor.admin !== true) return jsonError("Only a portal admin can combine existing orders.", 403);
      const result = await combineExistingPreorderOrders();
      return Response.json({ ok: true, combine: result });
    }

    if (operation === "update-locations") {
      if (actor.admin !== true) return jsonError("Only a portal admin can change preorder location settings.", 403);
      const locations = await setPreorderLocationSettings({
        AU: String(payload.AU ?? "").trim() || null,
        USA: String(payload.USA ?? "").trim() || null,
      }, actor.name);
      return Response.json({ ok: true, locations });
    }

    const supplierOrderId = parseId(payload.supplierOrderId);
    if (!supplierOrderId) return jsonError("Invalid production batch.", 400);

    if (operation === "activate-shopify") {
      const result = await activatePreorderSellingPlan({ supplierOrderId, actor, permissions });
      return Response.json({ ok: true, result });
    }

    // One-click flow for the per-row Status-cell control: set the customer
    // dispatch date, enable, and go live on Shopify in a single manage action.
    // If activation fails, roll the enable back so a row can never show "live"
    // without an actual Shopify selling plan behind it.
    // Enable: set the DB state (enabled + dispatch date), then let the RECONCILER
    // converge Shopify (it's the single writer — it activates the now-live batch).
    // If the activation fails, roll the enable back so a row can never show "live"
    // without a real selling plan behind it.
    if (operation === "row-enable") {
      const shipDate = parseOptionalDate(payload.shipDate);
      const setting = await setPreorderBatchEnabled({ supplierOrderId, enabled: true, shipDate, actor, permissions });
      const order = await prisma.supplierOrder.findUnique({ where: { id: supplierOrderId }, select: { shop: true, productId: true } });
      if (order?.shop && order.productId) {
        const rec = await reconcilePreorderForProduct(order.shop, order.productId);
        const activateError = rec.errors.find((e) => e.batchId === supplierOrderId && e.op === "activate");
        if (activateError) {
          await setPreorderBatchEnabled({ supplierOrderId, enabled: false, actor, permissions }).catch(() => undefined);
          await reconcilePreorderForProduct(order.shop, order.productId).catch(() => undefined);
          return jsonError(activateError.error || "Could not activate pre-order on Shopify.", 400);
        }
      }
      return Response.json({
        ok: true,
        state: { enabled: true, activated: true, shipDate: setting.shipDate?.toISOString() ?? null },
      });
    }

    // Turn off: set the DB state to disabled, then let the RECONCILER converge
    // Shopify (it deletes the now-not-live plan + restores DENY). The hourly
    // sweep is the backstop if the immediate reconcile hits a transient error.
    if (operation === "row-disable") {
      await setPreorderBatchEnabled({ supplierOrderId, enabled: false, actor, permissions });
      const order = await prisma.supplierOrder.findUnique({ where: { id: supplierOrderId }, select: { shop: true, productId: true } });
      if (order?.shop && order.productId) await reconcilePreorderForProduct(order.shop, order.productId).catch(() => undefined);
      return Response.json({ ok: true, state: { enabled: false, activated: false, shipDate: null } });
    }

    if (operation === "remove-shopify") {
      const result = await deactivatePreorderSellingPlan({ supplierOrderId, actor, permissions });
      return Response.json({ ok: true, result });
    }

    if (operation === "set-enabled") {
      const enabled = payload.enabled === true;
      if (!enabled) {
        // Remove the Shopify purchase option first so a portal pause can never
        // leave a customer-facing preorder selling plan accidentally active.
        await deactivatePreorderSellingPlan({ supplierOrderId, actor, permissions });
      }
      const setting = await setPreorderBatchEnabled({
        supplierOrderId,
        enabled,
        pausedReason: String(payload.pausedReason ?? "").trim() || null,
        actor,
        permissions,
      });
      return Response.json({
        ok: true,
        setting: {
          enabled: setting.enabled,
          shipDate: setting.shipDate?.toISOString() ?? null,
          safetyBufferPercent: setting.safetyBufferPercent,
          safetyBufferQty: setting.safetyBufferQty,
          pausedReason: setting.pausedReason,
        },
      });
    }

    if (operation === "update-settings") {
      const input: Parameters<typeof updatePreorderBatchSettings>[0] = {
        supplierOrderId,
        actor,
        permissions,
      };

      let shipDateChanged = false;
      if (Object.prototype.hasOwnProperty.call(payload, "shipDate")) {
        input.shipDate = parseOptionalDate(payload.shipDate);
        shipDateChanged = true;
      }
      if (Object.prototype.hasOwnProperty.call(payload, "safetyBufferPercent")) {
        input.safetyBufferPercent = Number(payload.safetyBufferPercent);
      }
      if (Object.prototype.hasOwnProperty.call(payload, "safetyBufferQty")) {
        const raw = payload.safetyBufferQty;
        input.safetyBufferQty = raw == null || String(raw).trim() === "" ? null : Number(raw);
      }

      const setting = await updatePreorderBatchSettings(input);
      // Keep the live Shopify selling plan's "Expected <date>" in sync so the
      // order line, cart and email show the new date (best-effort).
      if (shipDateChanged) await refreshPreorderSellingPlanDate(supplierOrderId).catch(() => undefined);
      return Response.json({
        ok: true,
        setting: {
          enabled: setting.enabled,
          shipDate: setting.shipDate?.toISOString() ?? null,
          safetyBufferPercent: setting.safetyBufferPercent,
          safetyBufferQty: setting.safetyBufferQty,
          pausedReason: setting.pausedReason,
        },
      });
    }

    return jsonError("Unknown operation.", 400);
  } catch (error) {
    if (error instanceof Response) return error;
    if (error instanceof PreorderPermissionError) return jsonError(error.message, 403);
    if (error instanceof PreorderEligibilityError) return jsonError(error.message, 400);
    if (error instanceof PreorderSellingPlanError) return jsonError(error.message, 400);
    console.error("[preorder manage] failed:", error);
    return jsonError("Could not update preorder settings.", 500);
  }
};
