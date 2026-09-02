import type { ActionFunctionArgs } from "react-router";
import {
  PreorderEligibilityError,
  PreorderPermissionError,
  setPreorderBatchEnabled,
  updatePreorderBatchSettings,
} from "../preorder/preorder-batch.service";
import { getPreorderPermissionContext } from "../preorder/preorder-permissions.server";
import { setPreorderLocationSettings } from "../preorder/preorder-locations.server";
import {
  PreorderSellingPlanError,
  activatePreorderSellingPlan,
  deactivatePreorderSellingPlan,
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

      if (Object.prototype.hasOwnProperty.call(payload, "shipDate")) {
        input.shipDate = parseOptionalDate(payload.shipDate);
      }
      if (Object.prototype.hasOwnProperty.call(payload, "safetyBufferPercent")) {
        input.safetyBufferPercent = Number(payload.safetyBufferPercent);
      }
      if (Object.prototype.hasOwnProperty.call(payload, "safetyBufferQty")) {
        const raw = payload.safetyBufferQty;
        input.safetyBufferQty = raw == null || String(raw).trim() === "" ? null : Number(raw);
      }

      const setting = await updatePreorderBatchSettings(input);
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
