import prisma from "../db.server";
import { getPreorderEligibility } from "./preorder-rules.server";
import { getPreorderLocationSettings, locationForMarket } from "./preorder-locations.server";
import { getPreorderPermissionContext } from "./preorder-permissions.server";
import { activatePreorderSellingPlan, deactivatePreorderSellingPlan } from "./preorder-selling-plan.service.server";
import { getPreorderSellingPlanRegistryEntries } from "./preorder-selling-plan-registry.server";
import { reconcilePreorderInventoryPolicyForVariants } from "./preorder-inventory-policy.server";

// ─── The pre-order reconciler (Phase 1) ─────────────────────────────────────────
//
// SINGLE WRITER of a product's Shopify-side pre-order state. Given a product, it
// reads the DESIRED state from the portal DB (which of its batches are live =
// open + pre-order enabled + eligible + market location set) and makes Shopify
// match it:
//   • a batch that SHOULD be live but has no selling plan  → activate
//   • a batch that has a selling plan but should NOT be live → deactivate
//   • then reconcile every variant's inventory policy (CONTINUE while capacity
//     remains, DENY when the batch is full — this is the quantity cap).
//
// Idempotent: running it when everything is already correct is a no-op. It
// orchestrates the existing, proven primitives (activate / deactivate /
// inventory-policy) so there is ONE place that decides Shopify's pre-order state
// instead of a dozen scattered one-off writes.
//
// Runs on events (and a manual "Reconcile now"); there is deliberately NO
// automatic background sweep.

// System identity for reconciler-driven activate/deactivate. The real
// permission check already happened when a staff member toggled the batch; the
// reconciler only converges Shopify to that already-decided DB state. admin:true
// satisfies canManagePreorders (see hasPermission).
const SYSTEM_ACTOR = { id: "system-reconciler", name: "System (pre-order reconciler)", admin: true, active: true } as const;
const numeric = (v: unknown) => String(v ?? "").replace(/\D/g, "");
const errMsg = (e: unknown) => (e instanceof Error ? e.message : String(e));

export type ReconcilePlanAction = { batchId: number; action: "activate" | "deactivate" | "none"; reason: string };
export type ReconcileResult = {
  ok: boolean;
  productId: string;
  dryRun: boolean;
  batches: ReconcilePlanAction[];
  activated: number;
  deactivated: number;
  errors: Array<{ batchId?: number; op: string; error: string }>;
};

export async function reconcilePreorderForProduct(
  shop: string,
  productId: unknown,
  opts: { dryRun?: boolean } = {},
): Promise<ReconcileResult> {
  const dryRun = opts.dryRun === true;
  const pnum = numeric(productId);
  const base: ReconcileResult = { ok: false, productId: pnum, dryRun, batches: [], activated: 0, deactivated: 0, errors: [] };
  if (!shop || !pnum) return { ...base, errors: [{ op: "input", error: "missing shop/productId" }] };

  // Batches for THIS product (narrow by contains, then exact numeric match).
  const candidates = await prisma.supplierOrder.findMany({
    where: { shop, productId: { contains: pnum } },
    select: { id: true, productId: true, status: true, supplierStatus: true, destination: true, lines: { select: { variantId: true } } },
  });
  const mine = candidates.filter((o) => numeric(o.productId) === pnum);
  if (!mine.length) return { ...base, ok: true };

  const ids = mine.map((o) => o.id);
  const [settings, registry, locations, permCtx] = await Promise.all([
    prisma.preorderBatchSetting.findMany({ where: { supplierOrderId: { in: ids } }, select: { supplierOrderId: true, enabled: true } }),
    getPreorderSellingPlanRegistryEntries(shop),
    getPreorderLocationSettings(),
    getPreorderPermissionContext(),
  ]);
  const enabledById = new Map(settings.map((s) => [s.supplierOrderId, s.enabled === true]));
  const activatedIds = new Set(registry.map((r) => r.supplierOrderId));
  const permissions = permCtx.permissions;

  // Desired: a batch is LIVE iff open + pre-order enabled + eligible + the
  // market's fulfilment location is configured (mirrors activate()'s own gates).
  const liveState = (o: typeof mine[number]): { live: boolean; reason: string } => {
    if (o.status !== "open") return { live: false, reason: `batch ${o.status} (not open)` };
    const elig = getPreorderEligibility({ supplierStatus: o.supplierStatus, destination: o.destination, preorderEnabled: enabledById.get(o.id) === true });
    if (!elig.eligible) return { live: false, reason: `not eligible (${elig.reason})` };
    if (elig.market && !locationForMarket(locations, elig.market)) return { live: false, reason: `no ${elig.market} location configured` };
    return { live: true, reason: "open + enabled + eligible" };
  };

  const plan: ReconcilePlanAction[] = mine.map((o) => {
    const { live, reason } = liveState(o);
    const hasPlan = activatedIds.has(o.id);
    if (live && !hasPlan) return { batchId: o.id, action: "activate", reason };
    if (!live && hasPlan) return { batchId: o.id, action: "deactivate", reason };
    return { batchId: o.id, action: "none", reason: hasPlan ? `already live — ${reason}` : reason };
  });
  base.batches = plan;

  if (dryRun) return { ...base, ok: true };

  // Deactivate stale plans first, then activate the live ones (so a product that
  // is both closing one batch and opening another ends in the right state).
  for (const p of plan) {
    if (p.action !== "deactivate") continue;
    try { await deactivatePreorderSellingPlan({ supplierOrderId: p.batchId, actor: SYSTEM_ACTOR, permissions }); base.deactivated += 1; }
    catch (e) { base.errors.push({ batchId: p.batchId, op: "deactivate", error: errMsg(e) }); }
  }
  for (const p of plan) {
    if (p.action !== "activate") continue;
    try { await activatePreorderSellingPlan({ supplierOrderId: p.batchId, actor: SYSTEM_ACTOR, permissions }); base.activated += 1; }
    catch (e) { base.errors.push({ batchId: p.batchId, op: "activate", error: errMsg(e) }); }
  }

  // Fine-tune each variant's inventory policy to remaining capacity (the cap).
  const variantIds = Array.from(new Set(mine.flatMap((o) => o.lines.map((l) => l.variantId)).filter(Boolean)));
  if (variantIds.length) {
    await reconcilePreorderInventoryPolicyForVariants(shop, variantIds).catch((e) => base.errors.push({ op: "inventory-policy", error: errMsg(e) }));
  }

  return { ...base, ok: base.errors.length === 0 };
}
