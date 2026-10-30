import prisma from "../db.server";
import { getPreorderEligibility } from "./preorder-rules.server";
import { getPreorderLocationSettings, locationForMarket } from "./preorder-locations.server";
import { getPreorderPermissionContext } from "./preorder-permissions.server";
import { activatePreorderSellingPlan, deactivatePreorderSellingPlan } from "./preorder-selling-plan.service.server";
import { getPreorderSellingPlanRegistryEntries } from "./preorder-selling-plan-registry.server";
import { reconcilePreorderInventoryPolicyForVariants, removeClosedPreorderSellingPlans } from "./preorder-inventory-policy.server";

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

// ─── Scheduled sweep (opt-in, observe-first) ─────────────────────────────────────
//
// Safety model:
//   • "off"     — scheduler does nothing.
//   • "observe" — (DEFAULT) runs the reconciler in DRY-RUN across all pre-order
//                 products and only LOGS what it WOULD change. Zero writes to
//                 Shopify. Use this to confirm it only ever touches genuinely
//                 out-of-sync state before letting it act.
//   • "apply"   — actually converges. Still safe by construction (idempotent,
//                 desired-state-driven: a live batch is never deactivated), but
//                 only turn this on after the observe logs look right.
// Flip the mode with setReconcileSweepMode() (exposed via /api/preorder-reconcile).
const SWEEP_MODE_KEY = "preorder-reconcile-sweep-mode-v1";
export type SweepMode = "off" | "observe" | "apply";

export async function getReconcileSweepMode(): Promise<SweepMode> {
  const s = await prisma.portalSetting.findUnique({ where: { key: SWEEP_MODE_KEY }, select: { value: true } }).catch(() => null);
  const raw = s?.value && typeof s.value === "object" && "mode" in (s.value as Record<string, unknown>) ? (s.value as Record<string, unknown>).mode : s?.value;
  const mode = String(raw ?? "").trim();
  return mode === "off" || mode === "apply" ? mode : "observe"; // default observe
}
export async function setReconcileSweepMode(mode: SweepMode, actorName: string): Promise<SweepMode> {
  const clean: SweepMode = mode === "off" || mode === "apply" ? mode : "observe";
  await prisma.portalSetting.upsert({
    where: { key: SWEEP_MODE_KEY },
    create: { key: SWEEP_MODE_KEY, value: { mode: clean, updatedBy: actorName, updatedAt: new Date().toISOString() } },
    update: { value: { mode: clean, updatedBy: actorName, updatedAt: new Date().toISOString() } },
  });
  return clean;
}

export type SweepSummary = {
  mode: SweepMode; dryRun: boolean; productsChecked: number; wouldChange: number;
  activated: number; deactivated: number; danglingRemoved: number;
  actions: Array<{ productId: string; action: string; batchId: number; reason: string }>;
  errors: Array<{ productId?: string; op: string; error: string }>;
};

// Reconcile EVERY product that currently has a pre-order footprint (an enabled
// batch, or a registered selling plan). dryRun reports intended actions only.
export async function reconcileAllPreorderProducts(opts: { dryRun?: boolean; shop?: string } = {}): Promise<SweepSummary> {
  const mode = await getReconcileSweepMode();
  const dryRun = opts.dryRun ?? mode !== "apply";
  const summary: SweepSummary = { mode, dryRun, productsChecked: 0, wouldChange: 0, activated: 0, deactivated: 0, danglingRemoved: 0, actions: [], errors: [] };

  const shop = opts.shop ?? (await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true } }))?.shop;
  if (!shop) { summary.errors.push({ op: "input", error: "no shop" }); return summary; }

  const [registry, enabledSettings] = await Promise.all([
    getPreorderSellingPlanRegistryEntries(shop),
    prisma.preorderBatchSetting.findMany({ where: { enabled: true }, select: { supplierOrderId: true } }),
  ]);
  const footprintIds = Array.from(new Set([...registry.map((r) => r.supplierOrderId), ...enabledSettings.map((s) => s.supplierOrderId)]));
  const orders = footprintIds.length
    ? await prisma.supplierOrder.findMany({ where: { id: { in: footprintIds }, shop }, select: { id: true, productId: true } })
    : [];
  const existingIds = new Set(orders.map((o) => o.id));
  const productNums = Array.from(new Set(orders.map((o) => numeric(o.productId)).filter(Boolean)));

  // Reconcile each product (throttled to be gentle on Shopify).
  for (const pnum of productNums) {
    try {
      const r = await reconcilePreorderForProduct(shop, pnum, { dryRun });
      summary.productsChecked += 1;
      for (const b of r.batches) {
        if (b.action === "none") continue;
        summary.wouldChange += 1;
        summary.actions.push({ productId: pnum, action: b.action, batchId: b.batchId, reason: b.reason });
      }
      summary.activated += r.activated;
      summary.deactivated += r.deactivated;
      for (const e of r.errors) summary.errors.push({ productId: pnum, op: e.op, error: e.error });
    } catch (e) {
      summary.errors.push({ productId: pnum, op: "reconcile", error: errMsg(e) });
    }
    await new Promise((res) => setTimeout(res, 150));
  }

  // Dangling: a registered plan whose batch no longer exists at all. These can't
  // map to a product, so handle them directly (apply mode only — delete the
  // orphaned plan group + clear the registry entry).
  const danglingIds = registry.filter((r) => !existingIds.has(r.supplierOrderId)).map((r) => r.supplierOrderId);
  if (danglingIds.length) {
    summary.wouldChange += danglingIds.length;
    for (const id of danglingIds) summary.actions.push({ productId: "(deleted batch)", action: "remove-dangling-plan", batchId: id, reason: "registered plan with no batch" });
    if (!dryRun) {
      await removeClosedPreorderSellingPlans(danglingIds).then(() => { summary.danglingRemoved = danglingIds.length; }).catch((e) => summary.errors.push({ op: "dangling", error: errMsg(e) }));
    }
  }

  return summary;
}

// Boot the sweep scheduler once per process. First pass ~90s after boot, then
// hourly. Reads the mode each run (off → skip, observe → dry-run+log, apply →
// converge). Safe to start always; it does nothing until mode is "observe"/"apply".
export function startPreorderReconcileSweepScheduler() {
  const g = globalThis as unknown as { __kePreorderReconcileSweepStarted?: boolean };
  if (g.__kePreorderReconcileSweepStarted) return;
  g.__kePreorderReconcileSweepStarted = true;
  const run = () => {
    void (async () => {
      try {
        const mode = await getReconcileSweepMode();
        if (mode === "off") return;
        const r = await reconcileAllPreorderProducts({});
        if (r.wouldChange || r.errors.length) {
          console.log(`[preorder sweep] mode=${r.mode} dryRun=${r.dryRun} checked=${r.productsChecked} wouldChange=${r.wouldChange} activated=${r.activated} deactivated=${r.deactivated} dangling=${r.danglingRemoved} errors=${r.errors.length}`);
          if (r.actions.length) console.log("[preorder sweep] actions:", JSON.stringify(r.actions.slice(0, 50)));
          if (r.errors.length) console.log("[preorder sweep] errors:", JSON.stringify(r.errors.slice(0, 20)));
        }
      } catch (e) {
        console.warn("[preorder sweep] cycle failed:", errMsg(e));
      }
    })();
  };
  setTimeout(run, 90_000);
  setInterval(run, 60 * 60 * 1000);
  console.log("[preorder sweep] scheduler started (hourly; mode defaults to observe)");
}
