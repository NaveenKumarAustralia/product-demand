import prisma from "../db.server";
import { calculatePreorderCapacity } from "./preorder-rules.server";
import { getPreorderSellingPlanRegistryEntries } from "./preorder-selling-plan-registry.server";
import { getOfflineToken } from "./preorder-fulfillment.server";

// Keeps a pre-order variant's Shopify inventory policy in lock-step with the
// batch's REMAINING capacity, so EVERY checkout path respects the same limit the
// pre-order button does:
//   • capacity remaining  → CONTINUE (button works; express/Shop Pay can buy,
//                           and a genuine out-of-stock oversell is caught+held);
//   • batch full (reserved ≥ incoming − buffer) → DENY → Shopify blocks ALL
//     paths (Shop Pay / PayPal / quick-add included) so they can't oversell
//     past the batch's real quantity and create un-holdable "ghosts".
//
// Without this, activation set CONTINUE once and left it on the whole time the
// batch was live, so express checkouts could oversell without limit. Driven off
// the same capacity figure the storefront button uses, so the two always agree.

const API_VERSION = "2025-10";
const numericId = (v: string) => String(v ?? "").split("/").pop()?.replace(/[^0-9]/g, "") ?? "";

async function setVariantsInventoryPolicy(
  shop: string,
  token: string,
  productId: string,
  variantGids: string[],
  policy: "CONTINUE" | "DENY",
): Promise<void> {
  if (!variantGids.length) return;
  const productGid = productId.startsWith("gid://") ? productId : `gid://shopify/Product/${productId}`;
  const variants = variantGids.map((id) => ({ id, inventoryPolicy: policy }));
  const res = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
    method: "POST",
    headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": token },
    body: JSON.stringify({
      query: `#graphql
        mutation KEPreorderPolicyReconcile($productId: ID!, $variants: [ProductVariantsBulkInput!]!) {
          productVariantsBulkUpdate(productId: $productId, variants: $variants) { userErrors { field message } }
        }`,
      variables: { productId: productGid, variants },
    }),
  });
  if (!res.ok) throw new Error(`Shopify HTTP ${res.status} updating inventory policy`);
  const json = await res.json() as { data?: { productVariantsBulkUpdate?: { userErrors?: Array<{ message?: string }> } }; errors?: Array<{ message?: string }> };
  const errs = json.errors?.map((e) => e.message) ?? json.data?.productVariantsBulkUpdate?.userErrors?.map((e) => e.message);
  if (errs?.length) throw new Error(errs.join("; "));
}

/**
 * Re-evaluate the inventory policy for the given variants across every LIVE
 * batch (enabled + activated) that contains them, and flip CONTINUE/DENY to
 * match remaining capacity. Best-effort and idempotent: a variant that isn't in
 * any live batch is left untouched (batch deactivation already restores DENY).
 */
export async function reconcilePreorderInventoryPolicyForVariants(shop: string, variantIdsRaw: string[]): Promise<void> {
  const wanted = new Set(variantIdsRaw.map(numericId).filter(Boolean));
  if (!wanted.size) return;
  const token = await getOfflineToken(shop);
  if (!token) return;

  const [enabledSettings, registry] = await Promise.all([
    prisma.preorderBatchSetting.findMany({ where: { enabled: true }, select: { supplierOrderId: true, safetyBufferPercent: true, safetyBufferQty: true } }),
    getPreorderSellingPlanRegistryEntries(shop),
  ]);
  const activatedIds = new Set(registry.map((r) => r.supplierOrderId));
  const liveSettings = enabledSettings.filter((s) => activatedIds.has(s.supplierOrderId));
  const liveIds = liveSettings.map((s) => s.supplierOrderId);
  if (!liveIds.length) return;
  const bufferById = new Map(liveSettings.map((s) => [s.supplierOrderId, { pct: s.safetyBufferPercent ?? 0, qty: s.safetyBufferQty ?? null }]));

  const [batches, reservations] = await Promise.all([
    prisma.supplierOrder.findMany({
      where: { id: { in: liveIds }, status: "open", destination: { in: ["send_to_au", "send_to_usa"] } },
      select: { id: true, productId: true, lines: { select: { variantId: true, qtyOrdered: true, qtyReceived: true } } },
    }),
    prisma.preorderReservation.findMany({ where: { supplierOrderId: { in: liveIds }, status: "reserved" }, select: { supplierOrderId: true, variantId: true, quantity: true } }),
  ]);
  const reservedByKey = new Map<string, number>();
  for (const r of reservations) {
    const key = `${r.supplierOrderId}:${numericId(r.variantId)}`;
    reservedByKey.set(key, (reservedByKey.get(key) ?? 0) + r.quantity);
  }

  // Sum each wanted variant's remaining capacity across every live batch it's in.
  const availByVariant = new Map<string, number>();
  const productByVariant = new Map<string, string>();
  for (const b of batches) {
    const buf = bufferById.get(b.id) ?? { pct: 0, qty: null };
    for (const line of b.lines) {
      const vnum = numericId(line.variantId);
      if (!wanted.has(vnum)) continue;
      const reserved = reservedByKey.get(`${b.id}:${vnum}`) ?? 0;
      const incomingRemaining = Math.max(0, (line.qtyOrdered ?? 0) - (line.qtyReceived ?? 0));
      const cap = calculatePreorderCapacity({ confirmedIncomingQty: incomingRemaining, reservedQty: reserved, safetyBufferPercent: buf.pct, safetyBufferQty: buf.qty });
      availByVariant.set(vnum, (availByVariant.get(vnum) ?? 0) + cap.availableToPreorder);
      if ((b.productId ?? "").trim() && !productByVariant.has(vnum)) productByVariant.set(vnum, String(b.productId));
    }
  }

  // Group the writes by product and target policy.
  const denyByProduct = new Map<string, string[]>();
  const contByProduct = new Map<string, string[]>();
  for (const vnum of wanted) {
    const product = productByVariant.get(vnum);
    if (!product) continue; // not in any live batch → leave its policy alone
    const target = (availByVariant.get(vnum) ?? 0) <= 0 ? denyByProduct : contByProduct;
    const arr = target.get(product) ?? [];
    arr.push(`gid://shopify/ProductVariant/${vnum}`);
    target.set(product, arr);
  }

  for (const [product, gids] of denyByProduct) {
    await setVariantsInventoryPolicy(shop, token, product, gids, "DENY")
      .catch((e) => console.warn("[preorder policy] DENY reconcile failed:", e instanceof Error ? e.message : e));
  }
  for (const [product, gids] of contByProduct) {
    await setVariantsInventoryPolicy(shop, token, product, gids, "CONTINUE")
      .catch((e) => console.warn("[preorder policy] CONTINUE reconcile failed:", e instanceof Error ? e.message : e));
  }
}

/**
 * Sweep EVERY variant in every live batch and align its inventory policy to
 * remaining capacity. Self-healing baseline for the scheduler: brings any batch
 * that's already full down to DENY even if no new order touches it, and reopens
 * ones that freed up. Best-effort; finds the shop from the offline session.
 */
export async function reconcileAllLivePreorderInventoryPolicies(): Promise<void> {
  const session = await prisma.session.findFirst({
    where: { isOnline: false, accessToken: { not: "" } },
    orderBy: { expires: "desc" },
    select: { shop: true },
  });
  if (!session?.shop) return;
  const [enabledSettings, registry] = await Promise.all([
    prisma.preorderBatchSetting.findMany({ where: { enabled: true }, select: { supplierOrderId: true } }),
    getPreorderSellingPlanRegistryEntries(session.shop),
  ]);
  const activatedIds = new Set(registry.map((r) => r.supplierOrderId));
  const liveIds = enabledSettings.map((s) => s.supplierOrderId).filter((id) => activatedIds.has(id));
  if (!liveIds.length) return;
  const batches = await prisma.supplierOrder.findMany({
    where: { id: { in: liveIds }, status: "open", destination: { in: ["send_to_au", "send_to_usa"] } },
    select: { lines: { select: { variantId: true } } },
  });
  const variantIds = batches.flatMap((b) => b.lines.map((l) => l.variantId)).filter(Boolean);
  if (variantIds.length) await reconcilePreorderInventoryPolicyForVariants(session.shop, variantIds);
}

export type PreorderPolicyReportRow = {
  variantId: string; productTitle: string; size: string | null; batchIds: number[];
  incoming: number; reserved: number; availableToPreorder: number;
  targetPolicy: "CONTINUE" | "DENY"; currentPolicy: string | null; needsChange: boolean;
};

// Read-only audit of every live pre-order variant: remaining capacity, the
// inventory policy Shopify currently has, and what it SHOULD be. `needsChange`
// flags a full batch still set to CONTINUE (i.e. still oversellable). Powers the
// dry-run of /api/preorder-reconcile-policies.
export async function reportLivePreorderInventoryPolicies(): Promise<{
  ok: boolean; shop?: string; rows: PreorderPolicyReportRow[];
  summary: { liveVariants: number; full: number; needChange: number; stillOversellable: number };
  error?: string;
}> {
  const empty = { liveVariants: 0, full: 0, needChange: 0, stillOversellable: 0 };
  const session = await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true } });
  if (!session?.shop) return { ok: false, rows: [], summary: empty, error: "No offline Shopify session." };
  const token = await getOfflineToken(session.shop);
  if (!token) return { ok: false, shop: session.shop, rows: [], summary: empty, error: "No offline Shopify token." };

  const [enabledSettings, registry] = await Promise.all([
    prisma.preorderBatchSetting.findMany({ where: { enabled: true }, select: { supplierOrderId: true, safetyBufferPercent: true, safetyBufferQty: true } }),
    getPreorderSellingPlanRegistryEntries(session.shop),
  ]);
  const activatedIds = new Set(registry.map((r) => r.supplierOrderId));
  const liveSettings = enabledSettings.filter((s) => activatedIds.has(s.supplierOrderId));
  const liveIds = liveSettings.map((s) => s.supplierOrderId);
  if (!liveIds.length) return { ok: true, shop: session.shop, rows: [], summary: empty };
  const bufferById = new Map(liveSettings.map((s) => [s.supplierOrderId, { pct: s.safetyBufferPercent ?? 0, qty: s.safetyBufferQty ?? null }]));

  const [batches, reservations] = await Promise.all([
    prisma.supplierOrder.findMany({
      where: { id: { in: liveIds }, status: "open", destination: { in: ["send_to_au", "send_to_usa"] } },
      select: { id: true, productTitle: true, lines: { select: { variantId: true, variantTitle: true, qtyOrdered: true, qtyReceived: true } } },
    }),
    prisma.preorderReservation.findMany({ where: { supplierOrderId: { in: liveIds }, status: "reserved" }, select: { supplierOrderId: true, variantId: true, quantity: true } }),
  ]);
  const reservedByKey = new Map<string, number>();
  for (const r of reservations) reservedByKey.set(`${r.supplierOrderId}:${numericId(r.variantId)}`, (reservedByKey.get(`${r.supplierOrderId}:${numericId(r.variantId)}`) ?? 0) + r.quantity);

  const agg = new Map<string, { productTitle: string; size: string | null; batchIds: number[]; incoming: number; reserved: number; avail: number }>();
  for (const b of batches) {
    const buf = bufferById.get(b.id) ?? { pct: 0, qty: null };
    for (const line of b.lines) {
      const vnum = numericId(line.variantId);
      if (!vnum) continue;
      const reserved = reservedByKey.get(`${b.id}:${vnum}`) ?? 0;
      const incomingRemaining = Math.max(0, (line.qtyOrdered ?? 0) - (line.qtyReceived ?? 0));
      const cap = calculatePreorderCapacity({ confirmedIncomingQty: incomingRemaining, reservedQty: reserved, safetyBufferPercent: buf.pct, safetyBufferQty: buf.qty });
      const cur = agg.get(vnum) ?? { productTitle: b.productTitle, size: line.variantTitle ?? null, batchIds: [], incoming: 0, reserved: 0, avail: 0 };
      cur.batchIds.push(b.id); cur.incoming += incomingRemaining; cur.reserved += reserved; cur.avail += cap.availableToPreorder;
      agg.set(vnum, cur);
    }
  }

  // Current Shopify inventory policy per variant (nodes in chunks of 100).
  const vnums = Array.from(agg.keys());
  const policyByVariant = new Map<string, string>();
  for (let i = 0; i < vnums.length; i += 100) {
    const ids = vnums.slice(i, i + 100).map((v) => `gid://shopify/ProductVariant/${v}`);
    try {
      const res = await fetch(`https://${session.shop}/admin/api/${API_VERSION}/graphql.json`, {
        method: "POST", headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": token },
        body: JSON.stringify({ query: `#graphql
          query KEPolicyReport($ids: [ID!]!) { nodes(ids: $ids) { ... on ProductVariant { id inventoryPolicy } } }`, variables: { ids } }),
      });
      const json = await res.json() as { data?: { nodes?: Array<{ id?: string; inventoryPolicy?: string } | null> } };
      for (const n of json.data?.nodes ?? []) { if (n?.id) policyByVariant.set(numericId(n.id), String(n.inventoryPolicy ?? "")); }
    } catch { /* leave current unknown for this chunk */ }
  }

  const rows: PreorderPolicyReportRow[] = vnums.map((vnum) => {
    const a = agg.get(vnum)!;
    const targetPolicy: "CONTINUE" | "DENY" = a.avail <= 0 ? "DENY" : "CONTINUE";
    const currentPolicy = policyByVariant.get(vnum) ?? null;
    return { variantId: vnum, productTitle: a.productTitle, size: a.size, batchIds: a.batchIds, incoming: a.incoming, reserved: a.reserved, availableToPreorder: a.avail, targetPolicy, currentPolicy, needsChange: currentPolicy != null && currentPolicy.toUpperCase() !== targetPolicy };
  }).sort((x, y) => x.productTitle.localeCompare(y.productTitle));

  const full = rows.filter((r) => r.availableToPreorder <= 0).length;
  const needChange = rows.filter((r) => r.needsChange).length;
  const stillOversellable = rows.filter((r) => r.availableToPreorder <= 0 && (r.currentPolicy ?? "").toUpperCase() === "CONTINUE").length;
  return { ok: true, shop: session.shop, rows, summary: { liveVariants: rows.length, full, needChange, stillOversellable } };
}

// ─── Stray "continue selling when out of stock" cleanup ─────────────────────────
// Bug sweep: many products ended up with variants set to inventoryPolicy CONTINUE
// ("continue selling when out of stock") WITHOUT pre-order being enabled — most
// likely because creating a product by DUPLICATING a pre-order product copies its
// CONTINUE policy. This scans every Shopify product, and for any product that is
// NOT a currently-enabled pre-order batch, flips its CONTINUE variants back to
// DENY. Dry-run by default (apply:false) so you can review first.
export type StrayContinueResult = {
  ok: boolean; shop?: string; error?: string;
  excludedPreorderProducts: number; scannedProducts: number;
  continueProductsTotal?: number; continueVariantsTotal?: number;
  affected: Array<{ productId: string; title: string; variants: string[] }>;
  totalVariants: number; applied: boolean; changedProducts: number; changedVariants: number;
  errors: string[]; truncated: boolean;
};
export async function sweepStrayContinueInventoryPolicies(opts: { apply?: boolean; maxPages?: number } = {}): Promise<StrayContinueResult> {
  const apply = opts.apply === true;
  const base: StrayContinueResult = { ok: false, excludedPreorderProducts: 0, scannedProducts: 0, affected: [], totalVariants: 0, applied: apply, changedProducts: 0, changedVariants: 0, errors: [], truncated: false };
  const session = await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true } });
  if (!session?.shop) return { ...base, error: "no_session" };
  const shop = session.shop;
  const token = await getOfflineToken(shop);
  if (!token) return { ...base, shop, error: "no_token" };

  // Products that legitimately may have CONTINUE: those with a pre-order batch
  // that is actually LIVE = enabled AND activated (selling plan registered). An
  // enabled-but-not-activated setting is NOT "on", so it isn't excluded. Keyed by
  // numeric product id.
  const [enabled, registry] = await Promise.all([
    prisma.preorderBatchSetting.findMany({ where: { enabled: true }, select: { supplierOrderId: true } }),
    getPreorderSellingPlanRegistryEntries(shop),
  ]);
  const activatedIds = new Set(registry.map((r) => r.supplierOrderId));
  const liveSupplierIds = enabled.map((e) => e.supplierOrderId).filter((id) => activatedIds.has(id));
  const orders = liveSupplierIds.length
    ? await prisma.supplierOrder.findMany({ where: { id: { in: liveSupplierIds } }, select: { productId: true } })
    : [];
  const excludeNums = new Set(orders.map((o) => numericId(o.productId ?? "")).filter(Boolean));
  let continueProductsTotal = 0;
  let continueVariantsTotal = 0;

  const gql = async (query: string, variables: Record<string, unknown>) => {
    const res = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
      method: "POST", headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": token }, body: JSON.stringify({ query, variables }),
    });
    if (!res.ok) throw new Error(`Shopify HTTP ${res.status}`);
    return res.json() as Promise<any>;
  };

  const affected: Array<{ productId: string; title: string; variantGids: string[]; variants: string[] }> = [];
  let scanned = 0;
  let cursor: string | null = null;
  const maxPages = opts.maxPages ?? 400;
  let truncated = false;
  for (let page = 0; page < maxPages; page++) {
    const j: any = await gql(
      `query($cursor:String){ products(first:40, after:$cursor){ pageInfo{ hasNextPage endCursor } nodes{ id title variants(first:100){ nodes{ id title inventoryPolicy } } } } }`,
      { cursor },
    ).catch((e) => { base.errors.push(`page ${page}: ${e instanceof Error ? e.message : String(e)}`); return null; });
    const conn: any = j?.data?.products;
    if (!conn) break;
    for (const p of conn.nodes ?? []) {
      scanned++;
      const continueVars = (p.variants?.nodes ?? []).filter((v: any) => String(v?.inventoryPolicy ?? "").toUpperCase() === "CONTINUE");
      if (continueVars.length) { continueProductsTotal++; continueVariantsTotal += continueVars.length; }
      if (excludeNums.has(numericId(p.id))) continue; // LIVE pre-order → leave alone
      if (continueVars.length) {
        affected.push({ productId: p.id, title: String(p.title ?? ""), variantGids: continueVars.map((v: any) => String(v.id)), variants: continueVars.map((v: any) => String(v.title ?? "")) });
      }
    }
    if (!conn.pageInfo?.hasNextPage) break;
    cursor = conn.pageInfo.endCursor;
    if (page === maxPages - 1) truncated = true;
    await new Promise((r) => setTimeout(r, 200));
  }

  const totalVariants = affected.reduce((a, x) => a + x.variantGids.length, 0);
  let changedProducts = 0, changedVariants = 0;
  if (apply) {
    for (const a of affected) {
      try {
        await setVariantsInventoryPolicy(shop, token, a.productId, a.variantGids, "DENY");
        changedProducts++; changedVariants += a.variantGids.length;
      } catch (e) {
        base.errors.push(`${a.title}: ${e instanceof Error ? e.message : String(e)}`);
      }
      await new Promise((r) => setTimeout(r, 120));
    }
  }

  return {
    ok: true, shop, excludedPreorderProducts: excludeNums.size, scannedProducts: scanned,
    continueProductsTotal, continueVariantsTotal,
    affected: affected.map((a) => ({ productId: a.productId, title: a.title, variants: a.variants })),
    totalVariants, applied: apply, changedProducts, changedVariants, errors: base.errors, truncated,
  };
}
