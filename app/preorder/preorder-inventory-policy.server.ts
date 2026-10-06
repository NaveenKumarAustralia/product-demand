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
