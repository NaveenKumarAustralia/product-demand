import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { getOfflineToken } from "../preorder/preorder-fulfillment.server";

// Admin "Unload packing-list inventory" — reverses an ACCIDENTAL "Load inventory"
// for ONE packing list. It subtracts from Shopify exactly what that load pushed
// (the per-line `shopifyLoadedQtys` snapshot), then clears the loaded flags so
// the list shows un-loaded and can be loaded again cleanly later.
//
// Scope: a single ?packingId= — it NEVER touches any other packing list.
// Dry-run by default; pass &apply=1 to actually adjust Shopify + clear flags.
//
// What it DOES reverse: the Shopify "available" quantity bump + the loaded-qty
// flags + the master "loaded at / by" stamp.
// What it does NOT touch (reported instead, because undoing them is lossy or
// risky): weighted-average COST pushes, and any supplier/restock orders this
// load may have closed. Those are surfaced so you can decide.

const API_VERSION = "2025-10";
const HQ_LOCATION_NAME = "karma east hq";

async function graphql<T>(shop: string, token: string, query: string, variables?: Record<string, unknown>): Promise<T | null> {
  try {
    const res = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": token },
      body: JSON.stringify({ query, variables }),
      signal: AbortSignal.timeout(15_000),
    });
    if (!res.ok) return null;
    return (await res.json()) as T;
  } catch {
    return null;
  }
}

const normalizeSize = (value: string) => value.trim().toLowerCase().replace(/\s+/g, "").replace(/-/g, "/");

function normalizeQtys(value: unknown): Record<string, number> {
  const out: Record<string, number> = {};
  if (value && typeof value === "object") {
    for (const [k, v] of Object.entries(value as Record<string, unknown>)) {
      const n = Number(v);
      if (Number.isFinite(n)) out[k] = n;
    }
  }
  return out;
}

type VariantRow = { id: string; title: string; inventoryItemId: string | null; available: number | null };

async function getVariants(shop: string, token: string, productId: string): Promise<VariantRow[]> {
  const json = await graphql<any>(shop, token, `
    query UnloadVariants($id: ID!) {
      product(id: $id) {
        variants(first: 100) {
          nodes {
            id title
            selectedOptions { name value }
            inventoryItem { id inventoryLevels(first: 20) { nodes { quantities(names: ["available"]) { name quantity } } } }
          }
        }
      }
    }
  `, { id: productId.startsWith("gid://") ? productId : `gid://shopify/Product/${productId.replace(/\D/g, "")}` });
  const nodes: any[] = json?.data?.product?.variants?.nodes ?? [];
  const isFreeSize = nodes.length === 1 && (() => {
    const opts = (nodes[0]?.selectedOptions ?? []) as { name?: string; value?: string }[];
    if (!opts.length) return true;
    if (opts.some((o) => (o?.name ?? "").trim().toLowerCase() === "size")) return false;
    return opts.every((o) => (o?.name ?? "") === "Title" && (o?.value ?? "") === "Default Title");
  })();
  return nodes.map((v) => {
    let available: number | null = null;
    for (const lvl of v.inventoryItem?.inventoryLevels?.nodes ?? []) {
      for (const q of lvl.quantities ?? []) {
        if (q?.name === "available" && Number.isFinite(Number(q.quantity))) available = (available ?? 0) + Number(q.quantity);
      }
    }
    return { id: String(v.id ?? ""), title: isFreeSize ? "Free Size" : String(v.title ?? ""), inventoryItemId: v.inventoryItem?.id ? String(v.inventoryItem.id) : null, available };
  }).filter((v) => v.id && v.title);
}

async function getHqLocationId(shop: string, token: string): Promise<string | null> {
  const json = await graphql<any>(shop, token, `query { locations(first: 50) { nodes { id name isActive } } }`);
  const locations: Array<{ id?: string; name?: string; isActive?: boolean }> = json?.data?.locations?.nodes ?? [];
  const norm = (s: unknown) => String(s ?? "").trim().toLowerCase();
  const chosen = locations.find((l) => norm(l.name) === HQ_LOCATION_NAME)
    ?? locations.find((l) => norm(l.name).includes("karma east"))
    ?? locations.find((l) => l.isActive)
    ?? locations[0];
  return chosen?.id ?? null;
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return new Response(JSON.stringify({ error: "Admin only." }), { status: 403, headers: { "Content-Type": "application/json" } });

  const url = new URL(request.url);
  const idParam = url.searchParams.get("packingId");
  const titleParam = (url.searchParams.get("title") ?? "").trim();
  const apply = url.searchParams.get("apply") === "1";

  // Resolve the ONE target list: by explicit id, or by an exact title/invoice
  // match that is unambiguous (exactly one list). Never a partial/broad match.
  let packingId = Number(idParam ?? "");
  if ((!Number.isFinite(packingId) || packingId <= 0) && titleParam) {
    const matches = await prisma.packingList.findMany({
      where: { OR: [{ title: { equals: titleParam, mode: "insensitive" } }, { invoiceNumber: { equals: titleParam, mode: "insensitive" } }] },
      select: { id: true, title: true, invoiceNumber: true },
    });
    if (matches.length === 1) packingId = matches[0].id;
    else if (matches.length > 1) {
      return new Response(JSON.stringify({ error: `"${titleParam}" matches ${matches.length} lists — pass ?packingId= to be exact.`, matches }), { status: 400, headers: { "Content-Type": "application/json" } });
    } else {
      return new Response(JSON.stringify({ error: `No packing list titled "${titleParam}". Pass ?packingId=<id> instead.` }), { status: 404, headers: { "Content-Type": "application/json" } });
    }
  }
  if (!Number.isFinite(packingId) || packingId <= 0) {
    return new Response(JSON.stringify({ error: "Pass ?packingId=<id> (or ?title=<exact name>). Add &apply=1 to actually unload." }), { status: 400, headers: { "Content-Type": "application/json" } });
  }

  const packingList = await prisma.packingList.findUnique({
    where: { id: packingId },
    include: { lines: { orderBy: [{ sortOrder: "asc" }, { id: "asc" }] } },
  });
  if (!packingList) return new Response(JSON.stringify({ error: `No packing list #${packingId}.` }), { status: 404, headers: { "Content-Type": "application/json" } });

  const shop = (await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true } }))?.shop ?? "";
  const token = shop ? await getOfflineToken(shop) : null;
  if (!shop || !token) return new Response(JSON.stringify({ error: "No Shopify session." }), { status: 500, headers: { "Content-Type": "application/json" } });

  const locationId = await getHqLocationId(shop, token);
  if (!locationId) return new Response(JSON.stringify({ error: "Could not resolve the Karma East HQ location." }), { status: 500, headers: { "Content-Type": "application/json" } });

  type Plan = { line: number; productId: string; productTitle: string; size: string; qty: number; inventoryItemId: string | null; currentAvailable: number | null; projectedAvailable: number | null; matched: boolean };
  const plan: Plan[] = [];
  const costPushedLines: Array<{ line: number; productTitle: string }> = [];
  const variantCache = new Map<string, VariantRow[]>();

  for (const line of packingList.lines) {
    if ((line as { costPushedAt?: Date | null }).costPushedAt) costPushedLines.push({ line: line.id, productTitle: line.productTitle });
    if (!line.productId || line.isCustom) continue;
    const loaded = normalizeQtys(line.shopifyLoadedQtys); // exactly what was pushed to Shopify
    const sizes = Object.entries(loaded).filter(([, q]) => q > 0);
    if (!sizes.length) continue;
    if (!variantCache.has(line.productId)) variantCache.set(line.productId, await getVariants(shop, token, line.productId));
    const variants = variantCache.get(line.productId) ?? [];
    for (const [size, qty] of sizes) {
      const v = variants.find((vr) => normalizeSize(vr.title) === normalizeSize(size)) ?? null;
      plan.push({
        line: line.id, productId: line.productId, productTitle: line.productTitle, size, qty,
        inventoryItemId: v?.inventoryItemId ?? null,
        currentAvailable: v?.available ?? null,
        projectedAvailable: v?.available != null ? v.available - qty : null,
        matched: Boolean(v?.inventoryItemId),
      });
    }
  }

  // Supplier/restock orders linked to THIS list that are already closed — the
  // load may have closed them. Reported only; not auto-reopened.
  const closedLinked = await prisma.supplierOrder.findMany({
    where: { packingListId: packingId, status: "closed" },
    select: { id: true, productTitle: true, productId: true },
  }).catch(() => [] as Array<{ id: number; productTitle: string | null; productId: string | null }>);

  const matched = plan.filter((p) => p.matched);
  const unmatched = plan.filter((p) => !p.matched);
  const totalUnits = matched.reduce((s, p) => s + p.qty, 0);

  if (!apply) {
    return new Response(JSON.stringify({
      mode: "dry-run",
      packingList: { id: packingList.id, title: packingList.invoiceNumber || packingList.title, masterInventoryLoadedAt: packingList.masterInventoryLoadedAt, loadedBy: packingList.loadedBy },
      willSubtract: { sizes: matched.length, totalUnits, location: locationId },
      plan: matched,
      unmatchedSizes: unmatched,
      warnings: {
        costPushedLines: costPushedLines.length ? costPushedLines : undefined,
        closedLinkedOrders: closedLinked.length ? closedLinked : undefined,
        note: "Dry-run only — nothing changed. Re-open this URL with &apply=1 to subtract the stock and clear the loaded flags. COST (WAC) and any closed orders are NOT auto-reverted; see warnings.",
      },
    }, null, 2), { headers: { "Content-Type": "application/json", "Cache-Control": "no-store" } });
  }

  // APPLY: subtract the exact deltas we pushed, at the HQ location.
  let adjusted = 0;
  const errors: string[] = [];
  if (matched.length) {
    const json = await graphql<any>(shop, token, `
      mutation UnloadInventory($input: InventoryAdjustQuantitiesInput!) {
        inventoryAdjustQuantities(input: $input) {
          userErrors { field message }
          inventoryAdjustmentGroup { id }
        }
      }
    `, {
      input: {
        name: "available",
        reason: "correction",
        changes: matched.map((p) => ({ delta: -p.qty, inventoryItemId: p.inventoryItemId, locationId })),
      },
    });
    const userErrors = json?.data?.inventoryAdjustQuantities?.userErrors ?? [];
    if (userErrors.length || !json?.data?.inventoryAdjustQuantities?.inventoryAdjustmentGroup) {
      errors.push(...(userErrors.map((e: { message?: string }) => e.message || "Shopify userError")));
      if (!userErrors.length) errors.push("Shopify returned no adjustment group (nothing was changed).");
      return new Response(JSON.stringify({ mode: "apply", ok: false, errors, note: "Shopify rejected the adjustment — NO flags were cleared, nothing changed. Safe to retry." }, null, 2), { status: 502, headers: { "Content-Type": "application/json" } });
    }
    adjusted = matched.length;
  }

  // Clear the loaded-qty flags on the affected lines, and the master stamp.
  const affectedLineIds = Array.from(new Set(matched.map((p) => p.line)));
  for (const lineId of affectedLineIds) {
    await prisma.packingListLine.update({ where: { id: lineId }, data: { shopifyLoadedQtys: {} } });
  }
  await prisma.packingList.update({ where: { id: packingId }, data: { masterInventoryLoadedAt: null, loadedBy: null } });

  try {
    await prisma.activityLog.create({
      data: {
        userName: actor.name ?? "Unknown", action: "Unloaded from Shopify", entity: "Packing List",
        entityId: String(packingId), entityName: packingList.invoiceNumber || packingList.title || `#${packingId}`,
        field: "Reverted accidental inventory load", toValue: `${adjusted} size(s), ${totalUnits} unit(s) subtracted`,
      },
    });
  } catch { /* logging is best-effort */ }

  return new Response(JSON.stringify({
    mode: "apply",
    ok: true,
    reverted: { sizes: adjusted, totalUnits, linesCleared: affectedLineIds.length },
    unmatchedSizes: unmatched.length ? unmatched : undefined,
    warnings: {
      costPushedLines: costPushedLines.length ? costPushedLines : undefined,
      closedLinkedOrders: closedLinked.length ? closedLinked : undefined,
      note: "Stock subtracted and loaded flags cleared for THIS list only. COST (WAC) was not restored and any closed orders were not reopened — tell me if you need those too.",
    },
  }, null, 2), { headers: { "Content-Type": "application/json", "Cache-Control": "no-store" } });
};
