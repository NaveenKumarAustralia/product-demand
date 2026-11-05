import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { getOfflineToken } from "../preorder/preorder-fulfillment.server";

// Admin "Unload packing-list inventory" — cleanly reverses an ACCIDENTAL "Load
// inventory" for ONE packing list. It subtracts from Shopify exactly what that
// load pushed (the per-line `shopifyLoadedQtys` snapshot), then clears the
// loaded flags so the list shows un-loaded and can be loaded again later.
//
// Design (deliberately gentle on the single-instance server):
//   • THROTTLED — a short pause before every Shopify call, so a big list never
//     fires a burst that exhausts Shopify's rate budget or ties up the box.
//   • PER-LINE — one inventoryAdjustQuantities call per packing line, mirroring
//     how the load applied it. Avoids Shopify's 250-changes-per-call cap and
//     its "same variant twice in one call" rejection (a variant can appear on
//     two lines).
//   • RESUMABLE — each line's loaded flag is cleared the instant ITS call
//     succeeds, so if anything is interrupted you just open the same &apply=1
//     URL again and it finishes the remaining lines (never double-subtracts).
//
// Scope: a single ?packingId= / ?title= — it NEVER touches any other list.
// Dry-run by default; &apply=1 executes. COST (WAC) and any orders this load
// closed are reported but NOT auto-reverted.

const API_VERSION = "2025-10";
const HQ_LOCATION_NAME = "karma east hq";
const THROTTLE_MS = 300; // pause before each Shopify call

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

async function graphql<T>(shop: string, token: string, query: string, variables?: Record<string, unknown>): Promise<T | null> {
  await sleep(THROTTLE_MS); // throttle every call — keep well under rate limits
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

  // Resolve the ONE target list: by explicit id, or by an exact (unambiguous)
  // title/invoice match. Never a partial/broad match.
  let packingId = Number(idParam ?? "");
  if ((!Number.isFinite(packingId) || packingId <= 0) && titleParam) {
    const matches = await prisma.packingList.findMany({
      where: { OR: [{ title: { equals: titleParam, mode: "insensitive" } }, { invoiceNumber: { equals: titleParam, mode: "insensitive" } }] },
      select: { id: true, title: true, invoiceNumber: true },
    });
    if (matches.length === 1) packingId = matches[0].id;
    else if (matches.length > 1) return new Response(JSON.stringify({ error: `"${titleParam}" matches ${matches.length} lists — pass ?packingId= to be exact.`, matches }), { status: 400, headers: { "Content-Type": "application/json" } });
    else return new Response(JSON.stringify({ error: `No packing list titled "${titleParam}". Pass ?packingId=<id> instead.` }), { status: 404, headers: { "Content-Type": "application/json" } });
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

  // Build a per-line plan. Only lines that still carry a `shopifyLoadedQtys`
  // snapshot are pending — so a re-run naturally skips lines already done.
  type Item = { size: string; qty: number; inventoryItemId: string | null; currentAvailable: number | null; projectedAvailable: number | null; matched: boolean };
  type LinePlan = { line: number; productId: string; productTitle: string; items: Item[]; allMatched: boolean };
  const linePlans: LinePlan[] = [];
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
    const items: Item[] = sizes.map(([size, qty]) => {
      const v = variants.find((vr) => normalizeSize(vr.title) === normalizeSize(size)) ?? null;
      return { size, qty, inventoryItemId: v?.inventoryItemId ?? null, currentAvailable: v?.available ?? null, projectedAvailable: v?.available != null ? v.available - qty : null, matched: Boolean(v?.inventoryItemId) };
    });
    linePlans.push({ line: line.id, productId: line.productId, productTitle: line.productTitle, items, allMatched: items.every((i) => i.matched) });
  }

  // Orders linked to THIS list that are already closed — the load may have
  // closed them. Reported only; not auto-reopened.
  const closedLinked = await prisma.supplierOrder.findMany({
    where: { packingListId: packingId, status: "closed" },
    select: { id: true, productTitle: true, productId: true },
  }).catch(() => [] as Array<{ id: number; productTitle: string | null; productId: string | null }>);

  const totalUnits = linePlans.reduce((s, lp) => s + lp.items.filter((i) => i.matched).reduce((a, i) => a + i.qty, 0), 0);
  const totalSizes = linePlans.reduce((s, lp) => s + lp.items.filter((i) => i.matched).length, 0);
  const linesWithUnmatched = linePlans.filter((lp) => !lp.allMatched).map((lp) => ({ line: lp.line, productTitle: lp.productTitle, sizes: lp.items.filter((i) => !i.matched).map((i) => i.size) }));

  if (!apply) {
    return new Response(JSON.stringify({
      mode: "dry-run",
      packingList: { id: packingList.id, title: packingList.invoiceNumber || packingList.title, masterInventoryLoadedAt: packingList.masterInventoryLoadedAt, loadedBy: packingList.loadedBy },
      willSubtract: { lines: linePlans.filter((lp) => lp.allMatched).length, sizes: totalSizes, totalUnits, location: locationId },
      plan: linePlans.map((lp) => ({ line: lp.line, productTitle: lp.productTitle, items: lp.items })),
      warnings: {
        linesWithUnmatchedSizes: linesWithUnmatched.length ? linesWithUnmatched : undefined,
        costPushedLines: costPushedLines.length ? costPushedLines : undefined,
        closedLinkedOrders: closedLinked.length ? closedLinked : undefined,
        note: "Dry-run only — nothing changed. Re-open with &apply=1 to subtract the stock and clear the loaded flags. Throttled + resumable: if a run is interrupted, just open the &apply=1 URL again to finish. COST (WAC) and any closed orders are NOT auto-reverted.",
      },
    }, null, 2), { headers: { "Content-Type": "application/json", "Cache-Control": "no-store" } });
  }

  // APPLY — one throttled call per line; clear each line's flag on success.
  let revertedSizes = 0;
  let revertedUnits = 0;
  let linesCleared = 0;
  const lineErrors: Array<{ line: number; productTitle: string; message: string }> = [];

  for (const lp of linePlans) {
    if (!lp.allMatched) {
      // Can't safely revert a line with an unmatched size — leave it, report it.
      lineErrors.push({ line: lp.line, productTitle: lp.productTitle, message: `Unmatched size(s): ${lp.items.filter((i) => !i.matched).map((i) => i.size).join(", ")}` });
      continue;
    }
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
        changes: lp.items.map((i) => ({ delta: -i.qty, inventoryItemId: i.inventoryItemId, locationId })),
      },
    });
    const userErrors = json?.data?.inventoryAdjustQuantities?.userErrors ?? [];
    const group = json?.data?.inventoryAdjustQuantities?.inventoryAdjustmentGroup;
    if (userErrors.length || !group) {
      lineErrors.push({ line: lp.line, productTitle: lp.productTitle, message: userErrors.map((e: { message?: string }) => e.message || "userError").join("; ") || "Shopify returned no adjustment group" });
      continue; // leave this line's flag set so a retry resumes here
    }
    await prisma.packingListLine.update({ where: { id: lp.line }, data: { shopifyLoadedQtys: {} } });
    revertedSizes += lp.items.length;
    revertedUnits += lp.items.reduce((s, i) => s + i.qty, 0);
    linesCleared += 1;
  }

  const fullyDone = lineErrors.length === 0;
  if (fullyDone) {
    await prisma.packingList.update({ where: { id: packingId }, data: { masterInventoryLoadedAt: null, loadedBy: null } });
  }

  try {
    await prisma.activityLog.create({
      data: {
        userName: actor.name ?? "Unknown", action: "Unloaded from Shopify", entity: "Packing List",
        entityId: String(packingId), entityName: packingList.invoiceNumber || packingList.title || `#${packingId}`,
        field: "Reverted accidental inventory load", toValue: `${revertedSizes} size(s), ${revertedUnits} unit(s) subtracted${fullyDone ? "" : ` (${lineErrors.length} line(s) failed)`}`,
      },
    });
  } catch { /* logging is best-effort */ }

  return new Response(JSON.stringify({
    mode: "apply",
    ok: fullyDone,
    reverted: { lines: linesCleared, sizes: revertedSizes, totalUnits: revertedUnits },
    failedLines: lineErrors.length ? lineErrors : undefined,
    retryHint: lineErrors.length ? "Some lines weren't reverted — the ones that succeeded are already cleared, so just open the same &apply=1 URL again to finish the rest." : undefined,
    warnings: {
      costPushedLines: costPushedLines.length ? costPushedLines : undefined,
      closedLinkedOrders: closedLinked.length ? closedLinked : undefined,
      note: fullyDone
        ? "Stock subtracted and loaded flags cleared for THIS list only. COST (WAC) was not restored and any closed orders were not reopened — tell me if you need those too."
        : "PARTIAL — see failedLines. Succeeded lines are done and won't be touched again; re-run the same URL to finish.",
    },
  }, null, 2), { status: fullyDone ? 200 : 207, headers: { "Content-Type": "application/json", "Cache-Control": "no-store" } });
};
