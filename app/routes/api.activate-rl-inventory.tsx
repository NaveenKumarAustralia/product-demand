import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { getOfflineToken } from "../preorder/preorder-fulfillment.server";

// Admin "Activate inventory at RENEWAL LOGISTICS (USA 3PL)" — one-time bulk job.
// Activates every inventory-TRACKED variant at the RENEWAL LOGISTICS location so
// the US 3PL can receive/restock returns. Newly activated levels start at 0.
//
// SAFE BY DESIGN:
//   • Dry-run by default — reports the count that WOULD be activated + the
//     resolved RL location. Changes nothing.
//   • &apply=1 executes. &apply=1&max=5 runs the mandated 5-variant test first.
//   • Resolves RENEWAL LOGISTICS BY NAME (never a hardcoded id).
//   • Only activates variants not already stocked at RL (idempotent) — a re-run
//     safely resumes and can never double-activate.
//   • Throttled + adaptive backoff on Shopify's throttleStatus.
//   • Only the RL location is sent to inventoryBulkToggleActivation, so no other
//     location's activation or quantity is touched. No prices/orders/fulfilment.
//   • Credentials come from the stored offline session server-side; no token is
//     ever printed.

const API_VERSION = "2025-10";
const RL_NAME_NEEDLE = "renewal"; // RENEWAL LOGISTICS, matched case-insensitively
const PAGE_SIZE = 250;
const BASE_THROTTLE_MS = 250;
const DEFAULT_MAX = 500;         // activations per apply run (keeps each request short)

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));
const norm = (s: unknown) => String(s ?? "").trim().toLowerCase();
const esc = (s: unknown) => String(s ?? "");

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return json({ error: "Admin only." }, 403);

  const url = new URL(request.url);
  const apply = url.searchParams.get("apply") === "1";
  const maxParam = Number(url.searchParams.get("max") ?? "");
  const maxActivations = apply ? (Number.isFinite(maxParam) && maxParam > 0 ? Math.floor(maxParam) : DEFAULT_MAX) : Infinity;
  let after = url.searchParams.get("after") || null;
  // Accumulators carried across resumed runs (so the final run shows a grand total).
  const accIn = {
    scanned: toInt(url.searchParams.get("accScanned")),
    would: toInt(url.searchParams.get("accWould")),
    already: toInt(url.searchParams.get("accAlready")),
    skipped: toInt(url.searchParams.get("accSkipped")),
    activated: toInt(url.searchParams.get("accActivated")),
    failed: toInt(url.searchParams.get("accFailed")),
  };

  const shop = (await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true } }))?.shop ?? "";
  const token = shop ? await getOfflineToken(shop) : null;
  if (!shop || !token) return json({ error: "No Shopify session." }, 500);

  // Per-request throttle state + GraphQL helper (adaptive on throttleStatus).
  let availableBudget = 2000;
  async function gql<T>(query: string, variables?: Record<string, unknown>): Promise<{ data?: T; errors?: Array<{ message?: string }>; extensions?: any } | null> {
    await sleep(availableBudget < 300 ? 1500 : BASE_THROTTLE_MS); // back off when the cost bucket is low
    try {
      const res = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
        method: "POST",
        headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": token! },
        body: JSON.stringify({ query, variables }),
        signal: AbortSignal.timeout(20_000),
      });
      if (!res.ok) return null;
      const body = (await res.json()) as any;
      const avail = body?.extensions?.cost?.throttleStatus?.currentlyAvailable;
      if (typeof avail === "number") availableBudget = avail;
      return body;
    } catch {
      return null;
    }
  }

  // 1) Resolve RENEWAL LOGISTICS strictly by name.
  const locJson = await gql<any>(`query { locations(first: 100, includeInactive: true) { nodes { id name isActive } } }`);
  const locations: Array<{ id?: string; name?: string; isActive?: boolean }> = locJson?.data?.locations?.nodes ?? [];
  if (!locations.length) return json({ error: "Could not list Shopify locations." }, 500);
  const rlMatches = locations.filter((l) => norm(l.name).includes(RL_NAME_NEEDLE));
  if (rlMatches.length === 0) return json({ error: `No location whose name contains "${RL_NAME_NEEDLE}". Locations found: ${locations.map((l) => l.name).join(", ")}` }, 404);
  if (rlMatches.length > 1) return json({ error: `"${RL_NAME_NEEDLE}" matched ${rlMatches.length} locations — too ambiguous to proceed.`, matches: rlMatches.map((l) => ({ id: l.id, name: l.name })) }, 400);
  const rl = rlMatches[0];
  const rlId = String(rl.id);

  // 2) Scan variants, classify, and (when applying) activate.
  const scanQuery = `
    query ScanVariants($after: String, $rl: ID!) {
      productVariants(first: ${PAGE_SIZE}, after: $after) {
        pageInfo { hasNextPage endCursor }
        nodes { id inventoryItem { id tracked inventoryLevel(locationId: $rl) { id } } }
      }
    }`;
  const activateMutation = `
    mutation ActivateAtLocation($inventoryItemId: ID!, $updates: [InventoryBulkToggleActivationInput!]!) {
      inventoryBulkToggleActivation(inventoryItemId: $inventoryItemId, inventoryItemUpdates: $updates) {
        inventoryItem { id }
        userErrors { field message code }
      }
    }`;

  const run = { scanned: 0, would: 0, already: 0, skipped: 0, activated: 0, failed: 0 };
  const failures: Array<{ inventoryItemId: string; variantId: string; message: string }> = [];
  const BUDGET_MS = (apply ? 45 : 40) * 1000;
  const started = Date.now();
  let done = false;
  let resumeAfter: string | null = null;
  let stopped = false;

  while (true) {
    const pageAfter = after; // cursor that fetched THIS page (resume point, idempotent)
    const scan = await gql<any>(scanQuery, { after, rl: rlId });
    if (!scan || !scan.data?.productVariants) {
      // Transient scan failure — stop cleanly; resuming from pageAfter is safe.
      resumeAfter = pageAfter;
      failures.push({ inventoryItemId: "-", variantId: "-", message: "Variant scan page failed (will resume on re-run)." });
      break;
    }
    const conn = scan.data.productVariants;
    for (const v of conn.nodes as Array<{ id?: string; inventoryItem?: { id?: string; tracked?: boolean; inventoryLevel?: { id?: string } | null } }>) {
      run.scanned += 1;
      const inv = v.inventoryItem;
      if (!inv?.id || inv.tracked !== true) { run.skipped += 1; continue; }  // not inventory-tracked
      if (inv.inventoryLevel) { run.already += 1; continue; }                 // already active at RL
      // Needs activation.
      if (!apply) { run.would += 1; continue; }
      if (run.activated >= maxActivations || Date.now() - started > BUDGET_MS) { stopped = true; break; }
      const res = await gql<any>(activateMutation, { inventoryItemId: inv.id, updates: [{ locationId: rlId, activate: true }] });
      const ue = res?.data?.inventoryBulkToggleActivation?.userErrors ?? [];
      if (!res || ue.length) {
        run.failed += 1;
        failures.push({ inventoryItemId: inv.id, variantId: String(v.id ?? ""), message: ue.map((e: { message?: string }) => e.message || "userError").join("; ") || "Shopify call failed" });
      } else {
        run.activated += 1;
      }
    }
    if (stopped) { resumeAfter = pageAfter; break; }          // re-do this page on resume (idempotent)
    if (!conn.pageInfo.hasNextPage) { done = true; break; }
    after = conn.pageInfo.endCursor;
    if (Date.now() - started > BUDGET_MS) { resumeAfter = after; break; } // time cap between pages
  }

  // Grand totals across resumed runs.
  const totals = {
    variantsScanned: accIn.scanned + run.scanned,
    wouldActivate: accIn.would + run.would,
    alreadyActiveAtRL: accIn.already + run.already,
    skippedNotTracked: accIn.skipped + run.skipped,
    activated: accIn.activated + run.activated,
    failed: accIn.failed + run.failed,
  };

  // Build a ready-to-click resume URL that carries the cursor + accumulators.
  let nextUrl: string | undefined;
  if (!done && resumeAfter != null) {
    const u = new URL(url.origin + url.pathname);
    if (apply) u.searchParams.set("apply", "1");
    if (apply && Number.isFinite(maxParam) && maxParam > 0) u.searchParams.set("max", String(Math.floor(maxParam)));
    u.searchParams.set("after", resumeAfter);
    u.searchParams.set("accScanned", String(totals.variantsScanned));
    u.searchParams.set("accWould", String(totals.wouldActivate));
    u.searchParams.set("accAlready", String(totals.alreadyActiveAtRL));
    u.searchParams.set("accSkipped", String(totals.skippedNotTracked));
    if (apply) { u.searchParams.set("accActivated", String(totals.activated)); u.searchParams.set("accFailed", String(totals.failed)); }
    nextUrl = u.toString();
  }

  if (apply && (run.activated > 0 || run.failed > 0)) {
    try {
      await prisma.activityLog.create({ data: {
        userName: actor.name ?? "Unknown", action: "Activate inventory @ RL", entity: "Inventory",
        entityId: rlId, entityName: esc(rl.name), field: "inventoryBulkToggleActivation",
        toValue: `activated ${run.activated}, failed ${run.failed} this run (${totals.activated} total)`,
      } });
    } catch { /* best-effort */ }
  }

  return json({
    mode: apply ? "apply" : "dry-run",
    location: { id: rlId, name: rl.name, isActive: rl.isActive },
    done,
    totals,
    thisRun: run,
    failures: failures.length ? failures.slice(0, 50) : undefined,
    nextUrl,
    note: !done
      ? (apply
        ? `Batch done (activated ${run.activated} this run). MORE REMAIN — open nextUrl to continue; it resumes safely and never re-activates.`
        : `Scan time-capped — open nextUrl to keep counting toward the grand total.`)
      : (apply
        ? `COMPLETE. ${totals.activated} variant(s) activated at ${rl.name} (started at 0 each), ${totals.failed} failed, ${totals.alreadyActiveAtRL} already active, ${totals.skippedNotTracked} not inventory-tracked. No other location was touched.`
        : `Dry-run COMPLETE — nothing changed. ${totals.wouldActivate} variant(s) WOULD be activated at ${rl.name}. ${totals.alreadyActiveAtRL} already active, ${totals.skippedNotTracked} not inventory-tracked, ${totals.variantsScanned} scanned. To TEST on 5: add &apply=1&max=5. For the full run after approval: &apply=1.`),
  }, done || apply ? 200 : 200);
};

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body, null, 2), { status, headers: { "Content-Type": "application/json", "Cache-Control": "no-store" } });
}
function toInt(v: string | null): number {
  const n = Number(v ?? "");
  return Number.isFinite(n) && n > 0 ? Math.floor(n) : 0;
}
