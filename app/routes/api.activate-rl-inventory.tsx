import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { getOfflineToken } from "../preorder/preorder-fulfillment.server";

// Admin "Activate inventory at RENEWAL LOGISTICS (USA 3PL)" — one-time bulk job.
// Activates every inventory-TRACKED variant at the RENEWAL LOGISTICS location so
// the US 3PL can receive/restock. Newly activated levels start at 0. Gets around
// Shopify admin's 50-products-at-a-time cap.
//
// MODES (all admin-only):
//   • (default)                 dry-run: count what WOULD activate + resolved RL.
//   • ?apply=1&max=5            live TEST on 5 variants (synchronous).
//   • ?apply=1&background=1     START the full run as a BACKGROUND job (one go).
//   • ?status=1                 watch the background job's progress.
//   • ?apply=1 (no background)  synchronous, resumable-via-nextUrl batch (manual).
//
// Safe by design: resolves RL strictly BY NAME; idempotent (skips variants
// already active at RL) so it can never double-activate; throttled with adaptive
// backoff on Shopify's throttleStatus; ONLY the RL location is sent to the
// mutation, so no other location's activation or quantity is touched. No prices,
// orders, fulfilment or shipping profiles. Credentials come from the stored
// offline session server-side; no token is printed.

const API_VERSION = "2025-10";
const RL_NAME_NEEDLE = "renewal"; // RENEWAL LOGISTICS, matched case-insensitively
const PAGE_SIZE = 250;
const BASE_THROTTLE_MS = 250;
const DEFAULT_MAX = 500; // activations per synchronous apply run

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));
const norm = (s: unknown) => String(s ?? "").trim().toLowerCase();

const SCAN_QUERY = `
  query ScanVariants($after: String, $rl: ID!) {
    productVariants(first: ${PAGE_SIZE}, after: $after) {
      pageInfo { hasNextPage endCursor }
      nodes { id inventoryItem { id tracked inventoryLevel(locationId: $rl) { id } } }
    }
  }`;
const ACTIVATE_MUTATION = `
  mutation ActivateAtLocation($inventoryItemId: ID!, $updates: [InventoryBulkToggleActivationInput!]!) {
    inventoryBulkToggleActivation(inventoryItemId: $inventoryItemId, inventoryItemUpdates: $updates) {
      inventoryItem { id }
      userErrors { field message code }
    }
  }`;

// ─── Background job state (one at a time; single Railway instance) ───────────
type JobState = {
  running: boolean; done: boolean; startedAt: string | null; finishedAt: string | null;
  location: { id: string; name: string } | null;
  scanned: number; activated: number; already: number; skipped: number; failed: number;
  lastError: string | null; recentFailures: Array<{ inventoryItemId: string; message: string }>;
};
let JOB: JobState = emptyJob();
function emptyJob(): JobState {
  return { running: false, done: false, startedAt: null, finishedAt: null, location: null, scanned: 0, activated: 0, already: 0, skipped: 0, failed: 0, lastError: null, recentFailures: [] };
}

// A throttled Admin GraphQL caller with adaptive backoff on the cost bucket.
type GqlResult = { data?: any; errors?: Array<{ message?: string }>; extensions?: any } | null;
function makeGql(shop: string, token: string): (query: string, variables?: Record<string, unknown>) => Promise<GqlResult> {
  let available = 2000;
  return async function gql(query: string, variables?: Record<string, unknown>): Promise<GqlResult> {
    await sleep(available < 300 ? 1500 : BASE_THROTTLE_MS);
    try {
      const res = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
        method: "POST",
        headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": token },
        body: JSON.stringify({ query, variables }),
        signal: AbortSignal.timeout(20_000),
      });
      if (!res.ok) return null;
      const body = (await res.json()) as any;
      const avail = body?.extensions?.cost?.throttleStatus?.currentlyAvailable;
      if (typeof avail === "number") available = avail;
      return body;
    } catch {
      return null;
    }
  };
}

async function runActivationJob(shop: string, token: string, rlId: string, rlName: string, actorName: string) {
  JOB = { ...emptyJob(), running: true, startedAt: new Date().toISOString(), location: { id: rlId, name: rlName } };
  const gql = makeGql(shop, token);
  let after: string | null = null;
  let consecutiveScanFails = 0;
  try {
    while (true) {
      const scan = await gql(SCAN_QUERY, { after, rl: rlId });
      const conn = scan?.data?.productVariants;
      if (!conn) {
        // Transient scan failure — back off and retry the same cursor a few times.
        if (++consecutiveScanFails > 5) { JOB.lastError = "Repeated scan failures; stopping. Re-run to resume (idempotent)."; break; }
        JOB.lastError = `Scan page failed (retry ${consecutiveScanFails})`;
        await sleep(3000);
        continue;
      }
      consecutiveScanFails = 0;
      for (const v of conn.nodes as Array<{ id?: string; inventoryItem?: { id?: string; tracked?: boolean; inventoryLevel?: { id?: string } | null } }>) {
        JOB.scanned += 1;
        const inv = v.inventoryItem;
        if (!inv?.id || inv.tracked !== true) { JOB.skipped += 1; continue; }
        if (inv.inventoryLevel) { JOB.already += 1; continue; }
        const res = await gql(ACTIVATE_MUTATION, { inventoryItemId: inv.id, updates: [{ locationId: rlId, activate: true }] });
        const ue = res?.data?.inventoryBulkToggleActivation?.userErrors ?? [];
        if (!res || ue.length) {
          JOB.failed += 1;
          if (JOB.recentFailures.length < 50) JOB.recentFailures.push({ inventoryItemId: inv.id, message: ue.map((e: { message?: string }) => e.message || "userError").join("; ") || "Shopify call failed" });
        } else {
          JOB.activated += 1;
        }
      }
      if (!conn.pageInfo.hasNextPage) { JOB.done = true; break; }
      after = conn.pageInfo.endCursor;
    }
  } catch (e) {
    JOB.lastError = e instanceof Error ? e.message : String(e);
  } finally {
    JOB.running = false;
    JOB.finishedAt = new Date().toISOString();
    try {
      await prisma.activityLog.create({ data: {
        userName: actorName, action: "Activate inventory @ RL (bulk)", entity: "Inventory",
        entityId: rlId, entityName: rlName, field: "inventoryBulkToggleActivation",
        toValue: `activated ${JOB.activated}, failed ${JOB.failed}, already ${JOB.already}, skipped ${JOB.skipped}${JOB.done ? "" : " (incomplete)"}`,
      } });
    } catch { /* best-effort */ }
  }
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return json({ error: "Admin only." }, 403);

  const url = new URL(request.url);

  // Watch the background job (no Shopify calls needed).
  if (url.searchParams.get("status") === "1") {
    return json({ mode: "status", job: JOB, note: JOB.running ? "Running — refresh to watch. Done when job.done is true." : (JOB.done ? "Finished." : "No job has run yet (or last run stopped — see job.lastError).") });
  }

  const apply = url.searchParams.get("apply") === "1";
  const background = url.searchParams.get("background") === "1";
  const maxParam = Number(url.searchParams.get("max") ?? "");
  const maxActivations = apply ? (Number.isFinite(maxParam) && maxParam > 0 ? Math.floor(maxParam) : DEFAULT_MAX) : Infinity;
  let after = url.searchParams.get("after") || null;
  const accIn = {
    scanned: toInt(url.searchParams.get("accScanned")), would: toInt(url.searchParams.get("accWould")),
    already: toInt(url.searchParams.get("accAlready")), skipped: toInt(url.searchParams.get("accSkipped")),
    activated: toInt(url.searchParams.get("accActivated")), failed: toInt(url.searchParams.get("accFailed")),
  };

  const shop = (await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true } }))?.shop ?? "";
  const token = shop ? await getOfflineToken(shop) : null;
  if (!shop || !token) return json({ error: "No Shopify session." }, 500);

  const gql = makeGql(shop, token);

  // Resolve RENEWAL LOGISTICS strictly by name.
  const locJson = await gql(`query { locations(first: 100, includeInactive: true) { nodes { id name isActive } } }`);
  const locations: Array<{ id?: string; name?: string; isActive?: boolean }> = locJson?.data?.locations?.nodes ?? [];
  if (!locations.length) return json({ error: "Could not list Shopify locations." }, 500);
  const rlMatches = locations.filter((l) => norm(l.name).includes(RL_NAME_NEEDLE));
  if (rlMatches.length === 0) return json({ error: `No location whose name contains "${RL_NAME_NEEDLE}". Found: ${locations.map((l) => l.name).join(", ")}` }, 404);
  if (rlMatches.length > 1) return json({ error: `"${RL_NAME_NEEDLE}" matched ${rlMatches.length} locations — too ambiguous.`, matches: rlMatches.map((l) => ({ id: l.id, name: l.name })) }, 400);
  const rl = rlMatches[0];
  const rlId = String(rl.id);

  // Background full run — one click, server works through the whole catalogue.
  if (apply && background) {
    if (JOB.running) return json({ started: false, message: "A background activation is already running — watch ?status=1.", job: JOB });
    void runActivationJob(shop, token, rlId, String(rl.name), actor.name ?? "Unknown");
    return json({
      started: true,
      location: { id: rlId, name: rl.name },
      message: "Background activation STARTED. It will work through the whole catalogue (throttled, idempotent). Watch progress at the statusUrl until job.done is true.",
      statusUrl: `https://${url.host}${url.pathname}?status=1`,
    });
  }

  // Synchronous path: dry-run (default) or a capped live apply (e.g. the 5-test).
  const scanQuery = SCAN_QUERY;
  const run = { scanned: 0, would: 0, already: 0, skipped: 0, activated: 0, failed: 0 };
  const failures: Array<{ inventoryItemId: string; variantId: string; message: string }> = [];
  const BUDGET_MS = (apply ? 45 : 40) * 1000;
  const started = Date.now();
  let done = false;
  let resumeAfter: string | null = null;
  let stopped = false;

  while (true) {
    const pageAfter = after;
    const scan = await gql(scanQuery, { after, rl: rlId });
    const conn = scan?.data?.productVariants;
    if (!conn) { resumeAfter = pageAfter; failures.push({ inventoryItemId: "-", variantId: "-", message: "Scan page failed (will resume on re-run)." }); break; }
    for (const v of conn.nodes as Array<{ id?: string; inventoryItem?: { id?: string; tracked?: boolean; inventoryLevel?: { id?: string } | null } }>) {
      run.scanned += 1;
      const inv = v.inventoryItem;
      if (!inv?.id || inv.tracked !== true) { run.skipped += 1; continue; }
      if (inv.inventoryLevel) { run.already += 1; continue; }
      if (!apply) { run.would += 1; continue; }
      if (run.activated >= maxActivations || Date.now() - started > BUDGET_MS) { stopped = true; break; }
      const res = await gql(ACTIVATE_MUTATION, { inventoryItemId: inv.id, updates: [{ locationId: rlId, activate: true }] });
      const ue = res?.data?.inventoryBulkToggleActivation?.userErrors ?? [];
      if (!res || ue.length) { run.failed += 1; failures.push({ inventoryItemId: inv.id, variantId: String(v.id ?? ""), message: ue.map((e: { message?: string }) => e.message || "userError").join("; ") || "Shopify call failed" }); }
      else run.activated += 1;
    }
    if (stopped) { resumeAfter = pageAfter; break; }
    if (!conn.pageInfo.hasNextPage) { done = true; break; }
    after = conn.pageInfo.endCursor;
    if (Date.now() - started > BUDGET_MS) { resumeAfter = after; break; }
  }

  const totals = {
    variantsScanned: accIn.scanned + run.scanned, wouldActivate: accIn.would + run.would,
    alreadyActiveAtRL: accIn.already + run.already, skippedNotTracked: accIn.skipped + run.skipped,
    activated: accIn.activated + run.activated, failed: accIn.failed + run.failed,
  };

  let nextUrl: string | undefined;
  if (!done && resumeAfter != null) {
    const u = new URL(`https://${url.host}${url.pathname}`);
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

  return json({
    mode: apply ? "apply" : "dry-run",
    location: { id: rlId, name: rl.name, isActive: rl.isActive },
    done, totals, thisRun: run,
    failures: failures.length ? failures.slice(0, 50) : undefined,
    nextUrl,
    note: !done
      ? (apply ? `Did ${run.activated} this run — open nextUrl to continue (resumes safely).` : `Scan time-capped — open nextUrl to keep counting, OR just note it's ~all tracked variants.`)
      : (apply ? `COMPLETE. ${totals.activated} activated at ${rl.name}, ${totals.failed} failed, ${totals.alreadyActiveAtRL} already active, ${totals.skippedNotTracked} not tracked.`
        : `Dry-run complete. ${totals.wouldActivate} would activate at ${rl.name}. For the full run in one go: &apply=1&background=1 then watch ?status=1. Test on 5 first: &apply=1&max=5.`),
  }, 200);
};

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body, null, 2), { status, headers: { "Content-Type": "application/json", "Cache-Control": "no-store" } });
}
function toInt(v: string | null): number {
  const n = Number(v ?? "");
  return Number.isFinite(n) && n > 0 ? Math.floor(n) : 0;
}
