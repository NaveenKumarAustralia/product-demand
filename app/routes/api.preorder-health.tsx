import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { getOfflineToken, getAvailableAtLocation } from "../preorder/preorder-fulfillment.server";
import { getPreorderLocationSettings, locationForMarket } from "../preorder/preorder-locations.server";

// Admin "Pre-order Health" page. Shows, at a glance, every pre-order order and
// its real state (HELD / RELEASED-ready-to-pick / RELEASED-but-no-stock /
// FULFILLED), plus allocation failures — and one-click fixes. Read-only except
// via the Fix buttons, which call /api/preorder-reprocess-order actions.
const esc = (s: unknown) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c] as string));
const num = (s: unknown) => String(s ?? "").replace(/\D/g, "");
const groupKey = (r: { supplierOrderId: number; variantId: string; market: string }) => `${r.supplierOrderId}::${r.variantId}::${r.market}`;

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return new Response("Admin only.", { status: 403, headers: { "Content-Type": "text/plain" } });

  const session = await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true } });
  const shop = session?.shop ?? "";

  const reservations = await prisma.preorderReservation.findMany({
    where: { status: "reserved" },
    select: { shopifyOrderId: true, shopifyOrderName: true, variantId: true, variantTitle: true, supplierOrderId: true, market: true, quantity: true, reservedAt: true, readyAt: true, releasedAt: true, fulfilledAt: true },
    orderBy: { reservedAt: "desc" },
  });

  type OrderAgg = { orderId: string; name: string; market: string; sizes: string[]; batches: Set<number>; reservedAt: Date | null; held: boolean; released: boolean; fulfilled: boolean; variantIds: Set<string>; stuck: boolean };
  const byOrder = new Map<string, OrderAgg>();
  for (const r of reservations) {
    const o = byOrder.get(r.shopifyOrderId) ?? { orderId: r.shopifyOrderId, name: r.shopifyOrderName || r.shopifyOrderId, market: r.market, sizes: [], batches: new Set<number>(), reservedAt: r.reservedAt, held: false, released: false, fulfilled: false, variantIds: new Set<string>(), stuck: false };
    if (r.variantTitle && !o.sizes.includes(r.variantTitle)) o.sizes.push(r.variantTitle);
    o.batches.add(r.supplierOrderId);
    o.variantIds.add(r.variantId);
    if (r.fulfilledAt) o.fulfilled = true;
    else if (r.readyAt) o.released = true;
    else o.held = true;
    if (r.reservedAt && (!o.reservedAt || r.reservedAt < o.reservedAt)) o.reservedAt = r.reservedAt;
    byOrder.set(r.shopifyOrderId, o);
  }

  // For RELEASED-but-unfulfilled orders, compute genuine spare to flag "no stock".
  const stillHeld = reservations.filter((r) => !r.readyAt && !r.releasedAt && !r.fulfilledAt);
  const heldByGroup = new Map<string, number>();
  for (const r of stillHeld) heldByGroup.set(groupKey(r), (heldByGroup.get(groupKey(r)) ?? 0) + r.quantity);
  const token = shop ? await getOfflineToken(shop) : null;
  if (token) {
    const locations = await getPreorderLocationSettings();
    const spareCache = new Map<string, number>();
    const releasedRows = reservations.filter((r) => r.readyAt && !r.releasedAt && !r.fulfilledAt);
    for (const r of releasedRows) {
      const locationId = locationForMarket(locations, r.market as "AU" | "USA");
      if (!locationId) continue;
      const key = `${num(r.variantId)}::${locationId}`;
      let spare = spareCache.get(key);
      if (spare === undefined) {
        try {
          const [onHand, committed] = await Promise.all([
            getAvailableAtLocation(shop, token, r.variantId, locationId, "on_hand"),
            getAvailableAtLocation(shop, token, r.variantId, locationId, "committed"),
          ]);
          spare = onHand - committed + (heldByGroup.get(groupKey(r)) ?? 0);
        } catch { spare = 1; }
        spareCache.set(key, spare);
      }
      if (spare <= 0) { const o = byOrder.get(r.shopifyOrderId); if (o) o.stuck = true; }
    }
  }

  const orders = Array.from(byOrder.values());
  const stuck = orders.filter((o) => o.stuck && !o.fulfilled);
  const heldOrders = orders.filter((o) => o.held && !o.released && !o.fulfilled);
  const releasedOk = orders.filter((o) => o.released && !o.stuck && !o.fulfilled);

  const failLogs = await prisma.activityLog.findMany({
    where: { action: { in: ["preorder_allocation_failed"] }, createdAt: { gte: new Date(Date.now() - 30 * 86400000) } },
    orderBy: { createdAt: "desc" }, take: 50, select: { entityName: true, entityId: true, toValue: true, createdAt: true },
  }).catch(() => []);

  const fmt = (d: Date | null) => d ? new Date(d).toLocaleString("en-AU", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" }) : "—";
  const row = (o: OrderAgg, tag: string) => `<tr>
    <td><b>${esc(o.name)}</b></td><td>${esc(o.sizes.join(", "))}</td><td>${esc(o.market)}</td>
    <td>${Array.from(o.batches).map((b) => `#${b}`).join(", ")}</td><td>${fmt(o.reservedAt)}</td>
    <td>${tag}</td>
    <td><button onclick="fix('${esc(o.orderId)}', this)">Reprocess</button></td></tr>`;

  const html = `<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Pre-order Health</title>
  <style>
    body{font-family:ui-sans-serif,system-ui,Arial;margin:0;padding:20px;background:#f8fafc;color:#0f172a}
    h1{font-size:20px;margin:0 0 4px} .sub{color:#64748b;font-size:13px;margin-bottom:16px}
    .cards{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:18px}
    .card{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:12px 16px;min-width:120px}
    .card .n{font-size:26px;font-weight:800} .card.bad .n{color:#b91c1c} .card .l{font-size:11px;text-transform:uppercase;color:#64748b;font-weight:700}
    .bar{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:18px}
    button{background:#0f766e;color:#fff;border:0;border-radius:7px;padding:8px 12px;font-size:13px;font-weight:700;cursor:pointer}
    button.sec{background:#fff;color:#0f766e;border:1px solid #0f766e}
    table{width:100%;border-collapse:collapse;background:#fff;border:1px solid #e2e8f0;border-radius:10px;overflow:hidden;margin-bottom:22px;font-size:13px}
    th,td{text-align:left;padding:8px 10px;border-bottom:1px solid #f1f5f9} th{background:#f8fafc;font-size:11px;text-transform:uppercase;color:#64748b}
    h2{font-size:15px;margin:18px 0 8px} .tag{padding:2px 8px;border-radius:999px;font-size:11px;font-weight:700}
    .t-held{background:#dbeafe;color:#1e40af}.t-rel{background:#dcfce7;color:#166534}.t-stuck{background:#fee2e2;color:#991b1b}
    .empty{color:#94a3b8;padding:14px}
  </style></head><body>
  <h1>Pre-order Health</h1>
  <div class="sub">${esc(shop)} · live snapshot of every pre-order order and its hold/release state.</div>
  <div class="cards">
    <div class="card"><div class="n">${heldOrders.length}</div><div class="l">Held (awaiting stock)</div></div>
    <div class="card"><div class="n">${releasedOk.length}</div><div class="l">Released · ready to pick</div></div>
    <div class="card ${stuck.length ? "bad" : ""}"><div class="n">${stuck.length}</div><div class="l">Released · NO stock</div></div>
    <div class="card ${failLogs.length ? "bad" : ""}"><div class="n">${failLogs.length}</div><div class="l">Allocation failures (30d)</div></div>
  </div>
  <div class="bar">
    <button onclick="act('rehold-premature', this)">Pull back released-but-no-stock &amp; hold</button>
    <button class="sec" onclick="act('rehold-all', this)">Re-hold any unheld</button>
    <button class="sec" onclick="location.reload()">Refresh</button>
  </div>
  <div id="msg"></div>

  <h2 style="color:#991b1b">⚠ Released but no stock — needs attention (${stuck.length})</h2>
  ${stuck.length ? `<table><tr><th>Order</th><th>Sizes</th><th>Market</th><th>Batch</th><th>Reserved</th><th>State</th><th></th></tr>${stuck.map((o) => row(o, '<span class="tag t-stuck">released · no stock</span>')).join("")}</table>` : '<div class="empty">None — nothing stuck. 🎉</div>'}

  <h2>Held — waiting for stock (${heldOrders.length})</h2>
  ${heldOrders.length ? `<table><tr><th>Order</th><th>Sizes</th><th>Market</th><th>Batch</th><th>Reserved</th><th>State</th><th></th></tr>${heldOrders.slice(0, 200).map((o) => row(o, '<span class="tag t-held">held</span>')).join("")}</table>` : '<div class="empty">None.</div>'}

  <h2>Released — ready to pick (${releasedOk.length})</h2>
  ${releasedOk.length ? `<table><tr><th>Order</th><th>Sizes</th><th>Market</th><th>Batch</th><th>Reserved</th><th>State</th><th></th></tr>${releasedOk.slice(0, 200).map((o) => row(o, '<span class="tag t-rel">released</span>')).join("")}</table>` : '<div class="empty">None.</div>'}

  <h2>Allocation failures (couldn't reserve — possible ghosts) (${failLogs.length})</h2>
  ${failLogs.length ? `<table><tr><th>Order</th><th>Reason</th><th>When</th><th></th></tr>${failLogs.map((f) => `<tr><td><b>${esc(f.entityName || f.entityId)}</b></td><td>${esc(f.toValue)}</td><td>${fmt(f.createdAt)}</td><td><button onclick="fix('${esc(num(f.entityId))}', this)">Reprocess</button></td></tr>`).join("")}</table>` : '<div class="empty">None in the last 30 days.</div>'}

  <script>
    async function call(u, btn){ const t=btn?btn.textContent:''; if(btn){btn.disabled=true;btn.textContent='Working…';} try{ const r=await fetch(u,{credentials:'same-origin'}); const j=await r.json(); document.getElementById('msg').innerHTML='<div class="card" style="margin-bottom:14px">'+JSON.stringify(j)+'</div>'; setTimeout(()=>location.reload(), 1200);}catch(e){ document.getElementById('msg').textContent=String(e);} finally{ if(btn){btn.disabled=false;btn.textContent=t;} } }
    function act(a, btn){ call('/api/preorder-reprocess-order?action='+a, btn); }
    function fix(id, btn){ call('/api/preorder-reprocess-order?orderId='+encodeURIComponent(id), btn); }
  </script>
  </body></html>`;

  return new Response(html, { headers: { "Content-Type": "text/html; charset=utf-8", "Cache-Control": "no-store" } });
};
