import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { listAffectedForFollowup } from "../preorder/preorder-missed-capture.server";

// Admin, staff-facing HTML report: EVERY order (AU + USA) that bought a live
// pre-order variant WITHOUT the selling plan (Shop Pay / PayPal / express / any
// no-plan path) and is still out of stock — i.e. the customer was NOT told it's a
// pre-order at checkout. Grouped by order, with a ready-to-send "wait or cancel"
// message per customer so staff can follow up.
//   GET /api/preorder-followup?days=90            (all markets)
//   GET /api/preorder-followup?days=90&market=usa (or au)
const esc = (s: unknown) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return new Response("Admin only.", { status: 403 });

  const url = new URL(request.url);
  const days = Math.max(1, Math.min(120, Math.floor(Number(url.searchParams.get("days")) || 90)));
  const marketParam = (url.searchParams.get("market") || "all").toLowerCase();
  const marketFilter = marketParam === "usa" ? "USA" : marketParam === "au" ? "AU" : "all";

  const session = await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true } });
  const shop = session?.shop ?? "";

  const { scannedOrders, affected, liveBatchCount, variantsTracked, queryErrors, truncated, oldestScannedIso } = await listAffectedForFollowup({ days });
  const shownLines = marketFilter === "all" ? affected : affected.filter((a) => a.market === marketFilter);
  const usaCount = new Set(affected.filter((a) => a.market === "USA").map((a) => a.order)).size;
  const auCount = new Set(affected.filter((a) => a.market === "AU").map((a) => a.order)).size;

  // Group the affected lines by order so each customer is one row.
  const byOrder = new Map<string, { order: string; orderId: string; email: string | null; customerName: string | null; market: string; lines: Array<{ product: string | null; size: string | null; qty: number; dispatch: string | null }> }>();
  for (const a of shownLines) {
    const g = byOrder.get(a.order) ?? { order: a.order, orderId: a.orderId, email: a.email, customerName: a.customerName, market: a.market, lines: [] };
    g.lines.push({ product: a.product, size: a.size, qty: a.qty, dispatch: a.dispatch });
    byOrder.set(a.order, g);
  }
  const orders = Array.from(byOrder.values());

  // Independent cross-check: which of these orders has our system ALREADY reserved
  // (captured + held)? A reservation is separate proof the order is a genuine
  // pre-order. "Not captured" = it slipped past auto-capture and still needs it.
  const orderIds = orders.map((o) => o.orderId).filter(Boolean);
  const reserved = orderIds.length
    ? await prisma.preorderReservation.findMany({ where: { shopifyOrderId: { in: orderIds }, status: "reserved" }, select: { shopifyOrderId: true } }).catch(() => [])
    : [];
  const capturedSet = new Set(reserved.map((r) => r.shopifyOrderId));
  const heldCount = orders.filter((o) => capturedSet.has(o.orderId)).length;
  const oldestScanned = oldestScannedIso ? new Date(oldestScannedIso).toISOString().slice(0, 10) : null;

  const messageFor = (g: (typeof orders)[number]) => {
    const name = (g.customerName ?? "").trim().split(" ")[0] || "there";
    const itemsText = g.lines.map((l) => `${l.product ?? "item"}${l.size ? ` (size ${l.size})` : ""}${l.dispatch ? ` — estimated dispatch ${l.dispatch}` : ""}`).join("; ");
    return `Hi ${name}, thank you for your order ${g.order} with Karma East! We wanted to let you know that the following item(s) in your order are pre-order pieces, currently being made: ${itemsText}. Any in-stock items are already on their way. Please reply to let us know if you're happy to wait for the pre-order item(s), or if you'd prefer we cancel and refund just that part of your order. Thank you so much for your patience — The Karma East Team`;
  };

  const rows = orders.map((g) => {
    const adminUrl = shop && g.orderId ? `https://${shop}/admin/orders/${g.orderId}` : "";
    const items = g.lines.map((l) => `<li>${esc(l.product ?? "item")}${l.size ? ` &middot; <strong>size ${esc(l.size)}</strong>` : ""} &times; ${esc(l.qty)}${l.dispatch ? ` <span class="disp">ships ~ ${esc(l.dispatch)}</span>` : ""}</li>`).join("");
    const msg = messageFor(g);
    const flag = g.market === "USA" ? "🇺🇸 USA" : "🇦🇺 AU";
    const held = capturedSet.has(g.orderId);
    const badge = held ? `<span class="badge ok" title="Our system has reserved &amp; held this order — confirmed pre-order">✓ holding</span>` : `<span class="badge warn" title="Not yet reserved by the system — run the capture-missed tool">⚠ not captured</span>`;
    return `
    <tr>
      <td class="ord">${adminUrl ? `<a href="${esc(adminUrl)}" target="_blank" rel="noopener">${esc(g.order)}</a>` : esc(g.order)}<div class="mkt">${flag}</div><div>${badge}</div></td>
      <td>${esc(g.customerName ?? "—")}<br><span class="email">${esc(g.email ?? "no email on order")}</span></td>
      <td><ul class="items">${items}</ul></td>
      <td class="msgcell">
        <button class="copy" data-msg="${esc(msg)}" onclick="copyMsg(this)">Copy message</button>
        <details><summary>preview</summary><div class="msg">${esc(msg)}</div></details>
      </td>
    </tr>`;
  }).join("");

  const scopeLabel = marketFilter === "all" ? "AU + USA" : marketFilter === "USA" ? "USA only" : "AU only";

  const html = `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Pre-order follow-ups</title>
  <style>
    :root{--teal:#006061;--ink:#1c2222;--muted:#6b7674;--line:#e3e6e5;--bg:#f6f5f2}
    *{box-sizing:border-box}
    body{margin:0;background:var(--bg);color:var(--ink);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Helvetica,Arial,sans-serif;font-size:14px;line-height:1.5}
    .wrap{max-width:1100px;margin:0 auto;padding:28px 20px 60px}
    h1{font-size:24px;margin:0 0 4px}
    .sub{color:var(--muted);margin:0 0 8px}
    .tabs{margin:0 0 20px;font-size:13px}
    .tabs a{display:inline-block;padding:5px 12px;border:1px solid var(--line);border-radius:999px;color:var(--teal);text-decoration:none;margin-right:6px;background:#fff}
    .tabs a.on{background:var(--teal);color:#fff;border-color:var(--teal)}
    .card{background:#fff;border:1px solid var(--line);border-radius:10px;padding:16px 18px;margin:0 0 18px}
    .card h2{font-size:14px;letter-spacing:.4px;text-transform:uppercase;color:var(--teal);margin:0 0 8px}
    table{width:100%;border-collapse:collapse;background:#fff;border:1px solid var(--line);border-radius:10px;overflow:hidden}
    th,td{text-align:left;padding:12px 14px;border-bottom:1px solid var(--line);vertical-align:top}
    th{background:#eef4f4;color:var(--teal);font-size:12px;text-transform:uppercase;letter-spacing:.5px}
    tr:last-child td{border-bottom:none}
    .ord a{color:var(--teal);font-weight:700;text-decoration:none}
    .mkt{font-size:11px;color:var(--muted);margin-top:3px}
    .email{color:var(--muted);font-size:12px}
    ul.items{margin:0;padding-left:16px}
    ul.items li{margin:2px 0}
    .disp{color:var(--teal);font-size:12px;white-space:nowrap}
    .copy{background:var(--teal);color:#fff;border:none;border-radius:6px;padding:7px 12px;font-size:12px;font-weight:700;cursor:pointer}
    .copy.done{background:#3f7d3f}
    details{margin-top:8px}
    summary{cursor:pointer;color:var(--muted);font-size:12px}
    .msg{margin-top:6px;background:#faf9f6;border:1px solid var(--line);border-radius:6px;padding:10px;font-size:13px;color:#333;white-space:pre-wrap}
    .badge{display:inline-block;margin-top:4px;font-size:11px;font-weight:700;padding:2px 7px;border-radius:999px}
    .badge.ok{background:#e5f2ea;color:#2f6f43}
    .badge.warn{background:#fdf0e3;color:#9a5b12}
    .verify{background:#eef4f4;border:1px solid #cfe0df}
    .verify li{margin:3px 0}
    .warnbox{color:#9a5b12;font-weight:600}
    .empty{background:#fff;border:1px solid var(--line);border-radius:10px;padding:28px;text-align:center;color:var(--muted)}
    .meta{color:var(--muted);font-size:12px;margin-top:24px}
    @media print{.copy,summary,.tabs{display:none}details .msg{display:block}body{background:#fff}}
  </style></head>
  <body><div class="wrap">
    <h1>Pre-order follow-ups</h1>
    <p class="sub">${orders.length} order${orders.length === 1 ? "" : "s"} shown (${esc(scopeLabel)}) · last ${days} days · totals: ${auCount} AU, ${usaCount} USA</p>
    <div class="tabs">
      <a href="?days=${days}&market=all" class="${marketFilter === "all" ? "on" : ""}">All (${auCount + usaCount})</a>
      <a href="?days=${days}&market=au" class="${marketFilter === "AU" ? "on" : ""}">🇦🇺 AU (${auCount})</a>
      <a href="?days=${days}&market=usa" class="${marketFilter === "USA" ? "on" : ""}">🇺🇸 USA (${usaCount})</a>
    </div>

    <div class="card">
      <h2>What to tell each customer</h2>
      <p style="margin:0">These customers bought an item that is actually a <strong>pre-order</strong> (out of stock, being made), but their order didn't flag it as a pre-order at checkout — so they don't know. For each order below: contact the customer, let them know the item is a pre-order with the estimated dispatch date, and ask whether they're happy to <strong>wait</strong> or would prefer we <strong>cancel &amp; refund</strong> just that item. Use the <em>Copy message</em> button for a ready-to-send note (edit as you like).</p>
    </div>

    <div class="card verify">
      <h2>How this list was checked</h2>
      <p style="margin:0 0 8px">Every order here independently passes all three tests — that's what "the customer didn't know" means:</p>
      <ul style="margin:0 0 10px">
        <li><strong>No Karma East selling plan</strong> on the line → Shopify's confirmation email had nothing to flag, so they weren't told.</li>
        <li><strong>Still unfulfilled</strong> → the item hasn't shipped; they're genuinely waiting (already-shipped orders are excluded).</li>
        <li><strong>Out of stock</strong> at the fulfilment location → it's awaiting the batch, not shippable now.</li>
      </ul>
      <p style="margin:0 0 8px">Independent cross-check against our own records: <strong>${heldCount} of ${orders.length}</strong> shown ${orders.length === 1 ? "order is" : "orders are"} already <span class="badge ok">✓ holding</span> (the system reserved &amp; held them — separate proof they're real pre-orders). Any <span class="badge warn">⚠ not captured</span> slipped past auto-capture and should be reserved — run <code>/api/preorder-capture-missed?days=${days}&amp;apply=1</code>.</p>
      <p style="margin:0" class="${truncated ? "warnbox" : ""}">Coverage: scanned <strong>${scannedOrders}</strong> paid orders${oldestScanned ? ` back to <strong>${esc(oldestScanned)}</strong>` : ""} in the last ${days} days. ${truncated ? "⚠ Hit the scan cap — orders OLDER than the date above were NOT checked, so some older ones could be missing. Tell me and I'll raise the cap further." : "✓ Full window covered — the scan reached the end of the window, nothing older was skipped."}</p>
    </div>

    ${orders.length === 0
      ? `<div class="empty">${
          queryErrors.length
            ? `⚠️ Couldn't read orders from Shopify:<br><strong>${esc(queryErrors.join("; "))}</strong>`
            : liveBatchCount === 0
              ? `No live pre-order batches right now, so there's nothing to check.`
              : scannedOrders === 0
                ? `Live batches found (${liveBatchCount}), but Shopify returned 0 paid orders in the last ${days} days.`
                : `🎉 No orders need follow-up (${esc(scopeLabel)}) in the last ${days} days.`
        }</div>`
      : `<table><thead><tr><th>Order</th><th>Customer</th><th>Pre-order item(s)</th><th>Message</th></tr></thead><tbody>${rows}</tbody></table>`}

    <p class="meta">Diagnostics — live batches: ${liveBatchCount} · variants tracked: ${variantsTracked} · paid orders scanned: ${scannedOrders}${queryErrors.length ? ` · errors: ${esc(queryErrors.join("; "))}` : ""}. Only lists items still out of stock (awaiting the batch) — if an item has since restocked it ships normally, so it's dropped. Live each load. Window: ?days=N (max 120). Filter: ?market=au|usa|all.</p>
  </div>
  <script>
    function copyMsg(btn){
      var t = btn.getAttribute('data-msg') || '';
      navigator.clipboard.writeText(t).then(function(){
        var old = btn.textContent; btn.textContent = 'Copied ✓'; btn.classList.add('done');
        setTimeout(function(){ btn.textContent = old; btn.classList.remove('done'); }, 1600);
      });
    }
  </script>
  </body></html>`;

  return new Response(html, { headers: { "Content-Type": "text/html; charset=utf-8", "Cache-Control": "no-store" } });
};
