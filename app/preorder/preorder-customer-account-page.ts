import type { CustomerPreorderLine } from "./preorder-customer-account.server";

// Escape for embedding in an app-proxy response served as `application/liquid`.
// As well as HTML-escaping, neutralise `{` / `}` so no customer-supplied text
// (order name, variant title, sku) can be interpreted as Liquid.
function esc(value: string | null | undefined): string {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;")
    .replace(/\{/g, "&#123;")
    .replace(/\}/g, "&#125;");
}

function formatDate(value: string | null): string {
  if (!value) return "Date to be confirmed";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "Date to be confirmed";
  return new Intl.DateTimeFormat("en-AU", { day: "numeric", month: "long", year: "numeric" }).format(date);
}

function statusChip(status: string): string {
  const map: Record<string, { label: string; bg: string; color: string }> = {
    reserved: { label: "Reserved", bg: "#eef2ff", color: "#3730a3" },
    fulfilled: { label: "Dispatched", bg: "#ecfdf5", color: "#065f46" },
    released: { label: "Cancelled", bg: "#f1f5f9", color: "#475569" },
  };
  const chip = map[status] ?? { label: status || "Reserved", bg: "#eef2ff", color: "#3730a3" };
  return `<span class="ke-po-chip" style="background:${chip.bg};color:${chip.color};">${esc(chip.label)}</span>`;
}

// Renders the logged-in customer's pre-orders as an HTML fragment. Served via the
// signed app proxy as `application/liquid`, so Shopify wraps it in the store's
// theme (header/footer/fonts) — it looks native with no theme editing required.
export function renderMyPreordersPage(preorders: CustomerPreorderLine[], loggedIn: boolean): string {
  const styles = `
    <style>
      .ke-po-wrap{max-width:820px;margin:0 auto;padding:32px 20px 56px;font-family:inherit;color:#1f2937;}
      .ke-po-h1{font-size:28px;font-weight:700;letter-spacing:-.01em;margin:0 0 6px;}
      .ke-po-sub{font-size:15px;color:#6b7280;margin:0 0 28px;line-height:1.5;}
      .ke-po-order{border:1px solid #e5e7eb;border-radius:14px;padding:18px 18px 6px;margin-bottom:16px;background:#fff;}
      .ke-po-ohead{display:flex;justify-content:space-between;align-items:baseline;gap:12px;flex-wrap:wrap;margin-bottom:6px;}
      .ke-po-oname{font-size:15px;font-weight:700;color:#111827;}
      .ke-po-odate{font-size:12px;color:#9ca3af;}
      .ke-po-line{display:flex;justify-content:space-between;align-items:center;gap:14px;flex-wrap:wrap;padding:12px 0;border-top:1px solid #f3f4f6;}
      .ke-po-item{font-size:14px;font-weight:600;color:#1f2937;}
      .ke-po-item small{display:block;font-weight:400;color:#9ca3af;margin-top:2px;font-size:12px;}
      .ke-po-right{display:flex;align-items:center;gap:14px;flex-wrap:wrap;}
      .ke-po-ship{font-size:13px;color:#374151;}
      .ke-po-ship strong{display:block;font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:#9ca3af;font-weight:700;}
      .ke-po-chip{font-size:11px;font-weight:700;padding:4px 10px;border-radius:999px;white-space:nowrap;}
      .ke-po-empty{border:1px dashed #d1d5db;border-radius:14px;padding:40px 24px;text-align:center;color:#6b7280;background:#fafafa;}
      .ke-po-empty a{color:#111827;font-weight:600;}
    </style>`;

  if (!loggedIn) {
    return `${styles}<div class="ke-po-wrap"><h1 class="ke-po-h1">My pre-orders</h1>
      <div class="ke-po-empty">Please <a href="/account/login">log in</a> to see your pre-orders.</div></div>`;
  }

  if (!preorders.length) {
    return `${styles}<div class="ke-po-wrap"><h1 class="ke-po-h1">My pre-orders</h1>
      <p class="ke-po-sub">Items you’ve reserved that are still on the way will appear here.</p>
      <div class="ke-po-empty">You have no active pre-orders right now. <a href="/collections/all">Continue shopping →</a></div></div>`;
  }

  // Group by order so multiple reserved sizes from one checkout show together.
  const byOrder = new Map<string, CustomerPreorderLine[]>();
  for (const line of preorders) {
    const key = line.orderName || "Pre-order";
    const list = byOrder.get(key) ?? [];
    list.push(line);
    byOrder.set(key, list);
  }

  const orders = Array.from(byOrder.entries()).map(([orderName, lines]) => {
    const reservedAt = lines[0]?.reservedAt ?? null;
    const rows = lines.map((line) => `
      <div class="ke-po-line">
        <div class="ke-po-item">${esc(line.variantTitle || "Item")} · ${esc(line.market)}<small>${line.sku ? "SKU " + esc(line.sku) + " · " : ""}Qty ${Math.max(1, Math.floor(line.quantity) || 1)}</small></div>
        <div class="ke-po-right">
          <div class="ke-po-ship"><strong>Expected dispatch</strong>${esc(formatDate(line.expectedShipDate))}</div>
          ${statusChip(line.status)}
        </div>
      </div>`).join("");
    return `
      <div class="ke-po-order">
        <div class="ke-po-ohead">
          <span class="ke-po-oname">${esc(orderName)}</span>
          <span class="ke-po-odate">Reserved ${esc(formatDate(reservedAt))}</span>
        </div>
        ${rows}
      </div>`;
  }).join("");

  return `${styles}<div class="ke-po-wrap">
    <h1 class="ke-po-h1">My pre-orders</h1>
    <p class="ke-po-sub">Items you’ve reserved that are being made. We’ll email you when each one dispatches.</p>
    ${orders}
  </div>`;
}
