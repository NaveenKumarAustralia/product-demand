import { useEffect, useMemo, useRef, useState } from "react";
import type {
  PreorderDashboardCustomerOrder,
  PreorderDashboardData,
} from "./preorder-dashboard.server";

function formatDate(value: string | null) {
  if (!value) return "Not set";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "Not set";
  return new Intl.DateTimeFormat("en-AU", { day: "numeric", month: "short", year: "numeric" }).format(date);
}

function formatDateTime(value: string | null) {
  if (!value) return "—";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "—";
  return new Intl.DateTimeFormat("en-AU", { day: "numeric", month: "short", year: "numeric", hour: "numeric", minute: "2-digit", hour12: true }).format(date);
}

type OrdersColId = "order" | "orderDate" | "customer" | "picture" | "product" | "batch" | "sku" | "qty" | "status" | "dispatch";
type OrdersSortCol = Exclude<OrdersColId, "picture">;
const ORDERS_COLUMNS: Array<{ id: OrdersColId; label: string; sortable: boolean; align?: "center" }> = [
  { id: "order", label: "Order", sortable: true },
  { id: "orderDate", label: "Order Date & Time", sortable: true },
  { id: "customer", label: "Customer", sortable: true },
  { id: "picture", label: "Picture", sortable: false },
  { id: "product", label: "Product", sortable: true },
  { id: "batch", label: "Batch", sortable: true },
  { id: "sku", label: "SKU", sortable: true },
  { id: "qty", label: "Qty", sortable: true, align: "center" },
  { id: "status", label: "Status", sortable: true },
  { id: "dispatch", label: "Dispatch", sortable: true },
];
const ORDERS_COL_DEF = new Map(ORDERS_COLUMNS.map((c) => [c.id, c]));
const ORDERS_COL_ALL: OrdersColId[] = ORDERS_COLUMNS.map((c) => c.id);
const ORDERS_COLS_KEY = "preorder-customer-orders-cols-v1";

function loadOrdersColPrefs(): { order: OrdersColId[]; hidden: OrdersColId[] } {
  if (typeof window === "undefined") return { order: ORDERS_COL_ALL, hidden: [] };
  try {
    const raw = window.localStorage.getItem(ORDERS_COLS_KEY);
    if (raw) {
      const p = JSON.parse(raw) as { order?: unknown; hidden?: unknown };
      const savedOrder = Array.isArray(p.order) ? (p.order as OrdersColId[]).filter((id) => ORDERS_COL_DEF.has(id)) : [];
      // Append any columns added since the prefs were saved so nothing is lost.
      const order = [...savedOrder, ...ORDERS_COL_ALL.filter((id) => !savedOrder.includes(id))];
      const hidden = Array.isArray(p.hidden) ? (p.hidden as OrdersColId[]).filter((id) => ORDERS_COL_DEF.has(id)) : [];
      return { order, hidden };
    }
  } catch { /* ignore */ }
  return { order: ORDERS_COL_ALL, hidden: [] };
}
function saveOrdersColPrefs(order: OrdersColId[], hidden: Set<OrdersColId>) {
  if (typeof window === "undefined") return;
  try { window.localStorage.setItem(ORDERS_COLS_KEY, JSON.stringify({ order, hidden: Array.from(hidden) })); } catch { /* ignore */ }
}

function statusPill(status: string) {
  const s2 = status.toLowerCase();
  const tone = s2 === "reserved" ? { bg: "#fef3c7", fg: "#92400e" }
    : s2 === "ready" || s2 === "dispatched" || s2 === "fulfilled" ? { bg: "#dcfce7", fg: "#166534" }
    : s2 === "released" ? { bg: "#e0e7ff", fg: "#3730a3" }
    : { bg: "#f1f5f9", fg: "#475569" };
  const label = s2 === "reserved" ? "Reserved" : s2 === "ready" ? "Ready" : s2 === "dispatched" ? "Dispatched" : s2 === "fulfilled" ? "Fulfilled" : s2 === "released" ? "Released" : status;
  return <span style={{ display: "inline-block", padding: "2px 9px", borderRadius: 999, fontSize: 11.5, fontWeight: 700, background: tone.bg, color: tone.fg }}>{label}</span>;
}

export function PreorderCustomerOrdersPanel({ orders, search = "" }: { orders: PreorderDashboardCustomerOrder[]; search?: string }) {
  const [sortCol, setSortCol] = useState<OrdersSortCol>("dispatch");
  const [sortDir, setSortDir] = useState<"asc" | "desc">("asc");
  // Column show/hide + order, remembered per browser.
  const [colOrder, setColOrder] = useState<OrdersColId[]>(ORDERS_COL_ALL);
  const [hidden, setHidden] = useState<Set<OrdersColId>>(new Set());
  const [colMenuOpen, setColMenuOpen] = useState(false);
  const [dragId, setDragId] = useState<OrdersColId | null>(null);
  const menuRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    const p = loadOrdersColPrefs();
    setColOrder(p.order);
    setHidden(new Set(p.hidden));
  }, []);
  useEffect(() => {
    if (!colMenuOpen) return;
    const onDown = (e: MouseEvent) => { if (menuRef.current && !menuRef.current.contains(e.target as Node)) setColMenuOpen(false); };
    document.addEventListener("mousedown", onDown);
    return () => document.removeEventListener("mousedown", onDown);
  }, [colMenuOpen]);

  const reorder = (drag: OrdersColId, target: OrdersColId) => {
    if (drag === target) return;
    setColOrder((prev) => {
      const next = prev.filter((id) => id !== drag);
      const idx = next.indexOf(target);
      if (idx < 0) return prev;
      next.splice(idx, 0, drag);
      saveOrdersColPrefs(next, hidden);
      return next;
    });
  };
  const toggleHidden = (id: OrdersColId) => {
    setHidden((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else { if (colOrder.filter((c) => !next.has(c)).length <= 1) return prev; next.add(id); }
      saveOrdersColPrefs(colOrder, next);
      return next;
    });
  };
  const visibleCols = colOrder.filter((id) => !hidden.has(id)).map((id) => ORDERS_COL_DEF.get(id)!);

  const rows = useMemo(() => {
    const flat = orders.flatMap((order) => order.lines.map((line) => ({
      key: line.reservationId,
      orderName: order.shopifyOrderName || order.shopifyOrderId,
      orderDate: order.reservedAt,
      customer: order.customerEmail || "—",
      market: order.market,
      imageUrl: line.imageUrl,
      product: line.productTitle || "—",
      size: line.variantTitle || "",
      batch: line.supplierOrderId,
      sku: line.sku || "",
      qty: line.quantity,
      status: line.status,
      dispatch: line.expectedShipDate,
    })));
    // Free-text search across everything on the row — order #, customer email,
    // product, size, batch #, SKU, status, market, order date and dispatch date.
    const q = search.trim().toLowerCase();
    const searched = q
      ? flat.filter((r) => [r.orderName, r.customer, r.product, r.size, `#${r.batch}`, String(r.batch), r.sku, r.status, r.market, formatDateTime(r.orderDate), formatDate(r.dispatch)]
          .some((f) => String(f ?? "").toLowerCase().includes(q)))
      : flat;
    const val = (r: typeof flat[number]): string | number => {
      switch (sortCol) {
        case "order": return r.orderName.toLowerCase();
        case "orderDate": return r.orderDate ? new Date(r.orderDate).getTime() : Number.POSITIVE_INFINITY;
        case "customer": return r.customer.toLowerCase();
        case "product": return r.product.toLowerCase();
        case "batch": return r.batch;
        case "sku": return r.sku.toLowerCase();
        case "qty": return r.qty;
        case "status": return r.status.toLowerCase();
        case "dispatch": return r.dispatch ? new Date(r.dispatch).getTime() : Number.POSITIVE_INFINITY;
      }
    };
    return [...searched].sort((a, b) => {
      const va = val(a), vb = val(b);
      const cmp = typeof va === "string" ? va.localeCompare(vb as string) : (va as number) - (vb as number);
      return sortDir === "asc" ? cmp : -cmp;
    });
  }, [orders, sortCol, sortDir, search]);

  if (!orders.length) {
    return <div style={s.empty}>No preorder reservations have been created yet. Customer orders will appear here once Shopify order allocation is connected.</div>;
  }

  const clickSort = (col: OrdersSortCol) => {
    if (sortCol === col) setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    else { setSortCol(col); setSortDir("asc"); }
  };
  const arrow = (col: OrdersSortCol) => (sortCol === col ? (sortDir === "asc" ? " ▲" : " ▼") : "");
  const th: React.CSSProperties = { position: "sticky", top: 0, zIndex: 2, background: "#f8fafc", borderBottom: "2px solid #e2e8f0", padding: "9px 10px", fontSize: 11, fontWeight: 800, color: "#64748b", textTransform: "uppercase", letterSpacing: ".03em", whiteSpace: "nowrap", textAlign: "left", userSelect: "none" };
  const td: React.CSSProperties = { padding: "8px 10px", fontSize: 13, borderBottom: "1px solid #eef2f7", verticalAlign: "middle", whiteSpace: "nowrap" };

  const renderCell = (id: OrdersColId, r: typeof rows[number]) => {
    switch (id) {
      case "order": return <td key={id} style={{ ...td, fontWeight: 700 }}>{r.orderName}</td>;
      case "orderDate": return <td key={id} style={{ ...td, color: "#475569" }}>{formatDateTime(r.orderDate)}</td>;
      case "customer": return <td key={id} style={{ ...td, color: "#64748b" }}>{r.customer}</td>;
      case "picture": return <td key={id} style={td}>{r.imageUrl ? <img src={r.imageUrl} alt="" style={{ width: 51, height: 63, objectFit: "cover", borderRadius: 5 }} /> : <div style={{ width: 51, height: 63, background: "#f1f5f9", borderRadius: 5 }} />}</td>;
      case "product": return <td key={id} style={td}>{r.product}{r.size ? <span style={{ color: "#94a3b8" }}> · {r.size}</span> : null}</td>;
      case "batch": return <td key={id} style={td}>#{r.batch}</td>;
      case "sku": return <td key={id} style={{ ...td, color: "#64748b" }}>{r.sku || "—"}</td>;
      case "qty": return <td key={id} style={{ ...td, textAlign: "center", fontWeight: 700 }}>{r.qty}</td>;
      case "status": return <td key={id} style={td}>{statusPill(r.status)}</td>;
      case "dispatch": return <td key={id} style={{ ...td, fontWeight: 700 }}>{formatDate(r.dispatch)}</td>;
    }
  };

  return (
    <div style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 12, overflow: "hidden" }}>
      {/* Toolbar — Columns selector */}
      <div style={{ display: "flex", justifyContent: "flex-end", alignItems: "center", padding: "8px 10px", borderBottom: "1px solid #eef2f7", position: "relative" }} ref={menuRef}>
        <button
          type="button"
          onClick={() => setColMenuOpen((o) => !o)}
          style={{ display: "inline-flex", alignItems: "center", gap: 6, background: colMenuOpen ? "#eef2ff" : "#fff", border: "1px solid #cbd5e1", borderRadius: 8, padding: "6px 11px", fontSize: 12.5, fontWeight: 700, color: "#334155", cursor: "pointer" }}
        >⚙ Columns</button>
        {colMenuOpen && (
          <div style={{ position: "absolute", top: "calc(100% + 4px)", right: 10, zIndex: 30, width: 250, background: "#fff", border: "1px solid #e2e8f0", borderRadius: 10, boxShadow: "0 12px 30px rgba(15,23,42,0.16)", padding: 8 }}>
            <div style={{ fontSize: 11, fontWeight: 800, color: "#64748b", textTransform: "uppercase", letterSpacing: ".03em", padding: "4px 6px 8px" }}>Show &amp; reorder columns</div>
            <div style={{ display: "flex", flexDirection: "column", gap: 1 }}>
              {colOrder.map((id) => {
                const col = ORDERS_COL_DEF.get(id)!;
                const isHidden = hidden.has(id);
                return (
                  <div
                    key={id}
                    draggable
                    onDragStart={() => setDragId(id)}
                    onDragOver={(e) => e.preventDefault()}
                    onDrop={() => { if (dragId) reorder(dragId, id); setDragId(null); }}
                    onDragEnd={() => setDragId(null)}
                    style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 6px", borderRadius: 7, cursor: "grab", background: dragId === id ? "#eef2ff" : "transparent" }}
                  >
                    <span style={{ color: "#cbd5e1", fontSize: 13, lineHeight: 1, cursor: "grab" }} title="Drag to reorder">⠿</span>
                    <label style={{ display: "flex", alignItems: "center", gap: 8, flex: 1, cursor: "pointer", fontSize: 13, color: "#1f2937" }}>
                      <input type="checkbox" checked={!isHidden} onChange={() => toggleHidden(id)} style={{ width: 15, height: 15, cursor: "pointer" }} />
                      {col.label}
                    </label>
                  </div>
                );
              })}
            </div>
            <div style={{ borderTop: "1px solid #f1f5f9", marginTop: 6, paddingTop: 6, textAlign: "right" }}>
              <button type="button" onClick={() => { setColOrder(ORDERS_COL_ALL); setHidden(new Set()); saveOrdersColPrefs(ORDERS_COL_ALL, new Set()); }} style={{ background: "none", border: "none", color: "#6366f1", fontSize: 12, fontWeight: 700, cursor: "pointer", padding: "2px 4px" }}>Reset to default</button>
            </div>
          </div>
        )}
      </div>
      <div style={{ overflow: "auto", maxHeight: "calc(100vh - 288px)", borderRadius: 12 }}>
        <table style={{ borderCollapse: "collapse", width: "100%", minWidth: 940 }}>
          <thead>
            <tr>
              {visibleCols.map((col) => (
                <th
                  key={col.id}
                  draggable
                  onDragStart={() => setDragId(col.id)}
                  onDragOver={(e) => e.preventDefault()}
                  onDrop={() => { if (dragId) reorder(dragId, col.id); setDragId(null); }}
                  onDragEnd={() => setDragId(null)}
                  onClick={() => { if (col.sortable) clickSort(col.id as OrdersSortCol); }}
                  title="Click to sort · drag to reorder"
                  style={{ ...th, textAlign: col.align === "center" ? "center" : "left", cursor: col.sortable ? "pointer" : "grab", opacity: dragId === col.id ? 0.4 : 1 }}
                >{col.label}{col.sortable ? arrow(col.id as OrdersSortCol) : ""}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr><td colSpan={visibleCols.length} style={{ ...td, textAlign: "center", color: "#94a3b8", padding: "28px 10px" }}>No customer orders match “{search.trim()}”. Search by order number, customer email, product, batch or SKU.</td></tr>
            ) : rows.map((r) => (
              <tr key={r.key}>{visibleCols.map((col) => renderCell(col.id, r))}</tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// One-glance activation readiness: aggregates the existing scope / webhook /
// Klaviyo checks + the AU/USA location config into a single checklist so staff
// can see exactly what's left before turning on customer preorders. These are
// internal checks only — activation still requires staff to enable + activate a
// batch. Blocker items must pass; optional items (Klaviyo, storefront) are
// warnings/manual verifications.
type ReadinessStatus = "loading" | "ok" | "fail" | "error" | "warn" | "manual";
const READINESS_ICON: Record<ReadinessStatus, string> = { loading: "…", ok: "✅", fail: "❌", error: "❌", warn: "⚠️", manual: "🔍" };

export function PreorderActivationReadinessPanel({ configuration }: { configuration: PreorderDashboardData["configuration"] }) {
  const [scopes, setScopes] = useState<{ status: ReadinessStatus; detail: string }>({ status: "loading", detail: "Checking Shopify permissions…" });
  const [webhooks, setWebhooks] = useState<{ status: ReadinessStatus; detail: string }>({ status: "loading", detail: "Checking webhooks…" });
  const [klaviyo, setKlaviyo] = useState<{ status: ReadinessStatus; detail: string }>({ status: "loading", detail: "Checking Klaviyo…" });

  useEffect(() => {
    let cancelled = false;
    fetch("/api/preorder-shopify-readiness", { credentials: "same-origin" }).then((r) => r.json()).then((d: { ok?: boolean; error?: string; shops?: Array<{ shop?: string; ok?: boolean; ready?: boolean; error?: string; missingScopes?: string[] }> }) => {
      if (cancelled) return;
      if (!d?.ok) { setScopes({ status: "error", detail: d?.error || "Could not check scopes." }); return; }
      const shops = d.shops ?? [];
      if (!shops.length) { setScopes({ status: "fail", detail: "No linked Shopify shop found to check scopes against." }); return; }
      const errored = shops.find((sh) => sh.ok === false);
      if (errored) { setScopes({ status: "error", detail: `Scope query failed for ${errored.shop || "shop"}: ${errored.error || "Shopify rejected the request"} — usually a stale offline token; re-auth at /force-auth?force=1.` }); return; }
      const ready = shops.every((sh) => sh.ready);
      const missing = Array.from(new Set(shops.flatMap((sh) => sh.missingScopes ?? [])));
      setScopes(ready ? { status: "ok", detail: "All required Shopify scopes granted on the install." } : { status: "fail", detail: missing.length ? `Declared but NOT granted on the install yet: ${missing.join(", ")}. Re-consent at /force-auth?force=1.` : "Some required scopes aren't granted on the install — re-consent at /force-auth?force=1." });
    }).catch(() => { if (!cancelled) setScopes({ status: "error", detail: "Network error checking scopes." }); });

    fetch("/api/preorder-webhook-status", { credentials: "same-origin" }).then((r) => r.json()).then((d: { ok?: boolean; healthy?: boolean; error?: string; shops?: Array<{ shop?: string; ok?: boolean; error?: string; subscriptions?: Array<{ label?: string; registered?: boolean }> }> }) => {
      if (cancelled) return;
      if (!d?.ok) { setWebhooks({ status: "error", detail: d?.error || "Could not check webhooks." }); return; }
      const shops = d.shops ?? [];
      const errored = shops.find((sh) => sh.ok === false);
      if (errored) { setWebhooks({ status: "error", detail: `Webhook query failed for ${errored.shop || "shop"}: ${errored.error || "Shopify rejected the request"} — usually a stale offline token; re-auth at /force-auth?force=1.` }); return; }
      if (d.healthy) { setWebhooks({ status: "ok", detail: "Order create / cancel / fulfilled webhooks registered." }); return; }
      const missing = shops.flatMap((sh) => (sh.subscriptions ?? []).filter((x) => !x.registered).map((x) => x.label ?? "")).filter(Boolean);
      setWebhooks({ status: "fail", detail: missing.length ? `Not registered: ${missing.join(", ")}. A redeploy re-runs registration; re-auth then reload.` : "Some order webhooks are not registered." });
    }).catch(() => { if (!cancelled) setWebhooks({ status: "error", detail: "Network error checking webhooks." }); });

    fetch("/api/preorder-klaviyo-status", { credentials: "same-origin" }).then((r) => r.json()).then((d: { ok?: boolean; configured?: boolean }) => {
      if (cancelled) return;
      setKlaviyo(d?.configured ? { status: "ok", detail: "Klaviyo API key present (Back-in-Stock notifications)." } : { status: "warn", detail: "Not connected — only needed for Back-in-Stock / notification emails." });
    }).catch(() => { if (!cancelled) setKlaviyo({ status: "warn", detail: "Could not check Klaviyo connection." }); });

    return () => { cancelled = true; };
  }, []);

  const auLoc = (configuration.locations.AU ?? "").trim();
  const usaLoc = (configuration.locations.USA ?? "").trim();
  // A market goes live only once ITS location is set. Launching one market
  // (e.g. AU) is fine — the other stays off until its location is added, then
  // plugs in automatically. Blocker only if NEITHER is set.
  const locations: { status: ReadinessStatus; detail: string } =
    !auLoc && !usaLoc ? { status: "fail", detail: "Set at least one region's Shopify location in Settings (AU to launch now; add USA later)." }
      : auLoc && usaLoc && auLoc === usaLoc ? { status: "fail", detail: "AU and USA are set to the same location — regional pools must differ." }
        : auLoc && usaLoc ? { status: "ok", detail: "AU + USA locations set and different." }
          : { status: "ok", detail: `${auLoc ? "AU" : "USA"} location set — ${auLoc ? "USA" : "AU"} preorders stay OFF until its location is added.` };

  const items: Array<{ key: string; label: string; blocker: boolean; status: ReadinessStatus; detail: string }> = [
    { key: "scopes", label: "Shopify permissions (scopes)", blocker: true, ...scopes },
    { key: "webhooks", label: "Order webhooks registered", blocker: true, ...webhooks },
    { key: "locations", label: "Fulfilment location(s) set", blocker: true, ...locations },
    { key: "klaviyo", label: "Klaviyo connected", blocker: false, ...klaviyo },
    { key: "storefront", label: "Storefront app-proxy + theme block", blocker: false, status: "manual", detail: "Verify on the live store: an out-of-stock variant shows Pre-order in the correct market only (not the other market's plan)." },
  ];

  const anyLoading = items.some((i) => i.status === "loading");
  const blockersFailing = items.filter((i) => i.blocker && (i.status === "fail" || i.status === "error")).length;
  const badgeColor = anyLoading ? { bg: "#f1f5f9", fg: "#475569" } : blockersFailing === 0 ? { bg: "#dcfce7", fg: "#166534" } : { bg: "#fef3c7", fg: "#92400e" };

  return (
    <div style={s.card}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}>
        <div style={s.title}>Activation readiness</div>
        <div style={{ ...s.badge, background: badgeColor.bg, color: badgeColor.fg }}>
          {anyLoading ? "Checking…" : blockersFailing === 0 ? "Core checks pass" : `${blockersFailing} blocker${blockersFailing === 1 ? "" : "s"} to resolve`}
        </div>
      </div>
      <div style={s.muted}>Internal checks only. Turning on customer preorders still requires enabling a batch and activating it on Shopify.</div>
      <div style={{ display: "flex", flexDirection: "column", gap: 8, marginTop: 10 }}>
        {items.map((it) => (
          <div key={it.key} style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
            <span style={{ fontSize: 14, width: 18, textAlign: "center", flexShrink: 0 }}>{READINESS_ICON[it.status]}</span>
            <div>
              <div style={{ fontWeight: 700, fontSize: 13 }}>{it.label}{!it.blocker ? <span style={{ fontWeight: 500, color: "#94a3b8" }}> · optional</span> : null}</div>
              <div style={{ ...s.muted, fontSize: 11.5 }}>{it.detail}</div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

type ShopifyLocationOption = { id: string; name: string; city?: string | null; country?: string | null };

// Dropdown of live Shopify locations (name · city, country) that stores the
// exact location GID. Falls back to a free-text input when the location list
// can't be loaded or the saved value isn't in the list, so an existing/raw
// value is never lost.
function LocationSelect({ label, value, onChange, locations, loaded, loadError }: {
  label: string;
  value: string;
  onChange: (next: string) => void;
  locations: ShopifyLocationOption[];
  loaded: boolean;
  loadError: string | null;
}) {
  const MANUAL = "__manual__";
  const inList = value !== "" && locations.some((l) => l.id === value);
  const [manual, setManual] = useState(false);
  const useText = !loaded || !!loadError || locations.length === 0 || manual;
  const optionLabel = (l: ShopifyLocationOption) => `${l.name}${l.city || l.country ? ` · ${[l.city, l.country].filter(Boolean).join(", ")}` : ""}`;
  return (
    <label style={s.label}>
      {label}
      {useText ? (
        <input value={value} onChange={(e) => onChange(e.target.value)} style={s.input} placeholder="gid://shopify/Location/..." />
      ) : (
        <select
          value={inList ? value : (value ? "__current__" : "")}
          onChange={(e) => { const v = e.target.value; if (v === MANUAL) { setManual(true); } else if (v !== "__current__") { onChange(v); } }}
          style={s.input}
        >
          <option value="">— select a location —</option>
          {value !== "" && !inList ? <option value="__current__">Current: {value}</option> : null}
          {locations.map((l) => <option key={l.id} value={l.id}>{optionLabel(l)}</option>)}
          <option value={MANUAL}>✎ Enter manually…</option>
        </select>
      )}
      {loadError ? <span style={{ ...s.muted, fontSize: 11 }}>Couldn’t load Shopify locations ({loadError}) — enter the GID manually.</span> : null}
      {useText && !loadError && loaded && locations.length > 0 ? (
        <button type="button" style={{ ...s.muted, fontSize: 11, background: "none", border: "none", padding: 0, cursor: "pointer", textAlign: "left", color: "#2563eb" }} onClick={() => setManual(false)}>↩ pick from list</button>
      ) : null}
    </label>
  );
}

export function PreorderSettingsPanel({ configuration }: { configuration: PreorderDashboardData["configuration"] }) {
  const [au, setAu] = useState(configuration.locations.AU ?? "");
  const [usa, setUsa] = useState(configuration.locations.USA ?? "");
  const [combineDays, setCombineDays] = useState(String(configuration.combineWindowDays ?? 8));
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState<{ kind: "error" | "success"; text: string } | null>(null);
  const [locations, setLocations] = useState<ShopifyLocationOption[]>([]);
  const [locLoaded, setLocLoaded] = useState(false);
  const [locError, setLocError] = useState<string | null>(null);
  useEffect(() => {
    let cancelled = false;
    fetch("/api/preorder-shopify-locations", { credentials: "same-origin" })
      .then((r) => r.json())
      .then((data: { ok?: boolean; error?: string; shops?: Array<{ ok?: boolean; error?: string; locations?: Array<{ id: string; name: string; address?: { city?: string | null; country?: string | null } | null }> }> }) => {
        if (cancelled) return;
        if (!data?.ok) { setLocError(data?.error || "unavailable"); setLocLoaded(true); return; }
        const flat: ShopifyLocationOption[] = [];
        const firstErr = (data.shops ?? []).find((sh) => sh.ok === false)?.error ?? null;
        for (const sh of data.shops ?? []) for (const l of sh.locations ?? []) flat.push({ id: l.id, name: l.name, city: l.address?.city ?? null, country: l.address?.country ?? null });
        setLocations(flat);
        if (!flat.length && firstErr) setLocError(firstErr);
        setLocLoaded(true);
      })
      .catch(() => { if (!cancelled) { setLocError("network error"); setLocLoaded(true); } });
    return () => { cancelled = true; };
  }, []);

  async function post(payload: Record<string, unknown>, successText: string) {
    setBusy(true);
    setNotice(null);
    try {
      const response = await fetch("/api/preorder-manage", {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const result = await response.json().catch(() => ({})) as { ok?: boolean; error?: string };
      if (!response.ok || result.ok !== true) throw new Error(result.error || "Could not save preorder settings.");
      setNotice({ kind: "success", text: successText });
    } catch (error) {
      setNotice({ kind: "error", text: error instanceof Error ? error.message : "Could not save preorder settings." });
    } finally {
      setBusy(false);
    }
  }


  return (
    <div style={s.stack}>
      {notice ? <div style={{ ...s.notice, ...(notice.kind === "error" ? s.noticeError : s.noticeSuccess) }}>{notice.text}</div> : null}

      <div style={s.card}>
        <div style={s.title}>Shopify locations</div>
        <div style={s.muted}>Keep AU and USA preorder capacity completely separate. Pick each region's Shopify location from the list (the exact location GID is stored){locLoaded && !locError && locations.length ? "" : "; or enter the numeric ID / full gid manually"}.</div>
        <div style={s.twoCols}>
          <LocationSelect label="Australia location" value={au} onChange={setAu} locations={locations} loaded={locLoaded} loadError={locError} />
          <LocationSelect label="USA location" value={usa} onChange={setUsa} locations={locations} loaded={locLoaded} loadError={locError} />
        </div>
        {au && usa && au === usa ? <div style={{ ...s.notice, ...s.noticeError }}>AU and USA are set to the SAME location — regional pools must be different.</div> : null}
        <div style={s.actions}><button type="button" disabled={busy} style={s.primary} onClick={() => post({ operation: "update-locations", AU: au, USA: usa }, "Shopify preorder locations saved.")}>Save locations</button></div>
      </div>
      <div style={s.card}>
        <div style={s.title}>Combine mixed orders (ship together)</div>
        <div style={s.muted}>
          When an order has both in-stock and pre-order items and the pre-order is due to dispatch within this many days,
          hold the whole order so it ships in <strong>one parcel</strong> when the batch lands (instead of shipping the in-stock
          part first). The order releases automatically the moment its stock is loaded into Shopify. Set to <strong>0</strong> to
          always ship in-stock items immediately.
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 10 }}>
          <input
            type="number" min={0} max={365} value={combineDays}
            onChange={(e) => setCombineDays(e.target.value.replace(/[^0-9]/g, ""))}
            style={{ width: 90, height: 36, boxSizing: "border-box", border: "1px solid #cbd5e1", borderRadius: 8, padding: "0 10px", fontSize: 14 }}
          />
          <span style={s.muted}>days</span>
        </div>
        <div style={{ ...s.actions, gap: 8, flexWrap: "wrap" }}>
          <button type="button" disabled={busy} style={s.primary} onClick={() => post({ operation: "set-combine-window", days: Number(combineDays) || 0 }, `Combine window saved (${Number(combineDays) || 0} days).`)}>Save combine window</button>
          <button type="button" disabled={busy} style={{ border: "1px solid #0f766e", background: "#fff", color: "#0f766e", borderRadius: 8, padding: "8px 12px", fontSize: 12, fontWeight: 800, cursor: busy ? "default" : "pointer" }} onClick={async () => {
            setBusy(true); setNotice(null);
            try {
              const r = await fetch("/api/preorder-manage", { method: "POST", credentials: "same-origin", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ operation: "combine-existing" }) });
              const d = await r.json().catch(() => ({})) as { ok?: boolean; error?: string; combine?: { heldOrders: number; scannedOrders: number; skippedNoScope?: boolean } };
              if (!r.ok || d.ok !== true) throw new Error(d.error || "Could not combine existing orders.");
              const c = d.combine;
              setNotice({ kind: "success", text: c ? `Held ${c.heldOrders} existing order${c.heldOrders === 1 ? "" : "s"} (of ${c.scannedOrders} due within the window)${c.skippedNoScope ? " — some skipped, re-auth the app for order/fulfilment scopes" : ""}. They'll release when their stock is loaded.` : "Done." });
            } catch (error) {
              setNotice({ kind: "error", text: error instanceof Error ? error.message : "Could not combine existing orders." });
            } finally { setBusy(false); }
          }}>Apply to existing orders now</button>
        </div>
        <div style={{ ...s.muted, marginTop: 6 }}>“Apply to existing orders” holds the in-stock items of orders already placed whose pre-order is due within the window above (raise the days to catch more). Already-shipped items aren’t affected.</div>
      </div>
      <div style={s.card}>
        <div style={s.title}>Staff permissions</div>
        <div style={s.muted}>Manage Pre-orders page access and action permissions in Production Portal → Settings → Users. This page no longer keeps a second permission list.</div>
      </div>
    </div>
  );
}

type PreorderReportResponse = {
  ok: boolean;
  error?: string;
  summary?: {
    customerOrders: number;
    reservationRows: number;
    quantities: Array<{ status: string; market: string; quantity: number; rows: number }>;
  };
  activeBatches?: Array<{ id: number; productTitle: string; supplier: string; destination: string | null; reservedQty: number; incomingQty?: number; fillPercent?: number }>;
  recentFailures?: Array<{ id: number; shopifyOrderId: string | null; orderName: string | null; message: string | null; createdAt: string }>;
};

export function PreorderReportsPanel() {
  const [report, setReport] = useState<PreorderReportResponse | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let active = true;
    fetch("/api/preorder-report", { credentials: "same-origin" })
      .then(async (response) => {
        const result = await response.json().catch(() => ({ ok: false, error: "Could not read report response." }));
        if (!response.ok || result.ok !== true) throw new Error(result.error || "Could not load preorder report.");
        return result as PreorderReportResponse;
      })
      .then((result) => { if (active) setReport(result); })
      .catch((error) => { if (active) setReport({ ok: false, error: error instanceof Error ? error.message : "Could not load preorder report." }); })
      .finally(() => { if (active) setLoading(false); });
    return () => { active = false; };
  }, []);

  const totals = useMemo(() => {
    const quantities = report?.summary?.quantities ?? [];
    const sum = (status: string, market?: string) => quantities
      .filter((row) => row.status === status && (!market || row.market === market))
      .reduce((total, row) => total + row.quantity, 0);
    return {
      reserved: sum("reserved"),
      fulfilled: sum("fulfilled"),
      released: sum("released"),
      au: sum("reserved", "AU"),
      usa: sum("reserved", "USA"),
    };
  }, [report]);

  if (loading) return <div style={s.empty}>Loading preorder report…</div>;
  if (!report?.ok) return <div style={{ ...s.notice, ...s.noticeError }}>{report?.error || "Could not load preorder report."}</div>;

  return (
    <div style={s.stack}>
      <div style={s.reportCards}>
        <ReportMetric label="Customer orders" value={report.summary?.customerOrders ?? 0} />
        <ReportMetric label="Reserved units" value={totals.reserved} />
        <ReportMetric label="Fulfilled units" value={totals.fulfilled} />
        <ReportMetric label="Released units" value={totals.released} />
        <ReportMetric label="AU active" value={totals.au} />
        <ReportMetric label="USA active" value={totals.usa} />
      </div>

      <div style={s.card}>
        <div style={s.title}>Most reserved active batches</div>
        <div style={s.muted}>Where current preorder commitments are concentrated.</div>
        {(report.activeBatches ?? []).length ? (
          <div style={s.lines}>
            {(report.activeBatches ?? []).slice(0, 20).map((batch) => {
              const incoming = batch.incomingQty ?? 0;
              const fill = batch.fillPercent ?? 0;
              const barColor = fill >= 90 ? "#dc2626" : fill >= 70 ? "#d97706" : "#0f766e";
              return (
                <div key={batch.id} style={s.reportBatchRow}>
                  <div>
                    <strong>{batch.productTitle}</strong>
                    <div style={s.small}>Batch #{batch.id} · {batch.supplier}</div>
                    <div style={s.fillTrack}><div style={{ ...s.fillBar, width: `${fill}%`, background: barColor }} /></div>
                  </div>
                  <div style={s.lineMeta}>{batch.destination === "send_to_usa" ? "USA" : "AU"}</div>
                  <div style={{ textAlign: "right" }}>
                    <div style={s.reportNumber}>{batch.reservedQty}<span style={s.reportDenom}> / {incoming}</span></div>
                    <div style={s.small}>{fill}% reserved</div>
                  </div>
                </div>
              );
            })}
          </div>
        ) : <div style={s.muted}>No active reservations yet.</div>}
      </div>

      <div style={s.card}>
        <div style={s.title}>Allocation exceptions</div>
        <div style={s.muted}>Orders where confirmed preorder capacity could not be allocated. These should be reviewed rather than silently cancelled.</div>
        {(report.recentFailures ?? []).length ? (
          <div style={s.lines}>
            {(report.recentFailures ?? []).map((failure) => (
              <div key={failure.id} style={s.failureRow}>
                <div><strong>{failure.orderName || failure.shopifyOrderId || "Shopify order"}</strong><div style={s.small}>{formatDate(failure.createdAt)}</div></div>
                <div style={s.failureMessage}>{failure.message || "Allocation failed"}</div>
              </div>
            ))}
          </div>
        ) : <div style={s.muted}>No allocation exceptions recorded.</div>}
      </div>
    </div>
  );
}

function ReportMetric({ label, value }: { label: string; value: number }) {
  return <div style={s.reportMetric}><div style={s.reportLabel}>{label}</div><div style={s.reportValue}>{value}</div></div>;
}

const s: Record<string, React.CSSProperties> = {
  stack: { display: "flex", flexDirection: "column", gap: 12 },
  card: { background: "white", border: "1px solid #e2e8f0", borderRadius: 14, padding: 18, boxShadow: "0 1px 3px rgba(15,23,42,.04)" },
  empty: { padding: 40, textAlign: "center", color: "#64748b", background: "white", border: "1px dashed #cbd5e1", borderRadius: 12 },
  rowBetween: { display: "flex", justifyContent: "space-between", gap: 16, alignItems: "flex-start" },
  title: { fontSize: 16, fontWeight: 800, color: "#0f172a" },
  muted: { fontSize: 12, color: "#64748b", marginTop: 4, lineHeight: 1.5 },
  small: { fontSize: 11, color: "#94a3b8", marginTop: 3 },
  badge: { padding: "5px 8px", borderRadius: 999, background: "#f1f5f9", color: "#475569", fontSize: 11, fontWeight: 700, whiteSpace: "nowrap" },
  lines: { marginTop: 14, borderTop: "1px solid #f1f5f9" },
  line: { display: "grid", gridTemplateColumns: "minmax(150px, 1.5fr) repeat(3, minmax(90px, .7fr))", gap: 10, alignItems: "center", padding: "10px 0", borderBottom: "1px solid #f8fafc", fontSize: 12 },
  lineMeta: { color: "#64748b" },
  twoCols: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 12, marginTop: 14 },
  label: { display: "flex", flexDirection: "column", gap: 5, fontSize: 12, fontWeight: 700, color: "#475569" },
  input: { border: "1px solid #cbd5e1", borderRadius: 8, padding: "9px 10px", fontSize: 13 },
  actions: { display: "flex", justifyContent: "flex-end", marginTop: 14 },
  primary: { border: "1px solid #0f766e", background: "#0f766e", color: "white", borderRadius: 8, padding: "8px 12px", fontSize: 12, fontWeight: 800, cursor: "pointer" },
  permissionTable: { marginTop: 14, overflowX: "auto", border: "1px solid #e2e8f0", borderRadius: 10 },
  permissionHeader: { display: "grid", gridTemplateColumns: "minmax(150px, 1.3fr) repeat(5, minmax(110px, 1fr))", gap: 8, padding: "9px 10px", background: "#f8fafc", fontSize: 10, fontWeight: 800, color: "#64748b", textTransform: "uppercase" },
  permissionRow: { display: "grid", gridTemplateColumns: "minmax(150px, 1.3fr) repeat(5, minmax(110px, 1fr))", gap: 8, alignItems: "center", padding: "9px 10px", borderTop: "1px solid #f1f5f9", fontSize: 12 },
  checkboxCell: { textAlign: "center" },
  notice: { padding: "10px 12px", borderRadius: 9, fontSize: 13, fontWeight: 700 },
  noticeError: { background: "#fef2f2", border: "1px solid #fecaca", color: "#991b1b" },
  noticeSuccess: { background: "#f0fdf4", border: "1px solid #bbf7d0", color: "#166534" },
  reportCards: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 10 },
  reportMetric: { background: "white", border: "1px solid #e2e8f0", borderRadius: 12, padding: 14 },
  reportLabel: { fontSize: 10, color: "#64748b", fontWeight: 800, textTransform: "uppercase", letterSpacing: ".04em" },
  reportValue: { marginTop: 4, fontSize: 24, color: "#0f172a", fontWeight: 800 },
  reportBatchRow: { display: "grid", gridTemplateColumns: "minmax(180px, 1fr) 70px 120px", gap: 10, alignItems: "center", padding: "10px 0", borderBottom: "1px solid #f8fafc", fontSize: 12 },
  reportNumber: { textAlign: "right", fontSize: 16, fontWeight: 800, color: "#0f172a" },
  reportDenom: { fontSize: 12, fontWeight: 600, color: "#94a3b8" },
  fillTrack: { marginTop: 6, height: 5, borderRadius: 999, background: "#f1f5f9", overflow: "hidden", maxWidth: 220 },
  fillBar: { height: "100%", borderRadius: 999 },
  failureRow: { display: "grid", gridTemplateColumns: "minmax(160px, .8fr) minmax(220px, 1.5fr)", gap: 16, padding: "10px 0", borderBottom: "1px solid #f8fafc", fontSize: 12 },
  failureMessage: { color: "#991b1b" },
};
