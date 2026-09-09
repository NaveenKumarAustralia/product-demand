import { useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import type { PreorderDashboardBatch, PreorderDashboardData } from "./preorder/preorder-dashboard.server";
import {
  PreorderActivationReadinessPanel,
  PreorderCustomerOrdersPanel,
  PreorderReportsPanel,
  PreorderSettingsPanel,
} from "./preorder/preorder-operations-panels";
import { PreorderWebhookStatusPanel } from "./preorder/preorder-webhook-status-panel";
import { PreorderShopifyReadinessPanel } from "./preorder/preorder-shopify-readiness-panel";
import { PreorderWaitlistPanel } from "./preorder/preorder-waitlist-panel";
import { PreorderNotificationsPanel } from "./preorder/preorder-notifications-panel";

type Props = {
  data: PreorderDashboardData;
  search?: string;
};

type TabId = "batches" | "orders" | "waitlist" | "notifications" | "reports" | "settings";

const TABS: Array<{ id: TabId; label: string }> = [
  { id: "batches", label: "Products & Batches" },
  { id: "orders", label: "Customer Orders" },
  { id: "waitlist", label: "Back in Stock" },
  { id: "notifications", label: "Notifications" },
  { id: "reports", label: "Reports" },
  { id: "settings", label: "Settings" },
];

function statusLabel(status: string) {
  return status
    .split("_")
    .filter(Boolean)
    .map((part) => part[0]?.toUpperCase() + part.slice(1))
    .join(" ");
}

function formatDate(value: string | null) {
  if (!value) return "Not set";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "Not set";
  return new Intl.DateTimeFormat("en-AU", { day: "numeric", month: "short", year: "numeric" }).format(date);
}

function dateInputValue(value: string | null) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return date.toISOString().slice(0, 10);
}

export function PreordersDashboard({ data, search: headerSearch = "" }: Props) {
  const [tab, setTab] = useState<TabId>("batches");
  const [market, setMarket] = useState<"ALL" | "AU" | "USA">("ALL");
  // Click a summary tile to filter the list below (e.g. "Active batches" → only
  // the active ones). Click the same tile again to clear.
  const [tileFilter, setTileFilter] = useState<"all" | "active" | "incoming" | "reserved" | "available" | "overallocated">("all");
  const [notice, setNotice] = useState<{ kind: "error" | "success"; text: string } | null>(null);
  const [busyBatchId, setBusyBatchId] = useState<number | null>(null);

  const batches = useMemo(() => {
    const q = headerSearch.trim().toLowerCase();
    return data.batches.filter((batch) => {
      if (market !== "ALL" && batch.market !== market) return false;
      if (tileFilter === "active" && !(batch.enabled && batch.eligible)) return false;
      if (tileFilter === "incoming" && !(batch.totalIncoming > 0)) return false;
      if (tileFilter === "reserved" && !(batch.totalReserved > 0)) return false;
      if (tileFilter === "available" && !(batch.totalAvailable > 0)) return false;
      if (tileFilter === "overallocated" && !(batch.totalReserved > batch.totalIncoming)) return false;
      if (q && !`${batch.productTitle} ${batch.supplier} ${batch.market ?? ""}`.toLowerCase().includes(q)) return false;
      return true;
    });
  }, [data.batches, market, headerSearch, tileFilter]);

  async function manageBatch(batchId: number, payload: Record<string, unknown>, successText: string) {
    setBusyBatchId(batchId);
    setNotice(null);
    try {
      const response = await fetch("/api/preorder-manage", {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ supplierOrderId: batchId, ...payload }),
      });
      const result = await response.json().catch(() => ({})) as { ok?: boolean; error?: string };
      if (!response.ok || result.ok !== true) {
        throw new Error(result.error || "Could not update this preorder batch.");
      }
      setNotice({ kind: "success", text: successText });
      window.setTimeout(() => window.location.reload(), 350);
    } catch (error) {
      setNotice({ kind: "error", text: error instanceof Error ? error.message : "Could not update this preorder batch." });
      setBusyBatchId(null);
    }
  }

  return (
    <div style={s.page}>
      {notice ? (
        <div style={{ ...s.notice, ...(notice.kind === "error" ? s.noticeError : s.noticeSuccess) }}>{notice.text}</div>
      ) : null}

      <div style={s.tabsRow}>
        <div style={s.tabs}>
          {TABS.map((item) => (
            <button
              key={item.id}
              type="button"
              onClick={() => setTab(item.id)}
              style={{ ...s.tab, ...(tab === item.id ? s.tabActive : {}) }}
            >
              {item.label}
            </button>
          ))}
        </div>
        {tab === "batches" ? (
          <div style={s.segmented}>
            {(["ALL", "AU", "USA"] as const).map((item) => (
              <button key={item} type="button" onClick={() => setMarket(item)} style={{ ...s.segmentButton, ...(market === item ? s.segmentActive : {}) }}>{item}</button>
            ))}
          </div>
        ) : null}
      </div>

      {tab === "batches" ? (
        <>
          <div style={s.cards}>
            <MetricCard label="Active batches" value={data.totals.activeBatches} hint="Enabled + eligible — click to filter" active={tileFilter === "active"} onClick={() => setTileFilter((f) => (f === "active" ? "all" : "active"))} />
            <MetricCard label="Incoming units" value={data.totals.incomingUnits} hint="Open production — click to filter" active={tileFilter === "incoming"} onClick={() => setTileFilter((f) => (f === "incoming" ? "all" : "incoming"))} />
            <MetricCard label="Reserved" value={data.totals.reservedUnits} hint="Has reservations — click to filter" active={tileFilter === "reserved"} onClick={() => setTileFilter((f) => (f === "reserved" ? "all" : "reserved"))} />
            <MetricCard label="Available capacity" value={data.totals.availableCapacity} hint="Has capacity — click to filter" active={tileFilter === "available"} onClick={() => setTileFilter((f) => (f === "available" ? "all" : "available"))} />
            <MetricCard label="Overallocated" value={data.totals.overallocatedUnits} hint="Needs attention — click to filter" danger={data.totals.overallocatedUnits > 0} active={tileFilter === "overallocated"} onClick={() => setTileFilter((f) => (f === "overallocated" ? "all" : "overallocated"))} />
          </div>

          {tileFilter !== "all" ? (
            <div style={s.toolbar}>
              <button type="button" onClick={() => setTileFilter("all")} style={s.clearFilter}>Showing {tileFilter} only · clear ✕</button>
            </div>
          ) : null}

          <div style={s.batchList}>
            {batches.length === 0 ? (
              <div style={s.empty}>No AU/USA production batches match this view yet.</div>
            ) : batches.map((batch) => (
              <div key={batch.id} style={s.batchCard}>
                <div style={s.batchTop}>
                  {/* Picture + title on the left, size grid to their right. */}
                  <div style={{ display: "flex", alignItems: "center", gap: 20, minWidth: 0, flexWrap: "wrap" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 14, minWidth: 0 }}>
                      {batch.imageUrl ? <img src={batch.imageUrl} alt="" style={{ width: 84, height: "auto", maxHeight: 130, borderRadius: 8, flexShrink: 0, display: "block" }} /> : <div style={{ width: 84, height: 104, background: "#f1f5f9", borderRadius: 8, flexShrink: 0 }} />}
                      <div style={{ minWidth: 0 }}>
                        <div style={s.productTitle}>{batch.productTitle}</div>
                        <div style={s.meta}>Batch #{batch.id} · {batch.supplier} · {batch.market ?? "No market"}</div>
                      </div>
                    </div>
                    {/* Table-like size grid: sizes across the top, Incoming / Reserved /
                        Available stacked down. */}
                    <div style={{ overflowX: "auto" }}>
                      <table style={{ borderCollapse: "collapse" }}>
                    <tbody>
                      <tr>
                        <td style={s.gridRowLabel} />
                        {batch.variants.map((v) => <td key={v.variantId} style={s.gridSizeHead}>{v.variantTitle || "Free"}</td>)}
                        <td style={s.gridTotalHead}>Total</td>
                      </tr>
                      <tr>
                        <td style={s.gridRowLabel}>Incoming</td>
                        {batch.variants.map((v) => <td key={v.variantId} style={s.gridCell}>{v.incomingRemaining}</td>)}
                        <td style={s.gridTotalCell}>{batch.totalIncoming}</td>
                      </tr>
                      <tr>
                        <td style={s.gridRowLabel}>Reserved</td>
                        {batch.variants.map((v) => <td key={v.variantId} style={{ ...s.gridCell, color: "#7c3aed" }}>{v.reservedQty}</td>)}
                        <td style={{ ...s.gridTotalCell, color: "#7c3aed" }}>{batch.totalReserved}</td>
                      </tr>
                      <tr>
                        <td style={s.gridRowLabel}>Available</td>
                        {batch.variants.map((v) => <td key={v.variantId} style={{ ...s.gridCell, color: v.availableToPreorder > 0 ? "#0f766e" : "#dc2626", fontWeight: 800 }}>{v.availableToPreorder}</td>)}
                        <td style={{ ...s.gridTotalCell, color: batch.totalAvailable > 0 ? "#0f766e" : "#dc2626" }}>{batch.totalAvailable}</td>
                      </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>
                  <div style={{ ...s.badge, ...(batch.enabled && batch.shopifySellingPlanActive ? s.badgeGreen : s.badgeGrey) }}>
                    {batch.enabled && batch.shopifySellingPlanActive ? "Pre-order active" : "Pre-order off"}
                  </div>
                </div>

                <BatchControls
                  batch={batch}
                  busy={busyBatchId === batch.id}
                  onManage={(payload, successText) => manageBatch(batch.id, payload, successText)}
                />
              </div>
            ))}
          </div>
        </>
      ) : tab === "orders" ? (
        <PreorderCustomerOrdersPanel orders={data.customerOrders} />
      ) : tab === "waitlist" ? (
        <>
          <NotifyBlockToggle enabled={data.configuration.notifyBlockEnabled} />
          <PreorderWaitlistPanel />
        </>
      ) : tab === "notifications" ? (
        <PreorderNotificationsPanel />
      ) : tab === "reports" ? (
        <PreorderReportsPanel />
      ) : tab === "settings" ? (
        <>
          <PreorderShopifyReadinessPanel />
          <PreorderWebhookStatusPanel />
          <PreorderSettingsPanel configuration={data.configuration} />
          <PreorderActivationReadinessPanel configuration={data.configuration} />
        </>
      ) : (
        <Placeholder tab={tab} />
      )}
    </div>
  );
}

function BatchControls({
  batch,
  busy,
  onManage,
}: {
  batch: PreorderDashboardBatch;
  busy: boolean;
  onManage: (payload: Record<string, unknown>, successText: string) => void;
}) {
  const [shipDate, setShipDate] = useState(dateInputValue(batch.shipDate));
  // "Live" = enabled + a Shopify selling plan active. Enabling does both in one
  // step (no separate "Activate on Shopify"); turning off removes both.
  const live = batch.enabled && batch.shopifySellingPlanActive;
  const canEnable = Boolean(batch.market); // destination (AUS/USA) is set

  return (
    <div style={s.controls}>
      <div style={s.controlRow}>
        <div style={s.controlField}>
          <span style={s.fieldCaption}>Batch</span>
          <span style={s.fieldValue}>#{batch.id}</span>
        </div>
        <div style={s.controlField}>
          <span style={s.fieldCaption}>Country</span>
          <span style={s.fieldValue}>{batch.market === "AU" ? "Australia" : batch.market === "USA" ? "USA" : "Not set"}</span>
        </div>
        <div style={s.controlField}>
          <span style={s.fieldCaption}>Dispatch date</span>
          <CalendarDatePicker value={shipDate} onChange={setShipDate} disabled={busy} />
        </div>
        <div style={{ flex: 1 }} />
        <button
          type="button"
          style={s.secondaryButton}
          disabled={busy}
          onClick={() => onManage({ operation: "update-settings", shipDate }, "Dispatch date saved.")}
        >
          {busy ? "Saving…" : "Save settings"}
        </button>
        <button
          type="button"
          style={{ ...s.primaryButton, ...(live ? s.pauseButton : {}) }}
          disabled={busy || (!live && !canEnable)}
          title={!live && !canEnable ? "Assign the batch to Send to AUS or Send to USA first." : undefined}
          onClick={() => (live
            ? onManage({ operation: "row-disable" }, "Pre-order turned off.")
            : onManage({ operation: "row-enable", shipDate }, "Pre-order enabled and live on Shopify."))}
        >
          {busy ? "Working…" : live ? "Turn off pre-order" : "Enable pre-order"}
        </button>
      </div>
      {!live && !canEnable ? (
        <div style={s.controlHint}>To enable, assign this batch to <strong>Send to AUS</strong> or <strong>Send to USA</strong>.</div>
      ) : null}
      {batch.pausedReason ? <div style={s.controlHint}>Paused reason: {batch.pausedReason}</div> : null}
    </div>
  );
}

function NotifyBlockToggle({ enabled }: { enabled: boolean }) {
  const [on, setOn] = useState(enabled);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const toggle = async () => {
    const next = !on;
    setBusy(true);
    setError(null);
    try {
      const response = await fetch("/api/preorder-manage", {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ operation: "set-notify-enabled", enabled: next }),
      });
      const result = await response.json().catch(() => ({})) as { ok?: boolean; error?: string; notifyBlockEnabled?: boolean };
      if (!response.ok || result.ok !== true) throw new Error(result.error || "Could not update the notify-me block.");
      setOn(result.notifyBlockEnabled ?? next);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not update the notify-me block.");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div style={{ border: "1px solid #e2e8f0", borderRadius: 12, padding: 16, marginBottom: 16, background: "#fff", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 16, flexWrap: "wrap" }}>
      <div style={{ maxWidth: 620 }}>
        <div style={{ fontSize: 15, fontWeight: 700, color: "#0f172a" }}>Storefront “Notify me” block</div>
        <div style={{ fontSize: 13, color: "#64748b", marginTop: 4 }}>
          Shows a “Notify me when available” form on out-of-stock variants. Turn this <strong>off</strong> if another
          back-in-stock app is running, so customers don’t see two forms. <strong>Pre-order is unaffected</strong> — only the
          notify-me fallback is hidden.
        </div>
        {error ? <div style={{ fontSize: 12, color: "#b91c1c", marginTop: 6 }}>{error}</div> : null}
      </div>
      <button
        type="button"
        onClick={toggle}
        disabled={busy}
        aria-pressed={on}
        title={on ? "Notify-me block is ON — click to turn off" : "Notify-me block is OFF — click to turn on"}
        style={{
          flexShrink: 0, display: "inline-flex", alignItems: "center", gap: 8,
          border: "none", borderRadius: 999, padding: "8px 16px", cursor: busy ? "default" : "pointer",
          fontSize: 13, fontWeight: 700, color: "#fff", opacity: busy ? 0.6 : 1,
          background: on ? "#16a34a" : "#9ca3af",
        }}
      >
        <span style={{ width: 8, height: 8, borderRadius: 999, background: "#fff" }} />
        {busy ? "Saving…" : on ? "ON — showing notify-me" : "OFF — hidden"}
      </button>
    </div>
  );
}

// A compact single-date calendar picker (nicer than the native date input),
// styled like the dashboard's: terracotta selected day, Monday-first grid.
// Value + onChange use YYYY-MM-DD. The popover is portaled to <body> so it can't
// be clipped by the card.
function CalendarDatePicker({ value, onChange, disabled }: { value: string; onChange: (v: string) => void; disabled?: boolean }) {
  const [open, setOpen] = useState(false);
  const btnRef = useRef<HTMLButtonElement | null>(null);
  const popRef = useRef<HTMLDivElement | null>(null);
  const [pos, setPos] = useState<{ top: number; left: number } | null>(null);
  const selected = value ? new Date(`${value}T00:00:00`) : null;
  const [view, setView] = useState(() => (selected && !Number.isNaN(selected.getTime()) ? new Date(selected) : new Date()));
  useEffect(() => {
    if (!open) return;
    const place = () => { const r = btnRef.current?.getBoundingClientRect(); if (r) setPos({ top: r.bottom + 6, left: r.left }); };
    place();
    window.addEventListener("resize", place);
    window.addEventListener("scroll", place, true);
    return () => { window.removeEventListener("resize", place); window.removeEventListener("scroll", place, true); };
  }, [open]);
  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => { const t = e.target as Node; if (btnRef.current?.contains(t) || popRef.current?.contains(t)) return; setOpen(false); };
    document.addEventListener("mousedown", onDown);
    return () => document.removeEventListener("mousedown", onDown);
  }, [open]);
  const fmt = (d: Date) => new Intl.DateTimeFormat("en-AU", { day: "numeric", month: "short", year: "numeric" }).format(d);
  const y = view.getFullYear(), m = view.getMonth();
  const startOffset = (new Date(y, m, 1).getDay() + 6) % 7; // Monday-first
  const daysInMonth = new Date(y, m + 1, 0).getDate();
  const cells: Array<number | null> = [];
  for (let i = 0; i < startOffset; i += 1) cells.push(null);
  for (let d = 1; d <= daysInMonth; d += 1) cells.push(d);
  const pick = (d: number) => {
    const iso = `${y}-${String(m + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
    onChange(iso);
    setOpen(false);
  };
  const isSel = (d: number) => selected && selected.getFullYear() === y && selected.getMonth() === m && selected.getDate() === d;
  return (
    <>
      <button ref={btnRef} type="button" disabled={disabled} onClick={() => setOpen((v) => !v)} style={s.dateButton}>
        <span>{selected && !Number.isNaN(selected.getTime()) ? fmt(selected) : "Pick a date"}</span>
        <span style={{ marginLeft: 10, color: "#94a3b8", fontSize: 11 }}>▾</span>
      </button>
      {open && pos && typeof document !== "undefined" && createPortal(
        <div ref={popRef} style={{ ...s.calPop, top: pos.top, left: pos.left }}>
          <div style={s.calHead}>
            <button type="button" style={s.calNav} onClick={() => setView(new Date(y, m - 1, 1))}>‹</button>
            <div style={{ fontWeight: 800, fontSize: 15, color: "#0f172a" }}>{new Intl.DateTimeFormat("en-AU", { month: "long", year: "numeric" }).format(view)}</div>
            <button type="button" style={s.calNav} onClick={() => setView(new Date(y, m + 1, 1))}>›</button>
          </div>
          <div style={s.calGrid}>
            {["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"].map((d) => <div key={d} style={s.calDow}>{d}</div>)}
            {cells.map((d, i) => (d === null ? <div key={`e${i}`} /> : (
              <button key={d} type="button" onClick={() => pick(d)} style={{ ...s.calDay, ...(isSel(d) ? s.calDaySel : {}) }}>{d}</button>
            )))}
          </div>
        </div>,
        document.body,
      )}
    </>
  );
}

function MetricCard({ label, value, hint, danger = false, active = false, onClick }: { label: string; value: number | string; hint: string; danger?: boolean; active?: boolean; onClick?: () => void }) {
  return (
    <div
      onClick={onClick}
      style={{ ...s.metricCard, ...(danger ? s.metricDanger : {}), ...(onClick ? { cursor: "pointer" } : {}), ...(active ? s.metricActive : {}) }}
    >
      <div style={s.metricLabel}>{label}</div>
      <div style={s.metricValue}>{value}</div>
      <div style={s.metricHint}>{hint}</div>
    </div>
  );
}

function MiniStat({ label, value }: { label: string; value: number | string }) {
  return (
    <div>
      <div style={s.miniLabel}>{label}</div>
      <div style={s.miniValue}>{value}</div>
    </div>
  );
}

function Placeholder({ tab }: { tab: TabId }) {
  const label = TABS.find((item) => item.id === tab)?.label ?? tab;
  return (
    <div style={s.placeholder}>
      <div style={s.placeholderIcon}>◌</div>
      <h2 style={{ margin: 0, fontSize: 18 }}>{label}</h2>
      <p style={{ margin: "6px 0 0", color: "#64748b", fontSize: 14 }}>
        This section is wired into the preorder area and will be filled as its backend phase is completed.
      </p>
    </div>
  );
}

const s: Record<string, React.CSSProperties> = {
  page: { padding: "4px 2px 40px" },
  header: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 20, marginBottom: 18 },
  title: { margin: 0, fontSize: 28, lineHeight: 1.15, color: "#0f172a" },
  subtitle: { margin: "6px 0 0", color: "#64748b", fontSize: 14 },
  headerPill: { padding: "7px 11px", borderRadius: 999, background: "#dbeafe", color: "#1d4ed8", fontWeight: 700, fontSize: 12 },
  notice: { marginBottom: 14, padding: "10px 12px", borderRadius: 9, fontSize: 13, fontWeight: 700 },
  noticeError: { background: "#fef2f2", border: "1px solid #fecaca", color: "#991b1b" },
  noticeSuccess: { background: "#f0fdf4", border: "1px solid #bbf7d0", color: "#166534" },
  tabsRow: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 20 },
  tabs: { display: "flex", gap: 8, flexWrap: "wrap" },
  tab: { border: "1px solid #e2e8f0", background: "#fff", borderRadius: 9, padding: "9px 15px", fontSize: 13, fontWeight: 700, color: "#475569", cursor: "pointer", boxShadow: "0 1px 2px rgba(15,23,42,0.05)" },
  tabActive: { background: "#C16452", color: "#fff", borderColor: "#C16452" },
  gridRowLabel: { padding: "4px 14px 4px 0", fontSize: 12, fontWeight: 700, color: "#64748b", textAlign: "right", whiteSpace: "nowrap" },
  gridSizeHead: { padding: "4px 10px", textAlign: "center", fontWeight: 800, fontSize: 13, borderBottom: "2px solid #e2e8f0", minWidth: 56, color: "#0f172a" },
  gridTotalHead: { padding: "4px 12px 4px 14px", textAlign: "center", fontWeight: 800, fontSize: 13, borderBottom: "2px solid #cbd5e1", borderLeft: "2px solid #cbd5e1", color: "#334155" },
  gridCell: { padding: "4px 10px", textAlign: "center", fontSize: 13, color: "#0f172a" },
  gridTotalCell: { padding: "4px 12px 4px 14px", textAlign: "center", fontSize: 13, fontWeight: 800, borderLeft: "2px solid #cbd5e1", color: "#334155" },
  controlRow: { display: "flex", alignItems: "flex-end", gap: 16, flexWrap: "wrap", marginTop: 3, paddingTop: 6, borderTop: "1px solid #f1f5f9" },
  controlField: { display: "flex", flexDirection: "column", gap: 5 },
  fieldCaption: { fontSize: 11, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", letterSpacing: ".04em" },
  fieldValue: { display: "flex", alignItems: "center", height: 40, fontSize: 15, fontWeight: 800, color: "#0f172a" },
  dateButton: { display: "inline-flex", alignItems: "center", height: 40, boxSizing: "border-box", border: "1px solid #cbd5e1", borderRadius: 10, padding: "0 14px", fontSize: 14, fontWeight: 700, background: "#fff", color: "#0f172a", cursor: "pointer" },
  calPop: { position: "fixed", zIndex: 4001, background: "#fff", border: "1px solid #e2e8f0", borderRadius: 14, boxShadow: "0 18px 44px rgba(15,23,42,0.22)", padding: 14, width: 300 },
  calHead: { display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 },
  calNav: { border: "1px solid #e2e8f0", background: "#fff", borderRadius: 8, width: 30, height: 30, cursor: "pointer", fontSize: 16, color: "#475569", lineHeight: 1 },
  calGrid: { display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 3 },
  calDow: { textAlign: "center", fontSize: 11, fontWeight: 700, color: "#94a3b8", padding: "2px 0" },
  calDay: { border: "none", background: "transparent", borderRadius: 8, height: 36, cursor: "pointer", fontSize: 14, fontWeight: 600, color: "#0f172a" },
  calDaySel: { background: "#C16452", color: "#fff", fontWeight: 800 },
  cards: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(165px, 1fr))", gap: 12, marginBottom: 18 },
  metricCard: { background: "white", border: "1px solid #e2e8f0", borderRadius: 12, padding: 16, boxShadow: "0 1px 2px rgba(15,23,42,0.04)" },
  metricDanger: { border: "1px solid #fecaca", background: "#fff7f7" },
  metricActive: { border: "2px solid #C16452", boxShadow: "0 2px 8px rgba(193,100,82,0.25)" },
  clearFilter: { border: "1px solid #C16452", background: "#fff", color: "#C16452", borderRadius: 999, padding: "7px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" },
  metricLabel: { fontSize: 12, fontWeight: 800, color: "#64748b", textTransform: "uppercase", letterSpacing: ".04em" },
  metricValue: { fontSize: 27, fontWeight: 800, color: "#0f172a", marginTop: 5 },
  metricHint: { fontSize: 11, color: "#94a3b8", marginTop: 4 },
  toolbar: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, marginBottom: 14, flexWrap: "wrap" },
  segmented: { display: "inline-flex", padding: 3, borderRadius: 9, background: "#f1f5f9" },
  segmentButton: { border: 0, borderRadius: 7, background: "transparent", padding: "7px 10px", fontWeight: 800, fontSize: 12, color: "#64748b", cursor: "pointer" },
  segmentActive: { background: "white", color: "#0f172a", boxShadow: "0 1px 2px rgba(15,23,42,.08)" },
  search: { minWidth: 250, flex: "0 1 360px", border: "1px solid #cbd5e1", borderRadius: 9, padding: "9px 11px", fontSize: 13, background: "white" },
  batchList: { display: "flex", flexDirection: "column", gap: 12 },
  batchCard: { background: "white", border: "1px solid #e2e8f0", borderRadius: 14, padding: 18, boxShadow: "0 1px 3px rgba(15,23,42,.04)" },
  batchTop: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 15 },
  productTitle: { fontSize: 17, fontWeight: 800, color: "#0f172a" },
  meta: { fontSize: 12, color: "#94a3b8", marginTop: 4 },
  badge: { borderRadius: 999, padding: "6px 9px", fontWeight: 800, fontSize: 11, whiteSpace: "nowrap" },
  badgeGreen: { background: "#dcfce7", color: "#166534" },
  badgeAmber: { background: "#fef3c7", color: "#92400e" },
  badgeGrey: { background: "#f1f5f9", color: "#475569" },
  batchStats: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(110px, 1fr))", gap: 12, marginTop: 17 },
  miniLabel: { fontSize: 11, color: "#94a3b8", fontWeight: 700, textTransform: "uppercase" },
  miniValue: { fontSize: 14, color: "#0f172a", fontWeight: 800, marginTop: 2 },
  progressTrack: { height: 7, background: "#e2e8f0", borderRadius: 999, marginTop: 15, overflow: "hidden" },
  progressFill: { height: "100%", background: "#0f766e", borderRadius: 999 },
  variantGrid: { marginTop: 14, borderTop: "1px solid #f1f5f9" },
  variantRow: { display: "grid", gridTemplateColumns: "minmax(120px, 1.4fr) repeat(3, minmax(90px, .7fr))", gap: 10, padding: "9px 2px", borderBottom: "1px solid #f8fafc", fontSize: 12, color: "#64748b" },
  variantName: { color: "#334155", fontWeight: 700 },
  controls: { marginTop: 16, paddingTop: 15, borderTop: "1px solid #e2e8f0" },
  controlsTitle: { fontSize: 12, fontWeight: 800, color: "#475569", textTransform: "uppercase", letterSpacing: ".04em", marginBottom: 10 },
  controlGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 10 },
  fieldLabel: { display: "flex", flexDirection: "column", gap: 5, fontSize: 11, fontWeight: 700, color: "#64748b" },
  optional: { fontWeight: 500, color: "#94a3b8" },
  input: { border: "1px solid #cbd5e1", borderRadius: 8, padding: "8px 9px", fontSize: 13, background: "white", color: "#0f172a" },
  controlActions: { display: "flex", gap: 8, justifyContent: "flex-end", flexWrap: "wrap", marginTop: 11 },
  secondaryButton: { display: "inline-flex", alignItems: "center", height: 40, boxSizing: "border-box", border: "1px solid #cbd5e1", background: "white", color: "#334155", borderRadius: 10, padding: "0 16px", fontSize: 13, fontWeight: 800, cursor: "pointer" },
  primaryButton: { display: "inline-flex", alignItems: "center", height: 40, boxSizing: "border-box", border: "1px solid #C16452", background: "#C16452", color: "white", borderRadius: 10, padding: "0 16px", fontSize: 13, fontWeight: 800, cursor: "pointer" },
  pauseButton: { border: "1px solid #b45309", background: "#b45309" },
  controlHint: { marginTop: 9, fontSize: 11, color: "#64748b" },
  empty: { padding: 36, textAlign: "center", color: "#64748b", background: "white", border: "1px dashed #cbd5e1", borderRadius: 12 },
  placeholder: { padding: 50, textAlign: "center", background: "white", border: "1px solid #e2e8f0", borderRadius: 14 },
  placeholderIcon: { width: 44, height: 44, borderRadius: 999, display: "grid", placeItems: "center", background: "#f1f5f9", color: "#64748b", fontSize: 24, margin: "0 auto 12px" },
};
