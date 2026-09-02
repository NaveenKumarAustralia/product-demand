import { useMemo, useState } from "react";
import type { PreorderDashboardBatch, PreorderDashboardData } from "./preorder/preorder-dashboard.server";
import { PreorderCustomerOrdersPanel, PreorderSettingsPanel } from "./preorder/preorder-operations-panels";

type Props = {
  data: PreorderDashboardData;
};

type TabId = "dashboard" | "batches" | "orders" | "waitlist" | "notifications" | "reports" | "settings";

const TABS: Array<{ id: TabId; label: string }> = [
  { id: "dashboard", label: "Dashboard" },
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

export function PreordersDashboard({ data }: Props) {
  const [tab, setTab] = useState<TabId>("dashboard");
  const [market, setMarket] = useState<"ALL" | "AU" | "USA">("ALL");
  const [search, setSearch] = useState("");
  const [notice, setNotice] = useState<{ kind: "error" | "success"; text: string } | null>(null);
  const [busyBatchId, setBusyBatchId] = useState<number | null>(null);

  const batches = useMemo(() => {
    const q = search.trim().toLowerCase();
    return data.batches.filter((batch) => {
      if (market !== "ALL" && batch.market !== market) return false;
      if (q && !`${batch.productTitle} ${batch.supplier} ${batch.market ?? ""}`.toLowerCase().includes(q)) return false;
      return true;
    });
  }, [data.batches, market, search]);

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
      <div style={s.header}>
        <div>
          <h1 style={s.title}>Pre-orders</h1>
          <p style={s.subtitle}>Incoming production capacity, customer reservations and preorder health.</p>
        </div>
        <div style={s.headerPill}>Operations preview</div>
      </div>

      {notice ? (
        <div style={{ ...s.notice, ...(notice.kind === "error" ? s.noticeError : s.noticeSuccess) }}>{notice.text}</div>
      ) : null}

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

      {tab === "dashboard" || tab === "batches" ? (
        <>
          <div style={s.cards}>
            <MetricCard label="Active batches" value={data.totals.activeBatches} hint="Enabled + eligible" />
            <MetricCard label="Incoming units" value={data.totals.incomingUnits} hint="AU + USA open production" />
            <MetricCard label="Reserved" value={data.totals.reservedUnits} hint="Live reservation ledger" />
            <MetricCard label="Available capacity" value={data.totals.availableCapacity} hint="After reservations + safety buffer" />
            <MetricCard label="Overallocated" value={data.totals.overallocatedUnits} hint="Needs attention" danger={data.totals.overallocatedUnits > 0} />
          </div>

          <div style={s.toolbar}>
            <div style={s.segmented}>
              {(["ALL", "AU", "USA"] as const).map((item) => (
                <button key={item} type="button" onClick={() => setMarket(item)} style={{ ...s.segmentButton, ...(market === item ? s.segmentActive : {}) }}>{item}</button>
              ))}
            </div>
            <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search product or supplier" style={s.search} />
          </div>

          <div style={s.batchList}>
            {batches.length === 0 ? (
              <div style={s.empty}>No AU/USA production batches match this view yet.</div>
            ) : batches.map((batch) => (
              <div key={batch.id} style={s.batchCard}>
                <div style={s.batchTop}>
                  <div>
                    <div style={s.productTitle}>{batch.productTitle}</div>
                    <div style={s.meta}>Batch #{batch.id} · {batch.supplier} · {batch.market ?? "No market"}</div>
                  </div>
                  <div style={{ ...s.badge, ...(batch.enabled && batch.eligible ? s.badgeGreen : batch.supplierStatus === "on_production" ? s.badgeAmber : s.badgeGrey) }}>
                    {batch.enabled && batch.eligible ? "Preorder active" : batch.supplierStatus === "on_production" ? "Eligible / Off" : statusLabel(batch.supplierStatus)}
                  </div>
                </div>

                <div style={s.batchStats}>
                  <MiniStat label="Incoming" value={batch.totalIncoming} />
                  <MiniStat label="Reserved" value={batch.totalReserved} />
                  <MiniStat label="Remaining" value={batch.totalAvailable} />
                  <MiniStat label="Safety" value={batch.safetyBufferQty != null ? batch.safetyBufferQty : `${batch.safetyBufferPercent}%`} />
                  <MiniStat label="Ship date" value={formatDate(batch.shipDate || batch.productionEta)} />
                </div>

                <div style={s.progressTrack}>
                  <div
                    style={{
                      ...s.progressFill,
                      width: `${Math.min(100, batch.totalIncoming > 0 ? (batch.totalReserved / batch.totalIncoming) * 100 : 0)}%`,
                    }}
                  />
                </div>

                <div style={s.variantGrid}>
                  {batch.variants.map((variant) => (
                    <div key={`${batch.id}-${variant.variantId}`} style={s.variantRow}>
                      <div style={s.variantName}>{variant.variantTitle || "Default"}</div>
                      <div>Incoming <strong>{variant.incomingRemaining}</strong></div>
                      <div>Reserved <strong>{variant.reservedQty}</strong></div>
                      <div>Available <strong>{variant.availableToPreorder}</strong></div>
                    </div>
                  ))}
                </div>

                {tab === "batches" ? (
                  <BatchControls
                    batch={batch}
                    busy={busyBatchId === batch.id}
                    onManage={(payload, successText) => manageBatch(batch.id, payload, successText)}
                  />
                ) : null}
              </div>
            ))}
          </div>
        </>
      ) : tab === "orders" ? (
        <PreorderCustomerOrdersPanel orders={data.customerOrders} />
      ) : tab === "settings" ? (
        <PreorderSettingsPanel configuration={data.configuration} />
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
  const [bufferPercent, setBufferPercent] = useState(String(batch.safetyBufferPercent));
  const [bufferQty, setBufferQty] = useState(batch.safetyBufferQty == null ? "" : String(batch.safetyBufferQty));
  const canActivate = batch.supplierStatus === "on_production" && (batch.destination === "send_to_au" || batch.destination === "send_to_usa");

  return (
    <div style={s.controls}>
      <div style={s.controlsTitle}>Staff controls</div>
      <div style={s.controlGrid}>
        <label style={s.fieldLabel}>
          Expected dispatch
          <input type="date" value={shipDate} onChange={(e) => setShipDate(e.target.value)} style={s.input} disabled={busy} />
        </label>
        <label style={s.fieldLabel}>
          Safety buffer %
          <input type="number" min={0} max={50} step={0.5} value={bufferPercent} onChange={(e) => setBufferPercent(e.target.value)} style={s.input} disabled={busy} />
        </label>
        <label style={s.fieldLabel}>
          Fixed buffer qty <span style={s.optional}>(optional override)</span>
          <input type="number" min={0} step={1} value={bufferQty} onChange={(e) => setBufferQty(e.target.value)} placeholder="Use %" style={s.input} disabled={busy} />
        </label>
      </div>
      <div style={s.controlActions}>
        <button
          type="button"
          style={s.secondaryButton}
          disabled={busy}
          onClick={() => onManage({
            operation: "update-settings",
            shipDate,
            safetyBufferPercent: bufferPercent,
            safetyBufferQty: bufferQty,
          }, "Preorder batch settings saved.")}
        >
          {busy ? "Saving…" : "Save settings"}
        </button>
        <button
          type="button"
          style={{ ...s.primaryButton, ...(batch.enabled ? s.pauseButton : {}) }}
          disabled={busy || (!batch.enabled && !canActivate)}
          title={!batch.enabled && !canActivate ? "Batch must be On Production and assigned to AUS or USA first." : undefined}
          onClick={() => onManage({ operation: "set-enabled", enabled: !batch.enabled }, batch.enabled ? "Preorder paused." : "Preorder enabled.")}
        >
          {busy ? "Working…" : batch.enabled ? "Pause preorder" : "Enable preorder"}
        </button>
      </div>
      {!batch.enabled && !canActivate ? (
        <div style={s.controlHint}>To enable, this production batch must be <strong>On Production</strong> and assigned to <strong>Send to AUS</strong> or <strong>Send to USA</strong>.</div>
      ) : null}
      {batch.pausedReason ? <div style={s.controlHint}>Paused reason: {batch.pausedReason}</div> : null}
    </div>
  );
}

function MetricCard({ label, value, hint, danger = false }: { label: string; value: number | string; hint: string; danger?: boolean }) {
  return (
    <div style={{ ...s.metricCard, ...(danger ? s.metricDanger : {}) }}>
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
  tabs: { display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 20, paddingBottom: 12, borderBottom: "1px solid #e2e8f0" },
  tab: { border: 0, background: "transparent", borderRadius: 8, padding: "8px 11px", fontSize: 13, fontWeight: 700, color: "#64748b", cursor: "pointer" },
  tabActive: { background: "#0f172a", color: "white" },
  cards: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(165px, 1fr))", gap: 12, marginBottom: 18 },
  metricCard: { background: "white", border: "1px solid #e2e8f0", borderRadius: 12, padding: 16, boxShadow: "0 1px 2px rgba(15,23,42,0.04)" },
  metricDanger: { border: "1px solid #fecaca", background: "#fff7f7" },
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
  secondaryButton: { border: "1px solid #cbd5e1", background: "white", color: "#334155", borderRadius: 8, padding: "8px 12px", fontSize: 12, fontWeight: 800, cursor: "pointer" },
  primaryButton: { border: "1px solid #0f766e", background: "#0f766e", color: "white", borderRadius: 8, padding: "8px 12px", fontSize: 12, fontWeight: 800, cursor: "pointer" },
  pauseButton: { border: "1px solid #b45309", background: "#b45309" },
  controlHint: { marginTop: 9, fontSize: 11, color: "#64748b" },
  empty: { padding: 36, textAlign: "center", color: "#64748b", background: "white", border: "1px dashed #cbd5e1", borderRadius: 12 },
  placeholder: { padding: 50, textAlign: "center", background: "white", border: "1px solid #e2e8f0", borderRadius: 14 },
  placeholderIcon: { width: 44, height: 44, borderRadius: 999, display: "grid", placeItems: "center", background: "#f1f5f9", color: "#64748b", fontSize: 24, margin: "0 auto 12px" },
};
