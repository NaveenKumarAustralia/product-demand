import { useEffect, useMemo, useState } from "react";
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

export function PreorderCustomerOrdersPanel({ orders }: { orders: PreorderDashboardCustomerOrder[] }) {
  if (!orders.length) {
    return <div style={s.empty}>No preorder reservations have been created yet. Customer orders will appear here once Shopify order allocation is connected.</div>;
  }

  return (
    <div style={s.stack}>
      {orders.map((order) => (
        <div key={order.shopifyOrderId} style={s.card}>
          <div style={s.rowBetween}>
            <div>
              <div style={s.title}>{order.shopifyOrderName || order.shopifyOrderId}</div>
              <div style={s.muted}>{order.customerEmail || "No customer email"} · {order.market} · {order.totalQuantity} unit{order.totalQuantity === 1 ? "" : "s"}</div>
            </div>
            <div style={s.badge}>Reserved {formatDate(order.reservedAt)}</div>
          </div>
          <div style={s.lines}>
            {order.lines.map((line) => (
              <div key={line.reservationId} style={s.line}>
                <div>
                  <strong>{line.variantTitle || line.sku || line.variantId}</strong>
                  <div style={s.small}>Batch #{line.supplierOrderId}{line.sku ? ` · ${line.sku}` : ""}</div>
                </div>
                <div style={s.lineMeta}>Qty <strong>{line.quantity}</strong></div>
                <div style={s.lineMeta}>{line.status}</div>
                <div style={s.lineMeta}>Dispatch <strong>{formatDate(line.expectedShipDate)}</strong></div>
              </div>
            ))}
          </div>
        </div>
      ))}
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
  activeBatches?: Array<{ id: number; productTitle: string; supplier: string; destination: string | null; reservedQty: number }>;
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
            {(report.activeBatches ?? []).slice(0, 20).map((batch) => (
              <div key={batch.id} style={s.reportBatchRow}>
                <div><strong>{batch.productTitle}</strong><div style={s.small}>Batch #{batch.id} · {batch.supplier}</div></div>
                <div style={s.lineMeta}>{batch.destination === "send_to_usa" ? "USA" : "AU"}</div>
                <div style={s.reportNumber}>{batch.reservedQty}</div>
              </div>
            ))}
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
  reportBatchRow: { display: "grid", gridTemplateColumns: "minmax(180px, 1fr) 90px 80px", gap: 10, alignItems: "center", padding: "10px 0", borderBottom: "1px solid #f8fafc", fontSize: 12 },
  reportNumber: { textAlign: "right", fontSize: 16, fontWeight: 800, color: "#0f172a" },
  failureRow: { display: "grid", gridTemplateColumns: "minmax(160px, .8fr) minmax(220px, 1.5fr)", gap: 16, padding: "10px 0", borderBottom: "1px solid #f8fafc", fontSize: 12 },
  failureMessage: { color: "#991b1b" },
};
