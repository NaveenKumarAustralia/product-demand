import { useEffect, useMemo, useState } from "react";

type WaitlistRow = {
  id: number;
  shop: string;
  email: string;
  productId: string | null;
  productTitle: string | null;
  variantId: string;
  variantTitle: string | null;
  sku: string | null;
  market: string;
  status: string;
  source: string;
  notifiedAt: string | null;
  convertedOrderId: string | null;
  createdAt: string;
};

type SummaryRow = { market: string; status: string; count: number };

type ApiResponse = {
  ok?: boolean;
  rows?: WaitlistRow[];
  summary?: SummaryRow[];
  error?: string;
};

function formatDate(value: string | null) {
  if (!value) return "—";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "—";
  return new Intl.DateTimeFormat("en-AU", { day: "numeric", month: "short", year: "numeric" }).format(date);
}

export function PreorderWaitlistPanel() {
  const [rows, setRows] = useState<WaitlistRow[]>([]);
  const [summary, setSummary] = useState<SummaryRow[]>([]);
  const [market, setMarket] = useState<"ALL" | "AU" | "USA">("ALL");
  const [status, setStatus] = useState<"ALL" | "waiting" | "notified" | "converted" | "removed">("ALL");
  const [search, setSearch] = useState("");
  const [loading, setLoading] = useState(true);
  const [busyId, setBusyId] = useState<number | null>(null);
  const [notice, setNotice] = useState<{ kind: "error" | "success"; text: string } | null>(null);

  async function load() {
    setLoading(true);
    setNotice(null);
    try {
      const params = new URLSearchParams();
      if (market !== "ALL") params.set("market", market);
      if (status !== "ALL") params.set("status", status);
      if (search.trim()) params.set("q", search.trim());
      const response = await fetch(`/api/preorder-waitlist?${params.toString()}`, { credentials: "same-origin" });
      const result = await response.json().catch(() => ({})) as ApiResponse;
      if (!response.ok || result.ok !== true) throw new Error(result.error || "Could not load waitlist.");
      setRows(result.rows ?? []);
      setSummary(result.summary ?? []);
    } catch (error) {
      setNotice({ kind: "error", text: error instanceof Error ? error.message : "Could not load waitlist." });
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    const timer = window.setTimeout(() => void load(), 180);
    return () => window.clearTimeout(timer);
  }, [market, status, search]);

  const totals = useMemo(() => {
    const count = (targetStatus: string, targetMarket?: string) => summary
      .filter((row) => row.status === targetStatus && (!targetMarket || row.market === targetMarket))
      .reduce((sum, row) => sum + row.count, 0);
    return {
      waiting: count("waiting"),
      auWaiting: count("waiting", "AU"),
      usaWaiting: count("waiting", "USA"),
      notified: count("notified"),
      converted: count("converted"),
    };
  }, [summary]);

  async function changeStatus(id: number, nextStatus: string) {
    setBusyId(id);
    setNotice(null);
    try {
      const response = await fetch("/api/preorder-waitlist", {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ id, status: nextStatus }),
      });
      const result = await response.json().catch(() => ({})) as ApiResponse;
      if (!response.ok || result.ok !== true) throw new Error(result.error || "Could not update waitlist entry.");
      setNotice({ kind: "success", text: "Waitlist entry updated." });
      await load();
    } catch (error) {
      setNotice({ kind: "error", text: error instanceof Error ? error.message : "Could not update waitlist entry." });
    } finally {
      setBusyId(null);
    }
  }

  return (
    <div style={s.stack}>
      {notice ? <div style={{ ...s.notice, ...(notice.kind === "error" ? s.noticeError : s.noticeSuccess) }}>{notice.text}</div> : null}

      <div style={s.cards}>
        <Metric label="Waiting" value={totals.waiting} hint="Ready for stock/preorder opportunity" />
        <Metric label="AU waiting" value={totals.auWaiting} hint="Australia production pool" />
        <Metric label="USA waiting" value={totals.usaWaiting} hint="USA production pool" />
        <Metric label="Notified" value={totals.notified} hint="Customer update recorded" />
        <Metric label="Converted" value={totals.converted} hint="Converted to an order" />
      </div>

      <div style={s.toolbar}>
        <div style={s.group}>
          {(["ALL", "AU", "USA"] as const).map((item) => (
            <button key={item} type="button" onClick={() => setMarket(item)} style={{ ...s.filter, ...(market === item ? s.filterActive : {}) }}>{item}</button>
          ))}
        </div>
        <select value={status} onChange={(e) => setStatus(e.target.value as typeof status)} style={s.select}>
          <option value="ALL">All statuses</option>
          <option value="waiting">Waiting</option>
          <option value="notified">Notified</option>
          <option value="converted">Converted</option>
          <option value="removed">Removed</option>
        </select>
        <input value={search} onChange={(e) => setSearch(e.target.value)} style={s.search} placeholder="Search email, product, size or SKU" />
      </div>

      <div style={s.card}>
        <div style={s.headerRow}>
          <div>
            <div style={s.title}>Back in Stock waitlist</div>
            <div style={s.muted}>Internal management view. Storefront signup is intentionally not enabled yet.</div>
          </div>
          <div style={s.badge}>{loading ? "Loading…" : `${rows.length} shown`}</div>
        </div>

        {loading ? (
          <div style={s.empty}>Loading waitlist…</div>
        ) : rows.length === 0 ? (
          <div style={s.empty}>No waitlist entries match this view yet.</div>
        ) : (
          <div style={s.tableWrap}>
            <div style={s.tableHeader}>
              <div>Customer</div><div>Product / variant</div><div>Market</div><div>Status</div><div>Added</div><div>Actions</div>
            </div>
            {rows.map((row) => (
              <div key={row.id} style={s.tableRow}>
                <div><strong>{row.email}</strong><div style={s.small}>{row.source}</div></div>
                <div><strong>{row.productTitle || "Unknown product"}</strong><div style={s.small}>{row.variantTitle || row.sku || row.variantId}</div></div>
                <div><span style={s.marketBadge}>{row.market}</span></div>
                <div><span style={{ ...s.statusBadge, ...(row.status === "waiting" ? s.waiting : row.status === "converted" ? s.converted : {}) }}>{row.status}</span></div>
                <div>{formatDate(row.createdAt)}{row.notifiedAt ? <div style={s.small}>Notified {formatDate(row.notifiedAt)}</div> : null}</div>
                <div style={s.actions}>
                  {row.status !== "waiting" ? <button disabled={busyId === row.id} style={s.secondary} onClick={() => void changeStatus(row.id, "waiting")}>Waiting</button> : null}
                  {row.status !== "notified" ? <button disabled={busyId === row.id} style={s.secondary} onClick={() => void changeStatus(row.id, "notified")}>Notified</button> : null}
                  {row.status !== "converted" ? <button disabled={busyId === row.id} style={s.primary} onClick={() => void changeStatus(row.id, "converted")}>Converted</button> : null}
                  {row.status !== "removed" ? <button disabled={busyId === row.id} style={s.danger} onClick={() => void changeStatus(row.id, "removed")}>Remove</button> : null}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function Metric({ label, value, hint }: { label: string; value: number; hint: string }) {
  return <div style={s.metric}><div style={s.metricLabel}>{label}</div><div style={s.metricValue}>{value}</div><div style={s.small}>{hint}</div></div>;
}

const s: Record<string, React.CSSProperties> = {
  stack: { display: "flex", flexDirection: "column", gap: 14 },
  cards: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 10 },
  metric: { background: "white", border: "1px solid #e2e8f0", borderRadius: 12, padding: 14 },
  metricLabel: { fontSize: 11, color: "#64748b", fontWeight: 800, textTransform: "uppercase" },
  metricValue: { fontSize: 25, fontWeight: 800, color: "#0f172a", marginTop: 3 },
  toolbar: { display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" },
  group: { display: "inline-flex", gap: 4, padding: 3, borderRadius: 9, background: "#f1f5f9" },
  filter: { border: 0, background: "transparent", padding: "7px 10px", borderRadius: 7, fontWeight: 800, fontSize: 12, cursor: "pointer", color: "#64748b" },
  filterActive: { background: "white", color: "#0f172a", boxShadow: "0 1px 2px rgba(15,23,42,.08)" },
  select: { border: "1px solid #cbd5e1", borderRadius: 8, padding: "8px 9px", background: "white", fontSize: 12 },
  search: { flex: "1 1 260px", minWidth: 220, border: "1px solid #cbd5e1", borderRadius: 8, padding: "8px 10px", fontSize: 12 },
  card: { background: "white", border: "1px solid #e2e8f0", borderRadius: 14, padding: 16 },
  headerRow: { display: "flex", justifyContent: "space-between", gap: 12, alignItems: "flex-start" },
  title: { fontSize: 16, fontWeight: 800, color: "#0f172a" },
  muted: { fontSize: 12, color: "#64748b", marginTop: 4 },
  small: { fontSize: 10, color: "#94a3b8", marginTop: 3 },
  badge: { padding: "5px 8px", borderRadius: 999, background: "#f1f5f9", color: "#475569", fontSize: 11, fontWeight: 700 },
  empty: { padding: 34, textAlign: "center", color: "#64748b" },
  tableWrap: { overflowX: "auto", marginTop: 14, border: "1px solid #f1f5f9", borderRadius: 10 },
  tableHeader: { display: "grid", gridTemplateColumns: "1.2fr 1.4fr .55fr .7fr .8fr 1.8fr", gap: 10, minWidth: 900, padding: "9px 10px", background: "#f8fafc", fontSize: 10, fontWeight: 800, color: "#64748b", textTransform: "uppercase" },
  tableRow: { display: "grid", gridTemplateColumns: "1.2fr 1.4fr .55fr .7fr .8fr 1.8fr", gap: 10, minWidth: 900, alignItems: "center", padding: "10px", borderTop: "1px solid #f1f5f9", fontSize: 11, color: "#475569" },
  marketBadge: { padding: "4px 7px", borderRadius: 999, background: "#e0f2fe", color: "#075985", fontSize: 10, fontWeight: 800 },
  statusBadge: { padding: "4px 7px", borderRadius: 999, background: "#f1f5f9", color: "#475569", fontSize: 10, fontWeight: 800, textTransform: "capitalize" },
  waiting: { background: "#fef3c7", color: "#92400e" },
  converted: { background: "#dcfce7", color: "#166534" },
  actions: { display: "flex", gap: 5, flexWrap: "wrap" },
  secondary: { border: "1px solid #cbd5e1", background: "white", color: "#334155", borderRadius: 7, padding: "5px 7px", fontSize: 10, fontWeight: 700, cursor: "pointer" },
  primary: { border: "1px solid #0f766e", background: "#0f766e", color: "white", borderRadius: 7, padding: "5px 7px", fontSize: 10, fontWeight: 700, cursor: "pointer" },
  danger: { border: "1px solid #fecaca", background: "#fff7f7", color: "#991b1b", borderRadius: 7, padding: "5px 7px", fontSize: 10, fontWeight: 700, cursor: "pointer" },
  notice: { padding: "9px 11px", borderRadius: 8, fontSize: 12, fontWeight: 700 },
  noticeError: { background: "#fef2f2", border: "1px solid #fecaca", color: "#991b1b" },
  noticeSuccess: { background: "#f0fdf4", border: "1px solid #bbf7d0", color: "#166534" },
};
