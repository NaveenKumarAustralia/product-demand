import { useSearchParams } from "react-router";
import type { ProductionCutsData } from "./routes/portal._index";

function fmtDate(iso: string) {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  return new Intl.DateTimeFormat("en-AU", { weekday: "short", day: "numeric", month: "short", year: "numeric" }).format(d);
}
function fmtDay(iso: string) {
  const d = new Date(`${iso}T00:00:00`);
  if (Number.isNaN(d.getTime())) return iso;
  return new Intl.DateTimeFormat("en-AU", { weekday: "short", day: "numeric", month: "short" }).format(d);
}

export function ProductionCutsPanel({ data }: { data: ProductionCutsData }) {
  const [params, setParams] = useSearchParams();
  const setRange = (patch: { from?: string; to?: string }) => {
    const next = new URLSearchParams(params);
    next.set("page", "production-cuts");
    if (patch.from !== undefined) { if (patch.from) next.set("from", patch.from); else next.delete("from"); }
    if (patch.to !== undefined) { if (patch.to) next.set("to", patch.to); else next.delete("to"); }
    setParams(next);
  };
  const maxDay = Math.max(1, ...data.byDay.map((d) => d.pieces));

  return (
    <div style={s.page}>
      {/* Summary cards */}
      <div style={s.cards}>
        <Card label="Cut today" pieces={data.totals.today.pieces} orders={data.totals.today.orders} accent="#0f766e" />
        <Card label="This week" pieces={data.totals.week.pieces} orders={data.totals.week.orders} accent="#1d4ed8" />
        <Card label="This month" pieces={data.totals.month.pieces} orders={data.totals.month.orders} accent="#7c3aed" />
        <Card label="Selected range" pieces={data.totals.range.pieces} orders={data.totals.range.orders} accent="#b45309" />
      </div>

      {/* Range picker */}
      <div style={s.toolbar}>
        <label style={s.field}><span style={s.fieldLabel}>From</span>
          <input type="date" value={data.from} max={data.to} onChange={(e) => setRange({ from: e.target.value })} style={s.dateInput} />
        </label>
        <label style={s.field}><span style={s.fieldLabel}>To</span>
          <input type="date" value={data.to} min={data.from} onChange={(e) => setRange({ to: e.target.value })} style={s.dateInput} />
        </label>
        <div style={{ flex: 1 }} />
        <span style={s.rangeHint}>{data.rows.length} cut{data.rows.length === 1 ? "" : "s"} in range · {data.totals.range.pieces} pieces</span>
      </div>

      <div style={s.split}>
        {/* Per-day totals */}
        <div style={s.panel}>
          <h3 style={s.panelTitle}>Pieces cut per day</h3>
          {data.byDay.length === 0 ? (
            <div style={s.empty}>No cuts in this range.</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 6, maxHeight: 460, overflowY: "auto" }}>
              {data.byDay.map((d) => (
                <div key={d.date} style={{ display: "grid", gridTemplateColumns: "118px 1fr 58px", alignItems: "center", gap: 10 }}>
                  <span style={{ fontSize: 12.5, color: "#475569", fontWeight: 600, whiteSpace: "nowrap" }}>{fmtDay(d.date)}</span>
                  <span style={{ height: 16, borderRadius: 5, background: "#e0f2fe", position: "relative", overflow: "hidden" }}>
                    <span style={{ position: "absolute", inset: 0, width: `${Math.round((d.pieces / maxDay) * 100)}%`, background: "#0ea5e9", borderRadius: 5 }} />
                  </span>
                  <span style={{ fontSize: 13, fontWeight: 800, color: "#0f172a", textAlign: "right", fontVariantNumeric: "tabular-nums" }}>{d.pieces}</span>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* The log */}
        <div style={s.panel}>
          <h3 style={s.panelTitle}>Cut log</h3>
          <div style={{ overflow: "auto", maxHeight: 460, border: "1px solid #eef2f7", borderRadius: 10 }}>
            <table style={{ borderCollapse: "collapse", width: "100%", minWidth: 520 }}>
              <thead>
                <tr>
                  <th style={s.th}>Cut date</th>
                  <th style={s.th}>Product</th>
                  <th style={s.th}>Type</th>
                  <th style={{ ...s.th, textAlign: "right" }}>Pieces</th>
                </tr>
              </thead>
              <tbody>
                {data.rows.length === 0 ? (
                  <tr><td colSpan={4} style={{ ...s.td, textAlign: "center", color: "#94a3b8", padding: "26px 10px" }}>No cuts recorded in this range.</td></tr>
                ) : data.rows.map((r) => (
                  <tr key={r.id}>
                    <td style={{ ...s.td, whiteSpace: "nowrap", color: "#475569" }}>{fmtDate(r.productionDate)}</td>
                    <td style={s.td}>
                      {r.productTitle}
                      {!r.live && <span title="The restock row was deleted — this record is kept for history" style={s.deletedTag}>row deleted</span>}
                      {r.supplier ? <span style={{ color: "#94a3b8" }}> · {r.supplier}</span> : null}
                    </td>
                    <td style={{ ...s.td, color: "#64748b" }}>{r.productType || "—"}</td>
                    <td style={{ ...s.td, textAlign: "right", fontWeight: 800, fontVariantNumeric: "tabular-nums" }}>{r.qty}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}

function Card({ label, pieces, orders, accent }: { label: string; pieces: number; orders: number; accent: string }) {
  return (
    <div style={{ ...s.card, borderTop: `3px solid ${accent}` }}>
      <div style={s.cardLabel}>{label}</div>
      <div style={{ ...s.cardValue, color: accent }}>{pieces.toLocaleString()}</div>
      <div style={s.cardHint}>pieces · {orders} order{orders === 1 ? "" : "s"}</div>
    </div>
  );
}

const s: Record<string, React.CSSProperties> = {
  page: { padding: "4px 2px 40px" },
  cards: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(170px, 1fr))", gap: 12, marginBottom: 16 },
  card: { background: "#fff", border: "1px solid #e2e8f0", borderRadius: 12, padding: "14px 16px", boxShadow: "0 1px 2px rgba(15,23,42,0.04)" },
  cardLabel: { fontSize: 12, fontWeight: 800, color: "#64748b", textTransform: "uppercase", letterSpacing: ".04em" },
  cardValue: { fontSize: 30, fontWeight: 800, marginTop: 4, fontVariantNumeric: "tabular-nums" },
  cardHint: { fontSize: 11.5, color: "#94a3b8", marginTop: 2 },
  toolbar: { display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", background: "#fff", border: "1px solid #e2e8f0", borderRadius: 12, padding: "10px 14px", marginBottom: 14, boxShadow: "0 1px 2px rgba(15,23,42,0.04)" },
  field: { display: "inline-flex", alignItems: "center", gap: 7 },
  fieldLabel: { fontSize: 11, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", letterSpacing: ".04em" },
  dateInput: { height: 34, boxSizing: "border-box", border: "1px solid #cbd5e1", borderRadius: 8, padding: "0 10px", fontSize: 13, fontWeight: 600, color: "#0f172a", background: "#fff" },
  rangeHint: { fontSize: 12.5, color: "#64748b", fontWeight: 600 },
  split: { display: "grid", gridTemplateColumns: "minmax(320px, 1fr) minmax(360px, 1.3fr)", gap: 14, alignItems: "start" },
  panel: { background: "#fff", border: "1px solid #e2e8f0", borderRadius: 12, padding: "14px 16px", boxShadow: "0 1px 2px rgba(15,23,42,0.04)" },
  panelTitle: { margin: "0 0 12px", fontSize: 15, fontWeight: 800, color: "#0f172a" },
  empty: { padding: 24, textAlign: "center", color: "#94a3b8", fontSize: 13 },
  th: { position: "sticky", top: 0, zIndex: 1, background: "#f8fafc", borderBottom: "2px solid #e2e8f0", padding: "9px 10px", fontSize: 11, fontWeight: 800, color: "#64748b", textTransform: "uppercase", letterSpacing: ".03em", textAlign: "left", whiteSpace: "nowrap" },
  td: { padding: "8px 10px", fontSize: 13, borderBottom: "1px solid #eef2f7", verticalAlign: "middle" },
  deletedTag: { marginLeft: 6, fontSize: 10.5, fontWeight: 700, color: "#9a3412", background: "#ffedd5", borderRadius: 5, padding: "1px 6px", whiteSpace: "nowrap" },
};
