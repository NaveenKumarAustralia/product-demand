import { useEffect, useMemo, useState } from "react";

type NotificationRow = {
  id: number;
  type: string;
  status: string;
  supplierOrderId: number | null;
  customerEmail: string | null;
  subject: string;
  body: string;
  createdByUserName: string | null;
  approvedByUserName: string | null;
  approvedAt: string | null;
  sentAt: string | null;
  createdAt: string;
};

type Batch = { id: number; shop: string; productTitle: string; supplier: string; destination: string | null };
type ApiResponse = { ok?: boolean; rows?: NotificationRow[]; batches?: Batch[]; error?: string };

function formatDate(value: string | null) {
  if (!value) return "—";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "—";
  return new Intl.DateTimeFormat("en-AU", { day: "numeric", month: "short", hour: "numeric", minute: "2-digit" }).format(date);
}

export function PreorderNotificationsPanel() {
  const [rows, setRows] = useState<NotificationRow[]>([]);
  const [batches, setBatches] = useState<Batch[]>([]);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState<{ kind: "error" | "success"; text: string } | null>(null);
  const [type, setType] = useState("eta_update");
  const [supplierOrderId, setSupplierOrderId] = useState("");
  const [customerEmail, setCustomerEmail] = useState("");
  const [subject, setSubject] = useState("");
  const [body, setBody] = useState("");

  async function load() {
    setLoading(true);
    try {
      const response = await fetch("/api/preorder-notifications", { credentials: "same-origin" });
      const result = await response.json().catch(() => ({})) as ApiResponse;
      if (!response.ok || result.ok !== true) throw new Error(result.error || "Could not load notifications.");
      setRows(result.rows ?? []);
      setBatches(result.batches ?? []);
    } catch (error) {
      setNotice({ kind: "error", text: error instanceof Error ? error.message : "Could not load notifications." });
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => { void load(); }, []);

  const counts = useMemo(() => ({
    draft: rows.filter((row) => row.status === "draft").length,
    approved: rows.filter((row) => row.status === "approved").length,
    sent: rows.filter((row) => row.status === "sent").length,
    cancelled: rows.filter((row) => row.status === "cancelled").length,
  }), [rows]);

  async function post(payload: Record<string, unknown>, successText: string) {
    setBusy(true);
    setNotice(null);
    try {
      const response = await fetch("/api/preorder-notifications", {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const result = await response.json().catch(() => ({})) as ApiResponse;
      if (!response.ok || result.ok !== true) throw new Error(result.error || "Could not update notification.");
      setNotice({ kind: "success", text: successText });
      await load();
    } catch (error) {
      setNotice({ kind: "error", text: error instanceof Error ? error.message : "Could not update notification." });
    } finally {
      setBusy(false);
    }
  }

  async function createDraft() {
    await post({
      operation: "create-draft",
      type,
      supplierOrderId: supplierOrderId || null,
      customerEmail,
      subject,
      body,
    }, "Notification draft created.");
    setSubject("");
    setBody("");
    setCustomerEmail("");
  }

  return (
    <div style={s.stack}>
      {notice ? <div style={{ ...s.notice, ...(notice.kind === "error" ? s.noticeError : s.noticeSuccess) }}>{notice.text}</div> : null}
      <div style={s.warning}><strong>Approval queue only.</strong> No email is sent from this screen yet. Approved messages are held until the delivery service is connected and tested.</div>

      <div style={s.cards}>
        <Metric label="Draft" value={counts.draft} />
        <Metric label="Approved" value={counts.approved} />
        <Metric label="Sent" value={counts.sent} />
        <Metric label="Cancelled" value={counts.cancelled} />
      </div>

      <div style={s.card}>
        <div style={s.title}>Create notification draft</div>
        <div style={s.muted}>Prepare customer wording first, then another authorised staff member can review/approve it.</div>
        <div style={s.formGrid}>
          <label style={s.label}>Type
            <select value={type} onChange={(e) => setType(e.target.value)} style={s.input}>
              <option value="eta_update">Expected dispatch update</option>
              <option value="delay">Production delay</option>
              <option value="waitlist_available">Back in stock / preorder available</option>
              <option value="general">General preorder update</option>
            </select>
          </label>
          <label style={s.label}>Production batch
            <select value={supplierOrderId} onChange={(e) => setSupplierOrderId(e.target.value)} style={s.input}>
              <option value="">No specific batch</option>
              {batches.map((batch) => <option key={batch.id} value={batch.id}>#{batch.id} · {batch.productTitle} · {batch.destination || "No destination"}</option>)}
            </select>
          </label>
          <label style={s.label}>Customer email <span style={s.optional}>(optional for batch-wide draft)</span>
            <input value={customerEmail} onChange={(e) => setCustomerEmail(e.target.value)} style={s.input} placeholder="customer@example.com" />
          </label>
        </div>
        <label style={s.label}>Subject<input value={subject} onChange={(e) => setSubject(e.target.value)} style={s.input} placeholder="Update on your Karma East pre-order" /></label>
        <label style={s.label}>Message<textarea value={body} onChange={(e) => setBody(e.target.value)} style={{ ...s.input, minHeight: 110, resize: "vertical" }} placeholder="Write the customer-facing message here…" /></label>
        <div style={s.actions}><button type="button" disabled={busy || !subject.trim() || !body.trim()} style={s.primary} onClick={() => void createDraft()}>{busy ? "Saving…" : "Save draft"}</button></div>
      </div>

      <div style={s.card}>
        <div style={s.title}>Notification approval queue</div>
        <div style={s.muted}>{loading ? "Loading…" : `${rows.length} recent notifications`}</div>
        {rows.length === 0 ? <div style={s.empty}>No notification drafts yet.</div> : (
          <div style={s.queue}>
            {rows.map((row) => (
              <div key={row.id} style={s.item}>
                <div style={s.itemTop}>
                  <div><strong>{row.subject}</strong><div style={s.small}>{row.type.replace(/_/g, " ")} · {row.customerEmail || "Batch-wide / no recipient yet"} · {formatDate(row.createdAt)}</div></div>
                  <span style={{ ...s.badge, ...(row.status === "approved" ? s.approved : row.status === "sent" ? s.sent : row.status === "cancelled" ? s.cancelled : {}) }}>{row.status}</span>
                </div>
                <div style={s.message}>{row.body}</div>
                <div style={s.small}>Created by {row.createdByUserName || "Unknown"}{row.approvedByUserName ? ` · Approved by ${row.approvedByUserName} ${formatDate(row.approvedAt)}` : ""}</div>
                <div style={s.actions}>
                  {row.status === "draft" ? <button disabled={busy} style={s.primary} onClick={() => void post({ operation: "approve", id: row.id }, "Notification approved and held for delivery.")}>Approve</button> : null}
                  {row.status === "approved" ? <button disabled={busy} style={s.secondary} onClick={() => void post({ operation: "return-to-draft", id: row.id }, "Notification returned to draft.")}>Return to draft</button> : null}
                  {row.status === "draft" || row.status === "approved" ? <button disabled={busy} style={s.danger} onClick={() => void post({ operation: "cancel", id: row.id }, "Notification cancelled.")}>Cancel</button> : null}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function Metric({ label, value }: { label: string; value: number }) {
  return <div style={s.metric}><div style={s.metricLabel}>{label}</div><div style={s.metricValue}>{value}</div></div>;
}

const s: Record<string, React.CSSProperties> = {
  stack: { display: "flex", flexDirection: "column", gap: 14 },
  warning: { padding: "10px 12px", border: "1px solid #fde68a", background: "#fffbeb", color: "#92400e", borderRadius: 9, fontSize: 12 },
  cards: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 10 },
  metric: { background: "white", border: "1px solid #e2e8f0", borderRadius: 12, padding: 14 },
  metricLabel: { fontSize: 11, textTransform: "uppercase", fontWeight: 800, color: "#64748b" },
  metricValue: { fontSize: 25, fontWeight: 800, color: "#0f172a", marginTop: 3 },
  card: { background: "white", border: "1px solid #e2e8f0", borderRadius: 14, padding: 16 },
  title: { fontSize: 16, fontWeight: 800, color: "#0f172a" },
  muted: { fontSize: 12, color: "#64748b", marginTop: 4, marginBottom: 12 },
  formGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(210px, 1fr))", gap: 10 },
  label: { display: "flex", flexDirection: "column", gap: 5, fontSize: 11, fontWeight: 700, color: "#475569", marginBottom: 10 },
  optional: { color: "#94a3b8", fontWeight: 500 },
  input: { border: "1px solid #cbd5e1", borderRadius: 8, padding: "8px 9px", fontSize: 12, background: "white", color: "#0f172a" },
  actions: { display: "flex", gap: 7, justifyContent: "flex-end", flexWrap: "wrap", marginTop: 9 },
  primary: { border: "1px solid #0f766e", background: "#0f766e", color: "white", borderRadius: 8, padding: "7px 10px", fontSize: 11, fontWeight: 800, cursor: "pointer" },
  secondary: { border: "1px solid #cbd5e1", background: "white", color: "#334155", borderRadius: 8, padding: "7px 10px", fontSize: 11, fontWeight: 800, cursor: "pointer" },
  danger: { border: "1px solid #fecaca", background: "#fff7f7", color: "#991b1b", borderRadius: 8, padding: "7px 10px", fontSize: 11, fontWeight: 800, cursor: "pointer" },
  empty: { padding: 30, textAlign: "center", color: "#64748b" },
  queue: { display: "flex", flexDirection: "column", gap: 9 },
  item: { border: "1px solid #e2e8f0", borderRadius: 10, padding: 12 },
  itemTop: { display: "flex", justifyContent: "space-between", gap: 12, alignItems: "flex-start" },
  small: { fontSize: 10, color: "#94a3b8", marginTop: 4 },
  message: { marginTop: 10, padding: 10, borderRadius: 8, background: "#f8fafc", color: "#475569", whiteSpace: "pre-wrap", fontSize: 12, lineHeight: 1.5 },
  badge: { padding: "4px 7px", borderRadius: 999, background: "#fef3c7", color: "#92400e", fontSize: 10, fontWeight: 800, textTransform: "capitalize" },
  approved: { background: "#dbeafe", color: "#1d4ed8" },
  sent: { background: "#dcfce7", color: "#166534" },
  cancelled: { background: "#f1f5f9", color: "#64748b" },
  notice: { padding: "9px 11px", borderRadius: 8, fontSize: 12, fontWeight: 700 },
  noticeError: { background: "#fef2f2", border: "1px solid #fecaca", color: "#991b1b" },
  noticeSuccess: { background: "#f0fdf4", border: "1px solid #bbf7d0", color: "#166534" },
};
