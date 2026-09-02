import { useState } from "react";
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

type PermissionKey = keyof PreorderDashboardData["configuration"]["permissions"];

const PERMISSIONS: Array<{ key: PermissionKey; label: string; help: string }> = [
  { key: "managePreorderUserIds", label: "Enable / pause preorder", help: "Can make eligible production batches available for customer preorder." },
  { key: "manageEtaUserIds", label: "Expected dispatch date", help: "Can change the customer-facing expected dispatch date." },
  { key: "manageSafetyBufferUserIds", label: "Safety buffer", help: "Can change percentage or fixed preorder safety buffers." },
  { key: "sendNotificationUserIds", label: "Customer notifications", help: "Can send or approve preorder customer updates when notification tools are enabled." },
  { key: "viewReportsUserIds", label: "Reports", help: "Can view preorder reporting and performance data." },
];

export function PreorderSettingsPanel({ configuration }: { configuration: PreorderDashboardData["configuration"] }) {
  const [au, setAu] = useState(configuration.locations.AU ?? "");
  const [usa, setUsa] = useState(configuration.locations.USA ?? "");
  const [permissions, setPermissions] = useState(configuration.permissions);
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState<{ kind: "error" | "success"; text: string } | null>(null);

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

  function togglePermission(key: PermissionKey, userId: string) {
    setPermissions((current) => {
      const ids = current[key];
      return {
        ...current,
        [key]: ids.includes(userId) ? ids.filter((id) => id !== userId) : [...ids, userId],
      };
    });
  }

  return (
    <div style={s.stack}>
      {notice ? <div style={{ ...s.notice, ...(notice.kind === "error" ? s.noticeError : s.noticeSuccess) }}>{notice.text}</div> : null}

      <div style={s.card}>
        <div style={s.title}>Shopify locations</div>
        <div style={s.muted}>Keep AU and USA preorder capacity completely separate. Enter either the numeric Shopify location ID or the full gid.</div>
        <div style={s.twoCols}>
          <label style={s.label}>Australia location<input value={au} onChange={(e) => setAu(e.target.value)} style={s.input} placeholder="gid://shopify/Location/..." /></label>
          <label style={s.label}>USA location<input value={usa} onChange={(e) => setUsa(e.target.value)} style={s.input} placeholder="gid://shopify/Location/..." /></label>
        </div>
        <div style={s.actions}><button type="button" disabled={busy} style={s.primary} onClick={() => post({ operation: "update-locations", AU: au, USA: usa }, "Shopify preorder locations saved.")}>Save locations</button></div>
      </div>

      <div style={s.card}>
        <div style={s.title}>Staff action permissions</div>
        <div style={s.muted}>The Pre-orders menu itself is still controlled in Production Portal → Settings → Users → Page access. These permissions only control sensitive actions inside Pre-orders. Portal admins always retain access.</div>
        <div style={s.permissionTable}>
          <div style={s.permissionHeader}><div>Staff</div>{PERMISSIONS.map((permission) => <div key={permission.key}>{permission.label}</div>)}</div>
          {configuration.users.map((user) => (
            <div key={user.id} style={s.permissionRow}>
              <div><strong>{user.name}</strong>{user.admin ? <div style={s.small}>Portal admin</div> : null}</div>
              {PERMISSIONS.map((permission) => (
                <div key={permission.key} style={s.checkboxCell} title={permission.help}>
                  <input
                    type="checkbox"
                    checked={user.admin || permissions[permission.key].includes(user.id)}
                    disabled={user.admin || busy}
                    onChange={() => togglePermission(permission.key, user.id)}
                  />
                </div>
              ))}
            </div>
          ))}
        </div>
        <div style={s.actions}><button type="button" disabled={busy} style={s.primary} onClick={() => post({ operation: "update-permissions", permissions }, "Preorder staff permissions saved.")}>Save permissions</button></div>
      </div>
    </div>
  );
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
};
