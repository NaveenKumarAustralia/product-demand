import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import { Form, redirect, useLoaderData, useNavigation } from "react-router";
import prisma from "../db.server";
import { getSession, commitSession, destroySession } from "../portal.session.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const session = await getSession(request.headers.get("Cookie"));
  const supplierId = session.get("supplierId");
  if (!supplierId) throw redirect("/portal/login");

  const account = await prisma.supplierAccount.findUnique({
    where: { id: Number(supplierId) },
  });
  if (!account || !account.active) {
    throw redirect("/portal/login", {
      headers: { "Set-Cookie": await destroySession(session) },
    });
  }

  const orders = await prisma.supplierOrder.findMany({
    where: { shop: account.shop, supplier: account.name, status: "open" },
    include: { lines: { orderBy: { id: "asc" } } },
    orderBy: { createdAt: "desc" },
  });

  return { supplierName: account.name, orders };
};

export const action = async ({ request }: ActionFunctionArgs) => {
  const session = await getSession(request.headers.get("Cookie"));
  const supplierId = session.get("supplierId");
  if (!supplierId) throw redirect("/portal/login");

  const form = await request.formData();
  const intent = form.get("intent");

  if (intent === "logout") {
    throw redirect("/portal/login", {
      headers: { "Set-Cookie": await destroySession(session) },
    });
  }

  if (intent === "update_status") {
    const orderId = Number(form.get("orderId"));
    const newStatus = String(form.get("newStatus"));
    const supplierNotes = String(form.get("supplierNotes") ?? "").trim();

    const account = await prisma.supplierAccount.findUnique({
      where: { id: Number(supplierId) },
    });
    if (!account) throw redirect("/portal/login");

    await prisma.supplierOrder.updateMany({
      where: { id: orderId, shop: account.shop, supplier: account.name },
      data: {
        supplierStatus: newStatus,
        supplierNotes: supplierNotes || null,
      },
    });
  }

  return null;
};

type Order = Awaited<ReturnType<typeof loader>>["orders"][number];

const STATUS_LABELS: Record<string, { label: string; color: string; bg: string }> = {
  pending:   { label: "Pending",   color: "#92400e", bg: "#fef3c7" },
  confirmed: { label: "Confirmed", color: "#1e40af", bg: "#dbeafe" },
  shipped:   { label: "Shipped",   color: "#065f46", bg: "#d1fae5" },
};

export default function PortalDashboard() {
  const { supplierName, orders } = useLoaderData<typeof loader>();
  const navigation = useNavigation();

  return (
    <div style={styles.page}>
      <header style={styles.header}>
        <div style={styles.headerInner}>
          <div>
            <span style={styles.logo}>Supplier Portal</span>
            <span style={styles.supplierBadge}>{supplierName}</span>
          </div>
          <Form method="post">
            <input type="hidden" name="intent" value="logout" />
            <button type="submit" style={styles.logoutBtn}>Sign out</button>
          </Form>
        </div>
      </header>

      <main style={styles.main}>
        <h2 style={styles.heading}>Your Orders</h2>

        {orders.length === 0 ? (
          <div style={styles.empty}>No open orders at the moment.</div>
        ) : (
          <div style={styles.grid}>
            {orders.map((order) => (
              <OrderCard key={order.id} order={order} submitting={navigation.state === "submitting"} />
            ))}
          </div>
        )}
      </main>
    </div>
  );
}

function OrderCard({ order, submitting }: { order: Order; submitting: boolean }) {
  const statusInfo = STATUS_LABELS[order.supplierStatus] ?? STATUS_LABELS.pending;
  const nextStatus = order.supplierStatus === "pending" ? "confirmed"
    : order.supplierStatus === "confirmed" ? "shipped"
    : null;
  const nextLabel = nextStatus === "confirmed" ? "Confirm Order"
    : nextStatus === "shipped" ? "Mark as Shipped"
    : null;

  return (
    <div style={styles.card}>
      <div style={styles.cardHeader}>
        <div>
          <div style={styles.productTitle}>{order.productTitle}</div>
          <div style={styles.meta}>
            {order.totalQty} units
            {order.eta && ` · ETA ${new Date(order.eta).toLocaleDateString("en-AU", { day: "numeric", month: "short", year: "numeric" })}`}
            {" · "}Placed {new Date(order.createdAt).toLocaleDateString("en-AU", { day: "numeric", month: "short" })}
          </div>
        </div>
        <span style={{ ...styles.badge, color: statusInfo.color, background: statusInfo.bg }}>
          {statusInfo.label}
        </span>
      </div>

      <table style={styles.table}>
        <thead>
          <tr>
            <th style={styles.th}>Variant</th>
            <th style={{ ...styles.th, textAlign: "right" }}>Qty</th>
          </tr>
        </thead>
        <tbody>
          {order.lines.map((line) => (
            <tr key={line.id}>
              <td style={styles.td}>
                {line.variantTitle}
                {line.sku && <span style={styles.sku}> · {line.sku}</span>}
              </td>
              <td style={{ ...styles.td, textAlign: "right", fontWeight: 600 }}>{line.qtyOrdered}</td>
            </tr>
          ))}
        </tbody>
      </table>

      {order.notes && (
        <div style={styles.notes}><strong>Notes:</strong> {order.notes}</div>
      )}

      {nextLabel && (
        <Form method="post" style={{ marginTop: 16 }}>
          <input type="hidden" name="intent" value="update_status" />
          <input type="hidden" name="orderId" value={order.id} />
          <input type="hidden" name="newStatus" value={nextStatus!} />
          <div style={styles.field}>
            <label style={styles.label}>Add a note (optional)</label>
            <input
              name="supplierNotes"
              defaultValue={order.supplierNotes ?? ""}
              style={styles.input}
              placeholder="e.g. estimated dispatch date, tracking info…"
            />
          </div>
          <button type="submit" disabled={submitting} style={styles.actionBtn}>
            {nextLabel}
          </button>
        </Form>
      )}

      {order.supplierStatus === "shipped" && (
        <div style={styles.shipped}>
          ✓ Marked as shipped{order.supplierNotes ? ` — ${order.supplierNotes}` : ""}
        </div>
      )}
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  page: {
    minHeight: "100vh",
    background: "#f4f6f8",
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
  },
  header: {
    background: "#fff",
    borderBottom: "1px solid #e5e7eb",
    position: "sticky",
    top: 0,
    zIndex: 10,
  },
  headerInner: {
    maxWidth: 900,
    margin: "0 auto",
    padding: "16px 24px",
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
  },
  logo: { fontSize: 18, fontWeight: 700, color: "#1a1a1a" },
  supplierBadge: {
    marginLeft: 12,
    background: "#ecfdf5",
    color: "#065f46",
    padding: "3px 10px",
    borderRadius: 20,
    fontSize: 13,
    fontWeight: 600,
  },
  logoutBtn: {
    background: "none",
    border: "1px solid #d1d5db",
    borderRadius: 8,
    padding: "6px 14px",
    fontSize: 13,
    cursor: "pointer",
    color: "#374151",
  },
  main: { maxWidth: 900, margin: "0 auto", padding: "32px 24px" },
  heading: { margin: "0 0 24px", fontSize: 22, fontWeight: 700, color: "#111827" },
  empty: {
    background: "#fff",
    borderRadius: 12,
    padding: 40,
    textAlign: "center",
    color: "#6b7280",
    fontSize: 15,
  },
  grid: { display: "flex", flexDirection: "column", gap: 20 },
  card: {
    background: "#fff",
    borderRadius: 12,
    padding: 24,
    boxShadow: "0 1px 4px rgba(0,0,0,0.06)",
  },
  cardHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    marginBottom: 16,
    gap: 12,
  },
  productTitle: { fontSize: 17, fontWeight: 700, color: "#111827" },
  meta: { fontSize: 13, color: "#6b7280", marginTop: 4 },
  badge: {
    padding: "4px 12px",
    borderRadius: 20,
    fontSize: 12,
    fontWeight: 600,
    whiteSpace: "nowrap",
  },
  table: { width: "100%", borderCollapse: "collapse", marginBottom: 8 },
  th: {
    textAlign: "left",
    fontSize: 12,
    fontWeight: 600,
    color: "#6b7280",
    textTransform: "uppercase",
    letterSpacing: "0.05em",
    padding: "6px 0",
    borderBottom: "1px solid #f3f4f6",
  },
  td: { padding: "8px 0", fontSize: 14, color: "#374151", borderBottom: "1px solid #f9fafb" },
  sku: { color: "#9ca3af" },
  notes: {
    marginTop: 12,
    fontSize: 13,
    color: "#6b7280",
    background: "#f9fafb",
    padding: "8px 12px",
    borderRadius: 6,
  },
  field: { display: "flex", flexDirection: "column", gap: 6, marginBottom: 10 },
  label: { fontSize: 13, fontWeight: 500, color: "#374151" },
  input: {
    padding: "9px 12px",
    borderRadius: 8,
    border: "1px solid #d1d5db",
    fontSize: 14,
    outline: "none",
    width: "100%",
    boxSizing: "border-box",
  },
  actionBtn: {
    padding: "10px 20px",
    background: "#008060",
    color: "#fff",
    border: "none",
    borderRadius: 8,
    fontSize: 14,
    fontWeight: 600,
    cursor: "pointer",
  },
  shipped: {
    marginTop: 16,
    fontSize: 14,
    color: "#065f46",
    background: "#ecfdf5",
    padding: "10px 14px",
    borderRadius: 8,
  },
};
