import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import { useFetcher, useLoaderData } from "react-router";
import prisma from "../db.server";

export const loader = async ({}: LoaderFunctionArgs) => {
  const orders = await prisma.supplierOrder.findMany({
    where: { status: "open" },
    include: { lines: { orderBy: { id: "asc" } } },
    orderBy: { createdAt: "desc" },
  });

  // Collect all unique size names across all orders, sorted logically
  const sizeOrder = ["XS","S","S/M","M","M/L","L","L/XL","XL","2XL","3XL","4XL","ONE SIZE"];
  const allSizes = [...new Set(orders.flatMap((o) => o.lines.map((l) => l.variantTitle)))];
  allSizes.sort((a, b) => {
    const ai = sizeOrder.indexOf(a.toUpperCase());
    const bi = sizeOrder.indexOf(b.toUpperCase());
    if (ai === -1 && bi === -1) return a.localeCompare(b);
    if (ai === -1) return 1;
    if (bi === -1) return -1;
    return ai - bi;
  });

  return { orders, sizes: allSizes };
};

export const action = async ({ request }: ActionFunctionArgs) => {
  const form = await request.formData();
  const intent = String(form.get("intent"));
  const orderId = Number(form.get("orderId"));

  const updates: Record<string, unknown> = {};

  if (intent === "update_status")        updates.supplierStatus = form.get("value");
  if (intent === "update_priority")      updates.priority = form.get("value");
  if (intent === "update_factory_notes") updates.factoryNotes = form.get("value");
  if (intent === "update_eta") {
    const raw = String(form.get("value") ?? "");
    updates.eta = raw ? new Date(raw) : null;
  }

  if (Object.keys(updates).length) {
    await prisma.supplierOrder.update({ where: { id: orderId }, data: updates });
  }
  return null;
};

// ─── Constants ───────────────────────────────────────────────────────────────

const STATUS_OPTIONS = [
  { value: "on_order",       label: "On Order" },
  { value: "on_production",  label: "On Production" },
  { value: "in_shipment",    label: "In Shipment" },
  { value: "arrived",        label: "Arrived" },
  { value: "arrived_loaded", label: "Arrived and Loaded" },
  { value: "cancelled",      label: "Cancelled" },
  { value: "ready_to_send",  label: "Ready To Send" },
];

const STATUS_COLORS: Record<string, string> = {
  on_order:       "#fef9c3",
  on_production:  "#dbeafe",
  in_shipment:    "#dcfce7",
  arrived:        "#bbf7d0",
  arrived_loaded: "#4ade80",
  cancelled:      "#fee2e2",
  ready_to_send:  "#ede9fe",
};

const PRIORITY_OPTIONS = [
  { value: "low",       label: "LOW",       bg: "#3b82f6", color: "#fff" },
  { value: "high",      label: "HIGH",      bg: "#7c3aed", color: "#fff" },
  { value: "urgent",    label: "URGENT",    bg: "#dc2626", color: "#fff" },
  { value: "cancelled", label: "Cancelled", bg: "#d97706", color: "#fff" },
];

// ─── Main component ───────────────────────────────────────────────────────────

type Order = Awaited<ReturnType<typeof loader>>["orders"][number];

export default function PortalDashboard() {
  const { orders, sizes } = useLoaderData<typeof loader>();

  return (
    <div style={s.page}>
      <header style={s.header}>
        <div style={s.headerInner}>
          <span style={s.logo}>Supplier Portal</span>
          <span style={s.count}>{orders.length} open order{orders.length !== 1 ? "s" : ""}</span>
        </div>
      </header>

      <main style={s.main}>
        {orders.length === 0 ? (
          <div style={s.empty}>No open orders at the moment.</div>
        ) : (
          <div style={s.tableWrap}>
            <table style={s.table}>
              <thead>
                <tr style={s.headerRow}>
                  <Th>Factory Notes</Th>
                  <Th>Picture</Th>
                  <Th>Name</Th>
                  <Th>SKU</Th>
                  {sizes.map((sz) => <Th key={sz} center>{sz}</Th>)}
                  <Th center>Total</Th>
                  <Th>Status</Th>
                  <Th>Notes</Th>
                  <Th>Priority</Th>
                  <Th>ETA</Th>
                </tr>
              </thead>
              <tbody>
                {orders.map((order) => (
                  <OrderRow key={order.id} order={order} sizes={sizes} />
                ))}
              </tbody>
            </table>
          </div>
        )}
      </main>
    </div>
  );
}

// ─── Row ─────────────────────────────────────────────────────────────────────

function OrderRow({ order, sizes }: { order: Order; sizes: string[] }) {
  const qtyBySize = Object.fromEntries(order.lines.map((l) => [l.variantTitle, l.qtyOrdered]));
  const allSkus = order.lines.map((l) => l.sku).filter(Boolean).join("\n");
  const etaValue = order.eta ? new Date(order.eta).toISOString().slice(0, 10) : "";

  return (
    <tr style={s.row}>
      {/* Factory notes */}
      <Td><NotesCell orderId={order.id} field="factory_notes" value={order.factoryNotes ?? ""} /></Td>

      {/* Picture */}
      <Td center>
        {order.productImageUrl
          ? <img src={order.productImageUrl} alt="" style={s.thumb} />
          : <div style={s.noImg}>—</div>}
      </Td>

      {/* Name */}
      <Td><span style={s.productName}>{order.productTitle}</span></Td>

      {/* SKU */}
      <Td><span style={s.sku}>{allSkus || "—"}</span></Td>

      {/* Size columns */}
      {sizes.map((sz) => (
        <Td key={sz} center>
          <span style={(qtyBySize[sz] ?? 0) > 0 ? s.qty : s.qtyZero}>
            {qtyBySize[sz] ?? 0}
          </span>
        </Td>
      ))}

      {/* Total */}
      <Td center><span style={s.total}>{order.totalQty}</span></Td>

      {/* Status */}
      <Td><StatusCell orderId={order.id} value={order.supplierStatus} /></Td>

      {/* Notes (from order) */}
      <Td><span style={s.noteText}>{order.notes || "—"}</span></Td>

      {/* Priority */}
      <Td><PriorityCell orderId={order.id} value={order.priority ?? ""} /></Td>

      {/* ETA */}
      <Td><EtaCell orderId={order.id} value={etaValue} /></Td>
    </tr>
  );
}

// ─── Editable cells ───────────────────────────────────────────────────────────

function StatusCell({ orderId, value }: { orderId: number; value: string }) {
  const fetcher = useFetcher();
  const current = fetcher.formData ? String(fetcher.formData.get("value")) : value;
  const bg = STATUS_COLORS[current] ?? "#f3f4f6";

  return (
    <fetcher.Form method="post">
      <input type="hidden" name="intent" value="update_status" />
      <input type="hidden" name="orderId" value={orderId} />
      <select
        name="value"
        value={current}
        onChange={(e) => fetcher.submit(e.currentTarget.form!)}
        style={{ ...s.select, background: bg }}
      >
        {STATUS_OPTIONS.map((o) => (
          <option key={o.value} value={o.value}>{o.label}</option>
        ))}
      </select>
    </fetcher.Form>
  );
}

function PriorityCell({ orderId, value }: { orderId: number; value: string }) {
  const fetcher = useFetcher();
  const current = fetcher.formData ? String(fetcher.formData.get("value")) : value;
  const opt = PRIORITY_OPTIONS.find((o) => o.value === current);

  return (
    <fetcher.Form method="post">
      <input type="hidden" name="intent" value="update_priority" />
      <input type="hidden" name="orderId" value={orderId} />
      <select
        name="value"
        value={current}
        onChange={(e) => fetcher.submit(e.currentTarget.form!)}
        style={{
          ...s.select,
          background: opt?.bg ?? "#f3f4f6",
          color: opt?.color ?? "#374151",
          fontWeight: 700,
        }}
      >
        <option value="">— Priority —</option>
        {PRIORITY_OPTIONS.map((o) => (
          <option key={o.value} value={o.value}>{o.label}</option>
        ))}
      </select>
    </fetcher.Form>
  );
}

function NotesCell({ orderId, field, value }: { orderId: number; field: string; value: string }) {
  const fetcher = useFetcher();
  return (
    <fetcher.Form method="post">
      <input type="hidden" name="intent" value={`update_${field}`} />
      <input type="hidden" name="orderId" value={orderId} />
      <textarea
        name="value"
        defaultValue={value}
        onBlur={(e) => fetcher.submit(e.currentTarget.form!)}
        rows={2}
        style={s.textarea}
        placeholder="Add note…"
      />
    </fetcher.Form>
  );
}

function EtaCell({ orderId, value }: { orderId: number; value: string }) {
  const fetcher = useFetcher();
  return (
    <fetcher.Form method="post">
      <input type="hidden" name="intent" value="update_eta" />
      <input type="hidden" name="orderId" value={orderId} />
      <input
        type="date"
        name="value"
        defaultValue={value}
        onBlur={(e) => fetcher.submit(e.currentTarget.form!)}
        style={s.dateInput}
      />
    </fetcher.Form>
  );
}

// ─── Table helpers ────────────────────────────────────────────────────────────

function Th({ children, center }: { children: React.ReactNode; center?: boolean }) {
  return <th style={{ ...s.th, textAlign: center ? "center" : "left" }}>{children}</th>;
}
function Td({ children, center }: { children: React.ReactNode; center?: boolean }) {
  return <td style={{ ...s.td, textAlign: center ? "center" : "left" }}>{children}</td>;
}

// ─── Styles ───────────────────────────────────────────────────────────────────

const s: Record<string, React.CSSProperties> = {
  page: { minHeight: "100vh", background: "#f3f4f6", fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif" },
  header: { background: "#fff", borderBottom: "1px solid #e5e7eb", position: "sticky", top: 0, zIndex: 10 },
  headerInner: { maxWidth: "100%", padding: "14px 24px", display: "flex", alignItems: "center", gap: 16 },
  logo: { fontSize: 17, fontWeight: 700, color: "#1a1a1a" },
  count: { fontSize: 13, color: "#6b7280" },
  main: { padding: "24px 16px" },
  empty: { background: "#fff", borderRadius: 12, padding: 40, textAlign: "center", color: "#6b7280" },
  tableWrap: {
    overflowX: "auto",
    background: "#fff",
    border: "1px solid #cbd5e1",
    boxShadow: "0 1px 2px rgba(15,23,42,0.08)",
  },
  table: {
    width: "100%",
    borderCollapse: "collapse",
    borderSpacing: 0,
    fontSize: 13,
    minWidth: 900,
    tableLayout: "auto",
  },
  headerRow: { background: "#eef2f7" },
  th: {
    padding: "8px 10px",
    fontWeight: 700,
    fontSize: 11,
    textTransform: "uppercase",
    letterSpacing: "0.05em",
    color: "#4b5563",
    border: "1px solid #cbd5e1",
    whiteSpace: "nowrap",
    background: "#eef2f7",
  },
  row: { background: "#fff" },
  td: {
    padding: "8px 10px",
    verticalAlign: "middle",
    color: "#374151",
    border: "1px solid #d1d5db",
    background: "#fff",
  },
  thumb: { width: 48, height: 64, objectFit: "cover", borderRadius: 2, display: "block", margin: "0 auto" },
  noImg: { color: "#d1d5db", textAlign: "center" },
  productName: { fontWeight: 600, color: "#111827", whiteSpace: "nowrap" },
  sku: { fontFamily: "monospace", fontSize: 11, color: "#6b7280", whiteSpace: "pre-line" },
  qty: { fontWeight: 700, color: "#111827" },
  qtyZero: { color: "#d1d5db" },
  total: { fontWeight: 700, fontSize: 14, color: "#111827" },
  noteText: { fontSize: 12, color: "#6b7280", maxWidth: 160, display: "block", whiteSpace: "pre-wrap" },
  select: {
    border: "1px solid #b6c0cc",
    borderRadius: 3,
    padding: "5px 8px",
    fontSize: 12,
    fontWeight: 600,
    cursor: "pointer",
    outline: "none",
    width: "100%",
  },
  textarea: {
    border: "1px solid #cbd5e1",
    borderRadius: 3,
    padding: "6px 8px",
    fontSize: 12,
    resize: "vertical",
    width: 140,
    fontFamily: "inherit",
    outline: "none",
    color: "#374151",
  },
  dateInput: {
    border: "1px solid #b6c0cc",
    borderRadius: 3,
    padding: "5px 8px",
    fontSize: 12,
    outline: "none",
    color: "#374151",
  },
};
