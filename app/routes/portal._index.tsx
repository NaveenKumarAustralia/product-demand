import { useState } from "react";
import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import { useFetcher, useLoaderData } from "react-router";
import prisma from "../db.server";

export const loader = async ({}: LoaderFunctionArgs) => {
  const [orders, columnWidthsSetting] = await Promise.all([
    prisma.supplierOrder.findMany({
      where: { status: "open" },
      include: { lines: { orderBy: { id: "asc" } } },
      orderBy: { createdAt: "desc" },
    }),
    prisma.portalSetting.findUnique({
      where: { key: COLUMN_WIDTHS_KEY },
      select: { value: true },
    }),
  ]);

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

  return {
    orders,
    sizes: allSizes,
    columnWidths: normalizeColumnWidths(columnWidthsSetting?.value),
  };
};

export const action = async ({ request }: ActionFunctionArgs) => {
  const form = await request.formData();
  const intent = String(form.get("intent"));
  const orderId = Number(form.get("orderId"));

  const updates: Record<string, unknown> = {};

  if (intent === "delete_order") {
    await prisma.supplierOrder.delete({ where: { id: orderId } });
    return null;
  }

  if (intent === "update_column_widths") {
    let columnWidths: Record<string, number>;
    try {
      columnWidths = normalizeColumnWidths(JSON.parse(String(form.get("value") ?? "{}")));
    } catch {
      return null;
    }

    await prisma.portalSetting.upsert({
      where: { key: COLUMN_WIDTHS_KEY },
      create: { key: COLUMN_WIDTHS_KEY, value: columnWidths },
      update: { value: columnWidths },
    });
    return null;
  }

  if (intent === "update_status")        updates.supplierStatus = form.get("value");
  if (intent === "update_priority")      updates.priority = form.get("value");
  if (intent === "update_factory_notes") updates.factoryNotes = form.get("value");
  if (intent === "update_notes")         updates.notes = form.get("value");
  if (intent === "update_eta") {
    const raw = String(form.get("value") ?? "");
    updates.eta = raw ? new Date(raw) : null;
  }

  if (intent === "update_qty") {
    const size = String(form.get("size") ?? "");
    const qtyOrdered = Math.max(0, Number(form.get("value") ?? 0) || 0);

    await prisma.$transaction(async (tx) => {
      const lines = await tx.orderLine.findMany({
        where: { orderId, variantTitle: size },
        orderBy: { id: "asc" },
        select: { id: true },
      });

      if (!lines.length) return;

      await tx.orderLine.update({
        where: { id: lines[0].id },
        data: { qtyOrdered },
      });

      if (lines.length > 1) {
        await tx.orderLine.updateMany({
          where: { id: { in: lines.slice(1).map((line) => line.id) } },
          data: { qtyOrdered: 0 },
        });
      }

      const allLines = await tx.orderLine.findMany({
        where: { orderId },
        select: { qtyOrdered: true },
      });
      await tx.supplierOrder.update({
        where: { id: orderId },
        data: { totalQty: allLines.reduce((sum, line) => sum + line.qtyOrdered, 0) },
      });
    });

    return null;
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

const COLUMN_WIDTHS_KEY = "supplier-portal-column-widths-v1";
const MIN_COLUMN_WIDTH = 52;
const DEFAULT_COLUMN_WIDTHS: Record<string, number> = {
  factoryNotes: 190,
  orderDate: 92,
  picture: 88,
  name: 260,
  sku: 115,
  total: 80,
  status: 210,
  notes: 150,
  priority: 160,
  eta: 145,
  delete: 82,
};

type ColumnDef = { id: string; label: string; center?: boolean };

function sizeColumnId(size: string) {
  return `size:${size}`;
}

function defaultColumnWidth(columnId: string) {
  return columnId.startsWith("size:") ? 58 : DEFAULT_COLUMN_WIDTHS[columnId] ?? 110;
}

function normalizeColumnWidths(value: unknown): Record<string, number> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};

  return Object.fromEntries(
    Object.entries(value)
      .map(([key, width]) => [key, Math.max(MIN_COLUMN_WIDTH, Number(width) || 0)] as const)
      .filter(([, width]) => width >= MIN_COLUMN_WIDTH),
  );
}

// ─── Main component ───────────────────────────────────────────────────────────

type Order = Awaited<ReturnType<typeof loader>>["orders"][number];

export default function PortalDashboard() {
  const { orders, sizes, columnWidths: savedColumnWidths } = useLoaderData<typeof loader>();
  const columnWidthsFetcher = useFetcher();
  const [columnWidths, setColumnWidths] = useState<Record<string, number>>(savedColumnWidths);
  const columns: ColumnDef[] = [
    { id: "factoryNotes", label: "Factory Notes" },
    { id: "orderDate", label: "Order Date" },
    { id: "picture", label: "Picture" },
    { id: "name", label: "Name" },
    { id: "sku", label: "SKU" },
    ...sizes.map((size) => ({ id: sizeColumnId(size), label: size, center: true })),
    { id: "total", label: "Total", center: true },
    { id: "status", label: "Status" },
    { id: "notes", label: "Notes" },
    { id: "priority", label: "Priority" },
    { id: "eta", label: "ETA" },
    { id: "delete", label: "Delete", center: true },
  ];

  const widthFor = (columnId: string) => columnWidths[columnId] ?? defaultColumnWidth(columnId);
  const tableWidth = columns.reduce((sum, column) => sum + widthFor(column.id), 0);

  const startResize = (columnId: string, event: React.MouseEvent<HTMLSpanElement>) => {
    event.preventDefault();
    const startX = event.clientX;
    const startWidth = widthFor(columnId);
    let nextColumnWidths = columnWidths;
    const previousCursor = document.body.style.cursor;
    const previousUserSelect = document.body.style.userSelect;

    document.body.style.cursor = "col-resize";
    document.body.style.userSelect = "none";

    const handleMove = (moveEvent: MouseEvent) => {
      const nextWidth = Math.max(MIN_COLUMN_WIDTH, startWidth + moveEvent.clientX - startX);
      nextColumnWidths = { ...nextColumnWidths, [columnId]: nextWidth };
      setColumnWidths(nextColumnWidths);
    };
    const handleUp = () => {
      document.body.style.cursor = previousCursor;
      document.body.style.userSelect = previousUserSelect;
      document.removeEventListener("mousemove", handleMove);
      document.removeEventListener("mouseup", handleUp);
      const formData = new FormData();
      formData.set("intent", "update_column_widths");
      formData.set("value", JSON.stringify(nextColumnWidths));
      columnWidthsFetcher.submit(formData, { method: "post" });
    };

    document.addEventListener("mousemove", handleMove);
    document.addEventListener("mouseup", handleUp);
  };

  const handleGridKeyDown = (event: React.KeyboardEvent<HTMLTableElement>) => {
    if (!["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"].includes(event.key)) return;

    const currentCell = (event.target as HTMLElement).closest<HTMLElement>("[data-grid-row][data-grid-col]");
    if (!currentCell) return;

    const row = Number(currentCell.dataset.gridRow);
    const col = Number(currentCell.dataset.gridCol);
    const next = {
      ArrowUp: [row - 1, col],
      ArrowDown: [row + 1, col],
      ArrowLeft: [row, col - 1],
      ArrowRight: [row, col + 1],
    }[event.key]!;
    const [nextRow, nextCol] = next;
    const nextCell = event.currentTarget.querySelector<HTMLElement>(
      `[data-grid-row="${nextRow}"][data-grid-col="${nextCol}"]`,
    );

    if (!nextCell) return;

    event.preventDefault();
    const focusTarget = nextCell.querySelector<HTMLElement>(
      "input, select, textarea, button, [tabindex]:not([tabindex='-1'])",
    ) ?? nextCell;
    focusTarget.focus();

    if (focusTarget instanceof HTMLInputElement) {
      focusTarget.select();
    }
  };

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
            <table style={{ ...s.table, width: tableWidth }} onKeyDown={handleGridKeyDown}>
              <colgroup>
                {columns.map((column) => (
                  <col key={column.id} style={{ width: widthFor(column.id) }} />
                ))}
              </colgroup>
              <thead>
                <tr style={s.headerRow}>
                  {columns.map((column) => (
                    <Th
                      key={column.id}
                      center={column.center}
                      onResizeStart={(event) => startResize(column.id, event)}
                    >
                      {column.label}
                    </Th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {orders.map((order, rowIndex) => (
                  <OrderRow key={order.id} order={order} rowIndex={rowIndex} sizes={sizes} />
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

function OrderRow({ order, rowIndex, sizes }: { order: Order; rowIndex: number; sizes: string[] }) {
  const qtyBySize = order.lines.reduce<Record<string, number>>((acc, line) => {
    acc[line.variantTitle] = (acc[line.variantTitle] ?? 0) + line.qtyOrdered;
    return acc;
  }, {});
  const allSkus = order.lines.map((l) => l.sku).filter(Boolean).join("\n");
  const etaValue = order.eta ? new Date(order.eta).toISOString().slice(0, 10) : "";
  const orderDate = new Date(order.createdAt).toLocaleDateString("en-AU", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
  });
  const totalCol = 5 + sizes.length;
  const statusCol = totalCol + 1;
  const notesCol = totalCol + 2;
  const priorityCol = totalCol + 3;
  const etaCol = totalCol + 4;
  const deleteCol = totalCol + 5;

  return (
    <tr style={s.row}>
      {/* Factory notes */}
      <Td rowIndex={rowIndex} colIndex={0}><NotesCell orderId={order.id} field="factory_notes" value={order.factoryNotes ?? ""} /></Td>

      {/* Order date */}
      <Td rowIndex={rowIndex} colIndex={1} center><span style={s.dateText}>{orderDate}</span></Td>

      {/* Picture */}
      <Td rowIndex={rowIndex} colIndex={2} center>
        {order.productImageUrl
          ? <img src={order.productImageUrl} alt="" style={s.thumb} />
          : <div style={s.noImg}>—</div>}
      </Td>

      {/* Name */}
      <Td rowIndex={rowIndex} colIndex={3}><span style={s.productName}>{order.productTitle}</span></Td>

      {/* SKU */}
      <Td rowIndex={rowIndex} colIndex={4}><span style={s.sku}>{allSkus || "—"}</span></Td>

      {/* Size columns */}
      {sizes.map((sz, sizeIndex) => (
        <Td key={sz} rowIndex={rowIndex} colIndex={5 + sizeIndex} center>
          <QtyCell orderId={order.id} size={sz} value={qtyBySize[sz] ?? 0} />
        </Td>
      ))}

      {/* Total */}
      <Td rowIndex={rowIndex} colIndex={totalCol} center><span style={s.total}>{order.totalQty}</span></Td>

      {/* Status */}
      <Td rowIndex={rowIndex} colIndex={statusCol}><StatusCell orderId={order.id} value={order.supplierStatus} /></Td>

      {/* Notes (from order) */}
      <Td rowIndex={rowIndex} colIndex={notesCol}><NotesCell orderId={order.id} field="notes" value={order.notes ?? ""} /></Td>

      {/* Priority */}
      <Td rowIndex={rowIndex} colIndex={priorityCol}><PriorityCell orderId={order.id} value={order.priority ?? ""} /></Td>

      {/* ETA */}
      <Td rowIndex={rowIndex} colIndex={etaCol}><EtaCell orderId={order.id} value={etaValue} /></Td>

      {/* Delete */}
      <Td rowIndex={rowIndex} colIndex={deleteCol} center><DeleteCell orderId={order.id} /></Td>
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

function QtyCell({ orderId, size, value }: { orderId: number; size: string; value: number }) {
  const fetcher = useFetcher();
  const current = fetcher.formData ? String(fetcher.formData.get("value")) : String(value);
  const numericCurrent = Number(current) || 0;
  const normalizeQty = (input: HTMLInputElement) => {
    input.value = input.value.replace(/\D/g, "");
  };

  return (
    <fetcher.Form method="post">
      <input type="hidden" name="intent" value="update_qty" />
      <input type="hidden" name="orderId" value={orderId} />
      <input type="hidden" name="size" value={size} />
      <input
        type="text"
        inputMode="numeric"
        pattern="[0-9]*"
        name="value"
        defaultValue={value}
        onChange={(e) => normalizeQty(e.currentTarget)}
        onBlur={(e) => fetcher.submit(e.currentTarget.form!)}
        style={{
          ...s.qtyInput,
          ...(numericCurrent > 0 ? s.qtyInputActive : s.qtyInputZero),
        }}
      />
    </fetcher.Form>
  );
}

function DeleteCell({ orderId }: { orderId: number }) {
  const fetcher = useFetcher();
  return (
    <fetcher.Form
      method="post"
      onSubmit={(e) => {
        if (!window.confirm("Are you sure you want to delete this order?")) {
          e.preventDefault();
        }
      }}
    >
      <input type="hidden" name="intent" value="delete_order" />
      <input type="hidden" name="orderId" value={orderId} />
      <button type="submit" style={s.deleteButton}>Delete</button>
    </fetcher.Form>
  );
}

// ─── Table helpers ────────────────────────────────────────────────────────────

function Th({
  children,
  center,
  onResizeStart,
}: {
  children: React.ReactNode;
  center?: boolean;
  onResizeStart: (event: React.MouseEvent<HTMLSpanElement>) => void;
}) {
  return (
    <th style={{ ...s.th, textAlign: center ? "center" : "left" }}>
      <span style={s.thContent}>{children}</span>
      <span
        role="separator"
        aria-orientation="vertical"
        aria-label={`Resize ${String(children)} column`}
        onMouseDown={onResizeStart}
        style={s.resizeHandle}
      />
    </th>
  );
}
function Td({
  children,
  center,
  rowIndex,
  colIndex,
}: {
  children: React.ReactNode;
  center?: boolean;
  rowIndex: number;
  colIndex: number;
}) {
  return (
    <td
      data-grid-row={rowIndex}
      data-grid-col={colIndex}
      tabIndex={0}
      style={{ ...s.td, textAlign: center ? "center" : "left" }}
    >
      {children}
    </td>
  );
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
    borderCollapse: "collapse",
    borderSpacing: 0,
    fontSize: 13,
    minWidth: 900,
    tableLayout: "fixed",
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
    position: "relative",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  thContent: { display: "block", overflow: "hidden", textOverflow: "ellipsis" },
  resizeHandle: {
    position: "absolute",
    top: 0,
    right: -3,
    width: 8,
    height: "100%",
    cursor: "col-resize",
    zIndex: 2,
    touchAction: "none",
  },
  row: { background: "#fff" },
  td: {
    padding: "8px 10px",
    verticalAlign: "middle",
    color: "#374151",
    border: "1px solid #d1d5db",
    background: "#fff",
    overflow: "hidden",
  },
  thumb: { width: 48, height: 64, objectFit: "cover", borderRadius: 2, display: "block", margin: "0 auto" },
  noImg: { color: "#d1d5db", textAlign: "center" },
  productName: { fontWeight: 600, color: "#111827", whiteSpace: "normal", overflowWrap: "anywhere", lineHeight: 1.35 },
  sku: { fontFamily: "monospace", fontSize: 11, color: "#6b7280", whiteSpace: "pre-line" },
  qty: { fontWeight: 700, color: "#111827" },
  qtyZero: { color: "#d1d5db" },
  dateText: { color: "#374151", fontWeight: 600, whiteSpace: "nowrap" },
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
    width: "100%",
    boxSizing: "border-box",
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
    width: "100%",
    boxSizing: "border-box",
  },
  qtyInput: {
    display: "block",
    width: "100%",
    border: "1px solid transparent",
    borderRadius: 3,
    padding: "4px 0",
    fontSize: 13,
    fontWeight: 700,
    textAlign: "center",
    outline: "none",
    background: "transparent",
    boxSizing: "border-box",
  },
  qtyInputActive: { color: "#111827" },
  qtyInputZero: { color: "#d1d5db" },
  deleteButton: {
    border: "1px solid #fecaca",
    borderRadius: 3,
    padding: "5px 8px",
    background: "#fee2e2",
    color: "#991b1b",
    fontSize: 12,
    fontWeight: 700,
    cursor: "pointer",
  },
};
