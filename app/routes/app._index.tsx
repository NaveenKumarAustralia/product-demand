import { useState, useEffect } from "react";
import type { HeadersFunction, LoaderFunctionArgs, ActionFunctionArgs } from "react-router";
import { useLoaderData, useFetcher } from "react-router";
import { boundary } from "@shopify/shopify-app-react-router/server";
import { authenticate } from "../shopify.server";
import prisma from "../db.server";

// ─── Types ───────────────────────────────────────────────────────────────────

type Variant = {
  id: string;
  title: string;
  sku: string;
  inventoryQuantity: number | null;
};

type Product = {
  id: string;
  title: string;
  vendor: string;
  status: string;
  featuredImage: { url: string; altText: string | null } | null;
  variants: Variant[];
};

type OpenOrder = {
  id: number;
  supplier: string;
  totalQty: number;
  eta: string | null;
  status: string;
};

// ─── Loader ──────────────────────────────────────────────────────────────────

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const { admin, session } = await authenticate.admin(request);

  const response = await admin.graphql(`
    #graphql
    query getProducts {
      products(first: 50, sortKey: TITLE) {
        edges {
          node {
            id
            title
            vendor
            status
            featuredImage { url altText }
            variants(first: 10) {
              edges {
                node { id title sku inventoryQuantity }
              }
            }
          }
        }
      }
    }
  `);

  const json = await response.json();
  const products: Product[] = json.data.products.edges.map((edge: any) => ({
    ...edge.node,
    variants: edge.node.variants.edges.map((v: any) => v.node),
  }));

  // Load open orders for this shop so we can show on-order status
  const openOrders = await prisma.supplierOrder.findMany({
    where: { shop: session.shop, status: "open" },
    select: {
      id: true,
      productId: true,
      supplier: true,
      totalQty: true,
      eta: true,
      status: true,
    },
  });

  // Map by productId for quick lookup
  const ordersByProduct: Record<string, OpenOrder> = {};
  for (const o of openOrders) {
    ordersByProduct[o.productId] = {
      id: o.id,
      supplier: o.supplier,
      totalQty: o.totalQty,
      eta: o.eta ? o.eta.toISOString() : null,
      status: o.status,
    };
  }

  return { products, ordersByProduct };
};

// ─── Action ──────────────────────────────────────────────────────────────────

export const action = async ({ request }: ActionFunctionArgs) => {
  const { session } = await authenticate.admin(request);
  const formData = await request.formData();
  const intent = formData.get("intent") as string;

  if (intent === "place_order") {
    const productId = formData.get("productId") as string;
    const productTitle = formData.get("productTitle") as string;
    const supplier = (formData.get("supplier") as string)?.trim();
    const eta = formData.get("eta") as string;
    const poNumber = formData.get("poNumber") as string;
    const notes = formData.get("notes") as string;
    const linesJson = formData.get("lines") as string;

    if (!supplier) return { ok: false, error: "Supplier name is required." };

    const lines: Array<{
      variantId: string;
      variantTitle: string;
      sku: string;
      qty: number;
      costPrice: string;
    }> = JSON.parse(linesJson);

    const orderLines = lines.filter((l) => l.qty > 0);
    if (orderLines.length === 0) {
      return { ok: false, error: "Enter a quantity for at least one variant." };
    }

    const totalQty = orderLines.reduce((sum, l) => sum + l.qty, 0);

    await prisma.supplierOrder.create({
      data: {
        shop: session.shop,
        productId,
        productTitle,
        supplier,
        poNumber: poNumber || null,
        notes: notes || null,
        eta: eta ? new Date(eta) : null,
        totalQty,
        lines: {
          create: orderLines.map((l) => ({
            variantId: l.variantId,
            variantTitle: l.variantTitle,
            sku: l.sku || null,
            qtyOrdered: l.qty,
            costPrice: l.costPrice ? parseFloat(l.costPrice) : null,
          })),
        },
      },
    });

    return { ok: true };
  }

  return { ok: false, error: "Unknown action." };
};

// ─── Helpers ─────────────────────────────────────────────────────────────────

function totalStock(variants: Variant[]) {
  return variants.reduce((sum, v) => sum + (v.inventoryQuantity ?? 0), 0);
}

function formatDate(iso: string) {
  return new Date(iso).toLocaleDateString("en-AU", {
    day: "numeric",
    month: "short",
    year: "numeric",
  });
}

// ─── Sub-components ──────────────────────────────────────────────────────────

function StockBadge({ stock }: { stock: number }) {
  if (stock === 0) return <span style={badge("critical")}>Out of stock</span>;
  if (stock < 10) return <span style={badge("warning")}>Low: {stock}</span>;
  return <span style={badge("success")}>In stock: {stock}</span>;
}

function OrderModal({
  product,
  onClose,
}: {
  product: Product;
  onClose: () => void;
}) {
  const fetcher = useFetcher<typeof action>();
  const isSubmitting = fetcher.state !== "idle";

  const [supplier, setSupplier] = useState(product.vendor || "");
  const [eta, setEta] = useState("");
  const [poNumber, setPoNumber] = useState("");
  const [notes, setNotes] = useState("");
  const [qtys, setQtys] = useState<Record<string, number>>(
    Object.fromEntries(product.variants.map((v) => [v.id, 0])),
  );
  const [costs, setCosts] = useState<Record<string, string>>(
    Object.fromEntries(product.variants.map((v) => [v.id, ""])),
  );

  // Close on success
  useEffect(() => {
    if (fetcher.data?.ok) onClose();
  }, [fetcher.data]);

  function handleSubmit(e: React.FormEvent<HTMLFormElement>) {
    e.preventDefault();
    const lines = product.variants.map((v) => ({
      variantId: v.id,
      variantTitle: v.title,
      sku: v.sku || "",
      qty: qtys[v.id] ?? 0,
      costPrice: costs[v.id] || "",
    }));

    const fd = new FormData();
    fd.append("intent", "place_order");
    fd.append("productId", product.id);
    fd.append("productTitle", product.title);
    fd.append("supplier", supplier);
    fd.append("eta", eta);
    fd.append("poNumber", poNumber);
    fd.append("notes", notes);
    fd.append("lines", JSON.stringify(lines));

    fetcher.submit(fd, { method: "POST" });
  }

  return (
    <div style={s.overlay} onClick={onClose}>
      <div style={s.modal} onClick={(e) => e.stopPropagation()}>
        {/* Header */}
        <div style={s.modalHeader}>
          <div>
            <div style={s.modalTitle}>Place supplier order</div>
            <div style={s.modalSubtitle}>{product.title}</div>
          </div>
          <button style={s.closeBtn} onClick={onClose}>✕</button>
        </div>

        <form onSubmit={handleSubmit}>
          <div style={s.modalBody}>
            {/* Error */}
            {fetcher.data && !fetcher.data.ok && (
              <div style={s.errorBox}>{fetcher.data.error}</div>
            )}

            {/* Supplier + ETA row */}
            <div style={s.row}>
              <div style={s.field}>
                <label style={s.label}>Supplier *</label>
                <input
                  style={s.input}
                  value={supplier}
                  onChange={(e) => setSupplier(e.target.value)}
                  placeholder="e.g. Aisling Fashion"
                  required
                />
              </div>
              <div style={s.field}>
                <label style={s.label}>Expected arrival (ETA)</label>
                <input
                  style={s.input}
                  type="date"
                  value={eta}
                  onChange={(e) => setEta(e.target.value)}
                />
              </div>
            </div>

            {/* PO Number + Notes row */}
            <div style={s.row}>
              <div style={s.field}>
                <label style={s.label}>PO number (optional)</label>
                <input
                  style={s.input}
                  value={poNumber}
                  onChange={(e) => setPoNumber(e.target.value)}
                  placeholder="e.g. PO-1042"
                />
              </div>
              <div style={s.field}>
                <label style={s.label}>Notes (optional)</label>
                <input
                  style={s.input}
                  value={notes}
                  onChange={(e) => setNotes(e.target.value)}
                  placeholder="Any notes for this order"
                />
              </div>
            </div>

            {/* Variant table */}
            <div style={{ marginTop: 20 }}>
              <label style={s.label}>Variants & quantities *</label>
              <div style={s.variantTable}>
                <div style={s.variantHeader}>
                  <div style={{ flex: 3 }}>Variant</div>
                  <div style={{ flex: 1, textAlign: "center" as const }}>Current stock</div>
                  <div style={{ flex: 1, textAlign: "center" as const }}>Qty to order</div>
                  <div style={{ flex: 1, textAlign: "center" as const }}>Cost price</div>
                </div>
                {product.variants.map((v) => (
                  <div key={v.id} style={s.variantRow}>
                    <div style={{ flex: 3 }}>
                      <span style={s.variantName}>
                        {v.title === "Default Title" ? product.title : v.title}
                      </span>
                      {v.sku && <span style={s.sku}> · SKU: {v.sku}</span>}
                    </div>
                    <div style={{ flex: 1, textAlign: "center" as const }}>
                      <span style={{ color: (v.inventoryQuantity ?? 0) < 5 ? "#c0392b" : "#333" }}>
                        {v.inventoryQuantity ?? 0}
                      </span>
                    </div>
                    <div style={{ flex: 1 }}>
                      <input
                        style={{ ...s.numInput, margin: "0 auto", display: "block" }}
                        type="number"
                        min={0}
                        value={qtys[v.id] ?? 0}
                        onChange={(e) =>
                          setQtys((prev) => ({
                            ...prev,
                            [v.id]: parseInt(e.target.value) || 0,
                          }))
                        }
                      />
                    </div>
                    <div style={{ flex: 1 }}>
                      <input
                        style={{ ...s.numInput, margin: "0 auto", display: "block" }}
                        type="number"
                        min={0}
                        step="0.01"
                        value={costs[v.id] ?? ""}
                        onChange={(e) =>
                          setCosts((prev) => ({ ...prev, [v.id]: e.target.value }))
                        }
                        placeholder="$0.00"
                      />
                    </div>
                  </div>
                ))}
              </div>

              {/* Total summary */}
              {Object.values(qtys).some((q) => q > 0) && (
                <div style={s.summary}>
                  Total units to order:{" "}
                  <strong>
                    {Object.values(qtys).reduce((a, b) => a + b, 0)}
                  </strong>
                </div>
              )}
            </div>
          </div>

          {/* Footer */}
          <div style={s.modalFooter}>
            <button type="button" style={s.cancelBtn} onClick={onClose}>
              Cancel
            </button>
            <button type="submit" style={s.submitBtn} disabled={isSubmitting}>
              {isSubmitting ? "Placing order…" : "Place order"}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

// ─── Main page ───────────────────────────────────────────────────────────────

export default function ProductList() {
  const { products, ordersByProduct } = useLoaderData<typeof loader>();
  const [selectedProduct, setSelectedProduct] = useState<Product | null>(null);

  return (
    <s-page heading="Product Demand">
      <s-section heading={`${products.length} products`}>
        <div style={s.tableWrapper}>
          {/* Column headers */}
          <div style={s.headerRow}>
            <div style={{ width: 56 }} />
            <div style={{ flex: 3 }}><s-text>Product</s-text></div>
            <div style={{ flex: 2 }}><s-text>Supplier</s-text></div>
            <div style={{ flex: 2 }}><s-text>Stock</s-text></div>
            <div style={{ flex: 2 }}><s-text>Order status</s-text></div>
            <div style={{ flex: 1 }} />
          </div>

          {/* Rows */}
          {products.map((product) => {
            const stock = totalStock(product.variants);
            const openOrder = ordersByProduct[product.id];

            const variantSummary = product.variants
              .filter((v) => v.title !== "Default Title")
              .map((v) => `${v.title}: ${v.inventoryQuantity ?? 0}`)
              .join(" · ");

            return (
              <div key={product.id} style={s.productRow}>
                {/* Thumbnail */}
                <div style={{ width: 56, flexShrink: 0 }}>
                  {product.featuredImage ? (
                    <img
                      src={product.featuredImage.url}
                      alt={product.featuredImage.altText || product.title}
                      style={s.thumbnail}
                    />
                  ) : (
                    <div style={s.noImage} />
                  )}
                </div>

                {/* Product + variants */}
                <div style={{ flex: 3 }}>
                  <div style={s.productTitle}>{product.title}</div>
                  {variantSummary && (
                    <div style={s.variantSummary}>{variantSummary}</div>
                  )}
                </div>

                {/* Supplier */}
                <div style={{ flex: 2 }}>
                  <span style={s.cell}>{product.vendor || "—"}</span>
                </div>

                {/* Stock badge */}
                <div style={{ flex: 2 }}>
                  <StockBadge stock={stock} />
                </div>

                {/* Order status */}
                <div style={{ flex: 2 }}>
                  {openOrder ? (
                    <div>
                      <span style={badge("info")}>On order: {openOrder.totalQty} units</span>
                      {openOrder.eta && (
                        <div style={s.eta}>ETA {formatDate(openOrder.eta)}</div>
                      )}
                    </div>
                  ) : (
                    <span style={badge("neutral")}>Not on order</span>
                  )}
                </div>

                {/* Action */}
                <div style={{ flex: 1 }}>
                  <button
                    style={openOrder ? s.reorderBtn : s.orderBtn}
                    onClick={() => setSelectedProduct(product)}
                  >
                    {openOrder ? "Reorder" : "Place order"}
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      </s-section>

      {/* Order modal */}
      {selectedProduct && (
        <OrderModal
          product={selectedProduct}
          onClose={() => setSelectedProduct(null)}
        />
      )}
    </s-page>
  );
}

// ─── Styles ──────────────────────────────────────────────────────────────────

function badge(tone: "success" | "warning" | "critical" | "neutral" | "info") {
  const map = {
    success: { bg: "#d3f5e2", color: "#1a6640" },
    warning: { bg: "#fff3cd", color: "#7d5c00" },
    critical: { bg: "#fde8e8", color: "#8c1515" },
    neutral: { bg: "#f1f2f3", color: "#616161" },
    info: { bg: "#ddeeff", color: "#1a4a7a" },
  };
  return {
    display: "inline-block",
    padding: "3px 10px",
    borderRadius: 12,
    fontSize: 12,
    fontWeight: 500,
    background: map[tone].bg,
    color: map[tone].color,
  };
}

const s = {
  tableWrapper: { display: "flex", flexDirection: "column" as const, gap: 2 },

  headerRow: {
    display: "flex",
    alignItems: "center",
    padding: "6px 12px",
    gap: 12,
    borderBottom: "1px solid #e1e3e5",
    color: "#616161",
    fontSize: 12,
  },

  productRow: {
    display: "flex",
    alignItems: "center",
    padding: "12px",
    gap: 12,
    borderBottom: "1px solid #f1f2f3",
    borderRadius: 8,
  },

  thumbnail: {
    width: 48,
    height: 48,
    objectFit: "cover" as const,
    borderRadius: 6,
    border: "1px solid #e1e3e5",
  },

  noImage: {
    width: 48,
    height: 48,
    borderRadius: 6,
    background: "#f1f2f3",
    border: "1px solid #e1e3e5",
  },

  productTitle: { fontWeight: 500, fontSize: 14, color: "#202223" },
  variantSummary: { fontSize: 12, color: "#616161", marginTop: 2 },
  cell: { fontSize: 14, color: "#202223" },
  eta: { fontSize: 11, color: "#616161", marginTop: 2 },

  orderBtn: {
    padding: "6px 14px",
    borderRadius: 6,
    border: "none",
    background: "#008060",
    color: "#fff",
    fontWeight: 600,
    fontSize: 13,
    cursor: "pointer",
    whiteSpace: "nowrap" as const,
  },

  reorderBtn: {
    padding: "6px 14px",
    borderRadius: 6,
    border: "1px solid #8c9196",
    background: "#fff",
    color: "#202223",
    fontWeight: 600,
    fontSize: 13,
    cursor: "pointer",
    whiteSpace: "nowrap" as const,
  },

  // Modal
  overlay: {
    position: "fixed" as const,
    inset: 0,
    background: "rgba(0,0,0,0.45)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 1000,
  },

  modal: {
    background: "#fff",
    borderRadius: 12,
    width: "min(740px, 95vw)",
    maxHeight: "90vh",
    display: "flex",
    flexDirection: "column" as const,
    boxShadow: "0 20px 60px rgba(0,0,0,0.2)",
    overflow: "hidden",
  },

  modalHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    padding: "20px 24px 16px",
    borderBottom: "1px solid #e1e3e5",
  },

  modalTitle: { fontSize: 18, fontWeight: 700, color: "#202223" },
  modalSubtitle: { fontSize: 13, color: "#616161", marginTop: 2 },

  closeBtn: {
    background: "none",
    border: "none",
    fontSize: 18,
    cursor: "pointer",
    color: "#616161",
    padding: 4,
    lineHeight: 1,
  },

  modalBody: {
    padding: "20px 24px",
    overflowY: "auto" as const,
    flex: 1,
  },

  modalFooter: {
    display: "flex",
    justifyContent: "flex-end",
    gap: 10,
    padding: "16px 24px",
    borderTop: "1px solid #e1e3e5",
    background: "#fafbfb",
  },

  row: { display: "flex", gap: 16, marginBottom: 16 },

  field: { flex: 1, display: "flex", flexDirection: "column" as const, gap: 4 },

  label: { fontSize: 13, fontWeight: 600, color: "#202223" },

  input: {
    padding: "8px 12px",
    borderRadius: 6,
    border: "1px solid #c9cccf",
    fontSize: 14,
    color: "#202223",
    outline: "none",
    width: "100%",
    boxSizing: "border-box" as const,
  },

  numInput: {
    width: 80,
    padding: "6px 8px",
    borderRadius: 6,
    border: "1px solid #c9cccf",
    fontSize: 14,
    textAlign: "center" as const,
    outline: "none",
  },

  variantTable: {
    border: "1px solid #e1e3e5",
    borderRadius: 8,
    overflow: "hidden",
    marginTop: 8,
  },

  variantHeader: {
    display: "flex",
    gap: 12,
    padding: "8px 16px",
    background: "#f6f6f7",
    fontSize: 12,
    fontWeight: 600,
    color: "#616161",
  },

  variantRow: {
    display: "flex",
    alignItems: "center",
    gap: 12,
    padding: "10px 16px",
    borderTop: "1px solid #f1f2f3",
  },

  variantName: { fontSize: 14, color: "#202223", fontWeight: 500 },
  sku: { fontSize: 12, color: "#8c9196" },

  summary: {
    marginTop: 12,
    padding: "8px 16px",
    background: "#f6f6f7",
    borderRadius: 6,
    fontSize: 13,
    color: "#202223",
  },

  errorBox: {
    background: "#fde8e8",
    color: "#8c1515",
    borderRadius: 6,
    padding: "10px 14px",
    fontSize: 13,
    marginBottom: 16,
  },

  cancelBtn: {
    padding: "8px 20px",
    borderRadius: 6,
    border: "1px solid #c9cccf",
    background: "#fff",
    fontSize: 14,
    cursor: "pointer",
    color: "#202223",
  },

  submitBtn: {
    padding: "8px 24px",
    borderRadius: 6,
    border: "none",
    background: "#008060",
    color: "#fff",
    fontSize: 14,
    fontWeight: 600,
    cursor: "pointer",
  },
};

export const headers: HeadersFunction = (headersArgs) => {
  return boundary.headers(headersArgs);
};
