import { useEffect, useState, useCallback } from "react";
import {
  reactExtension,
  useApi,
  BlockStack,
  InlineStack,
  Box,
  Text,
  Button,
  Banner,
  Divider,
  Badge,
  ProgressIndicator,
  TextField,
  Select,
} from "@shopify/ui-extensions-react/admin";

const TARGET = "admin.product-details.block.render";
const APP_URL = "https://product-demand-production.up.railway.app";

export default reactExtension(TARGET, () => <ProductOrderBlock />);

type OrderLineStatus = { variantId: string; qtyOrdered: number };
type OrderStatusItem = {
  id: number;
  supplier: string;
  totalQty: number;
  productType?: string | null;
  eta: string | null;
  supplierStatus?: string | null;
  priority?: string | null;
  notes?: string | null;
  lines?: OrderLineStatus[];
};
type OrderStatus = OrderStatusItem | null;
type Variant = { id: string; title: string; sku: string; stockQty: number; onOrderQty: number; qtyOrdered: string };

const ORDER_LIMIT = 2;
const GAP = "small" as const;
const STATUS_OPTIONS = [
  { value: "on_order", label: "On Order" },
  { value: "on_production", label: "On Production" },
  { value: "in_shipment", label: "In Shipment" },
  { value: "arrived", label: "Arrived" },
  { value: "arrived_loaded", label: "Arrived and Loaded" },
  { value: "cancelled", label: "Cancelled" },
  { value: "ready_to_send", label: "Ready To Send" },
];
const PRIORITY_OPTIONS = [
  { value: "", label: "— Priority —" },
  { value: "low", label: "LOW" },
  { value: "high", label: "HIGH" },
  { value: "urgent", label: "URGENT" },
  { value: "cancelled", label: "Cancelled" },
];
const BASE_PRODUCT_GROUP_OPTIONS = [
  "Dresses",
  "Tops",
  "Skirts",
  "Pants",
  "Corduroy",
];

function labelFor(options: Array<{ value: string; label: string }>, value?: string | null) {
  return options.find((option) => option.value === value)?.label ?? "On Order";
}

function formatShortDate(value?: string | null) {
  if (!value) return "No ETA";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "No ETA";
  return date.toLocaleDateString("en-AU", { day: "2-digit", month: "2-digit", year: "2-digit" });
}

function tableWidths(orderCount: number) {
  return orderCount <= 1
    ? { size: "16%", stock: "16%", order: "18%", addOrder: "24%", total: "16%" }
    : { size: "14%", stock: "12%", order: "13%", addOrder: "21%", total: "12%" };
}

function productGroupOptions(currentGroup: string) {
  const options = [currentGroup, ...BASE_PRODUCT_GROUP_OPTIONS]
    .map((value) => value.trim())
    .filter(Boolean);
  return [
    { value: "", label: "— Product group —" },
    ...Array.from(new Set(options)).map((value) => ({ value, label: value })),
  ];
}

function Col({ w, align = "start", children }: { w: string; align?: "start" | "end" | "center"; children: React.ReactNode }) {
  return (
    <Box inlineSize={w as any}>
      <Box inlineSize="100%">
        <InlineStack inlineAlignment={align} blockAlignment="center">{children}</InlineStack>
      </Box>
    </Box>
  );
}

function CenterCell({ children }: { children: React.ReactNode }) {
  return (
    <Box inlineSize="100%">
      <InlineStack inlineAlignment="center" blockAlignment="center">
        {children}
      </InlineStack>
    </Box>
  );
}

function AddOrderCell({ value, onChange }: { value: string; onChange: (value: string) => void }) {
  return (
    <CenterCell>
      <Box inlineSize="78%">
        <TextField
          label=" "
          value={value}
          onChange={onChange}
        />
      </Box>
    </CenterCell>
  );
}

function ProductOrderBlock() {
  const { data, auth, query } = useApi(TARGET);
  const productGid: string | undefined = data.selected?.[0]?.id;

  const [shop, setShop] = useState<string | null>(null);
  const [productTitle, setProductTitle] = useState("");
  const [productGroup, setProductGroup] = useState("");
  const [productVendor, setProductVendor] = useState("");
  const [productImageUrl, setProductImageUrl] = useState<string | null>(null);
  const [order, setOrder] = useState<OrderStatus>(null);
  const [orders, setOrders] = useState<OrderStatusItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [showForm, setShowForm] = useState(false);
  const [variants, setVariants] = useState<Variant[]>([]);
  const [formError, setFormError] = useState<string | null>(null);
  const [submitting, setSubmitting] = useState(false);
  const [savingPriority, setSavingPriority] = useState(false);
  const [savingProductGroup, setSavingProductGroup] = useState(false);
  const [successMsg, setSuccessMsg] = useState<string | null>(null);
  const [notes, setNotes] = useState("");
  const [orderPriority, setOrderPriority] = useState("");
  const [orderProductGroup, setOrderProductGroup] = useState("");

  const applyOrderStatus = useCallback((nextOrder: OrderStatus, nextOrders: OrderStatusItem[]) => {
    const visibleOrders = nextOrders.slice(0, ORDER_LIMIT);
    setOrder(nextOrder);
    setOrders(visibleOrders);
    setOrderPriority(nextOrder?.priority || "");
    setOrderProductGroup((current) => nextOrder?.productType || current || productGroup || "");

    const onOrderByVariant = new Map<string, number>();
    for (const item of visibleOrders) {
      for (const line of item.lines ?? []) {
        onOrderByVariant.set(
          line.variantId,
          (onOrderByVariant.get(line.variantId) ?? 0) + line.qtyOrdered,
        );
      }
    }
    setVariants((prev) => prev.map((variant) => ({
      ...variant,
      onOrderQty: onOrderByVariant.get(variant.id) ?? 0,
    })));
  }, [productGroup]);

  const refreshOrderStatus = useCallback(async (shopDomain: string, productId: string) => {
    const token = await auth.idToken();
    if (!token) throw new Error("No auth token");
    const res = await fetch(
      `${APP_URL}/api/order-status?productId=${encodeURIComponent(productId)}&shop=${encodeURIComponent(shopDomain)}`,
      { headers: { Authorization: `Bearer ${token}` } },
    );
    const json = await res.json();
    if (!res.ok) throw new Error(json.error ?? "Could not load order status");
    applyOrderStatus(json.order ?? null, json.orders ?? []);
  }, [applyOrderStatus, auth]);

  useEffect(() => {
    if (!productGid) { setLoading(false); return; }
    async function init() {
      try {
        const result = await query<{
          shop: { myshopifyDomain: string };
          product: {
            title: string; vendor: string;
            productType: string;
            featuredImage: { url: string } | null;
            variants: { nodes: Array<{ id: string; title: string; sku: string; inventoryQuantity: number }> };
          };
        }>(`{
          shop { myshopifyDomain }
          product(id: "${productGid}") {
            title vendor productType
            featuredImage { url }
            variants(first: 100) { nodes { id title sku inventoryQuantity } }
          }
        }`);
        const shopDomain = result.data?.shop?.myshopifyDomain;
        if (!shopDomain) throw new Error("Could not resolve shop");
        setShop(shopDomain);
        const p = result.data?.product;
        if (p) {
          setProductTitle(p.title);
          setProductGroup(p.productType ?? "");
          setOrderProductGroup((current) => current || p.productType || "");
          setProductVendor(p.vendor ?? "");
          setProductImageUrl(p.featuredImage?.url ?? null);
          setVariants((p.variants?.nodes ?? []).map((v) => ({
            id: v.id, title: v.title, sku: v.sku ?? "",
            stockQty: v.inventoryQuantity ?? 0, onOrderQty: 0, qtyOrdered: "",
          })));
        }
        await refreshOrderStatus(shopDomain, productGid!);
      } catch (e: any) {
        setErrorMsg(`Error: ${e?.message ?? String(e)}`);
      } finally {
        setLoading(false);
      }
    }
    init();
  }, [productGid, query, refreshOrderStatus]);

  const updateQty = useCallback((idx: number, val: string) => {
    setVariants((prev) => prev.map((v, i) => i === idx ? { ...v, qtyOrdered: val.replace(/\D/g, "") } : v));
  }, []);

  async function updatePriority(nextPriority: string) {
    setOrderPriority(nextPriority);
    if (!order || !shop) return;
    setSavingPriority(true);
    setFormError(null);
    try {
      const token = await auth.idToken();
      if (!token) throw new Error("No auth token");
      const res = await fetch(`${APP_URL}/api/update-order`, {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({
          shop,
          orderId: order.id,
          priority: nextPriority,
        }),
      });
      const json = await res.json();
      if (!res.ok) { setFormError(`Error ${res.status}: ${json.error ?? "unknown"}`); return; }
      setOrder((current) => current ? { ...current, priority: json.order.priority } : current);
      setOrders((current) => current.map((item) => (
        item.id === json.order.id ? { ...item, priority: json.order.priority } : item
      )));
    } catch (e: any) {
      setFormError(`Error: ${e?.message ?? String(e)}`);
    } finally {
      setSavingPriority(false);
    }
  }

  async function updateProductGroup(nextProductGroup: string) {
    setOrderProductGroup(nextProductGroup);
    if (!order || !shop) return;
    setSavingProductGroup(true);
    setFormError(null);
    try {
      const token = await auth.idToken();
      if (!token) throw new Error("No auth token");
      const res = await fetch(`${APP_URL}/api/update-order`, {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({
          shop,
          orderId: order.id,
          productType: nextProductGroup,
        }),
      });
      const json = await res.json();
      if (!res.ok) { setFormError(`Error ${res.status}: ${json.error ?? "unknown"}`); return; }
      setOrder((current) => current ? { ...current, productType: json.order.productType } : current);
      setOrders((current) => current.map((item) => (
        item.id === json.order.id ? { ...item, productType: json.order.productType } : item
      )));
    } catch (e: any) {
      setFormError(`Error: ${e?.message ?? String(e)}`);
    } finally {
      setSavingProductGroup(false);
    }
  }

  async function handleSubmit(mode: "existing" | "new") {
    const orderedLines = variants.filter((v) => Number(v.qtyOrdered) > 0);
    const trimmedNotes = notes.trim();
    if (!orderedLines.length && !trimmedNotes) {
      setFormError("Enter a quantity or add an order note");
      return;
    }
    setFormError(null);
    setSubmitting(true);
    try {
      const token = await auth.idToken();
      if (!token) throw new Error("No auth token");
      const res = await fetch(`${APP_URL}/api/place-order`, {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({
          shop, productId: productGid, productTitle,
          productType: orderProductGroup || undefined,
          productImageUrl: productImageUrl || undefined,
          supplier: productVendor || "Unknown",
          notes: trimmedNotes || undefined,
          priority: orderPriority || undefined,
          existingOrderId: mode === "existing" ? order?.id : undefined,
          lines: orderedLines.map((v) => ({
            variantId: v.id, variantTitle: v.title,
            sku: v.sku || undefined, qtyOrdered: Number(v.qtyOrdered),
          })),
        }),
      });
      const json = await res.json();
      if (!res.ok) { setFormError(`Error ${res.status}: ${json.error ?? "unknown"}`); return; }
      const total = orderedLines.reduce((s, v) => s + Number(v.qtyOrdered), 0);
      setSuccessMsg(
        total > 0
          ? `${mode === "existing" ? "Existing order updated" : "New order created"} — ${total} units`
          : `${mode === "existing" ? "Existing order note added" : "New order note created"}`,
      );
      if (shop && productGid) await refreshOrderStatus(shop, productGid);
      setShowForm(false);
      setNotes("");
      setVariants((prev) => prev.map((v) => ({ ...v, qtyOrdered: "" })));
    } catch (e: any) {
      setFormError(`Error: ${e?.message ?? String(e)}`);
    } finally {
      setSubmitting(false);
    }
  }

  if (loading) return <InlineStack inlineAlignment="center"><ProgressIndicator size="small-200" /></InlineStack>;
  if (errorMsg) return <Banner tone="warning">{errorMsg}</Banner>;

  const statusRow = (
    <InlineStack gap="small" blockAlignment="center">
      {order ? (
        <>
          <Badge tone="success">On order</Badge>
          <Text>
            {order.totalQty} units · {order.supplier} · {labelFor(STATUS_OPTIONS, order.supplierStatus)}
          </Text>
        </>
      ) : (
        <Badge>Not on order</Badge>
      )}
    </InlineStack>
  );

  if (!showForm) {
    return (
      <BlockStack gap="base">
        {successMsg && <Banner tone="success">{successMsg}</Banner>}
        {statusRow}
        <Button onPress={() => setShowForm(true)}>Add order or order note</Button>
      </BlockStack>
    );
  }

  const totalStock = variants.reduce((s, v) => s + v.stockQty, 0);
  const orderColumns = orders.slice(0, ORDER_LIMIT);
  const W = tableWidths(orderColumns.length);
  const getOrderQty = (item: OrderStatusItem, variantId: string) =>
    (item.lines ?? [])
      .filter((line) => line.variantId === variantId)
      .reduce((sum, line) => sum + line.qtyOrdered, 0);
  const totalOnOrder = variants.reduce((s, v) => s + v.onOrderQty, 0);
  const totalAddOrder = variants.reduce((s, v) => s + (Number(v.qtyOrdered) || 0), 0);

  return (
    <BlockStack gap="small">
      {statusRow}
      <Divider />
      {formError && <Banner tone="critical">{formError}</Banner>}

      <InlineStack gap="base" blockAlignment="end">
        <Box inlineSize="30%">
          <Select
            label={savingPriority ? "Priority saving..." : "Priority"}
            value={orderPriority}
            options={PRIORITY_OPTIONS}
            onChange={updatePriority}
          />
        </Box>
        <Box inlineSize="30%">
          <Select
            label={savingProductGroup ? "Group saving..." : "Product group"}
            value={orderProductGroup}
            options={productGroupOptions(productGroup || orderProductGroup)}
            onChange={updateProductGroup}
          />
        </Box>
      </InlineStack>
      <TextField
        label="Notes for supplier portal"
        value={notes}
        onChange={setNotes}
      />
      <Divider />

      {/* Header */}
      <InlineStack gap={GAP} blockAlignment="center">
        <Col w={W.size}><Text fontWeight="bold">Size</Text></Col>
        <Col w={W.stock} align="center"><Text fontWeight="bold">In stock</Text></Col>
        {orderColumns.map((item, idx) => (
          <Col key={item.id} w={W.order} align="center">
            <BlockStack gap="none">
              <Text fontWeight="bold">{formatShortDate(item.eta)}</Text>
              <Text fontWeight="bold">On order {idx + 1}</Text>
            </BlockStack>
          </Col>
        ))}
        <Col w={W.addOrder} align="center"><Text fontWeight="bold">Add order</Text></Col>
        <Col w={W.total} align="center"><Text fontWeight="bold">Total</Text></Col>
      </InlineStack>
      <Divider />

      {/* Rows */}
      {variants.map((v, idx) => {
        const rowTotal = v.stockQty + v.onOrderQty + (Number(v.qtyOrdered) || 0);
        return (
          <InlineStack key={v.id} gap={GAP} blockAlignment="center">
            <Col w={W.size}><Text>{v.title}</Text></Col>
            <Col w={W.stock} align="center"><CenterCell><Text>{v.stockQty}</Text></CenterCell></Col>
            {orderColumns.map((item) => (
              <Col key={item.id} w={W.order} align="center"><CenterCell><Text>{getOrderQty(item, v.id)}</Text></CenterCell></Col>
            ))}
            <Col w={W.addOrder} align="center">
              <AddOrderCell value={v.qtyOrdered} onChange={(val) => updateQty(idx, val)} />
            </Col>
            <Col w={W.total} align="center"><CenterCell><Text fontWeight="bold">{rowTotal}</Text></CenterCell></Col>
          </InlineStack>
        );
      })}

      <Divider />

      {/* Totals */}
      <InlineStack gap={GAP} blockAlignment="center">
        <Col w={W.size}><Text fontWeight="bold">Total</Text></Col>
        <Col w={W.stock} align="center"><Text fontWeight="bold">{totalStock}</Text></Col>
        {orderColumns.map((item) => (
          <Col key={item.id} w={W.order} align="center"><Text fontWeight="bold">{item.totalQty}</Text></Col>
        ))}
        <Col w={W.addOrder} align="center"><Text fontWeight="bold">{totalAddOrder}</Text></Col>
        <Col w={W.total} align="center"><Text fontWeight="bold">{totalStock + totalOnOrder + totalAddOrder}</Text></Col>
      </InlineStack>

      <Divider />
      <InlineStack gap="base">
        {order && (
          <Button variant="primary" onPress={() => handleSubmit("existing")}>
            {submitting ? "Saving..." : "Add to existing order"}
          </Button>
        )}
        {orders.length < ORDER_LIMIT ? (
          <Button variant={order ? "secondary" : "primary"} onPress={() => handleSubmit("new")}>
            {submitting ? "Saving..." : "Create new order"}
          </Button>
        ) : (
          <Text>Maximum 2 open orders. Add to existing order.</Text>
        )}
        <Button variant="tertiary" onPress={() => { setShowForm(false); setFormError(null); }}>Cancel</Button>
      </InlineStack>
    </BlockStack>
  );
}
