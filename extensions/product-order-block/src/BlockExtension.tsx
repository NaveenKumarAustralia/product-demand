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
} from "@shopify/ui-extensions-react/admin";

const TARGET = "admin.product-details.block.render";
const APP_URL = "https://product-demand-production.up.railway.app";

export default reactExtension(TARGET, () => <ProductOrderBlock />);

type OrderLineStatus = { variantId: string; qtyOrdered: number };
type OrderStatus = {
  id: number;
  supplier: string;
  totalQty: number;
  eta: string | null;
  lines?: OrderLineStatus[];
} | null;
type Variant = { id: string; title: string; sku: string; stockQty: number; onOrderQty: number; qtyOrdered: string };

const W = { size: "22%", stock: "15%", onOrder: "17%", addOrder: "25%", total: "15%", gap: "base" } as const;

function Col({ w, align = "start", children }: { w: string; align?: "start" | "end" | "center"; children: React.ReactNode }) {
  return (
    <Box inlineSize={w as any}>
      <InlineStack inlineAlignment={align} blockAlignment="center">{children}</InlineStack>
    </Box>
  );
}

function ProductOrderBlock() {
  const { data, auth, query } = useApi(TARGET);
  const productGid: string | undefined = data.selected?.[0]?.id;

  const [shop, setShop] = useState<string | null>(null);
  const [productTitle, setProductTitle] = useState("");
  const [productVendor, setProductVendor] = useState("");
  const [productImageUrl, setProductImageUrl] = useState<string | null>(null);
  const [order, setOrder] = useState<OrderStatus>(null);
  const [loading, setLoading] = useState(true);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [showForm, setShowForm] = useState(false);
  const [variants, setVariants] = useState<Variant[]>([]);
  const [formError, setFormError] = useState<string | null>(null);
  const [submitting, setSubmitting] = useState(false);
  const [successMsg, setSuccessMsg] = useState<string | null>(null);
  const [notes, setNotes] = useState("");

  useEffect(() => {
    if (!productGid) { setLoading(false); return; }
    async function init() {
      try {
        const result = await query<{
          shop: { myshopifyDomain: string };
          product: {
            title: string; vendor: string;
            featuredImage: { url: string } | null;
            variants: { nodes: Array<{ id: string; title: string; sku: string; inventoryQuantity: number }> };
          };
        }>(`{
          shop { myshopifyDomain }
          product(id: "${productGid}") {
            title vendor
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
          setProductVendor(p.vendor ?? "");
          setProductImageUrl(p.featuredImage?.url ?? null);
          setVariants((p.variants?.nodes ?? []).map((v) => ({
            id: v.id, title: v.title, sku: v.sku ?? "",
            stockQty: v.inventoryQuantity ?? 0, onOrderQty: 0, qtyOrdered: "",
          })));
        }
        const token = await auth.idToken();
        if (!token) throw new Error("No auth token");
        const res = await fetch(
          `${APP_URL}/api/order-status?productId=${encodeURIComponent(productGid!)}&shop=${encodeURIComponent(shopDomain)}`,
          { headers: { Authorization: `Bearer ${token}` } },
        );
        const json = await res.json();
        if (res.ok) {
          const nextOrder = json.order ?? null;
          setOrder(nextOrder);
          const onOrderByVariant = new Map(
            (nextOrder?.lines ?? []).map((line: OrderLineStatus) => [line.variantId, line.qtyOrdered]),
          );
          setVariants((prev) => prev.map((variant) => ({
            ...variant,
            onOrderQty: onOrderByVariant.get(variant.id) ?? 0,
          })));
        }
      } catch (e: any) {
        setErrorMsg(`Error: ${e?.message ?? String(e)}`);
      } finally {
        setLoading(false);
      }
    }
    init();
  }, [productGid]);

  const updateQty = useCallback((idx: number, val: string) => {
    setVariants((prev) => prev.map((v, i) => i === idx ? { ...v, qtyOrdered: val.replace(/\D/g, "") } : v));
  }, []);

  async function handleSubmit() {
    const orderedLines = variants.filter((v) => Number(v.qtyOrdered) > 0);
    if (!orderedLines.length) { setFormError("Enter a quantity for at least one variant"); return; }
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
          productImageUrl: productImageUrl || undefined,
          supplier: productVendor || "Unknown",
          notes: notes.trim() || undefined,
          lines: orderedLines.map((v) => ({
            variantId: v.id, variantTitle: v.title,
            sku: v.sku || undefined, qtyOrdered: Number(v.qtyOrdered),
          })),
        }),
      });
      const json = await res.json();
      if (!res.ok) { setFormError(`Error ${res.status}: ${json.error ?? "unknown"}`); return; }
      const total = orderedLines.reduce((s, v) => s + Number(v.qtyOrdered), 0);
      setSuccessMsg(`Order placed — ${total} units`);
      const nextLines = orderedLines.map((v) => ({
        variantId: v.id,
        qtyOrdered: Number(v.qtyOrdered),
      }));
      setOrder({
        id: json.order.id,
        supplier: productVendor,
        totalQty: (order?.totalQty ?? 0) + total,
        eta: order?.eta ?? null,
        lines: nextLines,
      });
      setShowForm(false);
      setNotes("");
      setVariants((prev) => prev.map((v) => ({
        ...v,
        onOrderQty: v.onOrderQty + (nextLines.find((line) => line.variantId === v.id)?.qtyOrdered ?? 0),
        qtyOrdered: "",
      })));
    } catch (e: any) {
      setFormError(`Error: ${e?.message ?? String(e)}`);
    } finally {
      setSubmitting(false);
    }
  }

  if (loading) return <InlineStack inlineAlignment="center"><ProgressIndicator size="small-200" /></InlineStack>;
  if (errorMsg) return <Banner status="warning">{errorMsg}</Banner>;

  const statusRow = (
    <InlineStack gap="small" blockAlignment="center">
      {order ? (
        <>
          <Badge tone="success">On order</Badge>
          <Text>{order.totalQty} units · {order.supplier}{order.eta ? ` · ETA ${new Date(order.eta).toLocaleDateString("en-AU", { day: "numeric", month: "short", year: "numeric" })}` : ""}</Text>
        </>
      ) : (
        <Badge>Not on order</Badge>
      )}
    </InlineStack>
  );

  if (!showForm) {
    return (
      <BlockStack gap="base">
        {successMsg && <Banner status="success">{successMsg}</Banner>}
        {statusRow}
        <Button onPress={() => setShowForm(true)}>{order ? "Add order" : "Place order"}</Button>
      </BlockStack>
    );
  }

  const totalStock = variants.reduce((s, v) => s + v.stockQty, 0);
  const totalOnOrder = variants.reduce((s, v) => s + v.onOrderQty, 0);
  const totalAddOrder = variants.reduce((s, v) => s + (Number(v.qtyOrdered) || 0), 0);

  return (
    <BlockStack gap="small">
      {statusRow}
      <Divider />
      {formError && <Banner status="critical">{formError}</Banner>}

      {/* Header */}
      <InlineStack gap={W.gap} blockAlignment="center">
        <Col w={W.size}><Text tone="subdued" fontWeight="bold">Size</Text></Col>
        <Col w={W.stock} align="center"><Text tone="subdued" fontWeight="bold">In stock</Text></Col>
        <Col w={W.onOrder} align="center"><Text tone="subdued" fontWeight="bold">On order</Text></Col>
        <Col w={W.addOrder} align="center"><Text tone="subdued" fontWeight="bold">Add order</Text></Col>
        <Col w={W.total} align="center"><Text tone="subdued" fontWeight="bold">Total</Text></Col>
      </InlineStack>
      <Divider />

      {/* Rows */}
      {variants.map((v, idx) => {
        const rowTotal = v.stockQty + v.onOrderQty + (Number(v.qtyOrdered) || 0);
        return (
          <InlineStack key={v.id} gap={W.gap} blockAlignment="center">
            <Col w={W.size}><Text>{v.title}</Text></Col>
            <Col w={W.stock} align="center"><Text>{v.stockQty}</Text></Col>
            <Col w={W.onOrder} align="center"><Text>{v.onOrderQty}</Text></Col>
            <Col w={W.addOrder} align="center">
              <TextField
                label=" "
                value={v.qtyOrdered}
                onChange={(val) => updateQty(idx, val)}
              />
            </Col>
            <Col w={W.total} align="center"><Text fontWeight="bold">{rowTotal}</Text></Col>
          </InlineStack>
        );
      })}

      <Divider />

      {/* Totals */}
      <InlineStack gap={W.gap} blockAlignment="center">
        <Col w={W.size}><Text fontWeight="bold">Total</Text></Col>
        <Col w={W.stock} align="center"><Text fontWeight="bold">{totalStock}</Text></Col>
        <Col w={W.onOrder} align="center"><Text fontWeight="bold">{totalOnOrder}</Text></Col>
        <Col w={W.addOrder} align="center"><Text fontWeight="bold">{totalAddOrder}</Text></Col>
        <Col w={W.total} align="center"><Text fontWeight="bold">{totalStock + totalOnOrder + totalAddOrder}</Text></Col>
      </InlineStack>

      <Divider />
      <TextField
        label="Notes for supplier portal"
        value={notes}
        onChange={setNotes}
      />
      <Divider />
      <InlineStack gap="base">
        <Button variant="primary" onPress={handleSubmit} loading={submitting}>{order ? "Add Order" : "Place Order"}</Button>
        <Button variant="plain" onPress={() => { setShowForm(false); setFormError(null); }}>Cancel</Button>
      </InlineStack>
    </BlockStack>
  );
}
