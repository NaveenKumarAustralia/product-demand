import { useEffect, useState, useCallback } from "react";
import {
  reactExtension,
  useApi,
  BlockStack,
  InlineStack,
  Text,
  Button,
  Banner,
  Divider,
  Badge,
  ProgressIndicator,
  TextField,
  NumberField,
  Select,
} from "@shopify/ui-extensions-react/admin";

const TARGET = "admin.product-details.block.render";
const APP_URL = "https://product-demand-production.up.railway.app";

export default reactExtension(TARGET, () => <ProductOrderBlock />);

type OrderStatus = {
  id: number;
  supplier: string;
  totalQty: number;
  eta: string | null;
  status: string;
} | null;

type Variant = {
  id: string;
  title: string;
  sku: string;
  qtyOrdered: number;
};

function ProductOrderBlock() {
  const { data, auth, query } = useApi(TARGET);
  const productGid: string | undefined = data.selected?.[0]?.id;

  const [shop, setShop] = useState<string | null>(null);
  const [productTitle, setProductTitle] = useState("");
  const [order, setOrder] = useState<OrderStatus>(null);
  const [loading, setLoading] = useState(true);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  const [variants, setVariants] = useState<Variant[]>([]);
  const [supplierOptions, setSupplierOptions] = useState<Array<{ label: string; value: string }>>([]);
  const [supplier, setSupplier] = useState("");
  const [customSupplier, setCustomSupplier] = useState("");
  const [eta, setEta] = useState("");
  const [notes, setNotes] = useState("");
  const [formError, setFormError] = useState<string | null>(null);
  const [submitting, setSubmitting] = useState(false);
  const [successMsg, setSuccessMsg] = useState<string | null>(null);

  useEffect(() => {
    if (!productGid) {
      setLoading(false);
      setErrorMsg("No product ID in extension data");
      return;
    }

    async function init() {
      try {
        const result = await query<{
          shop: { myshopifyDomain: string };
          product: {
            title: string;
            variants: { nodes: Array<{ id: string; title: string; sku: string }> };
          };
        }>(`{
          shop { myshopifyDomain }
          product(id: "${productGid}") {
            title
            variants(first: 50) { nodes { id title sku } }
          }
        }`);

        const shopDomain = result.data?.shop?.myshopifyDomain;
        if (!shopDomain) throw new Error("Could not resolve shop domain");
        setShop(shopDomain);

        const p = result.data?.product;
        if (p) {
          setProductTitle(p.title);
          setVariants(
            (p.variants?.nodes ?? []).map((v) => ({
              id: v.id,
              title: v.title,
              sku: v.sku ?? "",
              qtyOrdered: 0,
            })),
          );
        }

        const token = await auth.idToken();
        if (!token) throw new Error("No auth token");

        // Fetch existing suppliers and order status in parallel
        const [statusRes, suppliersRes] = await Promise.all([
          fetch(
            `${APP_URL}/api/order-status?productId=${encodeURIComponent(productGid!)}&shop=${encodeURIComponent(shopDomain)}`,
            { headers: { Authorization: `Bearer ${token}` } },
          ),
          fetch(
            `${APP_URL}/api/suppliers?shop=${encodeURIComponent(shopDomain)}`,
            { headers: { Authorization: `Bearer ${token}` } },
          ),
        ]);

        const statusJson = await statusRes.json();
        if (statusRes.ok) setOrder(statusJson.order ?? null);

        const suppliersJson = await suppliersRes.json();
        const names: string[] = suppliersJson.suppliers ?? [];
        const opts = names.map((s) => ({ label: s, value: s }));
        opts.push({ label: "Add new supplier…", value: "__new__" });
        setSupplierOptions(opts);
        if (names.length > 0) setSupplier(names[0]);
        else setSupplier("__new__");
      } catch (e: any) {
        setErrorMsg(`Error: ${e?.message ?? String(e)}`);
      } finally {
        setLoading(false);
      }
    }

    init();
  }, [productGid]);

  const updateQty = useCallback((idx: number, val: number) => {
    setVariants((prev) => prev.map((v, i) => (i === idx ? { ...v, qtyOrdered: val } : v)));
  }, []);

  async function handleSubmit() {
    const resolvedSupplier = supplier === "__new__" ? customSupplier.trim() : supplier;
    if (!resolvedSupplier) {
      setFormError("Supplier name is required");
      return;
    }
    const orderedLines = variants.filter((v) => Number(v.qtyOrdered) > 0);
    if (!orderedLines.length) {
      setFormError("Enter a quantity for at least one variant");
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
          shop,
          productId: productGid,
          productTitle,
          supplier: resolvedSupplier,
          eta: eta || undefined,
          notes: notes.trim() || undefined,
          lines: orderedLines.map((v) => ({
            variantId: v.id,
            variantTitle: v.title,
            sku: v.sku || undefined,
            qtyOrdered: Number(v.qtyOrdered),
          })),
        }),
      });

      const json = await res.json();
      if (!res.ok) {
        setFormError(`Error ${res.status}: ${json.error ?? "unknown"}`);
        return;
      }

      const totalOrdered = orderedLines.reduce((s, v) => s + Number(v.qtyOrdered), 0);
      setSuccessMsg(`Order placed — ${totalOrdered} units from ${resolvedSupplier}`);
      setOrder({ id: json.order.id, supplier: resolvedSupplier, totalQty: totalOrdered, eta: eta || null, status: "open" });

      // Add new supplier to the dropdown
      if (supplier === "__new__" && resolvedSupplier) {
        setSupplierOptions((prev) => [
          { label: resolvedSupplier, value: resolvedSupplier },
          ...prev.filter((o) => o.value !== "__new__"),
          { label: "Add new supplier…", value: "__new__" },
        ]);
        setSupplier(resolvedSupplier);
        setCustomSupplier("");
      }

      // Reset quantities
      setVariants((prev) => prev.map((v) => ({ ...v, qtyOrdered: 0 })));
      setEta("");
      setNotes("");
    } catch (e: any) {
      setFormError(`Submit error: ${e?.message ?? String(e)}`);
    } finally {
      setSubmitting(false);
    }
  }

  if (loading) {
    return (
      <BlockStack gap="base">
        <Text fontWeight="bold">Supplier Ordering</Text>
        <InlineStack inlineAlignment="center">
          <ProgressIndicator size="small-200" />
        </InlineStack>
      </BlockStack>
    );
  }

  if (errorMsg) {
    return (
      <BlockStack gap="base">
        <Text fontWeight="bold">Supplier Ordering</Text>
        <Banner status="warning">{errorMsg}</Banner>
      </BlockStack>
    );
  }

  return (
    <BlockStack gap="base">
      <Text fontWeight="bold">Supplier Ordering</Text>
      <Divider />

      {/* Current order status */}
      {successMsg ? (
        <Banner status="success">{successMsg}</Banner>
      ) : order ? (
        <BlockStack gap="extraSmall">
          <InlineStack gap="small" blockAlignment="center">
            <Badge tone="info">On order</Badge>
            <Text>{order.totalQty} units · {order.supplier}</Text>
          </InlineStack>
          {order.eta && (
            <Text tone="subdued">
              ETA {new Date(order.eta).toLocaleDateString("en-AU", { day: "numeric", month: "short", year: "numeric" })}
            </Text>
          )}
        </BlockStack>
      ) : (
        <InlineStack gap="small" blockAlignment="center">
          <Badge>Not on order</Badge>
          <Text tone="subdued">No open supplier orders</Text>
        </InlineStack>
      )}

      <Divider />
      <Text fontWeight="bold">{order ? "Place another order" : "Place order"}</Text>

      {formError && <Banner status="critical">{formError}</Banner>}

      {supplierOptions.length > 0 ? (
        <Select
          label="Supplier"
          options={supplierOptions}
          value={supplier}
          onChange={setSupplier}
        />
      ) : (
        <TextField label="Supplier" value={customSupplier} onChange={setCustomSupplier} />
      )}

      {supplier === "__new__" && supplierOptions.length > 0 && (
        <TextField label="New supplier name" value={customSupplier} onChange={setCustomSupplier} />
      )}

      <TextField label="ETA (YYYY-MM-DD)" value={eta} onChange={setEta} />
      <TextField label="Notes" value={notes} onChange={setNotes} />

      <Divider />

      {variants.map((v, idx) => (
        <InlineStack key={v.id} gap="base" blockAlignment="center">
          <BlockStack gap="none">
            <Text>{v.title}</Text>
            {v.sku ? <Text tone="subdued">{v.sku}</Text> : null}
          </BlockStack>
          <NumberField
            label="Qty"
            value={v.qtyOrdered}
            min={0}
            onChange={(val) => updateQty(idx, val)}
          />
        </InlineStack>
      ))}

      <Divider />
      <Button variant="primary" onPress={handleSubmit} loading={submitting}>
        Place Order
      </Button>
    </BlockStack>
  );
}
