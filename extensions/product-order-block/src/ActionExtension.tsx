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
  TextField,
  NumberField,
  ProgressIndicator,
} from "@shopify/ui-extensions-react/admin";

const TARGET = "admin.product-details.action.render";
const APP_URL = "https://product-demand-production.up.railway.app";

export default reactExtension(TARGET, () => <ProductOrderAction />);

type Variant = {
  id: string;
  title: string;
  sku: string;
  qtyOrdered: number;
  costPrice: string;
};

type ShopProduct = {
  id: string;
  title: string;
  variants: Array<{ id: string; title: string; sku: string }>;
};

function ProductOrderAction() {
  const { data, auth, query, close } = useApi(TARGET);
  const productGid: string | undefined = data.selected?.[0]?.id;

  const [shop, setShop] = useState<string | null>(null);
  const [product, setProduct] = useState<ShopProduct | null>(null);
  const [loading, setLoading] = useState(true);
  const [submitting, setSubmitting] = useState(false);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [successMsg, setSuccessMsg] = useState<string | null>(null);

  const [supplier, setSupplier] = useState("");
  const [poNumber, setPoNumber] = useState("");
  const [eta, setEta] = useState("");
  const [notes, setNotes] = useState("");
  const [variants, setVariants] = useState<Variant[]>([]);

  useEffect(() => {
    if (!productGid) {
      setLoading(false);
      setErrorMsg("No product selected");
      return;
    }

    async function init() {
      try {
        const result = await query<{
          shop: { myshopifyDomain: string };
          product: { id: string; title: string; variants: { nodes: Array<{ id: string; title: string; sku: string }> } };
        }>(`{
          shop { myshopifyDomain }
          product(id: "${productGid}") {
            id
            title
            variants(first: 50) {
              nodes { id title sku }
            }
          }
        }`);

        const shopDomain = result.data?.shop?.myshopifyDomain;
        if (!shopDomain) throw new Error("Could not resolve shop domain");
        setShop(shopDomain);

        const p = result.data?.product;
        if (!p) throw new Error("Could not load product");

        const variantNodes = p.variants?.nodes ?? [];
        setProduct({ id: p.id, title: p.title, variants: variantNodes });
        setVariants(
          variantNodes.map((v) => ({
            id: v.id,
            title: v.title,
            sku: v.sku ?? "",
            qtyOrdered: 0,
            costPrice: "",
          })),
        );
      } catch (e: any) {
        setErrorMsg(`Load error: ${e?.message ?? String(e)}`);
      } finally {
        setLoading(false);
      }
    }

    init();
  }, [productGid]);

  const updateVariant = useCallback(
    (idx: number, field: "qtyOrdered" | "costPrice", value: string | number) => {
      setVariants((prev) =>
        prev.map((v, i) => (i === idx ? { ...v, [field]: value } : v)),
      );
    },
    [],
  );

  async function handleSubmit() {
    if (!supplier.trim()) {
      setErrorMsg("Supplier name is required");
      return;
    }

    const orderedLines = variants.filter((v) => (v.qtyOrdered ?? 0) > 0);
    if (!orderedLines.length) {
      setErrorMsg("Enter a quantity for at least one variant");
      return;
    }

    setErrorMsg(null);
    setSubmitting(true);

    try {
      const token = await auth.idToken();
      if (!token) throw new Error("No auth token");

      const res = await fetch(`${APP_URL}/api/place-order`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          shop,
          productId: product!.id,
          productTitle: product!.title,
          supplier: supplier.trim(),
          poNumber: poNumber.trim() || undefined,
          eta: eta || undefined,
          notes: notes.trim() || undefined,
          lines: orderedLines.map((v) => ({
            variantId: v.id,
            variantTitle: v.title,
            sku: v.sku || undefined,
            qtyOrdered: Number(v.qtyOrdered),
            costPrice: v.costPrice ? Number(v.costPrice) : undefined,
          })),
        }),
      });

      const json = await res.json();
      if (!res.ok) {
        setErrorMsg(`Error ${res.status}: ${json.error ?? "unknown"}`);
        return;
      }

      setSuccessMsg(
        `Order placed! ${orderedLines.reduce((s, v) => s + Number(v.qtyOrdered), 0)} units from ${supplier.trim()}`,
      );
      setTimeout(() => close(), 2000);
    } catch (e: any) {
      setErrorMsg(`Submit error: ${e?.message ?? String(e)}`);
    } finally {
      setSubmitting(false);
    }
  }

  if (loading) {
    return (
      <BlockStack gap="base">
        <InlineStack inlineAlignment="center">
          <ProgressIndicator size="small-200" />
        </InlineStack>
      </BlockStack>
    );
  }

  if (successMsg) {
    return (
      <BlockStack gap="base">
        <Banner status="success">{successMsg}</Banner>
      </BlockStack>
    );
  }

  return (
    <BlockStack gap="base">
      {product && (
        <Text fontWeight="bold">{product.title}</Text>
      )}
      <Divider />

      {errorMsg && <Banner status="critical">{errorMsg}</Banner>}

      <TextField
        label="Supplier"
        value={supplier}
        onChange={setSupplier}
      />
      <InlineStack gap="base">
        <TextField
          label="PO Number"
          value={poNumber}
          onChange={setPoNumber}
        />
        <TextField
          label="ETA (YYYY-MM-DD)"
          value={eta}
          onChange={setEta}
        />
      </InlineStack>
      <TextField
        label="Notes"
        value={notes}
        onChange={setNotes}
      />

      <Divider />
      <Text fontWeight="bold">Variants</Text>

      {variants.map((v, idx) => (
        <BlockStack key={v.id} gap="extraSmall">
          <Text>
            {v.title}{v.sku ? ` · ${v.sku}` : ""}
          </Text>
          <InlineStack gap="base">
            <NumberField
              label="Qty to order"
              value={v.qtyOrdered}
              min={0}
              onChange={(val) => updateVariant(idx, "qtyOrdered", val)}
            />
            <TextField
              label="Cost price"
              value={v.costPrice}
              onChange={(val) => updateVariant(idx, "costPrice", val)}
            />
          </InlineStack>
        </BlockStack>
      ))}

      <Divider />

      <InlineStack inlineAlignment="end" gap="base">
        <Button onPress={() => close()} variant="plain">
          Cancel
        </Button>
        <Button onPress={handleSubmit} variant="primary" loading={submitting}>
          Place Order
        </Button>
      </InlineStack>
    </BlockStack>
  );
}
