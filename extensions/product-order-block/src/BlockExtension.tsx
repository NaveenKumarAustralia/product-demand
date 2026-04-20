import { useEffect, useState } from "react";
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

function ProductOrderBlock() {
  const { data, sessionToken, navigation } = useApi(TARGET);

  const productGid: string | undefined = data.selected?.[0]?.id;

  const [order, setOrder] = useState<OrderStatus>(null);
  const [loading, setLoading] = useState(true);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  useEffect(() => {
    if (!productGid) {
      setLoading(false);
      setErrorMsg(`No product ID found in extension data`);
      return;
    }

    async function fetchStatus() {
      try {
        const token = await sessionToken.get();
        const url = `${APP_URL}/api/order-status?productId=${encodeURIComponent(productGid!)}`;
        const res = await fetch(url, {
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json",
          },
        });
        const json = await res.json();
        if (!res.ok) {
          setErrorMsg(`API ${res.status}: ${json.error ?? "unknown"}`);
          return;
        }
        setOrder(json.order ?? null);
      } catch (e: any) {
        setErrorMsg(`Fetch error: ${e?.message ?? String(e)}`);
      } finally {
        setLoading(false);
      }
    }

    fetchStatus();
  }, [productGid]);

  function goToApp() {
    navigation.navigate("shopify://admin/apps/product-demand");
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

      {order ? (
        <BlockStack gap="small">
          <InlineStack gap="small" blockAlignment="center">
            <Badge tone="info">On order</Badge>
            <Text>{order.totalQty} units · {order.supplier}</Text>
          </InlineStack>
          {order.eta && (
            <Text tone="subdued">
              ETA{" "}
              {new Date(order.eta).toLocaleDateString("en-AU", {
                day: "numeric",
                month: "short",
                year: "numeric",
              })}
            </Text>
          )}
        </BlockStack>
      ) : (
        <InlineStack gap="small" blockAlignment="center">
          <Badge>Not on order</Badge>
          <Text tone="subdued">No open supplier orders</Text>
        </InlineStack>
      )}

      <Button onPress={goToApp}>
        {order ? "View / Reorder" : "Place order"}
      </Button>
    </BlockStack>
  );
}
