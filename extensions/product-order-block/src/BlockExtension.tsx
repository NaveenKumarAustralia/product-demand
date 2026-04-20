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
  const [error, setError] = useState(false);

  useEffect(() => {
    if (!productGid) {
      setLoading(false);
      return;
    }

    async function fetchStatus() {
      try {
        const token = await sessionToken.get();
        const res = await fetch(
          `${APP_URL}/api/order-status?productId=${encodeURIComponent(productGid!)}`,
          {
            headers: {
              Authorization: `Bearer ${token}`,
              "Content-Type": "application/json",
            },
          },
        );
        if (!res.ok) throw new Error("Failed to fetch");
        const json = await res.json();
        setOrder(json.order ?? null);
      } catch {
        setError(true);
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

  if (error) {
    return (
      <BlockStack gap="base">
        <Text fontWeight="bold">Supplier Ordering</Text>
        <Banner status="warning">
          Could not load order status. Check app is running.
        </Banner>
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
