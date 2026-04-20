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
  const { data, auth, query, navigation } = useApi(TARGET);

  const productGid: string | undefined = data.selected?.[0]?.id;

  const [order, setOrder] = useState<OrderStatus>(null);
  const [loading, setLoading] = useState(true);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  useEffect(() => {
    if (!productGid) {
      setLoading(false);
      setErrorMsg("No product ID in extension data");
      return;
    }

    async function fetchStatus() {
      try {
        // Get shop domain via authenticated GraphQL query
        const shopResult = await query<{
          shop: { myshopifyDomain: string };
        }>(`{ shop { myshopifyDomain } }`);

        const shop = shopResult.data?.shop?.myshopifyDomain;
        if (!shop) throw new Error("Could not resolve shop domain");

        // Get ID token for backend authentication
        const token = await auth.idToken();
        if (!token) throw new Error("No auth token");

        const res = await fetch(
          `${APP_URL}/api/order-status?productId=${encodeURIComponent(productGid!)}&shop=${encodeURIComponent(shop)}`,
          {
            headers: {
              Authorization: `Bearer ${token}`,
              "Content-Type": "application/json",
            },
          },
        );

        const json = await res.json();
        if (!res.ok) {
          setErrorMsg(`API ${res.status}: ${json.error ?? "unknown"}`);
          return;
        }
        setOrder(json.order ?? null);
      } catch (e: any) {
        setErrorMsg(`Error: ${e?.message ?? String(e)}`);
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
            <Text>
              {order.totalQty} units · {order.supplier}
            </Text>
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
