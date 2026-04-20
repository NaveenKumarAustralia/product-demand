import type { HeadersFunction, LoaderFunctionArgs } from "react-router";
import { useLoaderData } from "react-router";
import { boundary } from "@shopify/shopify-app-react-router/server";
import { authenticate } from "../shopify.server";

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

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const { admin } = await authenticate.admin(request);

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
            featuredImage {
              url
              altText
            }
            variants(first: 10) {
              edges {
                node {
                  id
                  title
                  sku
                  inventoryQuantity
                }
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

  return { products };
};

function totalStock(variants: Variant[]) {
  return variants.reduce((sum, v) => sum + (v.inventoryQuantity ?? 0), 0);
}

function StockBadge({ stock }: { stock: number }) {
  if (stock === 0) {
    return (
      <span style={styles.badge("critical")}>Out of stock</span>
    );
  }
  if (stock < 10) {
    return (
      <span style={styles.badge("warning")}>Low: {stock}</span>
    );
  }
  return (
    <span style={styles.badge("success")}>In stock: {stock}</span>
  );
}

export default function ProductList() {
  const { products } = useLoaderData<typeof loader>();

  return (
    <s-page heading="Product Demand">
      <s-section heading={`${products.length} Products`}>
        <div style={styles.tableWrapper}>
          {/* Header row */}
          <div style={styles.headerRow}>
            <div style={{ ...styles.col, width: 56 }} />
            <div style={{ ...styles.col, flex: 3 }}>
              <s-text>Product</s-text>
            </div>
            <div style={{ ...styles.col, flex: 2 }}>
              <s-text>Supplier</s-text>
            </div>
            <div style={{ ...styles.col, flex: 2 }}>
              <s-text>Stock</s-text>
            </div>
            <div style={{ ...styles.col, flex: 2 }}>
              <s-text>Order status</s-text>
            </div>
            <div style={{ ...styles.col, flex: 1 }} />
          </div>

          {/* Product rows */}
          {products.map((product) => {
            const stock = totalStock(product.variants);
            const variantSummary = product.variants
              .map((v) =>
                v.title !== "Default Title"
                  ? `${v.title}: ${v.inventoryQuantity ?? 0}`
                  : null,
              )
              .filter(Boolean)
              .join(" · ");

            return (
              <div key={product.id} style={styles.productRow}>
                {/* Image */}
                <div style={{ ...styles.col, width: 56 }}>
                  {product.featuredImage ? (
                    <img
                      src={product.featuredImage.url}
                      alt={product.featuredImage.altText || product.title}
                      style={styles.thumbnail}
                    />
                  ) : (
                    <div style={styles.noImage} />
                  )}
                </div>

                {/* Title + variants */}
                <div style={{ ...styles.col, flex: 3 }}>
                  <s-text>{product.title}</s-text>
                  {variantSummary && (
                    <div style={{ marginTop: 4 }}>
                      <s-text>{variantSummary}</s-text>
                    </div>
                  )}
                </div>

                {/* Supplier */}
                <div style={{ ...styles.col, flex: 2 }}>
                  <s-text>{product.vendor || "—"}</s-text>
                </div>

                {/* Stock */}
                <div style={{ ...styles.col, flex: 2 }}>
                  <StockBadge stock={stock} />
                </div>

                {/* Order status */}
                <div style={{ ...styles.col, flex: 2 }}>
                  <span style={styles.badge("neutral")}>Not on order</span>
                </div>

                {/* Action */}
                <div style={{ ...styles.col, flex: 1 }}>
                  <s-button variant="primary">Place order</s-button>
                </div>
              </div>
            );
          })}
        </div>
      </s-section>
    </s-page>
  );
}

const styles = {
  tableWrapper: {
    display: "flex",
    flexDirection: "column" as const,
    gap: 1,
  },
  headerRow: {
    display: "flex",
    alignItems: "center",
    padding: "8px 12px",
    gap: 12,
    borderBottom: "1px solid #e1e3e5",
    opacity: 0.6,
  },
  productRow: {
    display: "flex",
    alignItems: "center",
    padding: "12px",
    gap: 12,
    borderBottom: "1px solid #f1f2f3",
    borderRadius: 8,
  },
  col: {
    display: "flex",
    flexDirection: "column" as const,
    justifyContent: "center" as const,
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
  badge: (tone: "success" | "warning" | "critical" | "neutral") => {
    const colors = {
      success: { bg: "#d3f5e2", color: "#1a6640" },
      warning: { bg: "#fff3cd", color: "#7d5c00" },
      critical: { bg: "#fde8e8", color: "#8c1515" },
      neutral: { bg: "#f1f2f3", color: "#4a4a4a" },
    };
    return {
      display: "inline-block",
      padding: "3px 10px",
      borderRadius: 12,
      fontSize: 12,
      fontWeight: 500,
      background: colors[tone].bg,
      color: colors[tone].color,
    };
  },
};

export const headers: HeadersFunction = (headersArgs) => {
  return boundary.headers(headersArgs);
};
