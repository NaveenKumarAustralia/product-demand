import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// Per-variant, per-country units sold for one product over a date range, for the
// Reorder Planner's expanded breakdown. Uses ShopifyQL (same as the dashboard),
// fetched on demand when a product row is expanded so nothing hangs.
// Response: { ok, rows: [{ variant, country, units }] }.
async function runShopifyQL(shop: string, token: string, shopifyql: string): Promise<Array<Record<string, string>>> {
  const graphql = `{ shopifyqlQuery(query: ${JSON.stringify(shopifyql)}) { tableData { columns { name } rows } parseErrors } }`;
  const res = await fetch(`https://${shop}/admin/api/unstable/graphql.json`, {
    method: "POST",
    headers: { "X-Shopify-Access-Token": token, "Content-Type": "application/json" },
    body: JSON.stringify({ query: graphql }),
  });
  if (!res.ok) throw new Error(`ShopifyQL HTTP ${res.status}`);
  const json = await res.json() as { data?: { shopifyqlQuery?: { tableData?: { columns?: Array<{ name: string }>; rows?: unknown[] }; parseErrors?: unknown[] } } };
  const result = json.data?.shopifyqlQuery;
  if (result?.parseErrors?.length) throw new Error(`ShopifyQL parse error: ${JSON.stringify(result.parseErrors)}`);
  const cols = (result?.tableData?.columns ?? []).map((c) => c.name);
  const rows = result?.tableData?.rows ?? [];
  return rows.map((row) => {
    const obj: Record<string, string> = {};
    if (Array.isArray(row)) cols.forEach((name, i) => { obj[name] = String((row as unknown[])[i] ?? ""); });
    else Object.assign(obj, row as Record<string, string>);
    return obj;
  });
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const url = new URL(request.url);
  const productId = (url.searchParams.get("productId") ?? "").trim();
  const since = (url.searchParams.get("since") ?? "").trim();
  const until = (url.searchParams.get("until") ?? "").trim();
  if (!productId || !since || !until) return Response.json({ ok: false, rows: [] });

  const session = await prisma.session.findFirst({ where: { accessToken: { not: "" } }, orderBy: { isOnline: "asc" } }).catch(() => null);
  if (!session?.shop || !session.accessToken) return Response.json({ ok: false, rows: [], error: "no_session" });

  const numeric = productId.replace(/[^0-9]/g, "");
  if (!numeric) return Response.json({ ok: true, rows: [] });

  try {
    const rows = await runShopifyQL(
      session.shop, session.accessToken,
      `FROM sales SHOW net_items_sold WHERE product_id = ${numeric} GROUP BY product_variant_title, shipping_country SINCE ${since} UNTIL ${until}`,
    );
    const out = rows
      .map((r) => ({
        variant: (r.product_variant_title || "").trim(),
        country: (r.shipping_country || "Unknown").trim() || "Unknown",
        units: Math.max(0, parseInt(r.net_items_sold || "0", 10) || 0),
      }))
      .filter((r) => r.units > 0);
    return Response.json({ ok: true, rows: out });
  } catch (e) {
    console.warn("[reorder-country-sales]", e);
    return Response.json({ ok: false, rows: [], error: "unavailable" });
  }
};
