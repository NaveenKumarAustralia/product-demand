import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { authorizeApiRequest } from "../api-auth.server";

// Units sold for one product, broken down by shipping country, over a date
// range. Powers the "Sales by country" block on the Shopify product page.
// Uses ShopifyQL (same analytics engine as Shopify's own reports), so it needs
// the read_reports scope (this app has it).
const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type, X-Api-Key",
};

async function runShopifyQL(
  shop: string,
  token: string,
  shopifyql: string,
): Promise<Array<Record<string, string>>> {
  const graphql = `{ shopifyqlQuery(query: ${JSON.stringify(shopifyql)}) { tableData { columns { name } rows } parseErrors } }`;
  const res = await fetch(`https://${shop}/admin/api/unstable/graphql.json`, {
    method: "POST",
    headers: { "X-Shopify-Access-Token": token, "Content-Type": "application/json" },
    body: JSON.stringify({ query: graphql }),
  });
  if (!res.ok) throw new Error(`ShopifyQL HTTP ${res.status}`);
  const json = await res.json() as {
    data?: { shopifyqlQuery?: { tableData?: { columns?: Array<{ name: string }>; rows?: unknown[] }; parseErrors?: Array<unknown> } };
  };
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
  if (request.method === "OPTIONS") {
    return new Response(null, { status: 204, headers: CORS });
  }

  const url = new URL(request.url);
  const productId = (url.searchParams.get("productId") ?? "").trim();
  const shop = (url.searchParams.get("shop") ?? "").trim();
  const since = (url.searchParams.get("since") ?? "").trim();  // YYYY-MM-DD
  const until = (url.searchParams.get("until") ?? "").trim();  // YYYY-MM-DD
  if (!productId || !shop || !since || !until) {
    return Response.json({ error: "productId, shop, since and until are required" }, { status: 400, headers: CORS });
  }

  const unauthorized = authorizeApiRequest(request, CORS);
  if (unauthorized) return unauthorized;

  try {
    const session = await prisma.session.findFirst({
      where: { shop, accessToken: { not: "" } },
      orderBy: { isOnline: "asc" },
    }).catch(() => null);
    if (!session?.accessToken) {
      return Response.json({ error: "No Shopify session" }, { status: 200, headers: CORS });
    }

    const numericId = productId.replace(/[^0-9]/g, "");
    if (!numericId) return Response.json({ ok: true, total: 0, byCountry: [] }, { headers: CORS });

    const rows = await runShopifyQL(
      shop,
      session.accessToken,
      `FROM sales SHOW net_items_sold WHERE product_id = ${numericId} GROUP BY shipping_country SINCE ${since} UNTIL ${until}`,
    );
    const byCountry = rows
      .map((r) => ({ country: (r.shipping_country || "Unknown").trim() || "Unknown", unitsSold: Math.max(0, parseInt(r.net_items_sold || "0", 10) || 0) }))
      .filter((r) => r.unitsSold > 0)
      .sort((a, b) => b.unitsSold - a.unitsSold);
    const total = byCountry.reduce((sum, r) => sum + r.unitsSold, 0);
    return Response.json({ ok: true, total, byCountry }, { headers: CORS });
  } catch (err) {
    console.error("sold-by-country error:", err);
    return Response.json({ error: "Analytics unavailable" }, { status: 200, headers: CORS });
  }
};
