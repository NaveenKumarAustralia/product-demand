import prisma from "../db.server";

// Read-only customer-account view: the logged-in customer's preorder
// reservations + expected ship dates. Served via the signed app proxy, so the
// customer id is validated by Shopify. This NEVER writes and never exposes
// another customer's data — it matches only the resolved logged-in customer.
const API_VERSION = "2025-10";

export type CustomerPreorderLine = {
  orderName: string | null;
  variantTitle: string | null;
  sku: string | null;
  quantity: number;
  market: string;
  status: string;
  expectedShipDate: string | null;
  reservedAt: string;
};

async function offlineSessionToken(shop: string): Promise<string | null> {
  const session = await prisma.session.findFirst({
    where: { shop, isOnline: false, accessToken: { not: "" } },
    orderBy: { expires: "desc" },
    select: { accessToken: true },
  }).catch(() => null);
  return session?.accessToken ?? null;
}

async function resolveCustomerEmail(shop: string, token: string, customerId: string): Promise<string | null> {
  const numeric = customerId.replace(/[^0-9]/g, "");
  if (!numeric) return null;
  try {
    const res = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
      method: "POST",
      signal: AbortSignal.timeout(8000),
      headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": token },
      body: JSON.stringify({
        query: "#graphql\n query CustomerEmail($id: ID!) { customer(id: $id) { email } }",
        variables: { id: `gid://shopify/Customer/${numeric}` },
      }),
    });
    if (!res.ok) return null;
    const json = await res.json() as { data?: { customer?: { email?: string | null } } };
    const email = json.data?.customer?.email?.trim().toLowerCase();
    return email || null;
  } catch {
    return null;
  }
}

export async function getCustomerPreorders(input: { shop: string; customerId: string | null }): Promise<{ ok: true; preorders: CustomerPreorderLine[] }> {
  const { shop, customerId } = input;
  // Not logged in / unresolvable → empty, never an error (public surface).
  if (!shop || !customerId) return { ok: true, preorders: [] };
  const token = await offlineSessionToken(shop);
  if (!token) return { ok: true, preorders: [] };
  const email = await resolveCustomerEmail(shop, token, customerId);
  if (!email) return { ok: true, preorders: [] };

  const rows = await prisma.preorderReservation.findMany({
    where: { shop, customerEmail: { equals: email, mode: "insensitive" }, status: { not: "released" } },
    orderBy: [{ expectedShipDate: "asc" }, { reservedAt: "desc" }],
    select: {
      shopifyOrderName: true, variantTitle: true, sku: true, quantity: true,
      market: true, status: true, expectedShipDate: true, reservedAt: true,
    },
  }).catch(() => [] as Array<{
    shopifyOrderName: string | null; variantTitle: string | null; sku: string | null;
    quantity: number; market: string; status: string; expectedShipDate: Date | null; reservedAt: Date;
  }>);

  return {
    ok: true,
    preorders: rows.map((r) => ({
      orderName: r.shopifyOrderName ?? null,
      variantTitle: r.variantTitle ?? null,
      sku: r.sku ?? null,
      quantity: r.quantity,
      market: r.market,
      status: r.status,
      expectedShipDate: r.expectedShipDate ? r.expectedShipDate.toISOString() : null,
      reservedAt: r.reservedAt.toISOString(),
    })),
  };
}
