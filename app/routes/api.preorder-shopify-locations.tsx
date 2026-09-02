import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

const API_VERSION = "2026-07";

type LocationNode = {
  id: string;
  name: string;
  isActive: boolean;
  fulfillsOnlineOrders: boolean;
  hasActiveInventory: boolean;
  address?: { city?: string | null; country?: string | null; countryCode?: string | null } | null;
};

async function loadLocations(shop: string, accessToken: string) {
  const response = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
    method: "POST",
    signal: AbortSignal.timeout(8000),
    headers: {
      "Content-Type": "application/json",
      "X-Shopify-Access-Token": accessToken,
    },
    body: JSON.stringify({
      query: `#graphql
        query PreorderLocations {
          locations(first: 100, includeInactive: false) {
            nodes {
              id
              name
              isActive
              fulfillsOnlineOrders
              hasActiveInventory
              address { city country countryCode }
            }
          }
        }
      `,
    }),
  });
  if (!response.ok) throw new Error(`Shopify returned HTTP ${response.status}.`);
  const json = await response.json() as {
    data?: { locations?: { nodes?: LocationNode[] } };
    errors?: Array<{ message?: string }>;
  };
  if (json.errors?.length) throw new Error(json.errors.map((error) => error.message || "Shopify GraphQL error").join("; "));
  return json.data?.locations?.nodes ?? [];
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  try {
    const actor = await requirePreorderPortalUser(request);
    if (actor.admin !== true) {
      return Response.json({ ok: false, error: "Only a portal admin can view Shopify locations." }, { status: 403 });
    }

    const shopRows = await prisma.supplierOrder.findMany({ select: { shop: true }, distinct: ["shop"] });
    const shops = Array.from(new Set(shopRows.map((row) => row.shop).filter(Boolean)));
    const result = [];

    for (const shop of shops) {
      const session = await prisma.session.findFirst({
        where: { shop, isOnline: false },
        orderBy: { expires: "desc" },
        select: { accessToken: true },
      });
      if (!session?.accessToken) {
        result.push({ shop, ok: false, error: "Offline Shopify session missing.", locations: [] });
        continue;
      }
      try {
        const locations = await loadLocations(shop, session.accessToken);
        result.push({
          shop,
          ok: true,
          error: null,
          locations: locations
            .filter((location) => location.isActive)
            .sort((a, b) => a.name.localeCompare(b.name)),
        });
      } catch (error) {
        result.push({ shop, ok: false, error: error instanceof Error ? error.message : "Could not query Shopify locations.", locations: [] });
      }
    }

    return Response.json({ ok: true, shops: result }, { headers: { "Cache-Control": "private, max-age=30" } });
  } catch (error) {
    if (error instanceof Response) return error;
    console.error("[preorder Shopify locations] failed:", error);
    return Response.json({ ok: false, error: "Could not load Shopify locations." }, { status: 500 });
  }
};
