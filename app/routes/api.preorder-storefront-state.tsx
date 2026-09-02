import type { LoaderFunctionArgs } from "react-router";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { getStorefrontPreorderState } from "../preorder/preorder-storefront-state.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  try {
    await requirePreorderPortalUser(request);
    const url = new URL(request.url);
    const shop = String(url.searchParams.get("shop") ?? "").trim();
    const variantId = String(url.searchParams.get("variantId") ?? "").trim();
    const marketRaw = String(url.searchParams.get("market") ?? "AU").trim().toUpperCase();
    const market = marketRaw === "USA" ? "USA" : marketRaw === "AU" ? "AU" : null;
    if (!shop || !variantId || !market) {
      return Response.json({ ok: false, error: "shop, variantId and market=AU|USA are required." }, { status: 400 });
    }

    const result = await getStorefrontPreorderState({ shop, variantId, market });
    return Response.json(result, {
      headers: { "Cache-Control": "private, max-age=5" },
    });
  } catch (error) {
    if (error instanceof Response) return error;
    console.error("[preorder storefront state] failed:", error);
    return Response.json({ ok: false, error: error instanceof Error ? error.message : "Could not calculate storefront state." }, { status: 500 });
  }
};
