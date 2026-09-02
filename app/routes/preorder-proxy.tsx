import type { LoaderFunctionArgs } from "react-router";
import { getStorefrontPreorderState } from "../preorder/preorder-storefront-state.server";
import { authenticate } from "../shopify.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  try {
    const context = await authenticate.public.appProxy(request);
    const url = new URL(request.url);
    const shop = String(context.session?.shop || url.searchParams.get("shop") || "").trim();
    const variantId = String(url.searchParams.get("variantId") || "").trim();
    const marketRaw = String(url.searchParams.get("market") || "AU").trim().toUpperCase();
    const market = marketRaw === "USA" ? "USA" : marketRaw === "AU" ? "AU" : null;

    if (!shop || !variantId || !market) {
      return Response.json({ ok: false, error: "variantId and market=AU|USA are required." }, { status: 400 });
    }

    const result = await getStorefrontPreorderState({ shop, variantId, market });
    return Response.json(result, {
      headers: {
        "Cache-Control": "no-store",
        "Content-Type": "application/json; charset=utf-8",
      },
    });
  } catch (error) {
    if (error instanceof Response) return error;
    console.error("[preorder app proxy] failed:", error);
    return Response.json(
      { ok: false, error: "Could not calculate preorder availability." },
      { status: 500, headers: { "Cache-Control": "no-store" } },
    );
  }
};
