import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { getStorefrontPreorderState } from "../preorder/preorder-storefront-state.server";
import { getCustomerPreorders } from "../preorder/preorder-customer-account.server";
import { renderMyPreordersPage } from "../preorder/preorder-customer-account-page";
import { authenticate } from "../shopify.server";

function marketFrom(value: unknown) {
  const market = String(value ?? "AU").trim().toUpperCase();
  return market === "USA" ? "USA" : market === "AU" ? "AU" : null;
}

function text(value: unknown, max = 250) {
  return String(value ?? "").trim().slice(0, max);
}

function normalizeEmail(value: unknown) {
  const email = text(value, 320).toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return null;
  return email;
}

function proxyShop(context: Awaited<ReturnType<typeof authenticate.public.appProxy>>, request: Request) {
  const url = new URL(request.url);
  return String(context.session?.shop || url.searchParams.get("shop") || "").trim();
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  try {
    const context = await authenticate.public.appProxy(request);
    const url = new URL(request.url);
    const shop = proxyShop(context, request);

    // Customer-account "My preorders" — Shopify signs the request (including the
    // logged_in_customer_id it injects), so the customer is trusted. Read-only.
    const customerId = url.searchParams.get("logged_in_customer_id");
    if (url.searchParams.get("op") === "my-preorders") {
      const result = await getCustomerPreorders({ shop, customerId });
      return Response.json(result, {
        headers: { "Cache-Control": "no-store", "Content-Type": "application/json; charset=utf-8" },
      });
    }

    // Themed "My pre-orders" PAGE — visit /apps/karma-east-preorder?view=my-preorders
    // (link it from the storefront account menu). Returned as application/liquid so
    // Shopify renders it inside the store's theme (header/footer/fonts) — native
    // look, no theme editing required.
    if (url.searchParams.get("view") === "my-preorders") {
      const result = await getCustomerPreorders({ shop, customerId });
      const html = renderMyPreordersPage(result.preorders, Boolean(customerId));
      return new Response(html, {
        headers: { "Content-Type": "application/liquid", "Cache-Control": "no-store" },
      });
    }

    const variantId = text(url.searchParams.get("variantId"), 120);
    const market = marketFrom(url.searchParams.get("market"));

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

export const action = async ({ request }: ActionFunctionArgs) => {
  try {
    if (request.method !== "POST") {
      return Response.json({ ok: false, error: "Method not allowed." }, { status: 405 });
    }

    const context = await authenticate.public.appProxy(request);
    const shop = proxyShop(context, request);
    if (!shop) return Response.json({ ok: false, error: "Shop could not be verified." }, { status: 401 });

    const contentType = request.headers.get("content-type") || "";
    let body: Record<string, unknown> = {};
    if (contentType.includes("application/json")) {
      body = await request.json().catch(() => ({})) as Record<string, unknown>;
    } else {
      const form = await request.formData();
      body = Object.fromEntries(form.entries());
    }

    const email = normalizeEmail(body.email);
    const variantId = text(body.variantId, 120);
    const market = marketFrom(body.market);
    if (!email || !variantId || !market) {
      return Response.json({ ok: false, error: "A valid email, variant and market are required." }, { status: 400 });
    }

    const productId = text(body.productId, 120) || null;
    const productTitle = text(body.productTitle, 250) || null;
    const variantTitle = text(body.variantTitle, 250) || null;
    const sku = text(body.sku, 120) || null;

    const rows = await prisma.$queryRawUnsafe<Array<{ id: number }>>(
      `INSERT INTO "PreorderWaitlist"
        ("shop", "email", "productId", "productTitle", "variantId", "variantTitle", "sku", "market", "status", "source", "createdAt", "updatedAt")
       VALUES ($1, $2, $3, $4, $5, $6, $7, $8, 'waiting', 'storefront', CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
       ON CONFLICT ("shop", "email", "variantId", "market")
       DO UPDATE SET
         "productId" = COALESCE(EXCLUDED."productId", "PreorderWaitlist"."productId"),
         "productTitle" = COALESCE(EXCLUDED."productTitle", "PreorderWaitlist"."productTitle"),
         "variantTitle" = COALESCE(EXCLUDED."variantTitle", "PreorderWaitlist"."variantTitle"),
         "sku" = COALESCE(EXCLUDED."sku", "PreorderWaitlist"."sku"),
         "status" = 'waiting',
         "notifiedAt" = NULL,
         "convertedOrderId" = NULL,
         "source" = 'storefront',
         "updatedAt" = CURRENT_TIMESTAMP
       RETURNING "id"`,
      shop,
      email,
      productId,
      productTitle,
      variantId,
      variantTitle,
      sku,
      market,
    );

    return Response.json(
      { ok: true, waitlistId: rows[0]?.id ?? null },
      { headers: { "Cache-Control": "no-store" } },
    );
  } catch (error) {
    if (error instanceof Response) return error;
    console.error("[preorder app proxy waitlist] failed:", error);
    return Response.json(
      { ok: false, error: "Could not join the notification list." },
      { status: 500, headers: { "Cache-Control": "no-store" } },
    );
  }
};
