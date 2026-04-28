import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

async function hmacBase64(secret: string, payload: string): Promise<string> {
  const enc = new TextEncoder();
  const key = await crypto.subtle.importKey(
    "raw",
    enc.encode(secret),
    { name: "HMAC", hash: { name: "SHA-256" } },
    false,
    ["sign"],
  );
  const sig = await crypto.subtle.sign("HMAC", key, enc.encode(payload));
  // Same base64 encoder the shopify-api library uses
  const table = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
  let out = "";
  const bytes = new Uint8Array(sig);
  for (let i = 0; i < bytes.length; ) {
    const b1 = bytes[i++], b2 = bytes[i++], b3 = bytes[i++];
    const e1 = b1 >> 2;
    const e2 = ((b1 & 3) << 4) | (b2 >> 4);
    let e3 = ((b2 & 15) << 2) | (b3 >> 6);
    let e4 = b3 & 63;
    if (isNaN(b2)) e3 = 64;
    if (isNaN(b3)) e4 = 64;
    out += table[e1] + table[e2] + table[e3] + table[e4];
  }
  return out;
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const shop = "karma-east-au.myshopify.com";
  const apiKey = process.env.SHOPIFY_API_KEY!;
  const apiSecret = process.env.SHOPIFY_API_SECRET!;
  // Use origin of the incoming request so this works in any environment
  const appUrl = new URL(request.url).origin;
  const scopes =
    process.env.SCOPES ||
    "read_products,write_products,read_inventory,write_inventory,read_locations,read_orders,read_reports";

  // Only delete session if forced (?force=1) or if no valid session exists
  const forceParam = new URL(request.url).searchParams.get("force");
  const existing = await prisma.session.findFirst({ where: { shop, accessToken: { not: "" } } }).catch(() => null);
  if (forceParam === "1" || !existing) {
    await prisma.session.deleteMany({ where: { shop } }).catch(() => {});
  } else {
    // Valid session already exists — nothing to do
    return Response.json({ status: "session_ok", shop, sessionId: existing.id.substring(0, 10) + "..." });
  }

  // Generate a 15-digit nonce — same algorithm as @shopify/shopify-api
  const nonce = crypto.getRandomValues(new Uint8Array(15))
    .reduce((acc, b) => acc + String(b % 10), "");

  // Sign the nonce with the API secret — library uses HMAC-SHA256 Base64
  const sig = await hmacBase64(apiSecret, nonce);

  const callbackPath = "/auth/callback";
  const expires = new Date(Date.now() + 60_000).toUTCString();
  const cookieOpts = `; Path=${callbackPath}; SameSite=Lax; Secure; Expires=${expires}`;

  const params = new URLSearchParams({
    client_id: apiKey,
    scope: scopes,
    redirect_uri: `${appUrl}${callbackPath}`,
    state: nonce,
    "grant_options[]": "",
  });
  const oauthUrl = `https://${shop}/admin/oauth/authorize?${params.toString()}`;

  const headers = new Headers();
  headers.set("Location", oauthUrl);
  // Library stores state as two cookies: value + HMAC signature
  headers.append("Set-Cookie", `shopify_app_state=${nonce}${cookieOpts}`);
  headers.append("Set-Cookie", `shopify_app_state.sig=${sig}${cookieOpts}`);

  // ?debug=1 shows the OAuth URL instead of redirecting — remove once working
  if (new URL(request.url).searchParams.get("debug") === "1") {
    return Response.json({
      redirect_uri: `${appUrl}${callbackPath}`,
      oauth_url: oauthUrl,
      client_id: apiKey,
    });
  }

  return new Response(null, { status: 302, headers });
};
