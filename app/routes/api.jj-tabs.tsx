import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// Serves the JJ order-tab list to the Shopify product-order extension so it can
// show a "which JJ tab" dropdown when placing an order. Auth mirrors
// api.place-order: a Shopify session JWT whose `aud` is our app's API key.

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type",
};

// Must match the key used in portal._index.tsx.
const JJ_TABS_KEY = "supplier-portal-jj-tabs-v1";
const JJ_DEFAULT_TABS = [{ id: "main", name: "Order 1" }];

type JJTab = { id: string; name: string; hidden?: boolean };

function normalizeJJTabs(value: unknown): JJTab[] {
  const arr = value && typeof value === "object" && Array.isArray((value as { tabs?: unknown }).tabs)
    ? (value as { tabs: unknown[] }).tabs
    : Array.isArray(value) ? value : [];
  const tabs = arr
    .map((t): JJTab | null => (t && typeof t === "object")
      ? { id: String((t as JJTab).id ?? "").trim(), name: String((t as JJTab).name ?? "").trim(), hidden: Boolean((t as JJTab).hidden) }
      : null)
    .filter((t): t is JJTab => Boolean(t && t.id && t.name));
  return tabs.length ? tabs : JJ_DEFAULT_TABS;
}

function decodeJwtPayload(token: string): Record<string, unknown> | null {
  try {
    const [, payloadB64] = token.split(".");
    return JSON.parse(Buffer.from(payloadB64, "base64url").toString("utf-8"));
  } catch {
    return null;
  }
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  if (request.method === "OPTIONS") {
    return new Response(null, { status: 204, headers: CORS });
  }

  const authHeader = request.headers.get("Authorization") ?? "";
  const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : null;
  if (!token) return Response.json({ error: "Missing token" }, { status: 401, headers: CORS });

  const payload = decodeJwtPayload(token);
  if (!payload) return Response.json({ error: "Invalid token" }, { status: 401, headers: CORS });

  const clientId = process.env.SHOPIFY_API_KEY;
  const aud = payload.aud;
  const audValid = aud === clientId || (Array.isArray(aud) && aud.includes(clientId));
  if (!audValid) return Response.json({ error: "Token audience mismatch" }, { status: 401, headers: CORS });

  const setting = await prisma.portalSetting.findUnique({ where: { key: JJ_TABS_KEY }, select: { value: true } }).catch(() => null);
  // Hidden tabs are excluded from the Shopify product-page dropdown.
  const tabs = normalizeJJTabs(setting?.value)
    .filter((t) => !t.hidden)
    .map(({ id, name }) => ({ id, name }));
  return Response.json({ tabs }, { headers: CORS });
};
