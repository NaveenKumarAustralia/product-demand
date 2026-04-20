import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type",
};

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

  if (!token) {
    return Response.json({ error: "Missing token" }, { status: 401, headers: CORS });
  }

  const payload = decodeJwtPayload(token);
  if (!payload) {
    return Response.json({ error: "Invalid token" }, { status: 401, headers: CORS });
  }

  const clientId = process.env.SHOPIFY_API_KEY;
  const aud = payload.aud;
  const audValid = aud === clientId || (Array.isArray(aud) && aud.includes(clientId));
  if (!audValid) {
    return Response.json({ error: "Token audience mismatch" }, { status: 401, headers: CORS });
  }

  const url = new URL(request.url);
  const shop = url.searchParams.get("shop");
  if (!shop) {
    return Response.json({ error: "shop is required" }, { status: 400, headers: CORS });
  }

  try {
    const rows = await prisma.supplierOrder.findMany({
      where: { shop },
      select: { supplier: true },
      distinct: ["supplier"],
      orderBy: { supplier: "asc" },
    });

    const suppliers = rows.map((r) => r.supplier).filter(Boolean);
    return Response.json({ suppliers }, { headers: CORS });
  } catch (err) {
    console.error("suppliers DB error:", err);
    return Response.json({ error: "Database error" }, { status: 500, headers: CORS });
  }
};
