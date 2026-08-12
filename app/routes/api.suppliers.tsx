import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { authorizeApiRequest } from "../api-auth.server";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type, X-Api-Key",
};

export const loader = async ({ request }: LoaderFunctionArgs) => {
  if (request.method === "OPTIONS") {
    return new Response(null, { status: 204, headers: CORS });
  }

  // Accepts either a Shopify session token from the embedded app or the
  // dashboard's shared secret. See api-auth.server.ts.
  const unauthorized = authorizeApiRequest(request, CORS);
  if (unauthorized) return unauthorized;

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
