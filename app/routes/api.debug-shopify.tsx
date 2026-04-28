import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { unauthenticated } from "../shopify.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const url = new URL(request.url);
  const query = url.searchParams.get("q") ?? "test";

  const session = await prisma.session.findFirst({
    where: { accessToken: { not: "" } },
    orderBy: { isOnline: "asc" },
  }).catch((err) => { return { error: String(err) }; });

  if (!session || "error" in session) {
    return Response.json({ step: "session", error: session && "error" in session ? session.error : "no session found" });
  }
  if (!("shop" in session) || !session.shop || !session.accessToken) {
    return Response.json({ step: "session", error: "session missing shop or accessToken", session });
  }

  const shop = session.shop;
  const gqlQuery = `{ products(first: 5, query: ${JSON.stringify(query)}) { edges { node { id title } } } }`;

  // Try unauthenticated.admin
  try {
    const { admin } = await unauthenticated.admin(shop);
    const resp = await admin.graphql(gqlQuery);
    const json = await resp.json();
    return Response.json({
      step: "unauthenticated.admin",
      shop,
      query,
      dataNULL: json?.data == null,
      errors: json?.errors ?? null,
      products: (json?.data?.products?.edges ?? []).map((e: any) => e.node.title),
    });
  } catch (err1) {
    // Fallback: direct fetch
    try {
      const resp = await fetch(`https://${shop}/admin/api/2025-10/graphql.json`, {
        method: "POST",
        headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": session.accessToken },
        body: JSON.stringify({ query: gqlQuery }),
      });
      const json = await resp.json();
      return Response.json({
        step: "directFetch",
        shop,
        query,
        httpStatus: resp.status,
        dataNULL: json?.data == null,
        errors: json?.errors ?? null,
        products: (json?.data?.products?.edges ?? []).map((e: any) => e.node.title),
        unauthError: String(err1),
      });
    } catch (err2) {
      return Response.json({
        step: "directFetch failed",
        shop,
        unauthError: String(err1),
        fetchError: String(err2),
      });
    }
  }
};
