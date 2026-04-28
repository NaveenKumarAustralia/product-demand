import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const shop = "karma-east-au.myshopify.com";

  // Delete the expired session so the auth flow creates a fresh one
  await prisma.session.deleteMany({ where: { shop } }).catch(() => {});

  // Redirect to the login route with shop param — the loader calls login(request)
  // which detects the GET+shop and immediately redirects to Shopify's install page
  return Response.redirect(
    `${new URL(request.url).origin}/auth/login?shop=${shop}`,
    302,
  );
};
