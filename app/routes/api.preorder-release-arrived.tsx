import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { releaseArrivedPreorders } from "../preorder/preorder-release.server";

const API_VERSION = "2025-10";
const REQUIRED_SCOPES = [
  "write_orders",
  "read_merchant_managed_fulfillment_orders",
  "write_merchant_managed_fulfillment_orders",
] as const;

// Admin: confirm the auto-release scopes are granted, then run the release check
// immediately (release Shopify holds + tag orders for Pick Pack for any batch
// whose stock has landed). Idempotent (readyAt). GET so it runs from the browser
// while logged into the portal.
async function checkScopes(): Promise<{ ok: boolean; granted: string[]; missing: string[]; error?: string }> {
  const session = await prisma.session.findFirst({
    where: { isOnline: false, accessToken: { not: "" } },
    orderBy: { expires: "desc" },
    select: { shop: true, accessToken: true },
  });
  if (!session?.accessToken) return { ok: false, granted: [], missing: [...REQUIRED_SCOPES], error: "No offline Shopify session." };
  try {
    const res = await fetch(`https://${session.shop}/admin/api/${API_VERSION}/graphql.json`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": session.accessToken },
      body: JSON.stringify({ query: "#graphql\n query { currentAppInstallation { accessScopes { handle } } }" }),
    });
    const json = await res.json() as { data?: { currentAppInstallation?: { accessScopes?: Array<{ handle?: string }> } } };
    const granted = (json.data?.currentAppInstallation?.accessScopes ?? []).map((s) => String(s.handle ?? ""));
    const grantedSet = new Set(granted);
    const missing = REQUIRED_SCOPES.filter((s) => !grantedSet.has(s));
    return { ok: missing.length === 0, granted, missing };
  } catch (error) {
    return { ok: false, granted: [], missing: [...REQUIRED_SCOPES], error: error instanceof Error ? error.message : String(error) };
  }
}

async function run(request: Request) {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const scopes = await checkScopes();
  const release = await releaseArrivedPreorders();
  return Response.json({
    ok: true,
    scopesReady: scopes.ok,
    missingScopes: scopes.missing,
    grantedScopes: scopes.granted,
    release,
  }, { headers: { "Cache-Control": "no-store" } });
}

export const loader = async ({ request }: LoaderFunctionArgs) => run(request);
export const action = async ({ request }: ActionFunctionArgs) => run(request);
