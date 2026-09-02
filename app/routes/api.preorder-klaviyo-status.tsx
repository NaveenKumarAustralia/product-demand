import type { LoaderFunctionArgs } from "react-router";
import { getKlaviyoConnectionStatus } from "../preorder/preorder-klaviyo.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  try {
    await requirePreorderPortalUser(request);
    const status = getKlaviyoConnectionStatus();
    const key = process.env.KLAVIYO_PRIVATE_API_KEY;
    const klaviyoVariableNames = Object.keys(process.env)
      .filter((name) => name.toUpperCase().includes("KLAVIYO"))
      .sort();

    return Response.json({
      ok: true,
      ...status,
      diagnostics: {
        expectedVariable: "KLAVIYO_PRIVATE_API_KEY",
        expectedVariablePresent: Object.prototype.hasOwnProperty.call(process.env, "KLAVIYO_PRIVATE_API_KEY"),
        valueLength: typeof key === "string" ? key.length : 0,
        klaviyoVariableNames,
      },
    });
  } catch (error) {
    if (error instanceof Response) return error;
    console.error("[preorder klaviyo] status check failed", error);
    return Response.json({
      ok: false,
      configured: false,
      error: error instanceof Error ? error.message : "Could not check Klaviyo connection.",
    }, { status: 500 });
  }
};
