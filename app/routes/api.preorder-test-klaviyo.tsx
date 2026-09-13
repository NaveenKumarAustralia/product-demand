import type { LoaderFunctionArgs } from "react-router";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";
import { getKlaviyoConnectionStatus, sendPreorderPlacedEvent, KLAVIYO_PREORDER_PLACED_METRIC } from "../preorder/preorder-klaviyo.server";

// Admin: fire ONE sample "Pre-order Placed" event so the metric appears in the
// Klaviyo flow builder (Klaviyo only lists a metric after it has received it once)
// and so you can preview/test the flow without waiting for a real order. Safe to
// run repeatedly — each call uses a fresh unique_id so Klaviyo won't dedupe it.
//   GET /api/preorder-test-klaviyo?email=you@example.com[&n=386999]
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const status = getKlaviyoConnectionStatus();
  if (!status.configured) {
    return Response.json(
      { ok: false, error: "Klaviyo is not connected. Add KLAVIYO_PRIVATE_API_KEY in Railway, redeploy, then try again.", metric: KLAVIYO_PREORDER_PLACED_METRIC },
      { status: 400 },
    );
  }

  const url = new URL(request.url);
  const email = (url.searchParams.get("email") || "").trim();
  if (!email || !email.includes("@")) {
    return Response.json({ ok: false, error: "Pass ?email=you@example.com so the test event has a recipient." }, { status: 400 });
  }
  // A unique, obviously-fake order number so it never collides with a real order.
  const orderName = url.searchParams.get("n") ? `#${url.searchParams.get("n")}TEST` : "#TEST-PREORDER";
  const orderId = `test-${orderName.replace(/[^0-9a-zA-Z]/g, "")}-${email.length}`;

  try {
    const res = await sendPreorderPlacedEvent({
      shop: "karma-east-test",
      orderId,
      orderName,
      email,
      market: "AU",
      items: [
        { title: "Pippa Dress Clematis", size: "2XL", dispatch: "12 Oct 2025" },
        { title: "Maddison Dress Shikari", size: "M", dispatch: "12 Oct 2025" },
      ],
    });
    return Response.json(
      { ok: true, sent: res, metric: KLAVIYO_PREORDER_PLACED_METRIC, hint: "In Klaviyo → Metrics you should now see 'Karma East Pre-order Placed'. Use it as the flow trigger." },
      { headers: { "Cache-Control": "no-store" } },
    );
  } catch (error) {
    return Response.json({ ok: false, error: error instanceof Error ? error.message : String(error) }, { status: 500 });
  }
};
