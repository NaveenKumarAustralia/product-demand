import type { ActionFunctionArgs } from "react-router";
import { randomUUID } from "node:crypto";
import { createKlaviyoEvent, KLAVIYO_BACK_IN_STOCK_METRIC } from "../preorder/preorder-klaviyo.server";
import { canSendPreorderNotifications, getPreorderPermissionContext } from "../preorder/preorder-permissions.server";
import { requirePreorderPortalUser, requireSameOrigin } from "../preorder/preorder-portal-auth.server";

const TEST_METRIC = "Karma East Klaviyo Integration Test";

export const action = async ({ request }: ActionFunctionArgs) => {
  try {
    requireSameOrigin(request);
    const actor = await requirePreorderPortalUser(request);
    const { permissions } = await getPreorderPermissionContext();
    if (!canSendPreorderNotifications(actor, permissions)) {
      return Response.json({ ok: false, error: "You do not have permission to send notification tests." }, { status: 403 });
    }

    const body = await request.json().catch(() => ({})) as { email?: unknown; kind?: unknown };
    const email = typeof body.email === "string" ? body.email.trim() : "";
    if (!email) return Response.json({ ok: false, error: "Enter a real email address for the Klaviyo test." }, { status: 400 });

    // "backinstock" fires a SAMPLE of the real back-in-stock metric so the
    // "Karma East Back In Stock Available" metric appears in Klaviyo (a metric
    // doesn't exist until its first event lands) and the flow/email template can
    // be built against realistic event properties. Default is the plain
    // connection test, which uses a separate throwaway metric.
    const kind = body.kind === "backinstock" ? "backinstock" : "connection";

    const uniqueId = `klaviyo-test:${randomUUID()}`;
    const result = await createKlaviyoEvent({
      email,
      metric: kind === "backinstock" ? KLAVIYO_BACK_IN_STOCK_METRIC : TEST_METRIC,
      uniqueId,
      properties: kind === "backinstock"
        ? {
            source: "karma-east-production-portal",
            purpose: "back-in-stock-template-sample",
            sent_by: actor.name,
            sent_at: new Date().toISOString(),
            waitlist_id: 0,
            shop: "sample",
            market: "AU",
            product_id: "sample-product",
            product_title: "Sample Product (test)",
            variant_title: "M",
            sku: "SAMPLE-SKU",
            product_url: "https://www.karmaeast.com.au/collections/new-arrivals",
            image: "https://d3k81ch9hvuctc.cloudfront.net/company/Th9m73/images/afa9f413-912d-4b29-a659-dc11e54d92b7.png",
            available_quantity: 3,
          }
        : {
            source: "karma-east-production-portal",
            purpose: "integration-test",
            sent_by: actor.name,
            sent_at: new Date().toISOString(),
          },
    });

    return Response.json({ ok: true, accepted: result.accepted, metric: result.metric, uniqueId: result.uniqueId });
  } catch (error) {
    if (error instanceof Response) return error;
    console.error("[preorder klaviyo] integration test failed", error);
    return Response.json({ ok: false, error: error instanceof Error ? error.message : "Klaviyo integration test failed." }, { status: 500 });
  }
};
