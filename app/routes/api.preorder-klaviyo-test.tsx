import type { ActionFunctionArgs } from "react-router";
import { randomUUID } from "node:crypto";
import { createKlaviyoEvent } from "../preorder/preorder-klaviyo.server";
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

    const body = await request.json().catch(() => ({})) as { email?: unknown };
    const email = typeof body.email === "string" ? body.email.trim() : "";
    if (!email) return Response.json({ ok: false, error: "Enter a real email address for the Klaviyo test." }, { status: 400 });

    const uniqueId = `klaviyo-test:${randomUUID()}`;
    const result = await createKlaviyoEvent({
      email,
      metric: TEST_METRIC,
      uniqueId,
      properties: {
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
