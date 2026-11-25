import { PassThrough } from "stream";
import { renderToPipeableStream } from "react-dom/server";
import { ServerRouter } from "react-router";
import { createReadableStreamFromReadable } from "@react-router/node";
import { type EntryContext } from "react-router";
import { isbot } from "isbot";
import { addDocumentResponseHeaders } from "./shopify.server";
import { startPreorderReleaseScheduler } from "./preorder/preorder-release.server";
import { startPreorderPlanReconcileScheduler } from "./preorder/preorder-selling-plan.service.server";
import { startMissedPreorderCaptureScheduler } from "./preorder/preorder-missed-capture.server";
import { startPreorderReconcileSweepScheduler } from "./preorder/preorder-reconcile.server";
import { startPreorderReholdScheduler } from "./preorder/preorder-fulfillment.server";
import { backgroundJobsEnabled } from "./background-jobs.server";

// Background jobs act on the LIVE Shopify store, so they only run in the real
// production environment. Isolated test copies (e.g. a Railway fork, whose
// RAILWAY_ENVIRONMENT_NAME is not "production") never start them, and
// PORTAL_BACKGROUND_JOBS=off switches them off anywhere. Off Railway (no
// RAILWAY_ENVIRONMENT_NAME) and in production, behaviour is unchanged.
if (backgroundJobsEnabled()) {
  // Start the pre-order auto-release timer once when the server boots (releases
  // Shopify holds + tags orders for Pick Pack when a batch's stock lands).
  startPreorderReleaseScheduler();
  // Daily safety net: keep every live selling plan's "Expected <date>" in sync
  // with its batch's dispatch date (in case an on-change rename ever failed).
  startPreorderPlanReconcileScheduler();
  // Every few hours: capture missed pre-orders (bought via Shop Pay / quick-add
  // with no selling plan) into the pre-order pipeline so they don't slip through.
  startMissedPreorderCaptureScheduler();
  // Hourly pre-order reconcile sweep. Converges every pre-order product's Shopify
  // state to the DB. Defaults to "observe" (dry-run + log only) until flipped to
  // "apply" from /api/preorder-reconcile — so it can't change anything unverified.
  startPreorderReconcileSweepScheduler();
  // Every 15 min: re-apply the fulfilment hold to any pre-order order that is
  // reserved but somehow not held (so a hold that hiccupped never leaves a
  // pre-order sitting as a normal unfulfilled order in the picking app).
  startPreorderReholdScheduler();
} else {
  console.warn(`[portal] background jobs NOT started (environment: ${process.env.RAILWAY_ENVIRONMENT_NAME ?? "unknown"}, PORTAL_BACKGROUND_JOBS=${process.env.PORTAL_BACKGROUND_JOBS ?? "unset"})`);
}

export const streamTimeout = 5000;

export default async function handleRequest(
  request: Request,
  responseStatusCode: number,
  responseHeaders: Headers,
  reactRouterContext: EntryContext
) {
  addDocumentResponseHeaders(request, responseHeaders);
  const userAgent = request.headers.get("user-agent");
  const callbackName = isbot(userAgent ?? '')
    ? "onAllReady"
    : "onShellReady";

  return new Promise((resolve, reject) => {
    const { pipe, abort } = renderToPipeableStream(
      <ServerRouter
        context={reactRouterContext}
        url={request.url}
      />,
      {
        [callbackName]: () => {
          const body = new PassThrough();
          const stream = createReadableStreamFromReadable(body);

          responseHeaders.set("Content-Type", "text/html");
          resolve(
            new Response(stream, {
              headers: responseHeaders,
              status: responseStatusCode,
            })
          );
          pipe(body);
        },
        onShellError(error) {
          reject(error);
        },
        onError(error) {
          responseStatusCode = 500;
          console.error(error);
        },
      }
    );

    // Automatically timeout the React renderer after 6 seconds, which ensures
    // React has enough time to flush down the rejected boundary contents
    setTimeout(abort, streamTimeout + 1000);
  });
}
