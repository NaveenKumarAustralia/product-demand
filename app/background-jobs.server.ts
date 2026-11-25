/**
 * Whether this server may run jobs that act on the live Shopify store
 * (pre-order release/holds/capture, webhook registration).
 *
 * Only the real production environment may. Test copies (a Railway fork has a
 * different RAILWAY_ENVIRONMENT_NAME) must not, and PORTAL_BACKGROUND_JOBS=off
 * turns them off anywhere. Keep in sync with scripts/register-preorder-webhooks.mjs.
 */
export function backgroundJobsEnabled(env: Record<string, string | undefined> = process.env): boolean {
  if (env.PORTAL_BACKGROUND_JOBS === "off") return false;
  const railwayEnv = env.RAILWAY_ENVIRONMENT_NAME;
  return railwayEnv === undefined || railwayEnv === "" || railwayEnv === "production";
}
