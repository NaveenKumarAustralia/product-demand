import prisma from "../db.server";

// Global on/off for the storefront "Notify me" (back-in-stock) block. Pre-order
// is unaffected — only the notify-me fallback is suppressed when this is off.
// Lets the merchant run a separate back-in-stock app without showing two forms.
export const PREORDER_NOTIFY_ENABLED_KEY = "preorder-storefront-notify-enabled-v1";

export async function getPreorderNotifyEnabled(): Promise<boolean> {
  const setting = await prisma.portalSetting.findUnique({
    where: { key: PREORDER_NOTIFY_ENABLED_KEY },
    select: { value: true },
  });
  const value: unknown = setting?.value;
  // Default ON when never set.
  if (value == null) return true;
  if (typeof value === "boolean") return value;
  if (typeof value === "object" && !Array.isArray(value) && "enabled" in (value as Record<string, unknown>)) {
    return (value as Record<string, unknown>).enabled !== false;
  }
  return true;
}

export async function setPreorderNotifyEnabled(enabled: boolean, actorName: string): Promise<boolean> {
  const value = { enabled, updatedBy: actorName, updatedAt: new Date().toISOString() };
  await prisma.portalSetting.upsert({
    where: { key: PREORDER_NOTIFY_ENABLED_KEY },
    create: { key: PREORDER_NOTIFY_ENABLED_KEY, value },
    update: { value },
  });
  return enabled;
}

// "Combine window" — if a MIXED order (in-stock + pre-order) has a pre-order
// item whose promised dispatch is within this many days at order time, we hold
// the WHOLE order (in-stock items too) so it all ships together when the batch
// lands, instead of shipping the in-stock part separately. Saves a second parcel
// when the pre-order stock is about to arrive anyway. 0 disables (always ship
// in-stock immediately). Default 0 = OFF — only the pre-order line is held, the
// in-stock items ship straight away. Set > 0 in settings to combine again.
export const PREORDER_COMBINE_WINDOW_KEY = "preorder-combine-window-days-v1";
export const PREORDER_COMBINE_WINDOW_DEFAULT = 0;

export async function getPreorderCombineWindowDays(): Promise<number> {
  const setting = await prisma.portalSetting.findUnique({
    where: { key: PREORDER_COMBINE_WINDOW_KEY },
    select: { value: true },
  });
  const value: unknown = setting?.value;
  if (value == null) return PREORDER_COMBINE_WINDOW_DEFAULT;
  const raw = typeof value === "object" && value && "days" in (value as Record<string, unknown>)
    ? (value as Record<string, unknown>).days
    : value;
  const days = Math.floor(Number(raw));
  if (!Number.isFinite(days) || days < 0) return PREORDER_COMBINE_WINDOW_DEFAULT;
  return Math.min(days, 365);
}

export async function setPreorderCombineWindowDays(days: number, actorName: string): Promise<number> {
  const clean = Math.max(0, Math.min(365, Math.floor(Number(days) || 0)));
  const value = { days: clean, updatedBy: actorName, updatedAt: new Date().toISOString() };
  await prisma.portalSetting.upsert({
    where: { key: PREORDER_COMBINE_WINDOW_KEY },
    create: { key: PREORDER_COMBINE_WINDOW_KEY, value },
    update: { value },
  });
  return clean;
}
