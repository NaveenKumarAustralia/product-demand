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
