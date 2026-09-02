import prisma from "../db.server";
import type { PreorderMarket } from "./preorder-rules.server";

export const PREORDER_LOCATION_SETTINGS_KEY = "preorder-location-settings-v1";

export type PreorderLocationSettings = {
  AU: string | null;
  USA: string | null;
};

const EMPTY_LOCATIONS: PreorderLocationSettings = { AU: null, USA: null };

export function normalizeShopifyLocationId(value: unknown): string | null {
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  if (/^gid:\/\/shopify\/Location\/\d+$/.test(raw)) return raw;
  const numeric = raw.replace(/[^0-9]/g, "");
  return numeric ? `gid://shopify/Location/${numeric}` : null;
}

export function normalizePreorderLocationSettings(value: unknown): PreorderLocationSettings {
  if (!value || typeof value !== "object" || Array.isArray(value)) return { ...EMPTY_LOCATIONS };
  const row = value as Record<string, unknown>;
  return {
    AU: normalizeShopifyLocationId(row.AU),
    USA: normalizeShopifyLocationId(row.USA),
  };
}

export async function getPreorderLocationSettings(): Promise<PreorderLocationSettings> {
  const setting = await prisma.portalSetting.findUnique({
    where: { key: PREORDER_LOCATION_SETTINGS_KEY },
    select: { value: true },
  });
  return normalizePreorderLocationSettings(setting?.value);
}

export function locationForMarket(settings: PreorderLocationSettings, market: PreorderMarket) {
  return settings[market];
}

export async function setPreorderLocationSettings(
  next: PreorderLocationSettings,
  actorName: string,
) {
  const normalized = normalizePreorderLocationSettings(next);
  await prisma.portalSetting.upsert({
    where: { key: PREORDER_LOCATION_SETTINGS_KEY },
    create: { key: PREORDER_LOCATION_SETTINGS_KEY, value: normalized },
    update: { value: normalized },
  });

  await prisma.activityLog.create({
    data: {
      userName: actorName || "Unknown",
      action: "preorder_location_settings_updated",
      entity: "preorder_settings",
      field: "locations",
      toValue: JSON.stringify(normalized),
    },
  }).catch((error) => console.warn("[preorder location audit] failed:", error));

  return normalized;
}
