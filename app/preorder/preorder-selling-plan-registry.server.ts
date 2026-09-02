import prisma from "../db.server";

export type PreorderSellingPlanRegistryEntry = {
  supplierOrderId: number;
  shop: string;
  sellingPlanGroupId: string;
  sellingPlanId: string;
  createdAt: string;
  updatedAt: string;
};

function key(shop: string, supplierOrderId: number) {
  return `preorder-selling-plan-v1:${shop}:${supplierOrderId}`;
}

function normalize(value: unknown): PreorderSellingPlanRegistryEntry | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  const supplierOrderId = Number(raw.supplierOrderId);
  const shop = String(raw.shop ?? "").trim();
  const sellingPlanGroupId = String(raw.sellingPlanGroupId ?? "").trim();
  const sellingPlanId = String(raw.sellingPlanId ?? "").trim();
  const createdAt = String(raw.createdAt ?? "").trim();
  const updatedAt = String(raw.updatedAt ?? "").trim();
  if (!Number.isInteger(supplierOrderId) || supplierOrderId <= 0 || !shop || !sellingPlanGroupId || !sellingPlanId) return null;
  return { supplierOrderId, shop, sellingPlanGroupId, sellingPlanId, createdAt, updatedAt };
}

export async function getPreorderSellingPlanRegistryEntry(shop: string, supplierOrderId: number) {
  const setting = await prisma.portalSetting.findUnique({ where: { key: key(shop, supplierOrderId) }, select: { value: true } });
  return normalize(setting?.value);
}

export async function getPreorderSellingPlanRegistryEntries(shop?: string) {
  const prefix = shop ? `preorder-selling-plan-v1:${shop}:` : "preorder-selling-plan-v1:";
  const rows = await prisma.portalSetting.findMany({
    where: { key: { startsWith: prefix } },
    select: { value: true },
  });
  return rows.map((row) => normalize(row.value)).filter(Boolean) as PreorderSellingPlanRegistryEntry[];
}

export async function savePreorderSellingPlanRegistryEntry(entry: Omit<PreorderSellingPlanRegistryEntry, "createdAt" | "updatedAt">) {
  const existing = await getPreorderSellingPlanRegistryEntry(entry.shop, entry.supplierOrderId);
  const now = new Date().toISOString();
  const value: PreorderSellingPlanRegistryEntry = {
    ...entry,
    createdAt: existing?.createdAt || now,
    updatedAt: now,
  };
  await prisma.portalSetting.upsert({
    where: { key: key(entry.shop, entry.supplierOrderId) },
    create: { key: key(entry.shop, entry.supplierOrderId), value },
    update: { value },
  });
  return value;
}

export async function removePreorderSellingPlanRegistryEntry(shop: string, supplierOrderId: number) {
  await prisma.portalSetting.delete({ where: { key: key(shop, supplierOrderId) } }).catch(() => undefined);
}
