import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { authorizeAislingRequest } from "../api-auth.server";

// Read-only product style feed for the Aisling Equation factory app: every
// style in the portal's Product Info, with its category and costing.
// Server-to-server only (AISLING_API_KEY).
//
// GET → { updatedAt, styles: [{ id, name, category, productType, imageUrl, costs… }] }
//
// Values are the ones SAVED in Product Info. For styles nobody has edited, the
// portal page fills gaps from built-in defaults (PRODUCT_STYLE_COSTING in
// portal._index.tsx); that module can't be imported from a resource route, so
// those gaps come through here as null rather than as a guessed number.

// Must match PRODUCT_INFO_KEY in portal._index.tsx.
const PRODUCT_INFO_KEY = "production-portal-product-info-v2";

const num = (v: unknown) => (v === null || v === undefined || v === "" || !Number.isFinite(Number(v)) ? null : Number(v));
const str = (v: unknown) => String(v ?? "").trim();

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const unauthorized = authorizeAislingRequest(request);
  if (unauthorized) return unauthorized;

  try {
    const setting = await prisma.portalSetting.findUnique({ where: { key: PRODUCT_INFO_KEY } });
    const value = setting?.value as { categories?: unknown } | null;
    const categories = Array.isArray(value?.categories) ? value!.categories as Record<string, unknown>[] : [];

    const styles = categories.flatMap((category) => {
      const categoryName = str(category?.name);
      const list = Array.isArray(category?.styles) ? category.styles as Record<string, unknown>[] : [];
      return list
        .filter((s) => s && str(s.id) && str(s.name))
        .map((s) => ({
          id: str(s.id),
          name: str(s.name),
          category: categoryName,
          productType: str(s.productType) || null,
          imageUrl: str(s.imageUrl) || null,
          hidden: s.hidden === true,
          averageMeters: num(s.averageMeters),
          averageTrimMeters: num(s.averageTrimMeters),
          zipButtonType: str(s.zipButtonType) || null,
          stitchingCost: num(s.stitchingCost),
          fabricCost: num(s.fabricCost),
          zipButtonsCost: num(s.zipButtonsCost),
          liningTrimCost: num(s.liningTrimCost),
          factoryCost: num(s.factoryCost),
          factoryProfit: num(s.factoryProfit),
          totalCost: num(s.totalCost),
          costingNotes: str(s.costingNotes) || null,
        }));
    });

    return Response.json({ updatedAt: setting?.updatedAt ?? null, styles });
  } catch (err) {
    console.error("aisling product-styles error:", err);
    return Response.json({ error: "Database error" }, { status: 500 });
  }
};
