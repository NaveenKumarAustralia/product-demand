import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// Serves the full bytes of a single CollectionImage by key. Used by
// the multi-image popup (which lazy-loads full quality) and the
// recompress action. Inline thumbnail in the row JSON is enough for
// the spreadsheet view itself, so this route is only hit when the
// user opens the manager modal or pushes to Shopify.
export const loader = async ({ params, request }: LoaderFunctionArgs) => {
  const key = (params.key ?? "").trim();
  if (!key) throw new Response("missing key", { status: 400 });

  const rows = await prisma.$queryRawUnsafe<Array<{ bytes: Buffer; mimeType: string }>>(
    `SELECT "bytes", "mimeType" FROM "CollectionImage" WHERE "key" = $1 LIMIT 1`,
    key,
  ).catch(() => []);
  if (rows.length === 0) throw new Response("not found", { status: 404 });

  const { bytes, mimeType } = rows[0];
  const buf = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes as unknown as ArrayBuffer);
  void request;
  return new Response(buf, {
    headers: {
      "Content-Type": mimeType || "application/octet-stream",
      "Content-Length": String(buf.length),
      // Immutable: the bytes never change for a given key — bumping
      // means a new key. Cache aggressively in the browser.
      "Cache-Control": "public, max-age=31536000, immutable",
    },
  });
};
