import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// Serves one fabric cell image as a real binary HTTP response. The fabric
// sheets blob keeps images as `data:image/...;base64,...` strings inside a
// 2D rows array; in the loader those cells are replaced with a URL pointing
// here, so the loader response no longer ships megabytes of base64.
//
// URL form: /portal/fabric-image/<gid>/<row>/<col>.png?v=<updatedAt>
// (The `.png` is ignored by the server — it's there so the existing
// isFabricImageValue regex on the client treats the cell as an image.)

const FABRIC_MANUAL_SHEETS_KEY = "production-portal-fabric-manual-sheets-v1";

export async function loader({ params }: LoaderFunctionArgs) {
  const gid = String(params.gid ?? "").trim();
  const row = Number(String(params.row ?? ""));
  const colRaw = String(params.col ?? "").replace(/\.(png|jpe?g|webp|gif|avif)$/i, "");
  const col = Number(colRaw);

  if (
    !gid
    || !Number.isFinite(row) || !Number.isInteger(row) || row < 0
    || !Number.isFinite(col) || !Number.isInteger(col) || col < 0
  ) {
    return new Response("Not found", { status: 404 });
  }

  let cell: string | null = null;
  try {
    // Extract just the one cell from the JSONB array directly in PG. Walks
    // the structure server-side rather than pulling the whole multi-MB blob
    // into Node memory.
    const rows = await prisma.$queryRaw<Array<{ cell: string | null }>>`
      SELECT (elem -> 'rows' -> ${row} ->> ${col}) AS cell
      FROM "PortalSetting", jsonb_array_elements(value) AS elem
      WHERE key = ${FABRIC_MANUAL_SHEETS_KEY} AND elem ->> 'gid' = ${gid}
      LIMIT 1
    `;
    cell = rows[0]?.cell ?? null;
  } catch {
    return new Response("Lookup failed", { status: 500 });
  }

  if (!cell) return new Response("No image", { status: 404 });

  const match = cell.match(/^data:([^;]+);base64,(.+)$/);
  if (!match) return new Response("Invalid image data", { status: 500 });

  const mimeType = match[1];
  let bytes: Uint8Array;
  try {
    const buf = Buffer.from(match[2], "base64");
    bytes = new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
  } catch {
    return new Response("Decode failed", { status: 500 });
  }

  // The `?v=<updatedAt>` cache buster on the URL changes whenever any cell
  // in the blob changes, so we can safely cache forever — the URL itself
  // signals invalidation.
  return new Response(bytes as BodyInit, {
    status: 200,
    headers: {
      "Content-Type": mimeType,
      "Cache-Control": "private, max-age=31536000, immutable",
      "Content-Length": String(bytes.byteLength),
    },
  });
}
