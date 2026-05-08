import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// Serves a single image at a specific index of an item's images array as a
// real binary image response with long-lived cache headers. Drawers and
// galleries reference this via <img src=...> so the browser handles parallel
// downloads, decoding, and caching natively — no big base64-in-JSON payload.
//
// URL form: /portal/image/vision/123/2  or  /portal/image/sample/456/0
// A `?v=<timestamp>` cache-buster is appended by the client when the image
// list changes (e.g. after a remove that shifted indices).
export async function loader({ params }: LoaderFunctionArgs) {
  const entity = params.entity;
  const id = Number(params.id);
  const index = Number(params.index);
  if (
    !id || !Number.isFinite(id) || !Number.isInteger(id)
    || !Number.isFinite(index) || !Number.isInteger(index) || index < 0
    || (entity !== "vision" && entity !== "visionV2" && entity !== "sample")
  ) {
    return new Response("Not found", { status: 404 });
  }

  // Project ONLY the requested image from the JSONB array — no full row, no
  // multi-MB transfer from the DB to Node.
  const table = entity === "vision"
    ? "VisionBoardItem"
    : entity === "visionV2"
      ? "VisionBoardV2Item"
      : "SampleIteration";
  let dataUrl: string | null = null;
  try {
    const rows = await prisma.$queryRawUnsafe<Array<{ image: string | null }>>(
      `SELECT images ->> ${index} AS image FROM "${table}" WHERE id = ${id}`,
    );
    dataUrl = rows[0]?.image ?? null;
  } catch {
    return new Response("Lookup failed", { status: 500 });
  }
  if (!dataUrl) return new Response("No image at index", { status: 404 });

  const match = dataUrl.match(/^data:([^;]+);base64,(.+)$/);
  if (!match) return new Response("Invalid image data", { status: 500 });
  const mimeType = match[1];
  let bytes: Uint8Array;
  try {
    const buf = Buffer.from(match[2], "base64");
    bytes = new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
  } catch {
    return new Response("Decode failed", { status: 500 });
  }

  return new Response(bytes as BodyInit, {
    status: 200,
    headers: {
      "Content-Type": mimeType,
      "Cache-Control": "private, max-age=31536000, immutable",
      "Content-Length": String(bytes.byteLength),
    },
  });
}
