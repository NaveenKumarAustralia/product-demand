import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";

// Serves a single thumbnail (or first image) as a real binary image response
// with long-lived cache headers. Cards reference this via <img src=...> so the
// browser can cache, decode, and load thumbnails in parallel via HTTP/2.
//
// URL form: /portal/thumbnail/vision/123 or /portal/thumbnail/sample/456
// A `?v=<timestamp>` cache-buster is appended by the client so an updated
// thumbnail invalidates the cache.
export async function loader({ params }: LoaderFunctionArgs) {
  const entity = params.entity;
  const id = Number(params.id);
  if (!id || (entity !== "vision" && entity !== "sample" && entity !== "collection")) {
    return new Response("Not found", { status: 404 });
  }

  let dataUrl: string | null = null;
  try {
    if (entity === "vision") {
      const item = await prisma.visionBoardItem.findUnique({
        where: { id },
        select: { thumbnail: true, images: true },
      });
      if (item) {
        if (item.thumbnail) dataUrl = item.thumbnail;
        else if (Array.isArray(item.images) && item.images.length > 0) {
          const first = (item.images as unknown[])[0];
          if (typeof first === "string") dataUrl = first;
        }
      }
    } else if (entity === "sample") {
      const it = await prisma.sampleIteration.findUnique({
        where: { id },
        select: { thumbnail: true, images: true },
      });
      if (it) {
        if (it.thumbnail) dataUrl = it.thumbnail;
        else if (Array.isArray(it.images) && it.images.length > 0) {
          const first = (it.images as unknown[])[0];
          if (typeof first === "string") dataUrl = first;
        }
      }
    } else {
      // collection
      const c = await prisma.collection.findUnique({
        where: { id },
        select: { thumbnail: true },
      });
      if (c?.thumbnail) dataUrl = c.thumbnail;
    }
  } catch {
    return new Response("Lookup failed", { status: 500 });
  }

  if (!dataUrl) return new Response("No image", { status: 404 });

  const match = dataUrl.match(/^data:([^;]+);base64,(.+)$/);
  if (!match) return new Response("Invalid image data", { status: 500 });
  const mimeType = match[1];
  const base64 = match[2];

  let binary: Buffer;
  try {
    binary = Buffer.from(base64, "base64");
  } catch {
    return new Response("Decode failed", { status: 500 });
  }

  // The cache-buster (?v=) means whenever the underlying image changes the
  // URL changes, so we can safely cache forever.
  return new Response(binary, {
    status: 200,
    headers: {
      "Content-Type": mimeType,
      "Cache-Control": "private, max-age=31536000, immutable",
      "Content-Length": String(binary.byteLength),
    },
  });
}
