import type { LoaderFunctionArgs } from "react-router";
import { isConfigured, thumbnail } from "../dropbox.server";

// Serves a Dropbox file's thumbnail as a real binary image, cached by the
// browser. Grid tiles reference this via <img src>. Keyed by path (+rev so it
// self-invalidates when the file changes).
//   /api/dropbox-thumb?path=<path>&rev=<rev>&size=w256h256
export const loader = async ({ request }: LoaderFunctionArgs) => {
  if (!isConfigured()) return new Response("Dropbox not configured", { status: 404 });
  const url = new URL(request.url);
  const path = url.searchParams.get("path") ?? "";
  const rev = url.searchParams.get("rev") ?? "";
  const size = url.searchParams.get("size") ?? "w256h256";
  if (!path) return new Response("Missing path", { status: 400 });
  try {
    const buf = await thumbnail(path, size, rev);
    return new Response(new Uint8Array(buf), {
      status: 200,
      headers: {
        "Content-Type": "image/jpeg",
        // rev is in the URL, so a changed file changes the URL — cache long.
        "Cache-Control": "private, max-age=86400",
        "Content-Length": String(buf.byteLength),
      },
    });
  } catch (e) {
    return new Response(`Thumbnail failed: ${(e as Error).message}`, { status: 502 });
  }
};
