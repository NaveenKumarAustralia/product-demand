import type { ActionFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

// Upload a MEDIA file (image or video) from the computer straight into the
// CollectionImage table, returning a { key } the gallery stores on the row.
// Videos are far too big to inline as a data URL in the row JSON, so they're
// uploaded here and only the key is kept; the bytes stream on demand from
// /portal/collection-image/<key>. Resource route → plain fetch gets JSON back.
//   POST /api/collection-media-upload   (multipart: collectionId, file)
const VIDEO_MAX = 100 * 1024 * 1024; // 100 MB
const IMAGE_MAX = 25 * 1024 * 1024;  // 25 MB

export const action = async ({ request }: ActionFunctionArgs) => {
  await requirePreorderPortalUser(request);
  const form = await request.formData();
  const collectionId = Number(form.get("collectionId"));
  const file = form.get("file");
  if (!collectionId || !(file instanceof File)) return Response.json({ ok: false, error: "bad_input" }, { status: 400 });

  const mime = file.type || "application/octet-stream";
  const isVideo = mime.startsWith("video/");
  const max = isVideo ? VIDEO_MAX : IMAGE_MAX;
  const buf = Buffer.from(await file.arrayBuffer());
  if (!buf.length) return Response.json({ ok: false, error: "The file was empty." }, { status: 400 });
  if (buf.length > max) return Response.json({ ok: false, error: isVideo ? "Video is over 100 MB — please trim or compress it." : "Image is over 25 MB." }, { status: 400 });

  const key = (typeof crypto !== "undefined" && "randomUUID" in crypto)
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2)}`;
  try {
    await prisma.$executeRawUnsafe(
      `INSERT INTO "CollectionImage" ("key", "collectionId", "mimeType", "bytes", "createdAt") VALUES ($1, $2, $3, $4, NOW())`,
      key, collectionId, mime, buf,
    );
  } catch (e) {
    console.warn("[collection-media-upload] insert failed:", e);
    return Response.json({ ok: false, error: "Couldn't save the file. Try again." }, { status: 500 });
  }
  return Response.json({ ok: true, key, mime, kind: isVideo ? "video" : "image", filename: file.name || undefined });
};
