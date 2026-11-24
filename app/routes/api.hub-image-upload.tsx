import type { ActionFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

// Upload an IMAGE for a collection hub note/task comment into the CollectionImage
// table, returning a { key } to embed in the message body as "[img:<key>]".
// Bytes stream back on demand from /portal/collection-image/<key>.
//   POST /api/hub-image-upload   (multipart: file, optional collectionId)
const IMAGE_MAX = 25 * 1024 * 1024; // 25 MB

export const action = async ({ request }: ActionFunctionArgs) => {
  await requirePreorderPortalUser(request);
  const form = await request.formData();
  const file = form.get("file");
  if (!(file instanceof File)) return Response.json({ ok: false, error: "bad_input" }, { status: 400 });

  const mime = file.type || "application/octet-stream";
  if (!mime.startsWith("image/")) return Response.json({ ok: false, error: "Only images can be attached." }, { status: 400 });
  const buf = Buffer.from(await file.arrayBuffer());
  if (!buf.length) return Response.json({ ok: false, error: "The file was empty." }, { status: 400 });
  if (buf.length > IMAGE_MAX) return Response.json({ ok: false, error: "Image is over 25 MB." }, { status: 400 });

  // Owner tag: the sheet's collection id when it's a sheet-scoped attachment,
  // else 0 (group-scoped). Only used for cleanup; serving is by key.
  const collectionId = Number(form.get("collectionId")) || 0;
  const key = (typeof crypto !== "undefined" && "randomUUID" in crypto)
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2)}`;
  try {
    await prisma.$executeRawUnsafe(
      `INSERT INTO "CollectionImage" ("key", "collectionId", "mimeType", "bytes", "createdAt") VALUES ($1, $2, $3, $4, NOW())`,
      key, collectionId, mime, buf,
    );
  } catch (e) {
    console.warn("[hub-image-upload] insert failed:", e);
    return Response.json({ ok: false, error: "Couldn't save the image. Try again." }, { status: 500 });
  }
  return Response.json({ ok: true, key, mime });
};
