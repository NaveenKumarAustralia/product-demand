import type { ActionFunctionArgs } from "react-router";
import sharp from "sharp";

// Normalises ANY uploaded image into a small, web-displayable JPEG — including
// Apple HEIC/HEIF photos (the format AirDrop and iPhones produce), which
// browsers CANNOT decode client-side, so the usual <canvas> resize path fails
// and the tile ends up blank. sharp decodes HEIC via libheif and downsizes, so
// the stored value is always a normal JPEG data URL and the saved blob stays
// small.
//
// POST multipart: file=<image>, optional maxEdge (px, default 900).
// Response: { ok, dataUrl, bytes } | { ok:false, error }.
export const action = async ({ request }: ActionFunctionArgs) => {
  try {
    const form = await request.formData();
    const file = form.get("file");
    if (!(file instanceof Blob)) return Response.json({ ok: false, error: "no_file" }, { status: 400 });
    const maxEdge = Math.max(200, Math.min(2000, Number(form.get("maxEdge")) || 900));
    const input = Buffer.from(await file.arrayBuffer());
    const out = await sharp(input, { failOn: "none" })
      .rotate() // honour EXIF orientation (phones store it in metadata)
      .resize(maxEdge, maxEdge, { fit: "inside", withoutEnlargement: true })
      .jpeg({ quality: 82, mozjpeg: true })
      .toBuffer();
    const dataUrl = `data:image/jpeg;base64,${out.toString("base64")}`;
    return Response.json({ ok: true, dataUrl, bytes: out.length });
  } catch (e) {
    console.warn("[api.image-normalize] convert failed:", e);
    return Response.json({ ok: false, error: "convert_failed" }, { status: 500 });
  }
};
