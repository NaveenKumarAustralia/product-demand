import type { ActionFunctionArgs } from "react-router";

// AI image ALT TEXT generation for Collections images. Sends the image to the
// Anthropic Messages API (vision) and returns concise ecommerce alt text.
// POST form fields:
//   image        = a data URL (data:image/...;base64,....) — the image to describe
//   productName  = optional product/style name for context
// Returns { ok: true, text } or { ok: false, error }.

const MODEL = process.env.ANTHROPIC_MODEL || "claude-sonnet-5";
const ALLOWED_MEDIA = new Set(["image/jpeg", "image/png", "image/gif", "image/webp"]);

export const action = async ({ request }: ActionFunctionArgs) => {
  const key = process.env.ANTHROPIC_API_KEY;
  if (!key) return Response.json({ ok: false, error: "AI isn't configured (missing ANTHROPIC_API_KEY)." });

  const form = await request.formData();
  const image = String(form.get("image") ?? "").trim();
  const productName = String(form.get("productName") ?? "").trim();

  const m = image.match(/^data:([^;]+);base64,(.+)$/s);
  if (!m) return Response.json({ ok: false, error: "No image (expected a base64 data URL)." });
  const mediaType = m[1].toLowerCase();
  const data = m[2];
  if (!ALLOWED_MEDIA.has(mediaType)) return Response.json({ ok: false, error: `Unsupported image type: ${mediaType}` });

  const system = "You write ALT TEXT for fashion product photos on an online store. Be concise and factual: describe the garment type, colour/print, and the view or how it's shown (e.g. flat lay, on a model, close-up). 125 characters max. No quotes, no 'image of' / 'photo of', no marketing language. Output only the alt text.";
  const user = `Write alt text for this product image.${productName ? ` The product is: ${productName}.` : ""}`;

  try {
    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json" },
      body: JSON.stringify({
        model: MODEL,
        max_tokens: 100,
        system,
        messages: [{
          role: "user",
          content: [
            { type: "image", source: { type: "base64", media_type: mediaType, data } },
            { type: "text", text: user },
          ],
        }],
      }),
    });
    if (!res.ok) {
      const detail = await res.text().catch(() => "");
      return Response.json({ ok: false, error: `AI request failed (${res.status}). ${detail.slice(0, 200)}` });
    }
    const json = await res.json() as { content?: Array<{ type?: string; text?: string }> };
    const text = (json.content ?? []).filter((c) => c.type === "text").map((c) => c.text ?? "").join("").trim().replace(/^["']|["']$/g, "");
    if (!text) return Response.json({ ok: false, error: "AI returned no text." });
    return Response.json({ ok: true, text });
  } catch (e) {
    return Response.json({ ok: false, error: (e as Error).message || "AI request errored." });
  }
};
