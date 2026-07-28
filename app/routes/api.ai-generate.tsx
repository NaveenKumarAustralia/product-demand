import type { ActionFunctionArgs } from "react-router";

// AI copy generation for the Collections page (description / SEO). Calls the
// Anthropic Messages API directly with the ANTHROPIC_API_KEY env var — no SDK
// dependency. POST form fields:
//   kind        = "description" | "seoTitle" | "seoDescription"
//   productName = the product / style name (required)
//   productType = optional, e.g. "Short Sleeve Dress"
//   source      = optional source text to rewrite (e.g. a duplicated product's
//                 description) — the model improves/rewrites it rather than
//                 inventing from scratch.
// Returns { ok: true, text } or { ok: false, error }.

const MODEL = process.env.ANTHROPIC_MODEL || "claude-sonnet-5";

function buildPrompt(kind: string, name: string, productType: string, source: string): { system: string; user: string; maxTokens: number } {
  const ctx = `Product name: ${name}${productType ? `\nProduct type: ${productType}` : ""}${source ? `\n\nExisting copy to rewrite/improve (keep the meaning, refresh the wording, do not copy verbatim):\n${source}` : ""}`;
  const brand = "You write for a women's fashion brand: warm, natural, confident, not salesy or clichéd. Never invent specific facts (materials, measurements, care) that aren't given. Output ONLY the requested copy — no preamble, quotes, or labels.";
  if (kind === "seoTitle") {
    return { system: brand, user: `Write an SEO page title (max 60 characters) for this product. Include the product name.\n\n${ctx}`, maxTokens: 60 };
  }
  if (kind === "seoDescription") {
    return { system: brand, user: `Write an SEO meta description (max 155 characters, one sentence or two) for this product.\n\n${ctx}`, maxTokens: 120 };
  }
  return { system: brand, user: `Write a product description of 2 short paragraphs (about 50-90 words total) for this product. Focus on the look, feel and versatility.\n\n${ctx}`, maxTokens: 400 };
}

export const action = async ({ request }: ActionFunctionArgs) => {
  const key = process.env.ANTHROPIC_API_KEY;
  if (!key) return Response.json({ ok: false, error: "AI isn't configured (missing ANTHROPIC_API_KEY)." }, { status: 200 });

  const form = await request.formData();
  const kind = String(form.get("kind") ?? "description");
  const productName = String(form.get("productName") ?? "").trim();
  const productType = String(form.get("productType") ?? "").trim();
  const source = String(form.get("source") ?? "").trim();
  if (!productName) return Response.json({ ok: false, error: "No product name." }, { status: 200 });

  const { system, user, maxTokens } = buildPrompt(kind, productName, productType, source);

  try {
    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "x-api-key": key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
      },
      body: JSON.stringify({
        model: MODEL,
        max_tokens: maxTokens,
        system,
        messages: [{ role: "user", content: user }],
      }),
    });
    if (!res.ok) {
      const detail = await res.text().catch(() => "");
      return Response.json({ ok: false, error: `AI request failed (${res.status}). ${detail.slice(0, 200)}` }, { status: 200 });
    }
    const json = await res.json() as { content?: Array<{ type?: string; text?: string }> };
    const text = (json.content ?? []).filter((c) => c.type === "text").map((c) => c.text ?? "").join("").trim();
    if (!text) return Response.json({ ok: false, error: "AI returned no text." }, { status: 200 });
    return Response.json({ ok: true, text });
  } catch (e) {
    return Response.json({ ok: false, error: (e as Error).message || "AI request errored." }, { status: 200 });
  }
};
