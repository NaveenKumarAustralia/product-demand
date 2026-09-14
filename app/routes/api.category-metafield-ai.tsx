import type { ActionFunctionArgs } from "react-router";

// AI suggestions for the Collections "Category metafields" editor. Given a product
// and its category attributes (each with the allowed value list), the model picks
// the most fitting value(s) per attribute — only from the allowed values.
// POST form fields:
//   productName, productType, fabric, description (context)
//   attributes = JSON: [{ key, label, isList, allowed: string[] }]
// Returns { ok: true, picks: { [key]: string[] } } (value NAMES from the allowed list).
const MODEL = process.env.ANTHROPIC_MODEL || "claude-sonnet-5";

export const action = async ({ request }: ActionFunctionArgs) => {
  const key = process.env.ANTHROPIC_API_KEY;
  if (!key) return Response.json({ ok: false, error: "AI isn't configured (missing ANTHROPIC_API_KEY)." }, { status: 200 });

  const form = await request.formData();
  const productName = String(form.get("productName") ?? "").trim();
  const productType = String(form.get("productType") ?? "").trim();
  const fabric = String(form.get("fabric") ?? "").trim();
  const description = String(form.get("description") ?? "").trim();
  let attributes: Array<{ key: string; label: string; isList: boolean; allowed: string[] }> = [];
  try { attributes = JSON.parse(String(form.get("attributes") ?? "[]")); } catch { /* ignore */ }
  if (!productName) return Response.json({ ok: false, error: "No product name." }, { status: 200 });
  if (!attributes.length) return Response.json({ ok: false, error: "No attributes." }, { status: 200 });

  // Only send attributes that have allowed values (choice fields).
  const choiceAttrs = attributes.filter((a) => Array.isArray(a.allowed) && a.allowed.length);
  const attrList = choiceAttrs.map((a) => `- "${a.key}" (${a.label})${a.isList ? " [can pick multiple]" : " [pick one]"}: ${a.allowed.join(" | ")}`).join("\n");

  const system = "You classify women's fashion products into Shopify category attributes. You ONLY ever choose from the allowed values given for each attribute — never invent a value. If you're not confident about an attribute, omit it. Base choices on the product name, type, fabric and description. Output ONLY valid JSON, no prose.";
  const user = `Product name: ${productName}${productType ? `\nType: ${productType}` : ""}${fabric ? `\nFabric: ${fabric}` : ""}${description ? `\nDescription: ${description.slice(0, 600)}` : ""}

Attributes and their allowed values:
${attrList}

Return a JSON object mapping each attribute key to an array of chosen value names (exact strings from its allowed list). For single-pick attributes use at most one value. Omit any attribute you're unsure about. Example: {"fabric":["Cotton"],"neckline":["Round"]}`;

  try {
    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json" },
      body: JSON.stringify({ model: MODEL, max_tokens: 900, system, messages: [{ role: "user", content: user }] }),
    });
    if (!res.ok) {
      const detail = await res.text().catch(() => "");
      return Response.json({ ok: false, error: `AI request failed (${res.status}). ${detail.slice(0, 200)}` }, { status: 200 });
    }
    const json = await res.json() as { content?: Array<{ type?: string; text?: string }> };
    const text = (json.content ?? []).filter((c) => c.type === "text").map((c) => c.text ?? "").join("").trim();
    // Extract the JSON object (the model may wrap it in ```).
    const m = text.match(/\{[\s\S]*\}/);
    if (!m) return Response.json({ ok: false, error: "AI returned no suggestions." }, { status: 200 });
    let raw: Record<string, unknown> = {};
    try { raw = JSON.parse(m[0]); } catch { return Response.json({ ok: false, error: "AI returned invalid JSON." }, { status: 200 }); }

    // Validate: keep only allowed value names per attribute.
    const allowedByKey = new Map(choiceAttrs.map((a) => [a.key, new Set(a.allowed.map((v) => v.toLowerCase()))]));
    const nameByKey = new Map(choiceAttrs.map((a) => [a.key, new Map(a.allowed.map((v) => [v.toLowerCase(), v]))]));
    const picks: Record<string, string[]> = {};
    for (const [k, v] of Object.entries(raw)) {
      const allow = allowedByKey.get(k);
      if (!allow) continue;
      const arr = (Array.isArray(v) ? v : [v]).map((x) => String(x));
      const valid = arr.filter((x) => allow.has(x.toLowerCase())).map((x) => nameByKey.get(k)!.get(x.toLowerCase())!);
      if (valid.length) picks[k] = Array.from(new Set(valid));
    }
    return Response.json({ ok: true, picks });
  } catch (e) {
    return Response.json({ ok: false, error: (e as Error).message || "AI request errored." }, { status: 200 });
  }
};
