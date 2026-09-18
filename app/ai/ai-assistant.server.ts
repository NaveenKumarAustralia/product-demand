import prisma from "../db.server";

// Karma East Portal AI assistant. General chat + "how do I…" portal help + live
// read-only data lookups (e.g. "how many Pippa dresses are on order"). Per-user
// history is kept permanently in a PortalSetting row keyed by the user id.

const MODEL = process.env.ANTHROPIC_MODEL || "claude-sonnet-5";
const PORTAL_USERS_KEY = "supplier-portal-users-v1";
const PORTAL_USER_COOKIE = "supplier_portal_user";
const HISTORY_KEY = (userId: string) => `ai-chat:${userId}`;
// How many past turns to send to the model (history is stored forever, but we only
// feed a recent window to keep each request affordable).
const CONTEXT_TURNS = 24;

export type AiChatMessage = { role: "user" | "assistant"; content: string; ts: number };

// ── Auth: identify the logged-in portal user from the request cookie ──────────
function cookie(request: Request, key: string): string {
  const raw = request.headers.get("cookie") ?? "";
  for (const part of raw.split(";")) {
    const [k, ...v] = part.trim().split("=");
    if (k === key) return decodeURIComponent(v.join("="));
  }
  return "";
}
export async function currentPortalUser(request: Request): Promise<{ id: string; name: string; admin: boolean } | null> {
  const userId = cookie(request, PORTAL_USER_COOKIE);
  if (!userId) return null;
  const setting = await prisma.portalSetting.findUnique({ where: { key: PORTAL_USERS_KEY }, select: { value: true } }).catch(() => null);
  const users = Array.isArray(setting?.value) ? (setting!.value as Array<Record<string, unknown>>) : [];
  const u = users.find((x) => String(x?.id ?? "") === userId && x?.active !== false);
  if (!u) return null;
  return { id: userId, name: String(u.name ?? "there"), admin: u.admin === true || u.role === "superadmin" };
}

// ── Per-user history (kept forever) ───────────────────────────────────────────
export async function loadHistory(userId: string): Promise<AiChatMessage[]> {
  const s = await prisma.portalSetting.findUnique({ where: { key: HISTORY_KEY(userId) }, select: { value: true } }).catch(() => null);
  const v = s?.value as { messages?: unknown } | null;
  const msgs = Array.isArray(v?.messages) ? v!.messages as AiChatMessage[] : [];
  return msgs.filter((m) => m && (m.role === "user" || m.role === "assistant") && typeof m.content === "string");
}
async function saveHistory(userId: string, messages: AiChatMessage[]): Promise<void> {
  await prisma.portalSetting.upsert({
    where: { key: HISTORY_KEY(userId) },
    create: { key: HISTORY_KEY(userId), value: { messages } },
    update: { value: { messages } },
  }).catch(() => {});
}
export async function clearHistory(userId: string): Promise<void> {
  await saveHistory(userId, []);
}

// ── Live read-only tools the model can call ───────────────────────────────────
const numeric = (s: unknown) => String(s ?? "").replace(/\D/g, "");
const sizeOrder = ["XS", "S", "M", "L", "XL", "2XL", "3XL", "S/M", "M/L", "L/XL", "Free Size"];
const bySizeKey = (a: string) => { const i = sizeOrder.indexOf(a); return i === -1 ? 999 : i; };

const TOOLS = [
  {
    name: "products_on_order",
    description: "Look up how many units of a product are currently on order (open supplier orders not yet fully received), broken down by size. Use for questions like 'how many Pippa dresses are on order' or 'what's on order for Nora Pants'.",
    input_schema: { type: "object", properties: { query: { type: "string", description: "Product name or partial name, e.g. 'Pippa' or 'Nora Pants'." } }, required: ["query"] },
  },
  {
    name: "preorder_reserved",
    description: "Look up how many pre-orders (reservations) are committed for a product against its incoming batch, by size, and how many units are coming in. Use for 'how many pre-orders for Pippa L' or 'is Nora Pants oversold'.",
    input_schema: { type: "object", properties: { query: { type: "string", description: "Product name or partial name." } }, required: ["query"] },
  },
  {
    name: "list_collections",
    description: "List the Collections (product boards) with how many product rows each has. Use for 'what collections do we have' or 'how many products in X collection'.",
    input_schema: { type: "object", properties: { query: { type: "string", description: "Optional: filter collections whose name contains this text." } }, required: [] },
  },
];

async function runTool(name: string, input: Record<string, unknown>): Promise<unknown> {
  try {
    if (name === "products_on_order") {
      const q = String(input.query ?? "").trim();
      if (!q) return { error: "No product name given." };
      const orders = await prisma.supplierOrder.findMany({
        where: { status: "open", productTitle: { contains: q, mode: "insensitive" } },
        select: { productTitle: true, supplier: true, destination: true, eta: true, productId: true, lines: { select: { variantTitle: true, qtyOrdered: true, qtyReceived: true } } },
      });
      if (!orders.length) return { found: 0, note: `No open orders match "${q}".` };
      const byProduct = new Map<string, { product: string; supplier: string; inShopify: boolean; eta: string | null; sizes: Record<string, { ordered: number; received: number; remaining: number }>; totalRemaining: number }>();
      for (const o of orders) {
        const key = (o.productTitle ?? "").trim() || "(untitled)";
        const p = byProduct.get(key) ?? { product: key, supplier: o.supplier ?? "", inShopify: !!numeric(o.productId), eta: o.eta ? o.eta.toISOString().slice(0, 10) : null, sizes: {}, totalRemaining: 0 };
        for (const l of o.lines) {
          const sz = (l.variantTitle ?? "").trim() || "-";
          const rem = Math.max(0, (l.qtyOrdered || 0) - (l.qtyReceived || 0));
          const cell = p.sizes[sz] ?? { ordered: 0, received: 0, remaining: 0 };
          cell.ordered += l.qtyOrdered || 0; cell.received += l.qtyReceived || 0; cell.remaining += rem;
          p.sizes[sz] = cell; p.totalRemaining += rem;
        }
        byProduct.set(key, p);
      }
      return { found: byProduct.size, products: Array.from(byProduct.values()).map((p) => ({ ...p, sizes: Object.fromEntries(Object.entries(p.sizes).sort((a, b) => bySizeKey(a[0]) - bySizeKey(b[0]))) })) };
    }

    if (name === "preorder_reserved") {
      const q = String(input.query ?? "").trim();
      if (!q) return { error: "No product name given." };
      const batches = await prisma.supplierOrder.findMany({
        where: { status: "open", productTitle: { contains: q, mode: "insensitive" } },
        select: { id: true, productTitle: true, lines: { select: { variantTitle: true, qtyOrdered: true, qtyReceived: true } } },
      });
      if (!batches.length) return { found: 0, note: `No batches match "${q}".` };
      const ids = batches.map((b) => b.id);
      const reserved = await prisma.preorderReservation.groupBy({ by: ["supplierOrderId", "variantTitle"], where: { supplierOrderId: { in: ids }, status: "reserved" }, _sum: { quantity: true } });
      const resByBatchSize = new Map<string, number>();
      for (const r of reserved) resByBatchSize.set(`${r.supplierOrderId}::${String(r.variantTitle ?? "").trim()}`, r._sum.quantity ?? 0);
      const out = batches.map((b) => {
        const sizes: Record<string, { incoming: number; reserved: number; available: number }> = {};
        for (const l of b.lines) {
          const sz = (l.variantTitle ?? "").trim() || "-";
          const incoming = Math.max(0, (l.qtyOrdered || 0) - (l.qtyReceived || 0));
          const res = resByBatchSize.get(`${b.id}::${sz}`) ?? 0;
          sizes[sz] = { incoming, reserved: res, available: Math.max(0, incoming - res) };
        }
        const anyReserved = Object.values(sizes).some((s) => s.reserved > 0);
        return { product: b.productTitle, batchId: b.id, anyPreorders: anyReserved, sizes: Object.fromEntries(Object.entries(sizes).sort((a, b2) => bySizeKey(a[0]) - bySizeKey(b2[0]))) };
      }).filter((b) => b.anyPreorders);
      return out.length ? { found: out.length, batches: out } : { found: 0, note: `No pre-orders reserved for "${q}".` };
    }

    if (name === "list_collections") {
      const q = String(input.query ?? "").trim().toLowerCase();
      const rows = await prisma.$queryRawUnsafe<Array<{ name: string; rowCount: number }>>(
        `SELECT name, CASE WHEN jsonb_typeof(rows) = 'array' THEN jsonb_array_length(rows) ELSE 0 END AS "rowCount" FROM "Collection" WHERE "kind" = 'collection' ORDER BY name ASC`,
      ).catch(() => [] as Array<{ name: string; rowCount: number }>);
      const list = rows.map((r) => ({ name: r.name, products: Number(r.rowCount) })).filter((r) => !q || r.name.toLowerCase().includes(q));
      return { count: list.length, collections: list };
    }

    return { error: `Unknown tool ${name}` };
  } catch (e) {
    return { error: e instanceof Error ? e.message : String(e) };
  }
}

// ── The portal how-to guide baked into the system prompt ──────────────────────
const PORTAL_GUIDE = `You are the Karma East Portal Assistant — a helpful in-app AI for the staff of Karma East (a women's fashion brand). You help with general questions, explain how the portal works, and can look up live data with your tools. Be warm, concise and practical. Use the staff member's name occasionally. Format answers with short paragraphs, bullet points and simple tables where helpful. If a data lookup returns nothing, say so plainly and suggest what to check (e.g. spelling of the product name).

WHAT THE PORTAL IS: an internal production + Shopify management portal. Main pages (left nav):
- Existing Products Restock: the master sheet of open production/restock orders for products already in Shopify. Columns per size, status, destination (keep in India / send to AU / send to USA), ETA, cost. A ▼ arrow by the name shows live Shopify inventory. Pre-order "reserved / available" badges show per size.
- JJ Order: the same style of sheet for the JJ supplier.
- JJ New Products: brand-new products (not yet in Shopify) coming from JJ orders — auto-filled here. Fill details and "Create in Shopify".
- Reorder Planner: suggests how much to reorder per size from Shopify stock + sell-through.
- Pre-orders: the pre-order dashboard — batches, reserved units, ship dates. Pre-orders are tracked as reservations against an incoming batch (NOT as a stock "-1"); Shopify inventory can go negative for pre-order items, which is normal.
- Fabric in stock, Packing Lists, Product Information, Samples, Vision Board, Collections, Dropbox.
- Collections: product "boards" where you build new products (duplicate from an existing Shopify product to copy its description/category/etc.), edit them, and "Create in Shopify". Rows can be grouped into folders. Lock/unlock controls two-way Shopify sync (locked = Shopify is source of truth; unlock to push portal edits).

HOW-TO EXAMPLES:
- Create a new product: Collections → Add Collection (or open one) → add a row → set "Duplicate from" a similar Shopify product to copy its description + category metafields → edit → "Create all in Shopify (DRAFT)".
- Change something on a product already in Shopify from the portal: the row is Locked by default (Shopify wins). Click "Unlock to edit" in the Name cell, make changes, then "Update in Shopify" (it re-locks after).
- See live Shopify stock for a restock product: click the ▼ arrow next to its name.
- Pre-order capacity: watch the 🔖 "reserved / available" badge on each size on the restock page — reserved must stay under what's coming in.

RULES: You can only READ data, never change it. Never invent numbers — if you don't have a tool result, say you can't see that yet. Keep answers focused. Today's data is live.`;

// ── The chat turn ─────────────────────────────────────────────────────────────
type ApiContentBlock = { type: string; text?: string; id?: string; name?: string; input?: Record<string, unknown>; tool_use_id?: string; content?: unknown };

export type Attachment = { name: string; mediaType: string; data: string; kind: "image" | "pdf" };

export async function runAssistantTurn(request: Request, userMessage: string, attachments: Attachment[] = []): Promise<{ ok: true; reply: string; history: AiChatMessage[] } | { ok: false; error: string }> {
  const key = process.env.ANTHROPIC_API_KEY;
  if (!key) return { ok: false, error: "AI isn't configured (missing ANTHROPIC_API_KEY)." };
  const user = await currentPortalUser(request);
  if (!user) return { ok: false, error: "Not signed in." };
  const text = userMessage.trim();
  if (!text && !attachments.length) return { ok: false, error: "Empty message." };

  const history = await loadHistory(user.id);
  const recent = history.slice(-CONTEXT_TURNS * 2);
  // This turn's user message = the attachments (image / PDF blocks) + the text.
  // Attachments are only sent for THIS turn (Claude reads them); history stores a
  // text note of what was attached, not the raw bytes (keeps storage small).
  const userContent: unknown[] = [];
  for (const att of attachments) {
    if (att.kind === "image") userContent.push({ type: "image", source: { type: "base64", media_type: att.mediaType, data: att.data } });
    else if (att.kind === "pdf") userContent.push({ type: "document", source: { type: "base64", media_type: "application/pdf", data: att.data } });
  }
  userContent.push({ type: "text", text: text || "Please look at the attached file(s)." });
  const apiMessages: Array<{ role: "user" | "assistant"; content: unknown }> = [
    ...recent.map((m) => ({ role: m.role, content: m.content })),
    { role: "user" as const, content: userContent },
  ];
  const today = new Date().toISOString().slice(0, 10);
  const system = `${PORTAL_GUIDE}\n\nYou are speaking with ${user.name}. Today is ${today}.`;

  const callApi = async (messages: Array<{ role: string; content: unknown }>) => {
    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json" },
      body: JSON.stringify({ model: MODEL, max_tokens: 1500, system, tools: TOOLS, messages }),
    });
    if (!res.ok) throw new Error(`AI request failed (${res.status}). ${(await res.text().catch(() => "")).slice(0, 200)}`);
    return res.json() as Promise<{ content?: ApiContentBlock[]; stop_reason?: string }>;
  };

  try {
    let reply = "";
    for (let step = 0; step < 6; step += 1) {
      const data = await callApi(apiMessages);
      const blocks = data.content ?? [];
      const toolUses = blocks.filter((b) => b.type === "tool_use");
      reply = blocks.filter((b) => b.type === "text").map((b) => b.text ?? "").join("").trim();
      if (data.stop_reason !== "tool_use" || !toolUses.length) break;
      // Execute each requested tool and feed the results back.
      apiMessages.push({ role: "assistant", content: blocks });
      const results = [];
      for (const t of toolUses) {
        const result = await runTool(String(t.name ?? ""), (t.input ?? {}) as Record<string, unknown>);
        results.push({ type: "tool_result", tool_use_id: t.id, content: JSON.stringify(result).slice(0, 12000) });
      }
      apiMessages.push({ role: "user", content: results });
    }
    if (!reply) reply = "Sorry — I couldn't put together an answer for that. Try rephrasing?";
    const now = Date.now();
    const attNote = attachments.length ? `${text ? "\n" : ""}📎 ${attachments.map((a) => a.name).join(", ")}` : "";
    const nextHistory: AiChatMessage[] = [...history, { role: "user", content: `${text}${attNote}`.trim(), ts: now }, { role: "assistant", content: reply, ts: now + 1 }];
    await saveHistory(user.id, nextHistory);
    return { ok: true, reply, history: nextHistory };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "AI request errored." };
  }
}
