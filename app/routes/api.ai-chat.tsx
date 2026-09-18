import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import { currentPortalUser, loadHistory, clearHistory, runAssistantTurn, type Attachment } from "../ai/ai-assistant.server";

const IMAGE_TYPES = new Set(["image/jpeg", "image/png", "image/gif", "image/webp"]);
const MAX_FILE_BYTES = 8 * 1024 * 1024; // 8MB per file

// Per-user AI assistant chat. Shared by the AI Assistant page (side nav) and the
// popup that opens from the AI icon next to the search bar — both read/write the
// SAME per-user conversation, so history is one continuous thread.
//   GET  /api/ai-chat                 → { ok, name, messages }
//   POST { message }                  → { ok, reply, messages }
//   POST { intent: "clear" }          → { ok, messages: [] }

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const user = await currentPortalUser(request);
  if (!user) return Response.json({ ok: false, error: "Not signed in." }, { status: 200 });
  const messages = await loadHistory(user.id);
  return Response.json({ ok: true, name: user.name, messages }, { headers: { "Cache-Control": "no-store" } });
};

export const action = async ({ request }: ActionFunctionArgs) => {
  const user = await currentPortalUser(request);
  if (!user) return Response.json({ ok: false, error: "Not signed in." }, { status: 200 });
  const form = await request.formData();
  const intent = String(form.get("intent") ?? "");
  if (intent === "clear") {
    await clearHistory(user.id);
    return Response.json({ ok: true, messages: [] });
  }
  const message = String(form.get("message") ?? "");
  // Attachments: images (jpeg/png/gif/webp) and PDFs. Read as base64 for the model.
  const attachments: Attachment[] = [];
  const skipped: string[] = [];
  for (const entry of form.getAll("files")) {
    if (!(entry instanceof File) || entry.size === 0) continue;
    if (entry.size > MAX_FILE_BYTES) { skipped.push(`${entry.name} (too big, max 8MB)`); continue; }
    const type = entry.type || "";
    const kind: "image" | "pdf" | null = IMAGE_TYPES.has(type) ? "image" : type === "application/pdf" ? "pdf" : null;
    if (!kind) { skipped.push(`${entry.name} (unsupported — use JPG/PNG/GIF/WebP or PDF)`); continue; }
    const data = Buffer.from(await entry.arrayBuffer()).toString("base64");
    attachments.push({ name: entry.name || (kind === "pdf" ? "file.pdf" : "image"), mediaType: type, data, kind });
  }
  if (!message.trim() && !attachments.length) {
    return Response.json({ ok: false, error: skipped.length ? `Couldn't use: ${skipped.join("; ")}` : "Empty message." }, { status: 200 });
  }
  const result = await runAssistantTurn(request, message, attachments);
  if (!result.ok) return Response.json({ ok: false, error: result.error }, { status: 200 });
  return Response.json({ ok: true, reply: result.reply, messages: result.history, skipped: skipped.length ? skipped : undefined });
};
