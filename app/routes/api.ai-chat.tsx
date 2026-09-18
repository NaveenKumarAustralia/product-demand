import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import { currentPortalUser, loadHistory, clearHistory, runAssistantTurn } from "../ai/ai-assistant.server";

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
  const result = await runAssistantTurn(request, message);
  if (!result.ok) return Response.json({ ok: false, error: result.error }, { status: 200 });
  return Response.json({ ok: true, reply: result.reply, messages: result.history });
};
