import { useEffect, useRef, useState } from "react";
import { useFetcher } from "react-router";

type Msg = { role: "user" | "assistant"; content: string; ts?: number };

// Very small markdown-ish renderer: paragraphs, "- " bullets and **bold**. Enough
// for the assistant's short answers without pulling in a markdown dependency.
function renderText(text: string) {
  const lines = (text || "").split(/\n/);
  const out: React.ReactNode[] = [];
  let bullets: string[] = [];
  const flush = (key: string) => {
    if (!bullets.length) return;
    out.push(<ul key={key} style={{ margin: "4px 0", paddingLeft: 18 }}>{bullets.map((b, i) => <li key={i} style={{ margin: "2px 0" }}>{inline(b)}</li>)}</ul>);
    bullets = [];
  };
  lines.forEach((ln, i) => {
    const m = ln.match(/^\s*[-*•]\s+(.*)$/);
    if (m) { bullets.push(m[1]); return; }
    flush(`ul${i}`);
    if (ln.trim()) out.push(<p key={i} style={{ margin: "4px 0" }}>{inline(ln)}</p>);
  });
  flush("ul-last");
  return out;
}
function inline(s: string): React.ReactNode {
  const parts = s.split(/(\*\*[^*]+\*\*)/g);
  return parts.map((p, i) => p.startsWith("**") && p.endsWith("**")
    ? <strong key={i}>{p.slice(2, -2)}</strong>
    : <span key={i}>{p}</span>);
}

export default function AiChat({ variant = "page", onClose }: { variant?: "page" | "popup"; onClose?: () => void }) {
  const loadFetcher = useFetcher<{ ok: boolean; name?: string; messages?: Msg[] }>();
  const sendFetcher = useFetcher<{ ok: boolean; reply?: string; messages?: Msg[]; error?: string }>();
  const clearFetcher = useFetcher<{ ok: boolean; messages?: Msg[] }>();
  const [messages, setMessages] = useState<Msg[]>([]);
  const [name, setName] = useState<string>("");
  const [input, setInput] = useState("");
  const [files, setFiles] = useState<File[]>([]);
  const [dragOver, setDragOver] = useState(false);
  const [pending, setPending] = useState<string | null>(null); // optimistic user message
  const scrollRef = useRef<HTMLDivElement | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const loadedRef = useRef(false);
  const addFiles = (list: FileList | File[] | null) => {
    if (!list) return;
    const ok = Array.from(list).filter((f) => f.type.startsWith("image/") || f.type === "application/pdf");
    if (ok.length) setFiles((prev) => [...prev, ...ok].slice(0, 6));
  };

  useEffect(() => { if (!loadedRef.current) { loadedRef.current = true; loadFetcher.load("/api/ai-chat"); } }, [loadFetcher]);
  useEffect(() => {
    if (loadFetcher.state === "idle" && loadFetcher.data?.ok) {
      setMessages(loadFetcher.data.messages ?? []);
      setName(loadFetcher.data.name ?? "");
    }
  }, [loadFetcher.state, loadFetcher.data]);
  useEffect(() => {
    if (sendFetcher.state === "idle" && sendFetcher.data) {
      if (sendFetcher.data.ok && sendFetcher.data.messages) setMessages(sendFetcher.data.messages);
      else if (sendFetcher.data.error) setMessages((m) => [...m, { role: "assistant", content: `⚠️ ${sendFetcher.data!.error}` }]);
      setPending(null);
    }
  }, [sendFetcher.state, sendFetcher.data]);
  useEffect(() => {
    if (clearFetcher.state === "idle" && clearFetcher.data?.ok) setMessages([]);
  }, [clearFetcher.state, clearFetcher.data]);

  const sending = sendFetcher.state !== "idle";
  useEffect(() => {
    const el = scrollRef.current; if (el) el.scrollTop = el.scrollHeight;
  }, [messages, pending, sending]);

  const send = () => {
    const text = input.trim();
    if ((!text && !files.length) || sending) return;
    setPending(`${text}${files.length ? `${text ? "\n" : ""}📎 ${files.map((f) => f.name).join(", ")}` : ""}`.trim());
    const fd = new FormData();
    fd.set("message", text);
    files.forEach((f) => fd.append("files", f));
    setInput("");
    setFiles([]);
    sendFetcher.submit(fd, { method: "post", action: "/api/ai-chat", encType: "multipart/form-data" });
  };

  const shell: React.CSSProperties = variant === "popup"
    ? { position: "fixed", right: 20, bottom: 20, width: 400, maxWidth: "calc(100vw - 40px)", height: 560, maxHeight: "calc(100vh - 120px)", zIndex: 2147483000, borderRadius: 14, boxShadow: "0 12px 40px rgba(0,0,0,0.28)", background: "#fff", display: "flex", flexDirection: "column", overflow: "hidden", border: "1px solid #e5e7eb" }
    : { position: "relative", display: "flex", flexDirection: "column", height: "100%", minHeight: 0, background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", overflow: "hidden" };

  return (
    <div
      style={shell}
      onDragOver={(e) => { e.preventDefault(); if (!dragOver) setDragOver(true); }}
      onDragLeave={(e) => { if (e.currentTarget === e.target) setDragOver(false); }}
      onDrop={(e) => { e.preventDefault(); setDragOver(false); addFiles(e.dataTransfer.files); }}
    >
      {dragOver && (
        <div style={{ position: "absolute", inset: 0, zIndex: 20, background: "rgba(15,118,110,0.12)", border: "2px dashed #0f766e", borderRadius: 12, display: "flex", alignItems: "center", justifyContent: "center", color: "#0f766e", fontWeight: 700, pointerEvents: "none" }}>Drop an image or PDF…</div>
      )}
      <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "10px 14px", borderBottom: "1px solid #eef0f2", background: "#0f766e", color: "#fff", flexShrink: 0 }}>
        <span style={{ fontSize: 16 }}>✨</span>
        <span style={{ fontWeight: 700, fontSize: 14 }}>AI Assistant</span>
        <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center" }}>
          <button type="button" onClick={() => { if (window.confirm("Clear your whole chat history? This can't be undone.")) clearFetcher.submit({ intent: "clear" }, { method: "post", action: "/api/ai-chat" }); }} title="Clear history" style={{ background: "transparent", border: "1px solid rgba(255,255,255,0.5)", color: "#fff", borderRadius: 6, padding: "3px 8px", fontSize: 11, fontWeight: 600, cursor: "pointer" }}>Clear</button>
          {variant === "popup" && onClose && (
            <button type="button" onClick={onClose} title="Close" style={{ background: "transparent", border: "none", color: "#fff", fontSize: 18, lineHeight: 1, cursor: "pointer", padding: 0 }}>×</button>
          )}
        </div>
      </div>

      <div ref={scrollRef} style={{ flex: 1, minHeight: 0, overflowY: "auto", padding: 14, display: "flex", flexDirection: "column", gap: 10, background: "#f8fafc" }}>
        {messages.length === 0 && !pending && (
          <div style={{ color: "#6b7280", fontSize: 13, lineHeight: 1.6, margin: "auto 0", textAlign: "center", padding: "0 12px" }}>
            <div style={{ fontSize: 30, marginBottom: 8 }}>✨</div>
            <div style={{ fontWeight: 700, color: "#374151", marginBottom: 6 }}>Hi{name ? ` ${name}` : ""} — how can I help?</div>
            Ask me anything, or try:
            <div style={{ marginTop: 8, display: "flex", flexDirection: "column", gap: 6 }}>
              {["How many Pippa dresses are on order?", "How do I create a new product in the portal?", "What collections do we have?"].map((s) => (
                <button key={s} type="button" onClick={() => setInput(s)} style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 8, padding: "7px 10px", fontSize: 12, color: "#374151", cursor: "pointer", textAlign: "left" }}>{s}</button>
              ))}
            </div>
          </div>
        )}
        {messages.map((m, i) => (
          <div key={i} style={{ alignSelf: m.role === "user" ? "flex-end" : "flex-start", maxWidth: "85%" }}>
            <div style={{ padding: "8px 12px", borderRadius: 12, fontSize: 13.5, lineHeight: 1.5, background: m.role === "user" ? "#0f766e" : "#fff", color: m.role === "user" ? "#fff" : "#111827", border: m.role === "user" ? "none" : "1px solid #e5e7eb", whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
              {m.role === "user" ? m.content : <div>{renderText(m.content)}</div>}
            </div>
          </div>
        ))}
        {pending && (
          <div style={{ alignSelf: "flex-end", maxWidth: "85%" }}>
            <div style={{ padding: "8px 12px", borderRadius: 12, fontSize: 13.5, background: "#0f766e", color: "#fff", opacity: 0.7, whiteSpace: "pre-wrap" }}>{pending}</div>
          </div>
        )}
        {sending && (
          <div style={{ alignSelf: "flex-start" }}>
            <div style={{ padding: "8px 12px", borderRadius: 12, fontSize: 13, background: "#fff", color: "#6b7280", border: "1px solid #e5e7eb" }}>Thinking…</div>
          </div>
        )}
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: 8, padding: 12, borderTop: "1px solid #eef0f2", flexShrink: 0, background: "#fff" }}>
        {files.length > 0 && (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
            {files.map((f, i) => (
              <span key={i} style={{ display: "inline-flex", alignItems: "center", gap: 6, background: "#f1f5f9", border: "1px solid #e2e8f0", borderRadius: 6, padding: "3px 8px", fontSize: 11, color: "#334155", maxWidth: 180 }}>
                <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{f.type === "application/pdf" ? "📄" : "🖼️"} {f.name}</span>
                <button type="button" onClick={() => setFiles((prev) => prev.filter((_, j) => j !== i))} style={{ background: "none", border: "none", color: "#94a3b8", cursor: "pointer", fontSize: 13, lineHeight: 1, padding: 0 }}>×</button>
              </span>
            ))}
          </div>
        )}
        <div style={{ display: "flex", gap: 8, alignItems: "flex-end" }}>
          <input ref={fileInputRef} type="file" accept="image/*,application/pdf" multiple style={{ display: "none" }} onChange={(e) => { addFiles(e.target.files); e.target.value = ""; }} />
          <button type="button" onClick={() => fileInputRef.current?.click()} title="Attach image or PDF" style={{ background: "#fff", border: "1px solid #d1d5db", borderRadius: 8, width: 38, height: 38, fontSize: 17, cursor: "pointer", color: "#475569", flexShrink: 0 }}>📎</button>
          <textarea
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); } }}
            onPaste={(e) => { const imgs = Array.from(e.clipboardData.files).filter((f) => f.type.startsWith("image/")); if (imgs.length) { e.preventDefault(); addFiles(imgs); } }}
            placeholder="Ask anything, or drop an image / PDF…"
            rows={1}
            style={{ flex: 1, resize: "none", border: "1px solid #d1d5db", borderRadius: 8, padding: "9px 11px", fontSize: 13.5, fontFamily: "inherit", outline: "none", maxHeight: 120 }}
          />
          <button type="button" onClick={send} disabled={sending || (!input.trim() && !files.length)} style={{ background: sending || (!input.trim() && !files.length) ? "#9ca3af" : "#0f766e", color: "#fff", border: "none", borderRadius: 8, padding: "0 16px", height: 38, fontSize: 14, fontWeight: 700, cursor: sending || (!input.trim() && !files.length) ? "default" : "pointer" }}>➤</button>
        </div>
      </div>
    </div>
  );
}
