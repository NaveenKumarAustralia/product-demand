import { useEffect, useState } from "react";

type StatusResponse = {
  ok: boolean;
  healthy?: boolean;
  origin?: string;
  error?: string;
  shops?: Array<{
    shop: string;
    ok: boolean;
    error: string | null;
    subscriptions: Array<{
      topic: string;
      label: string;
      uri: string;
      registered: boolean;
      id: string | null;
    }>;
  }>;
};

export function PreorderWebhookStatusPanel() {
  const [status, setStatus] = useState<StatusResponse | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let active = true;
    fetch("/api/preorder-webhook-status", { credentials: "same-origin" })
      .then(async (response) => {
        const result = await response.json().catch(() => ({ ok: false, error: "Could not read webhook status." }));
        if (!response.ok || result.ok !== true) throw new Error(result.error || "Could not check webhook status.");
        return result as StatusResponse;
      })
      .then((result) => { if (active) setStatus(result); })
      .catch((error) => { if (active) setStatus({ ok: false, error: error instanceof Error ? error.message : "Could not check webhook status." }); })
      .finally(() => { if (active) setLoading(false); });
    return () => { active = false; };
  }, []);

  if (loading) return <div style={s.card}><div style={s.title}>Shopify order connection</div><div style={s.muted}>Checking webhook registration…</div></div>;
  if (!status?.ok) {
    return <div style={s.card}><div style={s.title}>Shopify order connection</div><div style={s.error}>{status?.error || "Could not check webhook registration."}</div></div>;
  }

  return (
    <div style={s.card}>
      <div style={s.header}>
        <div>
          <div style={s.title}>Shopify order connection</div>
          <div style={s.muted}>Preorder reservations depend on these three verified Shopify order webhooks.</div>
        </div>
        <div style={{ ...s.badge, ...(status.healthy ? s.good : s.bad) }}>{status.healthy ? "Connected" : "Needs attention"}</div>
      </div>
      {(status.shops ?? []).map((shop) => (
        <div key={shop.shop} style={s.shop}>
          <div style={s.shopName}>{shop.shop}</div>
          {shop.error ? <div style={s.error}>{shop.error}</div> : (
            <div style={s.grid}>
              {shop.subscriptions.map((item) => (
                <div key={item.topic} style={s.item}>
                  <span style={{ ...s.dot, ...(item.registered ? s.dotGood : s.dotBad) }} />
                  <div><strong>{item.label}</strong><div style={s.small}>{item.registered ? "Registered" : "Missing"}</div></div>
                </div>
              ))}
            </div>
          )}
        </div>
      ))}
      {!status.shops?.length ? <div style={s.muted}>No production shop was found yet.</div> : null}
    </div>
  );
}

const s: Record<string, React.CSSProperties> = {
  card: { background: "white", border: "1px solid #e2e8f0", borderRadius: 14, padding: 18, boxShadow: "0 1px 3px rgba(15,23,42,.04)", marginBottom: 12 },
  header: { display: "flex", justifyContent: "space-between", gap: 14, alignItems: "flex-start" },
  title: { fontSize: 16, fontWeight: 800, color: "#0f172a" },
  muted: { marginTop: 4, fontSize: 12, color: "#64748b", lineHeight: 1.5 },
  badge: { padding: "5px 8px", borderRadius: 999, fontSize: 11, fontWeight: 800, whiteSpace: "nowrap" },
  good: { background: "#dcfce7", color: "#166534" },
  bad: { background: "#fef2f2", color: "#991b1b" },
  shop: { marginTop: 14, paddingTop: 12, borderTop: "1px solid #f1f5f9" },
  shopName: { fontSize: 12, fontWeight: 800, color: "#334155", marginBottom: 8 },
  grid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(170px, 1fr))", gap: 8 },
  item: { display: "flex", gap: 8, alignItems: "center", padding: "8px 9px", border: "1px solid #e2e8f0", borderRadius: 9, fontSize: 12 },
  dot: { width: 9, height: 9, borderRadius: 999, flex: "0 0 auto" },
  dotGood: { background: "#16a34a" },
  dotBad: { background: "#dc2626" },
  small: { marginTop: 2, fontSize: 10, color: "#94a3b8" },
  error: { marginTop: 6, fontSize: 12, color: "#991b1b" },
};
