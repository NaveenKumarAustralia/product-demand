import { useEffect, useState } from "react";

type ScopeItem = {
  handle: string;
  label: string;
  protected: boolean;
  granted: boolean;
};

type ShopReadiness = {
  shop: string;
  shopName?: string | null;
  myshopifyDomain?: string | null;
  ok: boolean;
  ready: boolean;
  missingScopes?: string[];
  scopes: ScopeItem[];
  error?: string | null;
};

type ResponseShape = {
  ok: boolean;
  ready?: boolean;
  shops?: ShopReadiness[];
  error?: string;
};

export function PreorderShopifyReadinessPanel() {
  const [data, setData] = useState<ResponseShape | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let active = true;
    fetch("/api/preorder-shopify-readiness", { credentials: "same-origin" })
      .then(async (response) => {
        const json = await response.json().catch(() => ({ ok: false, error: "Could not read Shopify readiness response." }));
        if (!response.ok || json.ok !== true) throw new Error(json.error || "Could not check Shopify preorder readiness.");
        return json as ResponseShape;
      })
      .then((result) => { if (active) setData(result); })
      .catch((error) => { if (active) setData({ ok: false, error: error instanceof Error ? error.message : "Could not check Shopify preorder readiness." }); })
      .finally(() => { if (active) setLoading(false); });
    return () => { active = false; };
  }, []);

  if (loading) return <div style={s.card}>Checking Shopify preorder permissions…</div>;
  if (!data?.ok) return <div style={{ ...s.card, ...s.error }}>{data?.error || "Could not check Shopify preorder readiness."}</div>;

  const shops = data.shops ?? [];
  return (
    <div style={s.card}>
      <div style={s.header}>
        <div>
          <div style={s.title}>Shopify preorder readiness</div>
          <div style={s.muted}>Live check of the Shopify API permissions required before customer-facing preorders can be activated.</div>
        </div>
        <div style={{ ...s.pill, ...(data.ready ? s.ready : s.blocked) }}>{data.ready ? "Ready" : "Action required"}</div>
      </div>

      {!shops.length ? <div style={s.notice}>No production Shopify shop was found yet.</div> : null}

      {shops.map((shop) => (
        <div key={shop.shop} style={s.shopBlock}>
          <div style={s.shopTop}>
            <div>
              <strong>{shop.shopName || shop.shop}</strong>
              <div style={s.small}>{shop.myshopifyDomain || shop.shop}</div>
            </div>
            <span style={{ ...s.smallPill, ...(shop.ready ? s.ready : s.blocked) }}>{shop.ready ? "Purchase options ready" : "Missing permissions"}</span>
          </div>

          {shop.error ? <div style={s.errorBox}>{shop.error}</div> : (
            <div style={s.scopeGrid}>
              {shop.scopes.map((scope) => (
                <div key={scope.handle} style={s.scopeRow}>
                  <span style={{ ...s.dot, ...(scope.granted ? s.dotOk : s.dotMissing) }}>{scope.granted ? "✓" : "!"}</span>
                  <div>
                    <div style={s.scopeLabel}>{scope.label}</div>
                    <div style={s.small}>{scope.handle}{scope.protected ? " · Shopify approval required" : ""}</div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      ))}

      {!data.ready ? (
        <div style={s.warning}>
          Customer-facing preorder selling stays OFF. Missing Shopify purchase-option permissions must be granted before the app creates selling plans.
        </div>
      ) : (
        <div style={s.success}>Shopify has all required deferred-purchase scopes. The next activation step can safely create Karma East preorder selling plans.</div>
      )}
    </div>
  );
}

const s: Record<string, React.CSSProperties> = {
  card: { background: "white", border: "1px solid #e2e8f0", borderRadius: 14, padding: 18, marginBottom: 12, boxShadow: "0 1px 3px rgba(15,23,42,.04)" },
  header: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 14 },
  title: { fontSize: 16, fontWeight: 800, color: "#0f172a" },
  muted: { fontSize: 12, color: "#64748b", marginTop: 4, lineHeight: 1.5 },
  small: { fontSize: 11, color: "#94a3b8", marginTop: 2 },
  pill: { borderRadius: 999, padding: "6px 9px", fontSize: 11, fontWeight: 800, whiteSpace: "nowrap" },
  smallPill: { borderRadius: 999, padding: "5px 8px", fontSize: 10, fontWeight: 800, whiteSpace: "nowrap" },
  ready: { background: "#dcfce7", color: "#166534" },
  blocked: { background: "#fef3c7", color: "#92400e" },
  shopBlock: { marginTop: 14, border: "1px solid #e2e8f0", borderRadius: 10, padding: 13 },
  shopTop: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12 },
  scopeGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 8, marginTop: 12 },
  scopeRow: { display: "flex", gap: 8, alignItems: "center", padding: 8, borderRadius: 8, background: "#f8fafc" },
  dot: { width: 22, height: 22, borderRadius: 999, display: "grid", placeItems: "center", fontWeight: 900, fontSize: 12, flex: "0 0 auto" },
  dotOk: { background: "#dcfce7", color: "#166534" },
  dotMissing: { background: "#fee2e2", color: "#991b1b" },
  scopeLabel: { fontSize: 12, fontWeight: 700, color: "#334155" },
  warning: { marginTop: 13, padding: "10px 12px", borderRadius: 9, background: "#fff7ed", color: "#9a3412", fontSize: 12, fontWeight: 700 },
  success: { marginTop: 13, padding: "10px 12px", borderRadius: 9, background: "#f0fdf4", color: "#166534", fontSize: 12, fontWeight: 700 },
  notice: { marginTop: 12, color: "#64748b", fontSize: 12 },
  error: { color: "#991b1b", background: "#fff7f7", borderColor: "#fecaca" },
  errorBox: { marginTop: 10, padding: 9, borderRadius: 8, background: "#fef2f2", color: "#991b1b", fontSize: 12 },
};
