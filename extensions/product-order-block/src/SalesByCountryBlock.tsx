import { useEffect, useState, useCallback } from "react";
import {
  reactExtension,
  useApi,
  BlockStack,
  InlineStack,
  Text,
  Divider,
  Badge,
  ProgressIndicator,
  TextField,
  Select,
} from "@shopify/ui-extensions-react/admin";

const TARGET = "admin.product-details.block.render";
const APP_URL = "https://product-demand-production.up.railway.app";

export default reactExtension(TARGET, () => <SalesByCountryBlock />);

type CountryRow = { country: string; unitsSold: number };

const RANGE_OPTIONS = [
  { value: "today", label: "Today" },
  { value: "yesterday", label: "Yesterday" },
  { value: "3d", label: "Last 3 days" },
  { value: "7d", label: "Last 7 days" },
  { value: "14d", label: "Last 14 days" },
  { value: "1m", label: "Last month" },
  { value: "3m", label: "Last 3 months" },
  { value: "1y", label: "Last year" },
  { value: "custom", label: "Custom range" },
];

const pad = (n: number) => String(n).padStart(2, "0");
const fmtISO = (d: Date) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const addDays = (d: Date, n: number) => { const x = new Date(d); x.setDate(x.getDate() + n); return x; };

// Resolve a range key to [since, until] as YYYY-MM-DD.
function rangeToDates(range: string, customFrom: string, customUntil: string): { since: string; until: string } | null {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const untilStr = fmtISO(today);
  switch (range) {
    case "today": return { since: untilStr, until: untilStr };
    case "yesterday": { const y = fmtISO(addDays(today, -1)); return { since: y, until: y }; }
    case "3d": return { since: fmtISO(addDays(today, -2)), until: untilStr };
    case "7d": return { since: fmtISO(addDays(today, -6)), until: untilStr };
    case "14d": return { since: fmtISO(addDays(today, -13)), until: untilStr };
    case "1m": { const s = new Date(today); s.setMonth(s.getMonth() - 1); return { since: fmtISO(s), until: untilStr }; }
    case "3m": { const s = new Date(today); s.setMonth(s.getMonth() - 3); return { since: fmtISO(s), until: untilStr }; }
    case "1y": { const s = new Date(today); s.setFullYear(s.getFullYear() - 1); return { since: fmtISO(s), until: untilStr }; }
    case "custom": {
      const okFrom = /^\d{4}-\d{2}-\d{2}$/.test(customFrom);
      const okUntil = /^\d{4}-\d{2}-\d{2}$/.test(customUntil);
      if (!okFrom || !okUntil) return null;
      return { since: customFrom, until: customUntil };
    }
    default: return { since: fmtISO(addDays(today, -6)), until: untilStr };
  }
}

function SalesByCountryBlock() {
  const { data, auth, query } = useApi(TARGET);
  const productGid: string | undefined = data.selected?.[0]?.id;

  const [shop, setShop] = useState<string | null>(null);
  const [range, setRange] = useState("7d");
  const [customFrom, setCustomFrom] = useState("");
  const [customUntil, setCustomUntil] = useState("");
  const [rows, setRows] = useState<CountryRow[]>([]);
  const [total, setTotal] = useState(0);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // Resolve the shop domain once.
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const result = await query<{ shop: { myshopifyDomain: string } }>(`{ shop { myshopifyDomain } }`);
        if (!cancelled) setShop(result.data?.shop?.myshopifyDomain ?? null);
      } catch { if (!cancelled) setShop(null); }
    })();
    return () => { cancelled = true; };
  }, [query]);

  const load = useCallback(async () => {
    if (!productGid || !shop) return;
    const dates = rangeToDates(range, customFrom, customUntil);
    if (!dates) { setError("Enter both custom dates as YYYY-MM-DD."); setLoading(false); return; }
    setLoading(true);
    setError(null);
    try {
      const token = await auth.idToken();
      if (!token) throw new Error("No auth token");
      const res = await fetch(
        `${APP_URL}/api/sold-by-country?productId=${encodeURIComponent(productGid)}&shop=${encodeURIComponent(shop)}&since=${dates.since}&until=${dates.until}`,
        { headers: { Authorization: `Bearer ${token}` } },
      );
      const json = await res.json();
      if (!res.ok || json.error) throw new Error(json.error ?? "Could not load sales");
      setRows(Array.isArray(json.byCountry) ? json.byCountry : []);
      setTotal(typeof json.total === "number" ? json.total : 0);
    } catch (e: any) {
      setError(`Error: ${e?.message ?? String(e)}`);
      setRows([]);
      setTotal(0);
    } finally {
      setLoading(false);
    }
  }, [productGid, shop, range, customFrom, customUntil, auth]);

  useEffect(() => { load(); }, [load]);

  const showCustom = range === "custom";

  return (
    <BlockStack gap="base">
      <Text fontWeight="bold">Sales by country</Text>

      <Select
        label="Date range"
        value={range}
        onChange={(v: string) => setRange(v)}
        options={RANGE_OPTIONS}
      />
      {showCustom && (
        <InlineStack gap="base">
          <TextField label="From (YYYY-MM-DD)" value={customFrom} onChange={(v: string) => setCustomFrom(v)} />
          <TextField label="To (YYYY-MM-DD)" value={customUntil} onChange={(v: string) => setCustomUntil(v)} />
        </InlineStack>
      )}

      <Divider />

      {loading ? (
        <InlineStack gap="base" blockAlignment="center">
          <ProgressIndicator size="small-200" />
          <Text>Loading…</Text>
        </InlineStack>
      ) : error ? (
        <Text fontWeight="bold">{error}</Text>
      ) : (
        <BlockStack gap="base">
          <InlineStack inlineAlignment="space-between">
            <Text fontWeight="bold">Total sold</Text>
            <Text fontWeight="bold">{total}</Text>
          </InlineStack>
          {rows.length === 0 ? (
            <Text>No sales in this range.</Text>
          ) : (
            rows.map((r) => (
              <InlineStack key={r.country} inlineAlignment="space-between" blockAlignment="center">
                <Text>{r.country}</Text>
                <Badge tone="info">{String(r.unitsSold)}</Badge>
              </InlineStack>
            ))
          )}
        </BlockStack>
      )}
    </BlockStack>
  );
}
