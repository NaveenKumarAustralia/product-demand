/**
 * Pure helpers for the MCP write tools (no database import, unit-tested).
 */

export const ACTIVE_SUPPLIER_STATUSES = ["on_order", "on_production", "ready", "in_shipment"] as const;
/** Orders that still represent "stock coming" for duplicate checks. */
export const DUPLICATE_BLOCKING_STATUSES = ["on_order", "on_production"] as const;
export const PENDING_CHANGE_TTL_MS = 30 * 60 * 1000;

const HYPHEN_TO_SLASH: Record<string, string> = { "S-M": "S/M", "M-L": "M/L", "L-XL": "L/XL", "XL-2XL": "XL/2XL" };

/** Normalise a size label for matching: trims, upper-cases, folds XXL→2XL etc. */
export function normalizeSize(label: string): string {
  let s = String(label ?? "").trim().toUpperCase().replace(/\s+/g, "");
  s = s.replace(/^XXXXL$/, "4XL").replace(/^XXXL$/, "3XL").replace(/^XXL$/, "2XL");
  if (HYPHEN_TO_SLASH[s]) s = HYPHEN_TO_SLASH[s];
  return s;
}

/** Find the product's real size label for a requested size, or null. */
export function matchSize(requested: string, available: string[]): string | null {
  const want = normalizeSize(requested);
  return available.find((a) => normalizeSize(a) === want) ?? null;
}

export type SizeQtyChange = { size: string; setTo?: number; add?: number };
export type ResolvedSizeChange = { size: string; before: number; after: number };

/**
 * Turn requested size edits into absolute before/after values. Rejects unknown
 * sizes, negative results and non-integers so a typo can never be applied.
 */
export function resolveQuantityChanges(
  current: Record<string, number>,
  availableSizes: string[],
  changes: SizeQtyChange[],
): { ok: true; changes: ResolvedSizeChange[] } | { ok: false; error: string } {
  if (!changes.length) return { ok: false, error: "No size changes given" };
  const out: ResolvedSizeChange[] = [];
  const seen = new Set<string>();
  for (const c of changes) {
    const size = matchSize(c.size, availableSizes);
    if (!size) return { ok: false, error: `Size "${c.size}" is not one of this product's sizes (${availableSizes.join(", ")})` };
    if (seen.has(size)) return { ok: false, error: `Size ${size} listed twice` };
    seen.add(size);
    const hasSet = c.setTo !== undefined && c.setTo !== null;
    const hasAdd = c.add !== undefined && c.add !== null;
    if (hasSet === hasAdd) return { ok: false, error: `Give exactly one of set_to or add for size ${size}` };
    const before = Math.max(0, Math.round(current[size] ?? 0));
    const value = Number(hasSet ? c.setTo : c.add);
    if (!Number.isInteger(value)) return { ok: false, error: `Quantity for ${size} must be a whole number` };
    const after = hasSet ? value : before + value;
    if (after < 0) return { ok: false, error: `Size ${size} would go below zero (${before} → ${after})` };
    if (after > 100000) return { ok: false, error: `Size ${size} quantity ${after} looks wrong (over 100,000)` };
    out.push({ size, before, after });
  }
  return { ok: true, changes: out };
}

/** Validate a new order's size quantities. */
export function resolveNewOrderQuantities(
  availableSizes: string[],
  quantities: Record<string, number>,
): { ok: true; qtys: Record<string, number>; total: number } | { ok: false; error: string } {
  const qtys: Record<string, number> = {};
  for (const [rawSize, rawQty] of Object.entries(quantities ?? {})) {
    const size = matchSize(rawSize, availableSizes);
    if (!size) return { ok: false, error: `Size "${rawSize}" is not one of this product's sizes (${availableSizes.join(", ")})` };
    const qty = Number(rawQty);
    if (!Number.isInteger(qty) || qty < 0) return { ok: false, error: `Quantity for ${size} must be a whole number ≥ 0` };
    if (qty > 100000) return { ok: false, error: `Quantity for ${size} looks wrong (over 100,000)` };
    if (qty > 0) qtys[size] = (qtys[size] ?? 0) + qty;
  }
  const total = Object.values(qtys).reduce((a, b) => a + b, 0);
  if (total <= 0) return { ok: false, error: "The order needs at least one piece" };
  return { ok: true, qtys, total };
}

export type OrderForDuplicateCheck = { id: number; supplierStatus: string; status: string; totalQty: number; createdAt: Date | string };

/** Open orders for the same product that are still on order / in production. */
export function findPotentialDuplicates(orders: OrderForDuplicateCheck[]): OrderForDuplicateCheck[] {
  return orders.filter((o) => o.status === "open" && (DUPLICATE_BLOCKING_STATUSES as readonly string[]).includes(o.supplierStatus));
}

/** Append a dated note line without losing what was there. */
export function appendNote(existing: string | null | undefined, note: string, date = new Date()): string {
  const stamp = date.toLocaleDateString("en-AU", { day: "2-digit", month: "2-digit", year: "numeric", timeZone: "Australia/Adelaide" });
  const line = `[ChatGPT ${stamp}] ${note.trim()}`;
  const prev = (existing ?? "").trimEnd();
  return prev ? `${prev}\n${line}` : line;
}

/**
 * Resolve requested staff identifiers (ids or names) against the portal user
 * list, for messaging. Case-insensitive; matches an exact id, then a full name,
 * then a unique first name. Never guesses: returns an error that names what was
 * missing or ambiguous rather than picking one, so a message can't go to the
 * wrong person. Only active users are considered.
 */
export function matchStaffRecipients(
  users: Array<{ id: string; name: string; active?: boolean }>,
  requested: string[],
): { ok: true; users: Array<{ id: string; name: string }> } | { ok: false; error: string } {
  const active = users.filter((u) => u.active !== false && u.name.trim());
  const picked = new Map<string, { id: string; name: string }>();
  for (const raw of requested) {
    const q = (raw ?? "").trim();
    if (!q) continue;
    const ql = q.toLowerCase();
    const byId = active.find((u) => u.id === q);
    if (byId) { picked.set(byId.id, { id: byId.id, name: byId.name }); continue; }
    const byFull = active.filter((u) => u.name.trim().toLowerCase() === ql);
    const byFirst = active.filter((u) => u.name.trim().toLowerCase().split(/\s+/)[0] === ql);
    const matches = byFull.length ? byFull : byFirst;
    if (matches.length === 0) return { ok: false, error: `No active staff member matches "${q}". Call list_staff_users for the exact names.` };
    if (matches.length > 1) return { ok: false, error: `"${q}" matches ${matches.length} staff (${matches.map((m) => m.name).join(", ")}). Use a full name or the id from list_staff_users.` };
    picked.set(matches[0].id, { id: matches[0].id, name: matches[0].name });
  }
  if (!picked.size) return { ok: false, error: "No recipients were given." };
  return { ok: true, users: [...picked.values()] };
}

/** Stable string for comparing a "before" snapshot at preview and at confirm time. */
export function stableStringify(value: unknown): string {
  if (value === null || typeof value !== "object") return JSON.stringify(value ?? null);
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
  const obj = value as Record<string, unknown>;
  return `{${Object.keys(obj).sort().map((k) => `${JSON.stringify(k)}:${stableStringify(obj[k])}`).join(",")}}`;
}

export function isValidIsoDate(value: string): boolean {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const d = new Date(`${value}T00:00:00Z`);
  return !Number.isNaN(d.getTime()) && d.toISOString().slice(0, 10) === value;
}

/** Trim very large strings and inline images out of portal data before sending to ChatGPT. */
export function compactForModel(value: unknown, depth = 0): unknown {
  if (depth > 8) return "…";
  if (typeof value === "string") {
    if (value.startsWith("data:")) return "[image omitted]";
    return value.length > 2000 ? `${value.slice(0, 2000)}…` : value;
  }
  if (Array.isArray(value)) return value.slice(0, 500).map((v) => compactForModel(v, depth + 1));
  if (value && typeof value === "object") {
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(value)) out[k] = compactForModel(v, depth + 1);
    return out;
  }
  return value;
}
