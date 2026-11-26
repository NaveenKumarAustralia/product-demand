/** Prepare stable row keys on a raw persisted snapshot without changing any other cells.
 * Rows can contain real stock in undeclared (UI-padded) columns. Never place a
 * key over one of those cells or shift its positional meaning. Such sheets
 * need explicit header reconciliation before backfill.
 * @param {unknown} value
 * @param {()=>string} newKey
 */
export function prepareFabricRowKeyBackfill(value, newKey) {
  if (!Array.isArray(value)) return { value, changed: false };
  let changed = false;
  const next = value.map(sheet => {
    if (!sheet || !Array.isArray(sheet.headers) || !Array.isArray(sheet.rows) ||
        !sheet.headers.every(h => typeof h === "string") ||
        !sheet.rows.every(r => Array.isArray(r) && r.every(c => typeof c === "string"))) return sheet;
    const keys = sheet.headers.flatMap((h,i) => h.trim() === "__rowKey" ? [i] : []);
    if (keys.length > 1 || !sheet.rows.length) return sheet;
    const idx = keys[0] ?? sheet.headers.length;
    if (!keys.length && sheet.rows.some(r => r.length > idx)) return sheet;
    if (keys.length && sheet.rows.every(r => (r[idx] ?? "").trim())) return sheet;
    const rows = sheet.rows.map(r => {
      if ((r[idx] ?? "").trim()) return r;
      const copy = [...r];
      while (copy.length <= idx) copy.push("");
      copy[idx] = newKey();
      return copy;
    });
    changed = true;
    return { ...sheet, headers: keys.length ? sheet.headers : [...sheet.headers, "__rowKey"], rows };
  });
  return { value: changed ? next : value, changed };
}
