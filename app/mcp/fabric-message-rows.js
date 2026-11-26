/** Read only: locate saved fabric note cells without backfilling IDs or touching sheets. message_ready refers to snapshot preparation only, not live sending. */
/** @param {unknown} value @param {{search:string, fabricType?:string, limit?:number}} options */
export function findFabricMessageRows(value, options) {
  const norm = (s) => typeof s === "string" ? s.trim().toLowerCase().replace(/\s+/g, " ") : "";
  const clean = (s) => typeof s === "string" ? s.trim() : "";
  const search = norm(options.search);
  const type = norm(options.fabricType);
  const limit = options.limit ?? 30;
  if (!search || !Number.isInteger(limit) || limit < 1 || limit > 100) throw new Error("A fabric name search and valid limit are required.");
  if (!Array.isArray(value) || !value.length) return { status: "manual_sheets_unavailable", rows: [], truncated: false };
  const sheets = value.filter((s) => s && typeof s === "object" && clean(s.gid) && clean(s.name) && Array.isArray(s.headers) && Array.isArray(s.rows));
  if (!sheets.length) return { status: "manual_sheets_unavailable", rows: [], truncated: false };
  const counts = new Map();
  const rows = [];
  let invalidLayout = sheets.length !== value.length;
  for (const sheet of sheets) {
    const cols = (predicate) => sheet.headers.flatMap((h,i) => predicate(h) ? [i] : []);
    const keys = cols(h => clean(h) === "__rowKey");
    const names = cols(h => norm(h) === "name");
    const types = cols(h => /^(fabric\s*type|type)$/.test(norm(h)));
    const notes = cols(h => /^notes?$/.test(norm(h)));
    if (keys.length > 1) invalidLayout = true;
    const hidden = ["new fabric on order", "fabric on order"].includes(norm(sheet.name));
    for (const row of sheet.rows) {
      if (!Array.isArray(row)) { invalidLayout = true; continue; }
      const key = clean(row[keys[0]]);
      if (key) counts.set(key, (counts.get(key) ?? 0) + 1);
      if (hidden || !names.length) continue;
      const name = clean(row[names[0]]);
      const fabricType = clean(row[types[0]]);
      if (!name || !norm(name).includes(search) || (type && norm(fabricType) !== type)) continue;
      const layoutBlocker = names.length !== 1 || types.length > 1 || notes.length > 1 ? "ambiguous_columns" : notes.length === 0 ? "missing_notes_column" : typeof (row[notes[0]] ?? "") !== "string" ? "invalid_note_value" : null;
      rows.push({ name, fabric_type: fabricType || null, sheet_name: clean(sheet.name), sheet_id: clean(sheet.gid), sheet_kind: clean(sheet.kind) || "stock", key, layoutBlocker });
    }
  }
  return {
    status: "ok",
    rows: rows.slice(0, limit).map(({ key, layoutBlocker, ...row }) => {
      const keyStatus = !key ? "missing" : counts.get(key) !== 1 ? "duplicate" : key.includes(":") ? "invalid" : "unique";
      const blocker = invalidLayout ? "invalid_sheet_layout" : keyStatus !== "unique" ? `key_${keyStatus}` : layoutBlocker;
      return { ...row, entity_type: "fabric_row", order_id: 0, field: "notes",
        entity_key: keyStatus === "unique" && !invalidLayout ? key : null,
        key_status: keyStatus, message_ready: blocker === null, message_blocker: blocker };
    }),
    truncated: rows.length > limit,
  };
}
