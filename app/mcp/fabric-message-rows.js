/** Read only: locate saved fabric note cells without backfilling IDs or touching sheets. */
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
  for (const sheet of sheets) {
    const keyIdx = sheet.headers.findIndex((h) => clean(h) === "__rowKey");
    const nameIdx = sheet.headers.findIndex((h) => norm(h) === "name");
    const typeIdx = sheet.headers.findIndex((h) => /^(fabric\s*type|type)$/.test(norm(h)));
    const hidden = ["new fabric on order", "fabric on order"].includes(norm(sheet.name));
    for (const row of sheet.rows) {
      if (!Array.isArray(row)) continue;
      const key = clean(row[keyIdx]);
      if (key) counts.set(key, (counts.get(key) ?? 0) + 1);
      if (hidden || nameIdx < 0) continue;
      const name = clean(row[nameIdx]);
      const fabricType = clean(row[typeIdx]);
      if (!name || !norm(name).includes(search) || (type && norm(fabricType) !== type)) continue;
      rows.push({ name, fabric_type: fabricType || null, sheet_name: clean(sheet.name), sheet_id: clean(sheet.gid), sheet_kind: clean(sheet.kind) || "stock", key });
    }
  }
  return {
    status: "ok",
    rows: rows.slice(0, limit).map(({ key, ...row }) => ({
      ...row,
      entity_type: "fabric_row", order_id: 0, field: "notes",
      entity_key: key && counts.get(key) === 1 ? key : null,
      key_status: !key ? "missing" : counts.get(key) !== 1 ? "duplicate" : "unique",
    })),
    truncated: rows.length > limit,
  };
}
