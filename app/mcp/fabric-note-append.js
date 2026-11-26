/** Pure preparation only. Must be called again at confirmation inside the shared fabric lock and transaction. Does not persist data or notify staff. */
/** @param {any[]} sheets @param {{entityKey:string,sheetId:string,name:string,fabricType?:string,note:string}} expected @param {string} message */
export function prepareFabricNoteAppend(sheets, expected, message) {
  const clean = (v) => typeof v === "string" ? v.trim() : "";
  const norm = (v) => clean(v).toLowerCase();
  if (!Array.isArray(sheets) || !expected || !clean(expected.entityKey) || expected.entityKey.includes(":") || !clean(expected.sheetId) || !clean(expected.name) || typeof expected.note !== "string" || typeof message !== "string" || !message.trim() || message.length > 2000) throw new Error("Invalid fabric message snapshot or body");
  const matches = [];
  for (const sheet of sheets) {
    if (!sheet || !Array.isArray(sheet.headers) || !Array.isArray(sheet.rows)) throw new Error("Invalid saved fabric sheet");
    const keys = sheet.headers.map((h,i)=>clean(h)==="__rowKey"?i:-1).filter(i=>i>=0);
    if(keys.length>1) throw new Error("Ambiguous fabric key columns");
    if(!keys.length) continue;
    for(let i=0;i<sheet.rows.length;i++){
      const row=sheet.rows[i];
      if(!Array.isArray(row)) throw new Error("Invalid saved fabric row");
      if(clean(row[keys[0]])===expected.entityKey) matches.push({sheet,rowIndex:i,row});
    }
  }
  if(matches.length!==1) throw new Error("Fabric row missing or duplicated");
  const {sheet,rowIndex,row}=matches[0];
  if(clean(sheet.gid)!==expected.sheetId || ["fabric on order","new fabric on order"].includes(norm(sheet.name))) throw new Error("Fabric row moved or is hidden; preview again");
  const col=(pattern)=>sheet.headers.map((h,i)=>pattern.test(norm(h))?i:-1).filter(i=>i>=0);
  const names=col(/^name$/), types=col(/^(fabric\s*type|type)$/), notes=col(/^notes?$/);
  if(names.length!==1 || types.length>1 || notes.length!==1) throw new Error("Fabric row has ambiguous or missing columns; preview again");
  const name=clean(row[names[0]]), fabricType=types.length?clean(row[types[0]]):"";
  const current=row[notes[0]]??"";
  if(name!==expected.name || fabricType!==(expected.fabricType??"") || typeof current!=="string" || current!==expected.note) throw new Error("Fabric or note changed since preview; preview again");
  const after=current+(current?"\n\n":"")+message.trim();
  return {sheetId:expected.sheetId,rowIndex,noteColumn:notes[0],before:current,after};
}
