/** Plan a fabric note edit without modifying any cell. Reject ambiguous keys and unsafe new-column layouts. */
/** @param {{headers:string[],rows:string[][]}[]} sheets @param {{headers:string[],rows:string[][]}} sheet @param {number} rowIndex @param {()=>string} newKey */
export function fabricNoteCellPlan(sheets, sheet, rowIndex, newKey) {
  const cols = (headers, pattern) => headers.flatMap((h,i)=>pattern.test(h.trim())?[i]:[]);
  if (!sheets.includes(sheet) || !sheet.rows[rowIndex]) throw new Error("Fabric row not found");
  const notes=cols(sheet.headers,/^notes?$/i), keys=cols(sheet.headers,/^__rowKey$/);
  if (notes.length>1 || keys.length>1) throw new Error("Ambiguous fabric note or key columns");
  if ((!notes.length || !keys.length) && sheet.rows.some(r=>r.length>sheet.headers.length)) throw new Error("Undeclared fabric cells; reconcile headers before adding note columns");
  for(const other of sheets) {
    if (cols(other.headers,/^__rowKey$/).length>1) throw new Error("Ambiguous fabric key columns");
  }
  const noteIndex=notes[0]??sheet.headers.length;
  const keyIndex=keys[0]??(sheet.headers.length+(notes.length?0:1));
  const row=sheet.rows[rowIndex];
  const key=(keys.length?row[keyIndex]??"":"").trim()||newKey();
  if(!key || key.includes(":")) throw new Error("Invalid fabric row key");
  let matches=0;
  for(const other of sheets){
    const idx=cols(other.headers,/^__rowKey$/)[0];
    if(idx===undefined)continue;
    for(let i=0;i<other.rows.length;i++) {
      if(other===sheet&&i===rowIndex)continue;
      if((other.rows[i][idx]??"").trim()===key) matches++;
    }
  }
  if(matches)throw new Error("Duplicate fabric row key");
  return {noteIndex,keyIndex,key,addNotes:!notes.length,addKey:!keys.length};
}

/** @param {{headers:string[],rows:string[][]}} sheet @param {number} rowIndex @param {ReturnType<typeof fabricNoteCellPlan>} plan @param {string} value */
export function applyFabricNoteCell(sheet,rowIndex,plan,value) {
  if(plan.addNotes)sheet.headers=[...sheet.headers,"Notes"];
  if(plan.addKey)sheet.headers=[...sheet.headers,"__rowKey"];
  for(const r of sheet.rows)while(r.length<sheet.headers.length)r.push("");
  const row=sheet.rows[rowIndex];
  row[plan.noteIndex]=value;
  row[plan.keyIndex]=plan.key;
}

/** @template {{headers:string[],rows:string[][]}} T @param {T[]} sheets @param {string} key */
export function findUniqueFabricThreadRow(sheets,key) {
  if(!key || key.includes(":"))throw new Error("Invalid fabric row key");
  const matches=[];
  for(const sheet of sheets) {
    const keys=sheet.headers.flatMap((h,i)=>h.trim()==="__rowKey"?[i]:[]);
    if(keys.length>1)throw new Error("Ambiguous fabric key columns");
    if(!keys.length)continue;
    sheet.rows.forEach((row,rowIndex)=>{if((row[keys[0]]??"").trim()===key)matches.push({sheet,rowIndex});});
  }
  if(matches.length!==1)throw new Error("Fabric row missing or duplicated");
  return matches[0];
}
