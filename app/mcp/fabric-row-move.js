/** Plan a row transfer without mutating either sheet or discarding populated cells. */
/** @param {{headers:string[],rows:string[][]}} sourceSheet @param {{headers:string[],rows:string[][]}} targetSheet @param {number} rowIndex @param {(h:string)=>string} normalize @param {(h:string)=>string} role */
export function planFabricRowMove(sourceSheet,targetSheet,rowIndex,normalize,role){
  const sourceRow=sourceSheet.rows[rowIndex];
  if(!sourceRow || sourceSheet===targetSheet) throw new Error("Fabric row or destination not found");
  const src=sourceSheet.headers, dst=[...targetSheet.headers];
  const hasValue=v=>String(v??"").trim()!=="";
  const keyIndices=h=>h.flatMap((x,i)=>x.trim()==="__rowKey"?[i]:[]);
  const sk=keyIndices(src), tk=keyIndices(dst);
  if(sk.length>1||tk.length>1)throw new Error("Ambiguous fabric row key columns");
  if(sourceRow.slice(src.length).some(hasValue))throw new Error("Fabric row has data outside declared columns; reconcile headers before moving");
  const key=sk.length?sourceRow[sk[0]]??"":"";
  if(hasValue(key)&&!tk.length){
    if(targetSheet.rows.some(r=>r.length>dst.length))throw new Error("Destination has undeclared fabric cells; cannot add row key safely");
    tk.push(dst.length);dst.push("__rowKey");
  }
  const used=new Set();
  const moved=dst.map(()=>"");
  const remaining=[];
  for(let i=0;i<dst.length;i++){
    if(i===tk[0])continue;
    const name=normalize(dst[i]);
    if(!name){remaining.push(i);continue;}
    const matches=src.flatMap((h,j)=>j!==sk[0]&&normalize(h)===name?[j]:[]);
    if(matches.length>1)throw new Error("Ambiguous source fabric column: "+dst[i]);
    if(matches.length===0){remaining.push(i);continue;}
    const j=matches[0];
    if(used.has(j))throw new Error("Duplicate destination fabric column: "+dst[i]);
    used.add(j);moved[i]=sourceRow[j]??"";
  }
  for(const i of remaining){
    const r=role(dst[i]);
    if(!r)continue;
    const matches=src.flatMap((h,j)=>j!==sk[0]&&role(h)===r?[j]:[]);
    if(matches.length>1)throw new Error("Ambiguous fabric column mapping: "+dst[i]);
    if(matches.length===0)continue;
    const j=matches[0];
    if(used.has(j))throw new Error("Fabric column maps to multiple destination columns: "+dst[i]);
    used.add(j);moved[i]=sourceRow[j]??"";
  }
  for(let j=0;j<src.length;j++){
    if(j!==sk[0]&&!used.has(j)&&hasValue(sourceRow[j]))throw new Error("Destination lacks a column for fabric data: "+src[j]);
  }
  if(tk.length)moved[tk[0]]=key;
  return {headers:dst,row:moved};
}
