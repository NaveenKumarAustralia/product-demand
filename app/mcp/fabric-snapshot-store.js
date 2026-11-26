/** Database-level optimistic concurrency for the entire fabric JSON snapshot. */
/** @param {any} store @param {string} key @param {{value:any,updatedAt:Date}|null|undefined} expected @param {any} value */
export async function commitFabricSnapshot(store,key,expected,value) {
  if (!expected || !Array.isArray(expected.value)) throw new Error("Fabric snapshot unavailable; reload before saving.");
  const result = await store.updateMany({where:{key,updatedAt:expected.updatedAt,value:{equals:expected.value}},data:{value}});
  if(result.count!==1) throw new Error("Fabric sheets changed while saving. Reload and retry your edit.");
}

/** Initialise only an absent setting; never replace a concurrent or existing snapshot. */
/** @param {any} store @param {string} key @param {any} value */
export async function createFabricSnapshotIfAbsent(store,key,value) {
  try { return await store.create({data:{key,value},select:{value:true,updatedAt:true}}); }
  catch(error) {
    if(error?.code!=="P2002") throw error;
    const latest=await store.findUnique({where:{key},select:{value:true,updatedAt:true}});
    if(!latest) throw new Error("Fabric sheets changed while initialising. Reload and retry.");
    return latest;
  }
}
