// @ts-nocheck
import { test } from "node:test";
import assert from "node:assert/strict";
import { commitFabricSnapshot, createFabricSnapshotIfAbsent } from "./fabric-snapshot-store.js";

test("fabric snapshot CAS rejects stale writes and initialization cannot overwrite existing stock", async () => {
  const copy = x => JSON.parse(JSON.stringify(x));
  let current = {key:"fabric",value:[{rows:[["100"]]}],updatedAt:"v1"};
  let writes = 0;
  const store = {
    updateMany: async ({where,data}) => {
      if (!current || current.key !== where.key || current.updatedAt !== where.updatedAt ||
          JSON.stringify(current.value) !== JSON.stringify(where.value.equals)) return {count:0};
      current = {...current,value:copy(data.value),updatedAt:"v"+(++writes+1)};
      return {count:1};
    },
    create: async ({data}) => {
      if (current) throw Object.assign(new Error("exists"),{code:"P2002"});
      current = {...copy(data),updatedAt:"new"};
      return copy(current);
    },
    findUnique: async () => current && copy(current),
  };
  const baseline = copy(current);
  await commitFabricSnapshot(store,"fabric",baseline,[{rows:[["80"]]}]);
  assert.equal(current.value[0].rows[0][0],"80");
  await assert.rejects(() => commitFabricSnapshot(store,"fabric",baseline,[{rows:[["120"]]}]), /changed while saving/);
  assert.equal(current.value[0].rows[0][0],"80");
  const sameTimestamp = copy(current);
  sameTimestamp.value = [{rows:[["50"]]}];
  await assert.rejects(() => commitFabricSnapshot(store,"fabric",sameTimestamp,[]), /changed while saving/);
  await assert.rejects(() => commitFabricSnapshot(store,"fabric",null,[]), /snapshot unavailable/);
  await assert.rejects(() => commitFabricSnapshot(store,"fabric",{value:{},updatedAt:"v2"},[]), /snapshot unavailable/);
  const existing = await createFabricSnapshotIfAbsent(store,"fabric",[{rows:[["999"]]}]);
  assert.equal(existing.value[0].rows[0][0],"80");
  current = null;
  await createFabricSnapshotIfAbsent(store,"fabric",[]);
  assert.deepEqual(current.value,[]);
  await assert.rejects(() => createFabricSnapshotIfAbsent({create:async()=>{throw Error("connection lost")}},"fabric",[]), /connection lost/);
  await assert.rejects(() => createFabricSnapshotIfAbsent({
    create:async()=>{throw Object.assign(Error("exists"),{code:"P2002"})},
    findUnique:async()=>null
  },"fabric",[]), /changed while initialising/);
});
