// @ts-nocheck
import { test } from "node:test";
import assert from "node:assert/strict";
import { planFabricRowMove } from "./fabric-row-move.js";
const n = h=>h.toLowerCase().trim().replace(/[^a-z0-9]+/g,"");
const role = h=>/^(meters|qty)$/i.test(h)?"quantity":n(h);
const sheet = (headers,rows)=>({headers,rows});
function runFabricRowMoveCases(plan,eq,throws){
  const normal=sheet(["Name","Qty","__rowKey"],[["Olive","40","id"]]);
  const target=sheet(["Meters","Name"],[["20","Daisy"]]);
  eq(plan(normal,target,0,n,role),{headers:["Meters","Name","__rowKey"],row:["40","Olive","id"]});
  eq(target,sheet(["Meters","Name"],[["20","Daisy"]]));
  eq(normal,sheet(["Name","Qty","__rowKey"],[["Olive","40","id"]]));
  eq(plan(normal,sheet(["__rowKey","Name","Qty"],[]),0,n,role).row,["id","Olive","40"]);
  throws(()=>plan(normal,sheet(["Name"],[]),0,n,role),/lacks a column/);
  throws(()=>plan(normal,sheet(["Meters","Name"],[["20","Daisy","hidden stock"]]),0,n,role),/undeclared/);
  throws(()=>plan(sheet(["Name"],[["Olive","40"]]),sheet(["Name"],[]),0,n,role),/outside declared/);
  throws(()=>plan(sheet(["Name","Name"],[["Olive","Daisy"]]),sheet(["Name"],[]),0,n,role),/Ambiguous source/);
  throws(()=>plan(normal,sheet(["Name","Qty","Meters"],[]),0,n,role),/multiple destination/);
  throws(()=>plan(sheet(["Name","Qty","Meters"],[["Olive","40","50"]]),sheet(["Name","Meters "],[]),0,n,role),/lacks a column/);
  throws(()=>plan(normal,sheet(["Name","Qty","__rowKey","__rowKey"],[]),0,n,role),/Ambiguous fabric row key/);
  eq(plan(sheet(["Name","Notes"],[["Olive",""]]),sheet(["Name"],[]),0,n,role).row,["Olive"]);
}
test("fabric moves preserve populated cells and stable keys or fail closed",()=>runFabricRowMoveCases(planFabricRowMove,assert.deepEqual,assert.throws));
