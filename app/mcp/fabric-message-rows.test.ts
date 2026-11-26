import test from 'node:test';
import assert from 'node:assert/strict';
import { findFabricMessageRows } from './fabric-message-rows.js';
const fixtures = ()=>[
{gid:"s1",name:"Stock",kind:"stock",headers:["Name","Fabric Type","Notes","__rowKey"],rows:[["Isha","40x40","private note","a"],["Isha","60x60","","b"],["Olive Plain","40x40","","c"],["Olive Plain","40x40","",""],["Other","40x40","","dup"]]},
{gid:"s2",name:"On Order",kind:"order",headers:["Name","Type","__rowKey"],rows:[["Olive Plain","40x40","d"],["Olive Plain","","dup"]]},
{gid:"s3",name:"Fabric on order",kind:"order",headers:["Name","Type","__rowKey"],rows:[["Olive Plain","40x40","c"]]}
];
test("exact fabric type does not mix Isha 40x40 and 60x60", () => assert.deepEqual((()=>findFabricMessageRows(fixtures(),{search:"isha",fabricType:"40x40"}).rows.map(x=>x.entity_key))(), ["a"]));
test("case-insensitive name search", () => assert.deepEqual((()=>findFabricMessageRows(fixtures(),{search:" ISHA "}).rows.length)(), 2));
test("hidden sheets do not appear", () => assert.deepEqual((()=>findFabricMessageRows(fixtures(),{search:"Olive"}).rows.map(x=>x.sheet_id))(), ["s1","s1","s2","s2"]));
test("duplicate keys including hidden sheet are blocked", () => assert.deepEqual((()=>findFabricMessageRows(fixtures(),{search:"Olive"}).rows.map(x=>[x.entity_key,x.key_status]))(), [[null,"duplicate"],[null,"missing"],["d","unique"],[null,"duplicate"]]));
test("blank type never inferred", () => assert.deepEqual((()=>findFabricMessageRows(fixtures(),{search:"Olive"}).rows[3].fabric_type)(), null));
test("raw note text is not returned", () => assert.deepEqual((()=>Object.keys(findFabricMessageRows(fixtures(),{search:"isha"}).rows[0]).includes("notes"))(), false));
test("missing data does not backfill", () => assert.deepEqual((()=>findFabricMessageRows(null,{search:"isha"}).status)(), "manual_sheets_unavailable"));
test("pagination bound", () => assert.deepEqual((()=>findFabricMessageRows(fixtures(),{search:"Olive",limit:1}).truncated)(), true));
test("does not mutate input", () => assert.deepEqual((()=>{const f=fixtures();const before=JSON.stringify(f);findFabricMessageRows(f,{search:"Olive"});return JSON.stringify(f)===before})(), true));
test("no blank search", () => assert.deepEqual((()=>{try{findFabricMessageRows(fixtures(),{search:" "});return false}catch{return true}})(), true));
