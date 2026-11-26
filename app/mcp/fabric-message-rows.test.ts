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

const readyFixture = (): any[] => [{gid:"s",name:"Stock",headers:["Name","Fabric Type","Notes","__rowKey"],rows:[["Olive","40x40","old","a"],["Olive","60x60","","b"]]}];
const readyOptions = {search:"Olive",fabricType:"40x40"};
test("valid fabric row can prepare a message", () => assert.equal(findFabricMessageRows(readyFixture(),readyOptions).rows[0].message_ready,true));
test("missing notes column is not message-ready", () => {const f=readyFixture();f[0].headers[2]="Other";assert.equal(findFabricMessageRows(f,readyOptions).rows[0].message_blocker,"missing_notes_column");});
test("duplicate notes columns are not message-ready", () => {const f=readyFixture();f[0].headers.push("Note");assert.equal(findFabricMessageRows(f,readyOptions).rows[0].message_blocker,"ambiguous_columns");});
test("duplicate key columns block identifiers", () => {const f=readyFixture();f[0].headers.push("__rowKey");assert.equal(findFabricMessageRows(f,readyOptions).rows[0].entity_key,null);});
test("colon-delimited key is rejected", () => {const f=readyFixture();f[0].rows[0][3]="a:b";assert.equal(findFabricMessageRows(f,readyOptions).rows[0].message_blocker,"key_invalid");});
test("malformed sheet prevents message preparation", () => {const f=readyFixture();f.push(null);assert.equal(findFabricMessageRows(f,readyOptions).rows[0].message_ready,false);});
test("non-string note is rejected", () => {const f=readyFixture();f[0].rows[0][2]=3;assert.equal(findFabricMessageRows(f,readyOptions).rows[0].message_blocker,"invalid_note_value");});
test("duplicate hidden key cannot prepare message", () => {const f=readyFixture();f.push({...f[0],gid:"h",name:"Fabric on order"});assert.equal(findFabricMessageRows(f,readyOptions).rows[0].message_blocker,"key_duplicate");});
