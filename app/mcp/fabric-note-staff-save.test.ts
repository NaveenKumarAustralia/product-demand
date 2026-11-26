import assert from "node:assert/strict";
import test from "node:test";
import { commitFabricNoteSnapshot } from "./fabric-note-staff-save.js";
const expected={value:[],updatedAt:new Date(0)};
const next=[{gid:"a"}];
const note={rowKey:"key",text:"@Girish ETA?",fromName:"Staff",productTitle:"Olive"};
const users=[{id:"g",name:"Girish"}];
const normalize=(value)=>value??[];
const tagged=(body,list)=>body.includes("@Girish")?list:[];
function fixture({conflict=false,failCreate=false,shortCreate=false,author="Staff"}={}){
  let state={value:expected.value,messages:[{id:1,body:"old"}]};
  const calls=[];
  const db={$transaction:async (work,opts)=>{
    assert.equal(opts.isolationLevel,"Serializable");
    const draft=JSON.parse(JSON.stringify(state));
    const result=await work({
      portalSetting:{
        updateMany:async args=>{calls.push(["save",args]);if(conflict)return {count:0};draft.value=args.data.value;return {count:1};},
        findUnique:async()=>({value:users})
      },
      portalMessage:{
        findFirst:async()=>({fromName:author}),
        deleteMany:async args=>{calls.push(["delete",args]);draft.messages=[];return {count:1};},
        createMany:async args=>{calls.push(["create",args]);if(failCreate)throw Error("notification failure");draft.messages=args.data;return {count:shortCreate?0:args.data.length};}
      }
    });
    state=draft;return result;
  }};
  return {db,calls,state:()=>state};
}
const save=(f,n=note)=>commitFabricNoteSnapshot(f.db,"fabric",expected,next,n,normalize,tagged);
test("staff note and exact notification committed together",async()=>{const f=fixture();await save(f);assert.deepEqual(f.state().value,next);assert.equal(f.state().messages[0].userId,"g");assert.equal(f.state().messages[0].entityKey,"key");assert.deepEqual(f.calls[1][1].where,{entityType:"fabric_row",orderId:0,entityKey:"key",field:"notes",parentMessageId:null,readAt:null});});
test("snapshot conflict never touches notifications",async()=>{const f=fixture({conflict:true});await assert.rejects(save(f),/changed while saving/);assert.equal(f.calls.length,1);assert.equal(f.state().messages[0].body,"old");});
test("notification failure rolls back note and old messages",async()=>{const f=fixture({failCreate:true});await assert.rejects(save(f),/notification failure/);assert.deepEqual(f.state().value,[]);assert.equal(f.state().messages[0].body,"old");});
test("incomplete notification insertion rolls back",async()=>{const f=fixture({shortCreate:true});await assert.rejects(save(f),/Incomplete/);assert.deepEqual(f.state().value,[]);assert.equal(f.state().messages[0].body,"old");});
test("explicit thread deletion stays inside transaction",async()=>{const f=fixture();await save(f,{...note,text:"",purgeThread:true,requireAuthor:true});assert.deepEqual(f.state().messages,[]);assert.deepEqual(f.calls[1][1].where,{entityType:"fabric_row",orderId:0,entityKey:"key",field:"notes"});});
test("author changed since initial UI check",async()=>{const f=fixture({author:"ChatGPT"});await assert.rejects(save(f,{...note,requireAuthor:true}),/not_author/);assert.equal(f.calls.length,0);});
test("purge cannot be requested for a non-empty note",async()=>{const f=fixture();await assert.rejects(save(f,{...note,purgeThread:true,requireAuthor:true}),/Invalid/);assert.equal(f.calls.length,0);});
test("no mentions clears only unread top-level notifications",async()=>{const f=fixture();await save(f,{...note,text:"No tags"});assert.equal(f.calls.length,2);assert.equal(f.calls[1][1].where.parentMessageId,null);assert.equal(f.calls[1][1].where.readAt,null);});
