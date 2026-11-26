// @ts-nocheck
import {test} from "node:test";
import assert from "node:assert/strict";
import {planNewFabricStaffMessage,commitNewFabricStaffMessage,FABRIC_MESSAGE_SHEETS_KEY} from "./fabric-staff-message-store.js";

const expected={entityKey:"stable",sheetId:"sheet",name:"Olive",fabricType:"Cotton",note:""};
const recipients=[{id:"g",name:"Girish"},{id:"m",name:"Manoj"}];
const users=[...recipients.map(r=>({...r,active:true})),{id:"x",name:"Other",active:true}];
const sheets=[{gid:"sheet",name:"Stock",headers:["Name","Fabric Type","Notes","__rowKey"],rows:[["Olive","Cotton","","stable","900"],["Rose","Cotton","old","rose","100"]],custom:{keep:true}}];
const copy=x=>JSON.parse(JSON.stringify(x));
const normalize=v=>Array.isArray(v)?v:[];
const tags=(text,staff)=>staff.filter(u=>u.active!==false && [...text.matchAll(/@([a-z0-9._-]+)/gi)].some(m=>m[1].toLowerCase()===u.name.replace(/\s+/g,"").toLowerCase()));
function fake({conflict=false,notifyFail=false,partial=false,existing=false,staff=users,value=sheets}={}){
  let state={value:copy(value),updatedAt:"v1",messages:existing?[{id:1,body:"old"}]:[]};
  let optionsSeen;
  return {
    get state(){return copy(state)}, get options(){return optionsSeen},
    async $transaction(fn,options){
      optionsSeen=options;
      const draft=copy(state);
      const result=await fn({
        portalSetting:{
          findUnique:async({where})=>where.key===FABRIC_MESSAGE_SHEETS_KEY?{value:copy(draft.value),updatedAt:draft.updatedAt}:{value:copy(staff)},
          updateMany:async({where,data})=>{
            if(conflict||where.updatedAt!==draft.updatedAt||JSON.stringify(where.value.equals)!==JSON.stringify(draft.value))return {count:0};
            draft.value=copy(data.value);draft.updatedAt="v2";return {count:1};
          }
        },
        portalMessage:{
          count:async()=>draft.messages.length,
          createMany:async({data})=>{
            draft.messages.push(...copy(data));
            if(notifyFail)throw Error("notification failure");
            return {count:partial?data.length-1:data.length};
          }
        }
      });
      state=draft;return result;
    }
  };
}
const run=db=>commitNewFabricStaffMessage(db,expected,"When will it arrive?",recipients,normalize,tags);
test("plan matches exact intended staff and does not mutate sheets",()=>{
  const value=copy(sheets);
  const p=planNewFabricStaffMessage(value,expected,"When?",recipients,users,tags);
  assert.deepEqual(value,sheets);
  assert.equal(p.anchor.entityKey,"stable");
  assert.deepEqual(p.notifications.map(n=>n.userId),["g","m"]);
});
test("existing notes cannot be replaced",()=>{
  assert.throws(()=>planNewFabricStaffMessage(sheets,{...expected,note:"old"},"When?",recipients,users,tags),/Existing fabric note/);
});
test("unapproved mentions are rejected",()=>{
  assert.throws(()=>planNewFabricStaffMessage(sheets,expected,"@Other When?",recipients,users,tags),/approved staff/);
});
test("inactive, renamed or duplicate staff cannot receive messages",()=>{
  for(const staff of [[{...users[0],active:false},...users.slice(1)],[{...users[0],name:"Renamed"},...users.slice(1)],[...users,users[0]]]){
    assert.throws(()=>planNewFabricStaffMessage(sheets,expected,"When?",recipients,staff,tags),/Staff changed/);
  }
});
test("duplicate requested recipient rejected",()=>{
  assert.throws(()=>planNewFabricStaffMessage(sheets,expected,"When?",[recipients[0],recipients[0]],users,tags),/duplicate staff/);
});
test("successful commit saves only target note and exact notifications",async()=>{
  const db=fake();const result=await run(db);
  const want=copy(sheets);want[0].rows[0][2]=result.body;
  assert.deepEqual(db.state.value,want);
  assert.deepEqual(db.state.messages.map(m=>m.userId),["g","m"]);
  assert.equal(db.state.messages[0].orderId,0);
  assert.equal(db.options.isolationLevel,"Serializable");
});
for(const [label,options,error] of [
  ["stale snapshot",{conflict:true},/changed while saving/],
  ["notification failure",{notifyFail:true},/notification failure/],
  ["incomplete notifications",{partial:true},/Incomplete staff notifications/],
  ["existing conversation",{existing:true},/Existing fabric conversation/],
  ["staff changed",{staff:[{...users[0],active:false},...users.slice(1)]},/Staff changed/],
  ["note changed",{value:[{...sheets[0],rows:[["Olive","Cotton","new note","stable","900"]]}]},/changed since preview/],
]){
  test(label+" leaves all state unchanged",async()=>{
    const db=fake(options);const before=db.state;
    await assert.rejects(()=>run(db),error);
    assert.deepEqual(db.state,before);
  });
}
