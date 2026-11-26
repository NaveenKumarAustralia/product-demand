import { commitFabricSnapshot } from "./fabric-snapshot-store.js";

/** Atomically save a staff fabric note and its notifications.
 * @param {any} db
 * @param {string} key
 * @param {{value:any,updatedAt:Date}} expected
 * @param {any[]} value
 * @param {{rowKey:string,text:string,fromName:string|null,productTitle:string|null,purgeThread?:boolean,requireAuthor?:boolean}} note
 * @param {(value:unknown)=>any[]} normalizeUsers
 * @param {(body:string,users:any[])=>any[]} taggedUsers
 */
export async function commitFabricNoteSnapshot(db,key,expected,value,note,normalizeUsers,taggedUsers) {
  const anchor={entityType:"fabric_row",orderId:0,entityKey:note.rowKey,field:"notes"};
  if(!note.rowKey || note.rowKey.includes(":"))throw new Error("Invalid fabric note key");
  if(note.purgeThread && (!note.requireAuthor || note.text.trim()))throw new Error("Invalid fabric thread purge");
  return db.$transaction(async tx=>{
    if(note.requireAuthor){
      const latest=await tx.portalMessage.findFirst({where:{...anchor,parentMessageId:null},select:{fromName:true},orderBy:{createdAt:"desc"}});
      if(latest && (latest.fromName??"").trim().toLowerCase()!==(note.fromName??"").trim().toLowerCase())throw new Error("not_author");
    }
    await commitFabricSnapshot(tx.portalSetting,key,expected,value);
    if(note.purgeThread){
      await tx.portalMessage.deleteMany({where:anchor});
      return;
    }
    const setting=await tx.portalSetting.findUnique({where:{key:"supplier-portal-users-v1"},select:{value:true}});
    const recipients=taggedUsers(note.text,normalizeUsers(setting?.value));
    // Keep existing staff edit semantics but never separate the note save
    // from mention replacement; preserve replies and already-read mentions.
    await tx.portalMessage.deleteMany({where:{...anchor,parentMessageId:null,readAt:null}});
    if(recipients.length){
      const created=await tx.portalMessage.createMany({data:recipients.map(user=>({...anchor,userId:user.id,userName:user.name,fromName:note.fromName,productTitle:note.productTitle,body:note.text.trim()}))});
      if(created.count!==recipients.length)throw new Error("Incomplete fabric note notifications");
    }
  },{isolationLevel:"Serializable"});
}
