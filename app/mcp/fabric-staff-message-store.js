import { prepareFabricNoteAppend } from "./fabric-note-append.js";
import { commitFabricSnapshot } from "./fabric-snapshot-store.js";

export const FABRIC_MESSAGE_SHEETS_KEY = "production-portal-fabric-manual-sheets-v1";
const USERS_KEY = "supplier-portal-users-v1";
const SENDER = "mcp-chatgpt";

/**
 * Prepare a NEW fabric note conversation only. Existing notes/threads must
 * use a reply, not a replacement of the UI's single top-level note.
 * @param {any} snapshot
 * @param {{entityKey:string,sheetId:string,name:string,fabricType?:string,note:string}} expected
 * @param {string} message
 * @param {{id:string,name:string}[]} recipients
 * @param {any[]} users
 * @param {(body:string, users:any[])=>any[]} taggedUsers
 */
export function planNewFabricStaffMessage(snapshot, expected, message, recipients, users, taggedUsers) {
  if (expected?.note !== "") throw new Error("Existing fabric note: reply in its conversation instead");
  if (typeof message !== "string" || !message.trim() || message.length > 1500) throw new Error("Invalid message");
  if (!Array.isArray(recipients) || recipients.length < 1 || recipients.length > 20 ||
      recipients.some(r => !r || typeof r.id !== "string" || !r.id.trim() || r.id === SENDER || typeof r.name !== "string" || !r.name.trim()) ||
      new Set(recipients.map(r => r.id)).size !== recipients.length) throw new Error("Invalid or duplicate staff recipients");
  if (!Array.isArray(users)) throw new Error("Staff list unavailable");
  for (const r of recipients) {
    const matches = users.filter(u => u.id === r.id);
    if (matches.length !== 1 || matches[0].active === false || matches[0].name !== r.name) throw new Error("Staff changed since preview");
  }
  const body = recipients.map(r => "@" + r.name.replace(/\s+/g, "")).join(" ") + " " + message.trim() + " — sent by ChatGPT for Koku";
  const actual = new Set(taggedUsers(body, users).map(u => u.id));
  if (actual.size !== recipients.length || recipients.some(r => !actual.has(r.id))) throw new Error("Message mentions do not match approved staff");
  const plan = prepareFabricNoteAppend(snapshot, expected, body);
  const anchor = { entityType: "fabric_row", orderId: 0, entityKey: expected.entityKey, field: "notes" };
  return { plan, body, anchor, notifications: recipients.map(r => ({
    ...anchor, userId: r.id, userName: r.name, fromName: "ChatGPT", productTitle: expected.name, body
  })) };
}

/**
 * Atomically save the new note and exact bell notifications. No existing
 * messages are deleted. A stale fabric snapshot or failed notification
 * rolls back the entire transaction. Never call this before approval.
 * @param {any} db
 * @param {{entityKey:string,sheetId:string,name:string,fabricType?:string,note:string}} expected
 * @param {string} message
 * @param {{id:string,name:string}[]} recipients
 * @param {(value:unknown)=>any[]} normalizeUsers
 * @param {(body:string, users:any[])=>any[]} taggedUsers
 */
export async function commitNewFabricStaffMessage(db, expected, message, recipients, normalizeUsers, taggedUsers) {
  return db.$transaction(async tx => {
    const snapshot = await tx.portalSetting.findUnique({where:{key:FABRIC_MESSAGE_SHEETS_KEY},select:{value:true,updatedAt:true}});
    if (!snapshot) throw new Error("Saved fabric sheets unavailable");
    const staff = await tx.portalSetting.findUnique({where:{key:USERS_KEY},select:{value:true}});
    const plan = planNewFabricStaffMessage(snapshot.value, expected, message, recipients, normalizeUsers(staff?.value), taggedUsers);
    if (await tx.portalMessage.count({where:plan.anchor})) throw new Error("Existing fabric conversation: preview a reply instead");
    // Preserve raw JSON, including undeclared columns and all sheet metadata.
    const next = snapshot.value.map(sheet => {
      if (sheet.gid.trim() !== plan.plan.sheetId) return sheet;
      const rows = sheet.rows.slice();
      const row = rows[plan.plan.rowIndex].slice();
      row[plan.plan.noteColumn] = plan.plan.after;
      rows[plan.plan.rowIndex] = row;
      return {...sheet, rows};
    });
    await commitFabricSnapshot(tx.portalSetting, FABRIC_MESSAGE_SHEETS_KEY, snapshot, next);
    const created = await tx.portalMessage.createMany({data:plan.notifications});
    if (created.count !== plan.notifications.length) throw new Error("Incomplete staff notifications");
    return {entity_key:expected.entityKey, sheet_id:expected.sheetId, body:plan.body, notified:recipients.map(r=>({id:r.id,name:r.name}))};
  }, {isolationLevel:"Serializable"});
}
