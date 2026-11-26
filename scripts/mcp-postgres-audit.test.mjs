// Real PostgreSQL integration check. Deliberately excluded from test:mcp.
// Run only with the disposable CI PostgreSQL service; never use a live database.
import assert from "node:assert/strict";
import test from "node:test";
import { PrismaClient } from "@prisma/client";
import { commitFabricSnapshot } from "../app/mcp/fabric-snapshot-store.js";
import { commitNewFabricStaffMessage, previewNewFabricStaffMessage, FABRIC_MESSAGE_SHEETS_KEY } from "../app/mcp/fabric-staff-message-store.js";

const localUrl = "postgresql://mcp_test:mcp_test@127.0.0.1:5432/mcp_audit_test?schema=public";
if (process.env.CI !== "true" || process.env.MCP_DISPOSABLE_DB !== "yes" || process.env.DATABASE_URL !== localUrl) {
  throw new Error("Refusing integration test outside the disposable CI database");
}
const db = new PrismaClient({ datasourceUrl: localUrl });
const key = "mcp-test-concurrency";
const staffKey = "supplier-portal-users-v1";
const expected = { entityKey: "test-fabric", sheetId: "test-sheet", name: "Synthetic fabric", fabricType: "Cotton", note: "" };
const recipients = [{ id: "test-staff", name: "TestStaff" }];
const staff = recipients.map(r => ({ ...r, active: true }));
const sheets = [{ gid: "test-sheet", name: "Synthetic", headers: ["Name", "Fabric Type", "Notes", "__rowKey"], rows: [["Synthetic fabric", "Cotton", "", "test-fabric", "900"]] }];
const normalize = value => Array.isArray(value) ? value : [];
const tags = (body, users) => users.filter(u => body.includes("@" + u.name));
const read = key => db.portalSetting.findUniqueOrThrow({ where: { key }, select: { value: true, updatedAt: true } });

test("real Postgres: compare-and-swap, read-only preview, atomic rollback and exact notifications", async () => {
  try {
    // Fail closed on reuse; no destructive cleanup or reset.
    assert.equal(await db.portalSetting.count(), 0, "expected a fresh disposable database");
    assert.equal(await db.portalMessage.count(), 0, "expected no existing conversations");
    await db.portalSetting.create({ data: { key, value: ["original"] } });
    const before = await read(key);
    const competing = await Promise.allSettled([
      commitFabricSnapshot(db.portalSetting, key, before, ["winner-a"]),
      commitFabricSnapshot(db.portalSetting, key, before, ["winner-b"])
    ]);
    assert.equal(competing.filter(r => r.status === "fulfilled").length, 1);
    assert.equal(competing.filter(r => r.status === "rejected").length, 1);
    assert.match(competing.find(r => r.status === "rejected").reason.message, /changed while saving/);
    assert.ok(["winner-a", "winner-b"].includes((await read(key)).value[0]));

    await db.portalSetting.createMany({ data: [{ key: FABRIC_MESSAGE_SHEETS_KEY, value: sheets }, { key: staffKey, value: staff }] });
    const snapshot = await read(FABRIC_MESSAGE_SHEETS_KEY);
    const args = [expected, "Test ETA?", recipients, normalize, tags];
    const preview = await previewNewFabricStaffMessage(db, ...args);
    assert.deepEqual(await read(FABRIC_MESSAGE_SHEETS_KEY), snapshot);
    assert.equal(await db.portalMessage.count(), 0);

    // Execute actual SQL writes inside Prisma's actual transaction, then
    // deliberately fail the notification step to verify rollback.
    const failingDb = { $transaction: (fn, opts) => db.$transaction(tx => fn({
      portalSetting: tx.portalSetting,
      portalMessage: {
        count: args => tx.portalMessage.count(args),
        createMany: async args => { await tx.portalMessage.createMany(args); throw new Error("injected notification failure"); }
      }
    }), opts) };
    await assert.rejects(commitNewFabricStaffMessage(failingDb, ...args), /injected notification failure/);
    assert.deepEqual(await read(FABRIC_MESSAGE_SHEETS_KEY), snapshot);
    assert.equal(await db.portalMessage.count(), 0);

    const committed = await commitNewFabricStaffMessage(db, ...args);
    assert.equal(committed.body, preview.body);
    const wanted = structuredClone(sheets);
    wanted[0].rows[0][2] = committed.body;
    assert.deepEqual((await read(FABRIC_MESSAGE_SHEETS_KEY)).value, wanted);
    const messages = await db.portalMessage.findMany();
    assert.equal(messages.length, 1);
    assert.equal(messages[0].userId, recipients[0].id);
    assert.equal(messages[0].entityKey, expected.entityKey);
    assert.equal(messages[0].body, committed.body);
    await assert.rejects(commitNewFabricStaffMessage(db, ...args), /changed since preview|Existing fabric/);
    assert.equal(await db.portalMessage.count(), 1);
  } finally {
    await db.$disconnect();
  }
});
