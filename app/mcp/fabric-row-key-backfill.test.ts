import { test } from "node:test";
import assert from "node:assert/strict";
import { prepareFabricRowKeyBackfill } from "./fabric-row-key-backfill.js";

const key = () => "new-key";
const sheet = { gid: "s", headers: ["Name", "Notes"], rows: [["Rose", "old"]], extra: "keep" };
test("fills missing key without mutating input or losing metadata", () => {
  const out = prepareFabricRowKeyBackfill([sheet], key);
  assert.equal(out.changed, true);
  assert.deepEqual(out.value, [{ ...sheet, headers: ["Name", "Notes", "__rowKey"], rows: [["Rose", "old", "new-key"]] }]);
  assert.deepEqual(sheet.rows, [["Rose", "old"]]);
  assert.equal(prepareFabricRowKeyBackfill(out.value, key).changed, false);
});
test("does not repurpose undeclared stock cells", () => {
  const trailing = { ...sheet, rows: [["Rose", "old", "900"]] };
  assert.equal(prepareFabricRowKeyBackfill([trailing], key).changed, false);
  const result = prepareFabricRowKeyBackfill([trailing, sheet], key);
  assert.equal(result.value[0], trailing);
  assert.equal(result.changed, true);
});
test("existing key column preserves trailing stock and current IDs", () => {
  const s = { ...sheet, headers: ["Name", "__rowKey"], rows: [["Rose", "existing", "900"], ["Leaf", "", "905"]] };
  assert.deepEqual(prepareFabricRowKeyBackfill([s], key).value[0].rows, [["Rose", "existing", "900"], ["Leaf", "new-key", "905"]]);
});
test("malformed or ambiguous snapshots are left alone", () => {
  assert.equal(prepareFabricRowKeyBackfill(null, key).changed, false);
  assert.equal(prepareFabricRowKeyBackfill([{ ...sheet, rows: [null] }], key).changed, false);
  assert.equal(prepareFabricRowKeyBackfill([{ ...sheet, headers: ["Name", "__rowKey", "__rowKey"] }], key).changed, false);
});
