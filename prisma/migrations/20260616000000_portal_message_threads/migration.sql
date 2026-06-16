-- Generalise PortalMessage so any "notes" cell across the portal can
-- create a thread, not just restock SupplierOrder.notes/factory_notes.
ALTER TABLE "PortalMessage"
  ADD COLUMN IF NOT EXISTS "entityType" TEXT NOT NULL DEFAULT 'supplier_order',
  ADD COLUMN IF NOT EXISTS "entityKey" TEXT,
  ADD COLUMN IF NOT EXISTS "parentMessageId" INTEGER,
  ADD COLUMN IF NOT EXISTS "editedAt" TIMESTAMP(3);

-- Backfill existing rows: every existing message was a restock order
-- mention, so the default already covers them. No-op left as a marker.
UPDATE "PortalMessage" SET "entityType" = 'supplier_order' WHERE "entityType" IS NULL;

-- Index thread fan-outs ("give me every message on this cell") and
-- reply walks. The existing [orderId, field] index already covers
-- supplier_order lookups; this one extends to the other entity types.
CREATE INDEX IF NOT EXISTS "PortalMessage_entityType_orderId_entityKey_field_idx"
  ON "PortalMessage" ("entityType", "orderId", "entityKey", "field");
CREATE INDEX IF NOT EXISTS "PortalMessage_parentMessageId_idx"
  ON "PortalMessage" ("parentMessageId");
