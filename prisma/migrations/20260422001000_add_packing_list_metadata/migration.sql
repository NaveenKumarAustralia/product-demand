ALTER TABLE "PackingList"
ADD COLUMN "invoiceNumber" TEXT,
ADD COLUMN "expectedLeaveFactoryDate" TIMESTAMP(3);

UPDATE "PackingList"
SET "expectedLeaveFactoryDate" = "shipmentDate"
WHERE "expectedLeaveFactoryDate" IS NULL
  AND "shipmentDate" IS NOT NULL;

UPDATE "PackingList"
SET "status" = 'still_packing'
WHERE "status" IN ('draft', 'sent_to_supplier', 'confirmed', 'checked', 'completed');

UPDATE "PackingList"
SET "status" = 'on_the_way'
WHERE "status" = 'in_shipment';
