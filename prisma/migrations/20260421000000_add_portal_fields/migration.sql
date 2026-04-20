ALTER TABLE "SupplierOrder" ADD COLUMN "factoryNotes" TEXT;
ALTER TABLE "SupplierOrder" ADD COLUMN "priority" TEXT;
ALTER TABLE "SupplierOrder" ADD COLUMN "productImageUrl" TEXT;

-- Migrate old supplierStatus values to new ones
UPDATE "SupplierOrder" SET "supplierStatus" = 'on_order' WHERE "supplierStatus" IN ('pending', 'confirmed');
UPDATE "SupplierOrder" SET "supplierStatus" = 'in_shipment' WHERE "supplierStatus" = 'shipped';
