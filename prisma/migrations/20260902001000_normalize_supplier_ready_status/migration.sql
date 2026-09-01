-- One-time cleanup previously executed on every Production Portal order-backed page load.
-- Keep the same behaviour while removing write work from the request path.
UPDATE "SupplierOrder"
SET "supplierStatus" = 'ready'
WHERE "supplierStatus" IN ('packed', 'ready_to_send');
