-- Fabric allocation on supplier orders: which fabric row an order is assigned
-- to (the pick that sets its price), meters per piece, and whether its fabric
-- has been deducted from physical stock (on production, reversible).
ALTER TABLE "SupplierOrder" ADD COLUMN IF NOT EXISTS "fabricKey" TEXT;
ALTER TABLE "SupplierOrder" ADD COLUMN IF NOT EXISTS "fabricName" TEXT;
ALTER TABLE "SupplierOrder" ADD COLUMN IF NOT EXISTS "metersPerPiece" DOUBLE PRECISION;
ALTER TABLE "SupplierOrder" ADD COLUMN IF NOT EXISTS "fabricConsumed" BOOLEAN NOT NULL DEFAULT false;
ALTER TABLE "SupplierOrder" ADD COLUMN IF NOT EXISTS "fabricConsumedMeters" DOUBLE PRECISION;
