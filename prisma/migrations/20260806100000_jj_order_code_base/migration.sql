-- Base SKU / barcode typed at order time for new JJ products (order-first flow).
ALTER TABLE "SupplierOrder" ADD COLUMN IF NOT EXISTS "skuBase" TEXT;
ALTER TABLE "SupplierOrder" ADD COLUMN IF NOT EXISTS "barcodeBase" TEXT;
