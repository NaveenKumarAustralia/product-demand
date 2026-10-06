-- AlterTable: date the order first went on production (cut). Editable; null until cut.
ALTER TABLE "SupplierOrder" ADD COLUMN "productionDate" TIMESTAMP(3);

-- CreateTable: permanent cut log, survives SupplierOrder deletion (no cascade).
CREATE TABLE "ProductionCutRecord" (
    "id" SERIAL NOT NULL,
    "shop" TEXT NOT NULL,
    "supplierOrderId" INTEGER,
    "productId" TEXT,
    "productTitle" TEXT NOT NULL,
    "supplier" TEXT,
    "productType" TEXT,
    "qty" INTEGER NOT NULL DEFAULT 0,
    "productionDate" TIMESTAMP(3) NOT NULL,
    "createdByName" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "ProductionCutRecord_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE UNIQUE INDEX "ProductionCutRecord_supplierOrderId_key" ON "ProductionCutRecord"("supplierOrderId");

-- CreateIndex
CREATE INDEX "ProductionCutRecord_shop_productionDate_idx" ON "ProductionCutRecord"("shop", "productionDate");
