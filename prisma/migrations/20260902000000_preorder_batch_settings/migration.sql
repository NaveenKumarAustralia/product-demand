CREATE TABLE "PreorderBatchSetting" (
    "id" SERIAL NOT NULL,
    "supplierOrderId" INTEGER NOT NULL,
    "shop" TEXT NOT NULL,
    "enabled" BOOLEAN NOT NULL DEFAULT false,
    "safetyBufferPercent" DOUBLE PRECISION NOT NULL DEFAULT 5,
    "safetyBufferQty" INTEGER,
    "shipDate" TIMESTAMP(3),
    "pausedReason" TEXT,
    "enabledByUserId" TEXT,
    "enabledByUserName" TEXT,
    "enabledAt" TIMESTAMP(3),
    "updatedByUserId" TEXT,
    "updatedByUserName" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "PreorderBatchSetting_pkey" PRIMARY KEY ("id")
);

CREATE UNIQUE INDEX "PreorderBatchSetting_supplierOrderId_key"
ON "PreorderBatchSetting"("supplierOrderId");

CREATE INDEX "PreorderBatchSetting_shop_enabled_idx"
ON "PreorderBatchSetting"("shop", "enabled");

CREATE INDEX "PreorderBatchSetting_supplierOrderId_enabled_idx"
ON "PreorderBatchSetting"("supplierOrderId", "enabled");
