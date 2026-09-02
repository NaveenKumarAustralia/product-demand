CREATE TABLE "PreorderNotification" (
  "id" SERIAL PRIMARY KEY,
  "shop" TEXT NOT NULL,
  "type" TEXT NOT NULL,
  "status" TEXT NOT NULL DEFAULT 'draft',
  "shopifyOrderId" TEXT,
  "waitlistId" INTEGER,
  "supplierOrderId" INTEGER,
  "customerEmail" TEXT,
  "subject" TEXT NOT NULL,
  "body" TEXT NOT NULL,
  "createdByUserId" TEXT,
  "createdByUserName" TEXT,
  "approvedByUserId" TEXT,
  "approvedByUserName" TEXT,
  "approvedAt" TIMESTAMP(3),
  "sentAt" TIMESTAMP(3),
  "cancelledAt" TIMESTAMP(3),
  "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  "updatedAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX "PreorderNotification_shop_status_createdAt_idx"
  ON "PreorderNotification"("shop", "status", "createdAt");
CREATE INDEX "PreorderNotification_supplierOrderId_status_idx"
  ON "PreorderNotification"("supplierOrderId", "status");
CREATE INDEX "PreorderNotification_shopifyOrderId_idx"
  ON "PreorderNotification"("shopifyOrderId");
