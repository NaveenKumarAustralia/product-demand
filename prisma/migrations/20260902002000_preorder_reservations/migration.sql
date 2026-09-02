CREATE TABLE "PreorderReservation" (
    "id" SERIAL NOT NULL,
    "shop" TEXT NOT NULL,
    "shopifyOrderId" TEXT NOT NULL,
    "shopifyOrderName" TEXT,
    "shopifyLineItemId" TEXT NOT NULL,
    "supplierOrderId" INTEGER NOT NULL,
    "productId" TEXT,
    "variantId" TEXT NOT NULL,
    "variantTitle" TEXT,
    "sku" TEXT,
    "market" TEXT NOT NULL,
    "quantity" INTEGER NOT NULL,
    "status" TEXT NOT NULL DEFAULT 'reserved',
    "customerEmail" TEXT,
    "expectedShipDate" TIMESTAMP(3),
    "reservedAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "releasedAt" TIMESTAMP(3),
    "fulfilledAt" TIMESTAMP(3),
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "PreorderReservation_pkey" PRIMARY KEY ("id")
);

CREATE UNIQUE INDEX "PreorderReservation_shop_shopifyLineItemId_supplierOrderId_key"
ON "PreorderReservation"("shop", "shopifyLineItemId", "supplierOrderId");

CREATE INDEX "PreorderReservation_shop_shopifyOrderId_status_idx"
ON "PreorderReservation"("shop", "shopifyOrderId", "status");

CREATE INDEX "PreorderReservation_supplierOrderId_variantId_status_idx"
ON "PreorderReservation"("supplierOrderId", "variantId", "status");

CREATE INDEX "PreorderReservation_variantId_market_status_idx"
ON "PreorderReservation"("variantId", "market", "status");

ALTER TABLE "PreorderReservation"
ADD CONSTRAINT "PreorderReservation_quantity_positive"
CHECK ("quantity" > 0);
