-- CreateTable
CREATE TABLE "SupplierOrder" (
    "id" SERIAL NOT NULL,
    "shop" TEXT NOT NULL,
    "poNumber" TEXT,
    "supplier" TEXT NOT NULL,
    "productId" TEXT NOT NULL,
    "productTitle" TEXT NOT NULL,
    "status" TEXT NOT NULL DEFAULT 'open',
    "eta" TIMESTAMP(3),
    "notes" TEXT,
    "totalQty" INTEGER NOT NULL DEFAULT 0,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,
    CONSTRAINT "SupplierOrder_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "OrderLine" (
    "id" SERIAL NOT NULL,
    "orderId" INTEGER NOT NULL,
    "variantId" TEXT NOT NULL,
    "variantTitle" TEXT NOT NULL,
    "sku" TEXT,
    "qtyOrdered" INTEGER NOT NULL,
    "qtyReceived" INTEGER NOT NULL DEFAULT 0,
    "costPrice" DOUBLE PRECISION,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT "OrderLine_pkey" PRIMARY KEY ("id")
);

-- AddForeignKey
ALTER TABLE "OrderLine" ADD CONSTRAINT "OrderLine_orderId_fkey"
    FOREIGN KEY ("orderId") REFERENCES "SupplierOrder"("id")
    ON DELETE CASCADE ON UPDATE CASCADE;

-- CreateIndex
CREATE INDEX "SupplierOrder_shop_productId_idx" ON "SupplierOrder"("shop", "productId");
CREATE INDEX "SupplierOrder_shop_status_idx" ON "SupplierOrder"("shop", "status");
