-- CreateTable
CREATE TABLE "SupplierAccount" (
    "id" SERIAL NOT NULL,
    "shop" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "email" TEXT NOT NULL,
    "passwordHash" TEXT NOT NULL,
    "active" BOOLEAN NOT NULL DEFAULT true,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,
    CONSTRAINT "SupplierAccount_pkey" PRIMARY KEY ("id")
);

CREATE UNIQUE INDEX "SupplierAccount_shop_email_key" ON "SupplierAccount"("shop", "email");

-- Add supplier-side status tracking to SupplierOrder
ALTER TABLE "SupplierOrder" ADD COLUMN "supplierStatus" TEXT NOT NULL DEFAULT 'pending';
ALTER TABLE "SupplierOrder" ADD COLUMN "supplierNotes" TEXT;
