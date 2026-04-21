CREATE TABLE "PackingList" (
  "id" SERIAL NOT NULL,
  "title" TEXT NOT NULL,
  "shipmentDate" TIMESTAMP(3),
  "status" TEXT NOT NULL DEFAULT 'draft',
  "notes" TEXT,
  "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  "updatedAt" TIMESTAMP(3) NOT NULL,

  CONSTRAINT "PackingList_pkey" PRIMARY KEY ("id")
);

CREATE TABLE "PackingListLine" (
  "id" SERIAL NOT NULL,
  "packingListId" INTEGER NOT NULL,
  "boxNumber" TEXT,
  "productId" TEXT,
  "productTitle" TEXT NOT NULL,
  "productImageUrl" TEXT,
  "fabricImageData" TEXT,
  "sku" TEXT,
  "isCustom" BOOLEAN NOT NULL DEFAULT false,
  "qtys" JSONB NOT NULL DEFAULT '{}',
  "priceRupees" DOUBLE PRECISION,
  "weight" DOUBLE PRECISION,
  "notes" TEXT,
  "sortOrder" INTEGER NOT NULL DEFAULT 0,
  "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  "updatedAt" TIMESTAMP(3) NOT NULL,

  CONSTRAINT "PackingListLine_pkey" PRIMARY KEY ("id")
);

ALTER TABLE "PackingListLine"
ADD CONSTRAINT "PackingListLine_packingListId_fkey"
FOREIGN KEY ("packingListId") REFERENCES "PackingList"("id")
ON DELETE CASCADE ON UPDATE CASCADE;
