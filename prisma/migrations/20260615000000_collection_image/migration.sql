CREATE TABLE IF NOT EXISTS "CollectionImage" (
  "key" TEXT NOT NULL,
  "collectionId" INTEGER NOT NULL,
  "mimeType" TEXT NOT NULL,
  "bytes" BYTEA NOT NULL,
  "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT "CollectionImage_pkey" PRIMARY KEY ("key")
);
CREATE INDEX IF NOT EXISTS "CollectionImage_collectionId_idx" ON "CollectionImage"("collectionId");
