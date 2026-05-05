CREATE TABLE "Collection" (
  "id" SERIAL NOT NULL,
  "name" TEXT NOT NULL DEFAULT 'Untitled collection',
  "sortOrder" INTEGER NOT NULL DEFAULT 0,
  "thumbnail" TEXT,
  "columns" JSONB NOT NULL DEFAULT '[]',
  "rows" JSONB NOT NULL DEFAULT '[]',
  "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  "updatedAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT "Collection_pkey" PRIMARY KEY ("id")
);
