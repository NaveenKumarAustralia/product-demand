-- New, isolated tables for the rebuilt Vision Board page. Same shape as the
-- original VisionBoard / VisionBoardItem so existing data can be copied in.

CREATE TABLE "VisionBoardV2" (
  "id" SERIAL NOT NULL,
  "name" TEXT NOT NULL,
  "sortOrder" INTEGER NOT NULL DEFAULT 0,
  "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  "updatedAt" TIMESTAMP(3) NOT NULL,
  CONSTRAINT "VisionBoardV2_pkey" PRIMARY KEY ("id")
);

CREATE TABLE "VisionBoardV2Item" (
  "id" SERIAL NOT NULL,
  "boardId" INTEGER NOT NULL,
  "name" TEXT NOT NULL DEFAULT '',
  "sortOrder" INTEGER NOT NULL DEFAULT 0,
  "images" JSONB NOT NULL DEFAULT '[]',
  "thumbnail" TEXT,
  "fields" JSONB NOT NULL DEFAULT '[]',
  "notes" TEXT,
  "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  "updatedAt" TIMESTAMP(3) NOT NULL,
  CONSTRAINT "VisionBoardV2Item_pkey" PRIMARY KEY ("id")
);

CREATE INDEX "VisionBoardV2Item_boardId_sortOrder_idx" ON "VisionBoardV2Item"("boardId", "sortOrder");

ALTER TABLE "VisionBoardV2Item"
  ADD CONSTRAINT "VisionBoardV2Item_boardId_fkey"
  FOREIGN KEY ("boardId") REFERENCES "VisionBoardV2"("id")
  ON DELETE CASCADE ON UPDATE CASCADE;
