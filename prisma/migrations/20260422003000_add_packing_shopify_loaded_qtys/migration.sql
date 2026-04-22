ALTER TABLE "PackingListLine"
ADD COLUMN "shopifyLoadedQtys" JSONB NOT NULL DEFAULT '{}';
