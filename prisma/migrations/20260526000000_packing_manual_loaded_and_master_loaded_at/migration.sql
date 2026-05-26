ALTER TABLE "PackingListLine" ADD COLUMN "manuallyLoadedQtys" JSONB NOT NULL DEFAULT '{}'::jsonb;
ALTER TABLE "PackingList" ADD COLUMN "masterInventoryLoadedAt" TIMESTAMP(3);
