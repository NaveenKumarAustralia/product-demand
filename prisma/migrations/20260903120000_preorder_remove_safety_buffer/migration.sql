-- Karma East does not use a preorder safety buffer (stock does not go missing),
-- so all incoming units are sellable as preorder. Drop the 5% default and clear
-- any buffer that was stored on existing batches.
ALTER TABLE "PreorderBatchSetting" ALTER COLUMN "safetyBufferPercent" SET DEFAULT 0;
UPDATE "PreorderBatchSetting" SET "safetyBufferPercent" = 0, "safetyBufferQty" = NULL;
