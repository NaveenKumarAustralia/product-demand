-- Who pressed "Load inventory" for a packing list (attribution).
ALTER TABLE "PackingList" ADD COLUMN IF NOT EXISTS "loadedBy" TEXT;
