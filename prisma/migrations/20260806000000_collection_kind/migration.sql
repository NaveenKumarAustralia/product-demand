-- Discriminates regular Collections from the JJ New Products sheet.
ALTER TABLE "Collection" ADD COLUMN IF NOT EXISTS "kind" TEXT NOT NULL DEFAULT 'collection';
