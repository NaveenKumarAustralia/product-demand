ALTER TABLE "SampleIteration" ADD COLUMN "name" TEXT;
ALTER TABLE "SampleIteration" ADD COLUMN "taggedUsers" JSONB NOT NULL DEFAULT '[]';
