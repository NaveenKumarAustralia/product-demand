CREATE TABLE IF NOT EXISTS "WarehouseDailyMetric" (
  "date" TIMESTAMP(3) NOT NULL,
  "staffCount" INTEGER NOT NULL DEFAULT 0,
  "staffHours" DOUBLE PRECISION NOT NULL DEFAULT 0,
  "ordersFulfilled" INTEGER NOT NULL DEFAULT 0,
  "unitsFulfilled" INTEGER NOT NULL DEFAULT 0,
  "deputyFetchedAt" TIMESTAMP(3),
  "shopifyFetchedAt" TIMESTAMP(3),
  "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  "updatedAt" TIMESTAMP(3) NOT NULL,
  CONSTRAINT "WarehouseDailyMetric_pkey" PRIMARY KEY ("date")
);
