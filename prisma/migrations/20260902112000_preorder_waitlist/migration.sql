CREATE TABLE "PreorderWaitlist" (
  "id" SERIAL PRIMARY KEY,
  "shop" TEXT NOT NULL,
  "email" TEXT NOT NULL,
  "productId" TEXT,
  "productTitle" TEXT,
  "variantId" TEXT NOT NULL,
  "variantTitle" TEXT,
  "sku" TEXT,
  "market" TEXT NOT NULL,
  "status" TEXT NOT NULL DEFAULT 'waiting',
  "source" TEXT NOT NULL DEFAULT 'storefront',
  "notifiedAt" TIMESTAMP(3),
  "convertedOrderId" TEXT,
  "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  "updatedAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE UNIQUE INDEX "PreorderWaitlist_shop_email_variant_market_key"
  ON "PreorderWaitlist"("shop", "email", "variantId", "market");
CREATE INDEX "PreorderWaitlist_shop_status_createdAt_idx"
  ON "PreorderWaitlist"("shop", "status", "createdAt");
CREATE INDEX "PreorderWaitlist_variant_market_status_idx"
  ON "PreorderWaitlist"("variantId", "market", "status");
