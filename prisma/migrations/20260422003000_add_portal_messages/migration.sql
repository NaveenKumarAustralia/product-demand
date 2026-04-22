CREATE TABLE "PortalMessage" (
  "id" SERIAL NOT NULL,
  "userId" TEXT NOT NULL,
  "userName" TEXT NOT NULL,
  "orderId" INTEGER NOT NULL,
  "field" TEXT NOT NULL,
  "fromName" TEXT,
  "productTitle" TEXT,
  "body" TEXT NOT NULL,
  "readAt" TIMESTAMP(3),
  "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

  CONSTRAINT "PortalMessage_pkey" PRIMARY KEY ("id")
);

CREATE INDEX "PortalMessage_userId_readAt_createdAt_idx" ON "PortalMessage"("userId", "readAt", "createdAt");
CREATE INDEX "PortalMessage_orderId_field_idx" ON "PortalMessage"("orderId", "field");
