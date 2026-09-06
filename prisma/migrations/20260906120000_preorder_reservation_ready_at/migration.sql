-- Marks when a preorder reservation's stock has landed in Shopify and we've
-- released the Shopify fulfilment hold + tagged the order for Pick Pack. Null =
-- still awaiting stock. Distinct from releasedAt (a cancellation).
ALTER TABLE "PreorderReservation" ADD COLUMN "readyAt" TIMESTAMP(3);
