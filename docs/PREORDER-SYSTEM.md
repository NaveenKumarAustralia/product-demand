# Preorder System — Foundation

Status: design/foundation only. This branch must not be deployed to production until reviewed and tested.

## Core eligibility rule

A SupplierOrder is eligible to supply preorder capacity only when ALL of these are true:

1. `supplierStatus` is `On Production` (using the portal's canonical stored value).
2. `destination` is `Send to AUS` or `Send to USA` (using the portal's canonical stored values).
3. Preorder is explicitly enabled by an authorised staff member.
4. The batch has remaining safe preorder capacity for the requested variant.

Changing production status or destination must never, by itself, expose preorder inventory to customers. It only makes the batch eligible. The explicit Preorder Enabled control is the final gate.

## Market mapping

- `Send to AUS` -> AU preorder pool / Australian fulfilment location.
- `Send to USA` -> USA preorder pool / US fulfilment location.
- `Keep at factory` and any other destination -> not preorder eligible.

AU and USA inventory/capacity must never be combined.

## Ship-date rule

Do not add or depend on a new `productionStartedAt` field for the first preorder version.

The customer-facing preorder ship/dispatch date is driven by the production status workflow and preorder configuration. The existing editable `eta` remains production information and is not the sole source of truth for preorder availability.

Implementation must keep ship-date calculation isolated in a preorder service so the rule can be changed later without changing the Production Portal's core order model.

## Explicit enable control

Add a `Preorder Enabled` control per eligible production batch.

Behaviour:

- Ineligible batch: control disabled/off with a reason.
- Eligible + off: no customer preorder availability.
- Eligible + on: batch contributes safe preorder capacity.
- If status leaves On Production: stop accepting new preorder reservations from that batch.
- If destination changes to an ineligible destination: stop accepting new preorder reservations from that batch.
- If destination changes AUS <-> USA: do not silently move existing customer allocations between markets; flag/reconcile them.
- Existing customer reservations must not be automatically cancelled by these state changes.

## Safety requirements

1. Do not modify existing Production Portal behaviour unless required for a reviewed preorder hook.
2. Do not deploy this development branch to Railway production.
3. Do not modify `main` while building/testing the preorder foundation.
4. Keep preorder business logic in separate modules/routes; do not add large blocks of logic to `portal._index.tsx`.
5. Preorder capacity must be variant-level and destination/location-level.
6. Shopify physical inventory and preorder reservations are separate concepts.
7. Never reduce physical stock receipts by the number of preorder reservations.
8. Never combine inventory across Shopify locations when determining AU/USA availability.
9. Capacity updates/reservations must eventually be transactional/idempotent to prevent overselling.

## Recommended first implementation phases

### Phase 1 — isolated foundation

- Add preorder database/settings models via a new Prisma migration.
- Add pure eligibility and capacity services.
- Add read-only preorder dashboard/routes.
- Add explicit batch enable/disable API.
- No storefront selling plans.
- No Shopify inventory mutation.
- No changes to existing staff production workflows beyond an isolated preorder control.

### Phase 2 — Shopify location awareness

- Map AU and USA destinations to Shopify location IDs.
- Refactor preorder inventory reads to be location-specific.
- Do not change existing portal inventory calculations until separately reviewed.

### Phase 3 — Shopify preorder integration

- Verify/obtain Shopify purchase-option/selling-plan permissions.
- Create/sync selling plans only for explicitly enabled eligible batches.
- Add required order/cancellation/refund webhooks.
- Add reservation allocation and release logic.

### Phase 4 — customer experience

- Product-page preorder messaging.
- Expected dispatch date.
- Back-in-stock/waitlist.
- Customer-order/admin reporting and notifications.

## Initial UI

Production Portal sidebar:

- PRE-ORDERS
  - Dashboard
  - Products & Batches
  - Customer Orders
  - Back in Stock
  - Notifications
  - Reports
  - Settings

For the first safe phase, Dashboard and Products & Batches can be implemented first; remaining pages may be placeholders until their backend behaviour exists.

## Important implementation principle

`On Production` + (`Send to AUS` or `Send to USA`) means **eligible for preorder**.

It does NOT mean **preorder automatically enabled**.

The explicit `Preorder Enabled` control is the final activation gate.
