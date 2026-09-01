# Preorder System — Foundation

Status: development foundation. Do not enable customer-facing preorder selling until the Shopify selling-plan and reservation phases are complete and verified.

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

## Staff permissions

Pre-orders are a restricted area of the Production Portal and reuse the portal's existing user and page-permission system.

### Menu/page access

- **Pre-orders** is an ordinary permission-controlled portal page.
- Admins manage access from **Production Portal -> Settings -> Users -> Page access -> Pre-orders**.
- If `Pre-orders` is not enabled for a user, that user does not see the Pre-orders menu item.
- A logged-in user without that permission who manually enters `?page=preorders` is redirected to a page they are allowed to access.
- Superadmins retain access for administration/recovery.
- There is no second staff directory and no separate duplicate "Pre-orders access" list inside the preorder module.

### Sensitive actions inside Pre-orders

Page access and action access are separate concepts. A staff member can be allowed to see the Pre-orders area without necessarily being allowed to change selling behaviour.

Additional action permissions can include:

- Manage preorder availability / enable / pause
- Change preorder ETA/lead days
- Change safety buffer
- Send customer delay notifications
- Manage back-in-stock notifications
- View sensitive preorder reports

These action permissions reuse existing portal staff identities, must be checked server-side, and should be editable only by an admin/owner permission group.

All enable, pause, destination-sensitive changes, and permission changes should be written to the existing ActivityLog/audit mechanism where possible.

## Safety requirements

1. Keep preorder business logic in separate modules/routes; do not add large blocks of business logic to `portal._index.tsx`.
2. Preorder capacity must be variant-level and destination/location-level.
3. Shopify physical inventory and preorder reservations are separate concepts.
4. Never reduce physical stock receipts by the number of preorder reservations.
5. Never combine inventory across Shopify locations when determining AU/USA availability.
6. Capacity updates/reservations must eventually be transactional/idempotent to prevent overselling.
7. All preorder menu/action/API permission checks must be enforced on the server.
8. Customer-facing preorder selling must remain off until selling-plan access, reservation allocation, cancellation/release logic and storefront behaviour are implemented and tested.

## Implementation phases

### Phase 1 — isolated foundation

Implemented/under validation:

- Preorder batch settings attached to existing SupplierOrder production batches.
- Pure eligibility and capacity services.
- Permission-controlled preorder dashboard integrated into the Production Portal.
- Existing portal Page access permissions control Pre-orders menu visibility.
- Server-side direct-page access guard.
- Server service for explicit batch enable/disable with action permission checks.
- AU/USA Shopify location mapping foundation.
- No storefront selling plans yet.
- No Shopify inventory mutation by the preorder module.
- Reservation totals remain zero until the real customer allocation ledger is implemented; the system must never invent reservation numbers.

### Phase 2 — Shopify location awareness

- Map AU and USA destinations to actual Shopify location IDs.
- Use location-specific inventory reads for preorder selling state.
- Existing product inventory resource now supports an optional Shopify `locationId` while retaining backwards-compatible all-location behaviour for existing portal callers.

### Phase 3 — Shopify preorder integration

- Verify/obtain Shopify purchase-option/selling-plan permissions.
- Create/sync selling plans only for explicitly enabled eligible batches.
- Add required order/cancellation/refund webhooks.
- Add transactional reservation allocation and release logic.
- Allocate customers FIFO to the earliest eligible region-matching production batch with safe capacity.

### Phase 4 — customer experience

- Product-page preorder messaging.
- Expected dispatch date/window.
- Mixed-cart communication.
- Back-in-stock/waitlist.
- Customer preorder account section.
- Customer-order/admin reporting and notifications.

## Initial portal UI

For authorised users, Production Portal contains one permission-controlled **Pre-orders** sidebar item. Inside the page are tabs for:

- Dashboard
- Products & Batches
- Customer Orders
- Back in Stock
- Notifications
- Reports
- Settings

Dashboard and Products & Batches are the first functional views; later tabs can be filled as their backend phases are completed.

## Performance work completed alongside the foundation

- Reorder Planner country/variant sales remain on-demand per expanded product rather than loading the whole Shopify report up front.
- Identical Reorder Planner ShopifyQL requests are briefly cached to avoid repeated calls during expand/collapse/reload usage.
- Product inventory reads are briefly cached and now support optional Shopify location filtering for future AU/USA separation.
- Repeated request-time `ALTER TABLE` checks were removed from normal Production Portal page loads because the corresponding Prisma migrations already exist.
- The legacy supplier-status cleanup was moved to a one-time migration rather than running an `updateMany` on ordinary staff navigation.
- Continue extracting heavy portal areas gradually rather than rewriting the 33,000-line portal route in one risky change.

## Important implementation principle

`On Production` + (`Send to AUS` or `Send to USA`) means **eligible for preorder**.

It does NOT mean **preorder automatically enabled**.

The explicit `Preorder Enabled` control is the final activation gate, and only authorised staff may operate it.
