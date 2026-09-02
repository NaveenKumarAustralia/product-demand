# Karma East Pre-order System — Claude Handoff

**Date:** 2 September 2026
**Repository:** `NaveenKumarAustralia/product-demand`
**Purpose:** This is the handoff from the ChatGPT development conversation so Claude Code can take over immediately. Treat the repository code as the source of truth where it differs from this document.

## Naveen's instruction

Continue building the preorder system toward safe activation. Work independently, make as much progress as possible, test before deployment, and deploy safe completed work to `main` so Railway auto-deploys it. Avoid repeatedly asking Naveen questions that can be answered from the code/repo. Do not turn on broad customer-facing preorders until the activation blockers below are resolved and tested.

## Product goal

Build a proper Shopify preorder system integrated directly into the existing Karma East Production Portal. This must use Shopify's real purchase-option / selling-plan preorder framework rather than simply enabling “continue selling when out of stock”.

Incoming production quantities and expected dispatch dates come from the existing Production Portal. Preorder capacity is tied to a specific production batch and capped by confirmed incoming quantity minus a safety buffer and existing reservations. Customers pay 100% upfront in v1.

There are separate AU and USA inventory/preorder pools. Never cross-allocate them.

The portal Pre-orders area has these tabs:
- Dashboard
- Products & Batches
- Customer Orders
- Back in Stock
- Notifications
- Reports
- Settings

A customer-account Pre-orders section showing expected shipping date is planned but not yet built.

## Business rules

### Batch eligibility
A production batch can only become eligible when:
1. `supplierStatus === on_production`
2. destination is `send_to_au` or `send_to_usa`
3. staff explicitly enables preorder for the batch.

Changing production status/destination alone must NEVER expose preorder inventory. `Preorder Enabled` is a deliberate final gate.

### Capacity
For each variant:
- incoming = `max(0, qtyOrdered - qtyReceived)`
- safety buffer percentage rounds up with `Math.ceil`
- fixed safety quantity overrides percentage when supplied
- remaining = incoming - safety - reserved
- available = `max(0, remaining)`
- overallocated = `max(0, -remaining)`

If incoming production is reduced below already reserved preorder quantity, do not silently cancel existing customers. Preserve reservations and surface an overallocated exception.

### Exact batch binding
A real Shopify preorder must remain tied to the exact production batch promised to the customer. Selling plan names include `Batch #123`. Order normalization parses that batch number and allocation requires `preferredSupplierOrderId`. Do not silently move an order to a later batch.

### Markets
- US customer/order => USA production pool/location.
- AU, NZ and ROW => AU production pool/location.
- Never cross allocate AU and USA.

### Storefront state
For a variant in a market:
1. If physical Shopify inventory at that market's configured location is > 0 => normal in-stock Add to Cart wins.
2. If physical stock is 0 and there is an eligible + explicitly enabled + Shopify-active selling plan with remaining capacity => Pre-order.
3. Otherwise => Notify Me / back-in-stock waitlist.

If several eligible batches exist, choose the earliest expected ship date, then lowest batch ID.

### Receiving
When stock arrives, receive the full physical quantity into Shopify. Example: 200 expected, 198 received, 126 preorder commitments => receive all 198; 126 are committed and 72 are ordinarily available. Do not load only 72 into Shopify inventory.

## Permissions

Pre-orders page access uses the existing Production Portal per-user page access system:
`Settings → Users → Page access → Pre-orders`.

Superadmin always has access. Direct URLs are server guarded.

Preorder action permissions are stored in each existing portal user record and edited in `Settings → Users`, not in a duplicate permissions UI inside Pre-orders. Permission categories currently cover management, ETA, safety buffer, notifications and reports.

## What has been built

### Foundation / allocation ledger
- Preorder batch settings.
- `PreorderReservation` model/migration.
- Transactional capacity reservation logic.
- FIFO ordinary allocation within the same market.
- Serializable transaction handling / P2034 retry.
- Cancellation releases reservations.
- Fulfillment marks reservations fulfilled.
- Production status rollback can pause new allocation without deleting existing commitments.
- Customer Orders ledger UI.
- Reports and allocation exception reporting.

### Shopify webhook processing
Shopify order create/cancel/fulfilled handlers exist. Only selling-plan lines whose plan name begins with `Karma East Pre-order` are treated as preorders.

Order create normalization reads the selling plan name and extracts the exact `Batch #ID`. A preorder line without the batch reference fails allocation rather than guessing another batch.

Webhook registration exists and is intentionally non-fatal during application startup. A previous production incident occurred because `/app/scripts/register-preorder-webhooks.mjs` was absent from the Docker image. This was fixed by copying scripts into Docker and making webhook registration non-fatal. Do not regress this.

### Shopify readiness
There is a protected admin readiness endpoint/panel that checks the installed app's live Shopify scopes.

Required scopes currently include:
- `write_products`
- `read_all_orders`
- `read_customer_payment_methods`
- `read_purchase_options`
- `write_purchase_options`
- `read_payment_mandate`
- `write_payment_mandate`
- `write_app_proxy`

### Selling plan builder / activation
A 100%-upfront Shopify selling-plan builder and activation service exist.

Selling plan configuration is intended as:
- PRE_ORDER
- 100% charged at checkout
- NO_REMAINING_BALANCE
- delivery UNKNOWN
- inventory policy ON_FULFILLMENT

Selling plan registry entries are stored in `PortalSetting` using key format:
`preorder-selling-plan-v1:${shop}:${supplierOrderId}`.

Staff controls in Products & Batches allow:
- Expected dispatch date
- safety buffer %
- fixed buffer qty override
- enable/pause preorder internally
- Activate on Shopify
- Remove from Shopify

Activation is blocked unless the batch is enabled, eligible and has available capacity.

### Regional storefront engine
A server-side resolver returns `in_stock`, `preorder`, or `notify_me` based on the market's physical inventory and live preorder candidates.

Physical stock always wins. AU and USA never cross.

### Storefront app proxy + theme extension
`shopify.app.toml` has app proxy configuration:
- backend `/preorder-proxy`
- storefront `/apps/karma-east-preorder`

Theme extension: `extensions/preorder-product-block`
- product app block
- detects AU vs USA
- displays Pre-order or Notify Me based on server state
- preorder cart add includes numeric `selling_plan`
- Notify Me posts to the signed app proxy
- does not modify Symmetry theme source files directly.

The native Symmetry sold-out button may still coexist with the app block; storefront/theme polish remains.

### Back in Stock waitlist
`PreorderWaitlist` is stored in DB/raw SQL support.
Storefront signup is connected through the signed app proxy.
Repeat signup upserts/reset waiting state.

Back in Stock portal UI can send a Klaviyo event manually for a waiting recipient.

### Klaviyo integration
Environment variable expected:
`KLAVIYO_PRIVATE_API_KEY`

Naveen created a Klaviyo Full Access private API key and added it to Railway under that exact variable name.

Server integration sends events to Klaviyo, never exposing the key to browser code.

Current event metrics:
- `Karma East Back In Stock Available`
- `Karma East Preorder Update`

Back-in-stock event uses a deterministic unique ID like:
`back-in-stock:${shop}:${waitlistId}`
so retries can be deduplicated by Klaviyo.

Important: the app currently creates a Klaviyo EVENT. An actual customer email requires a Klaviyo Flow triggered by metric `Karma East Back In Stock Available`. Do not claim the event itself sends the email unless a Flow is active.

First end-to-end test should use a real Naveen-controlled email, not `@test` or `@example`, because Klaviyo may drop test/example addresses.

### Klaviyo diagnostic issue
Naveen added/deployed the Railway variable but the portal initially still displayed “Not connected”. The original UI incorrectly mapped any status failure to a generic missing-key message.

A diagnostics hotfix was created and pushed to `main` at commit:
`44763f824bfee36fb60915914808cde93ae45b82`

It improves `/api/preorder-klaviyo-status` and the Back in Stock panel so it can distinguish:
- expected variable present/not present
- value length (never secret contents)
- names of runtime environment variables containing `KLAVIYO`
- API/status/auth failures.

The workflow run for that diagnostics branch succeeded before it was fast-forwarded to main. At handoff time, the Railway deployment status of this exact commit still needed to be rechecked. Do not assume it is live until Railway/GitHub deployment status confirms success.

There is also a public-ish runtime diagnostic route `api.preorder-klaviyo-runtime.tsx` in the repo that returns only configured/revision and no key. Review whether it should remain accessible before launch.

## Important activation blocker: selling-plan resource attachment

THIS IS A HIGH PRIORITY REVIEW BEFORE BROAD ACTIVATION.

The selling-plan activation builder was observed passing both:
- `productIds: order.productId ? [order.productId] : []`
- `variantIds`

This may attach the preorder selling-plan group to all variants of the product rather than only the variants that actually have confirmed incoming capacity in that production batch.

Before activating customer-facing preorders broadly:
1. Inspect the current Shopify `sellingPlanGroupCreate` resource semantics for the API version in use.
2. Prefer exact `variantIds` only unless every product variant is intentionally eligible.
3. Add tests that prove a batch containing only selected variants cannot expose preorder purchase options on unrelated variants.
4. Ensure removing/deactivating the plan cleans up correctly.

This is probably the most important code review item before real activation.

## Important activation blocker: AU/USA plan visibility

A product may have separate AU and USA production batches/selling plans. Shopify/theme native purchase-option rendering could expose more than one plan if both are attached to the same product/variant.

The custom storefront state engine selects the correct market plan, but real Symmetry storefront behavior still needs testing. Ensure a customer sees only the correct market preorder action and cannot select the wrong market's selling plan.

Do not blindly expose multiple AU/USA selling plans.

## Shopify locations

Backend endpoint `api.preorder-shopify-locations.tsx` already queries active Shopify locations. Settings UI still uses raw AU/USA location ID text inputs.

Next useful improvement: wire the location discovery endpoint into Settings as dropdowns showing human-readable location name/city/country, while storing the exact Shopify location GID. This reduces setup mistakes before activation.

## Customer notifications roadmap

There is already a Notifications draft/approval queue, but it is not yet fully wired to Klaviyo delivery.

Planned notification types include:
- Back in stock available
- preorder ETA changed
- production delay
- ready/fulfilled

Recommended rollout:
1. prove one manual Back in Stock Klaviyo event end-to-end
2. create/verify Klaviyo Flow
3. detect actual physical-stock/preorder-capacity availability by market/location
4. build an approval queue for eligible waitlist recipients
5. only later consider optional automatic sending.

Back-in-stock availability should be based on actual Shopify physical inventory at the correct regional location, or genuinely available preorder capacity as explicitly designed—not simply “production incoming exists”.

## Customer account

Still to build: a simple Shopify customer-account Pre-orders section showing preorder order/line and expected dispatch/shipping date. Keep it simple for v1.

## Pick & Pack integration later

Karma East has a separate custom Pick & Pack app. Future preorder integration should prevent held preorder lines from entering the pick queue until ready. Once stock is received and the preorder becomes ready, the line can enter the normal queue. Do not build a second warehouse system in the Production Portal.

## Existing warehouse app context
- Shopify embedded app on iPad
- staff passcode/name login
- choose date/country/service/single-multi
- Smart Tub capacity 40
- Express separate
- oldest orders first
- Starshipit packing.

## Core guardrails

1. Do not create a second production system; use `SupplierOrder` and existing production data.
2. Do not use editable `SupplierOrder.eta` as the only production timing signal. It is an editable expected date, not production-entry time.
3. Do not combine AU and USA inventory/capacity.
4. Keep major preorder logic in focused modules, not a giant portal index route.
5. Do not use negative Shopify inventory as the reservation ledger.
6. Capacity must be revalidated transactionally at reservation time, not only calculated on page load.
7. Do not silently unreserve customers when incoming production falls.
8. Do not cross allocate regions.
9. Do not expose activation actions without server-side permission checks.
10. Do not turn on broad customer-facing preorders until scopes, ledger, webhooks, market/storefront behavior and selling-plan attachment are tested.
11. Do not deploy a failing build.
12. Do not weaken TypeScript/compiler checks to make preorder code pass.
13. Optional webhook registration must never block application startup.
14. Keep Docker/runtime CI checks for preorder work.
15. Do not send customer notifications before intentional delivery integration/testing.
16. Never silently move a real preorder to a later production batch.
17. Never blindly expose multiple AU/USA selling plans.
18. Review the Phase 11 productId + variantIds attachment concern before broad launch.
19. Never expose Klaviyo private keys in browser responses, logs, GitHub or chat.

## CI / deployment

Workflow:
`.github/workflows/preorder-check.yml`

It has included checks such as:
- `npm ci`
- `prisma generate`
- preorder tests
- full TypeScript baseline
- production build
- theme extension checks
- Docker image build
- runtime webhook script checks/diagnostics.

Railway auto-deploys `main`.
Production portal URL is:
`https://product-demand-production.up.railway.app/auth-portal`

GitHub commit status context previously used for Railway:
`Product demand - product-demand`

At handoff, main was moved to:
`44763f824bfee36fb60915914808cde93ae45b82`
for the Klaviyo diagnostics hotfix. Verify current main and deployment status before starting new changes.

## Suggested immediate sequence for Claude

1. Pull latest `main` and run the complete preorder CI/test/build locally.
2. Verify commit `44763f824bfee36fb60915914808cde93ae45b82` is actually deployed/healthy on Railway; inspect current repo history because later commits may now exist.
3. Resolve the selling-plan resource attachment blocker: inspect activation code, switch to exact variant attachment if Shopify semantics confirm the concern, and add tests.
4. Review regional selling-plan visibility and theme extension behavior so AU customers cannot use USA plans and vice versa.
5. Wire Shopify location discovery into Settings dropdowns and validate that AU and USA locations differ.
6. Add a clear “Activation readiness” checklist/status in the portal that combines: Shopify scopes, webhook health, AU/USA locations, storefront proxy/theme block readiness, Klaviyo connection (for notifications), and critical configuration.
7. Test one controlled product/batch end-to-end before broad activation: enable batch → activate selling plan → storefront out-of-stock variant shows Pre-order only in correct market → checkout 100% → order webhook reserves exact batch → portal Customer Orders displays it → cancellation releases it.
8. Test Back in Stock with a real controlled email and verify Klaviyo event, then configure/verify the Klaviyo Flow before treating “Send via Klaviyo” as actual email delivery.
9. Build the simple customer-account preorder expected-date view.
10. Only after the controlled end-to-end test passes, prepare a safe limited activation rollout.

## Useful file map

Likely relevant files/modules include:
- `app/preorder/preorder-allocation.server.ts`
- `app/preorder/preorder-dashboard.server.ts`
- `app/preorder/preorder-rules.server.ts`
- `app/preorder/preorder-shopify-order-normalize.ts`
- `app/preorder/preorder-shopify-orders.server.ts`
- `app/preorder/preorder-storefront-state.ts`
- `app/preorder/preorder-selling-plan-registry.server.ts`
- selling plan builder/activation modules under `app/preorder/`
- `app/preorder/preorder-klaviyo.server.ts`
- `app/preorder/preorder-waitlist-panel.tsx`
- `app/preorder/preorder-operations-panels.tsx`
- `app/preorders-dashboard.tsx`
- `app/routes/api.preorder-manage.tsx`
- `app/routes/api.preorder-shopify-readiness.tsx`
- `app/routes/api.preorder-shopify-locations.tsx`
- `app/routes/api.preorder-klaviyo-status.tsx`
- `app/routes/api.preorder-klaviyo-runtime.tsx`
- `app/routes/api.preorder-waitlist-notify.tsx`
- `app/routes/preorder-proxy.tsx`
- `extensions/preorder-product-block/`
- `shopify.app.toml`
- `.github/workflows/preorder-check.yml`
- `Dockerfile`
- `scripts/register-preorder-webhooks.mjs`

Search the repository for `preorder` to get the authoritative complete file set.

## Final note to Claude

Naveen wants this project finished and activated soon, but safety is more important than prematurely exposing customer purchase options. Make concrete progress without unnecessary back-and-forth, keep changes small/testable, deploy passing work, and make the final activation a controlled real-product test before scaling it across the store.
