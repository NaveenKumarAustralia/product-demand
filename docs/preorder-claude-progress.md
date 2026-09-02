# Pre-order — Claude session progress

Companion to `CLAUDE_PREORDER_HANDOFF.md`. Records what Claude Code verified,
fixed, and built after taking the handoff. Newest first.

## Session 1 (2 Sep 2026)

### Verified baseline (repo is source of truth)
- `npm run test:preorder` → **18/18 pass**; `npm run build` → **OK**; typecheck
  back to the app baseline (53) with **no preorder-specific TS errors**.
- Local Prisma client had to be regenerated after the new models (`prisma generate`).
- `shopify.app.toml` already carries all 13 scopes + `write_app_proxy` + the
  `[app_proxy]` config (Claude's earlier staged toml edit was therefore dropped).

### Fixed (deployed to main)
1. **Eligibility was 100% broken** — `activatePreorderSellingPlan` passed
   `productionStatus` to `getPreorderEligibility`, which reads `supplierStatus`.
   It was always `undefined` → every batch rejected as `not_on_production`, so
   **no batch could ever activate**. Renamed the field. *(commit ce698dd)*
2. **Blocker #1 (selling-plan attachment)** — activation passed
   `productIds:[order.productId]` alongside `variantIds`, associating the plan
   with the product's whole variant set. Now attaches to **only the batch's
   incoming variants** (`productIds:[]`). Added a test asserting
   `resources.productIds` is empty. Verified **deactivation** deletes the whole
   selling-plan group (clean removal from all variants). *(commit ce698dd)*
3. **Actor type** widened `admin` to optional to match `PortalMessageUser`
   (fixed 3 TS errors in `api.preorder-manage.tsx`). *(commit ce698dd)*

### Built (deployed to main)
- **AU/USA market isolation test** — proves the server resolver returns each
  market only its own plan even when both are attached to one variant. Server
  side is safe; the *native* Symmetry selector still needs a live-store check. *(e806871)*
- **Location dropdowns** in Pre-orders → Settings — pick from live Shopify
  locations (name · city, country), stores the exact GID, falls back to manual
  entry, warns if AU==USA. *(dcb73fc)*
- **Activation-readiness checklist** at the top of the Pre-orders Dashboard —
  aggregates scopes / webhooks / AU-USA locations (blockers) + Klaviyo +
  storefront (optional/manual) into one badge. *(after dcb73fc)*
- **Gated `api.preorder-klaviyo-health`** behind portal auth (was public; only
  leaked a `configured` boolean, now not publicly readable). *(security pass)*
- **Customer-account "My preorders" endpoint** — `op=my-preorders` on the signed
  app proxy returns the logged-in customer's reservations + expected ship dates
  (read-only; resolves email from `logged_in_customer_id`). Theme render is the
  remaining manual step.

### Reviewed & confirmed correct (no change needed)
- **Order webhook → allocation**: only `Karma East Pre-order` selling-plan lines
  processed; batch reference required (no guessing); **idempotent** on duplicate
  Shopify deliveries; per-order rollback on any line failure; capacity errors →
  HTTP 200 + logged exception (no retry loop); serializable transaction + retry;
  market/destination enforced.

### Remaining — needs a live store / merchant action (cannot be done in code)
1. **Protected-scope approval** in the dev dashboard + update Railway `SCOPES` +
   re-auth. The readiness panel goes green when this is done.
2. **Theme AU/USA native-selector check** — confirm a customer only ever sees
   the correct market's plan on the real Symmetry storefront.
3. **Controlled single-batch end-to-end test** (enable → activate → storefront
   out-of-stock shows Pre-order → checkout 100% → order webhook reserves exact
   batch → Customer Orders shows it → cancel releases it).
4. **Klaviyo Flow** — the app fires the *event*; a Flow triggered by
   `Karma East Back In Stock Available` must exist for an email to send. Test
   with a real (non-@test) address.
5. **Theme render** of the customer-account "My preorders" endpoint.
