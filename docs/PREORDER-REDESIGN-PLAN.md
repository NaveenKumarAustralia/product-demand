# Pre-order System — Redesign Plan (v2, verified against Shopify APIs)

Status: **proposed plan, not yet built.** Supersedes v1. This version is grounded in what Shopify's Admin API *actually* supports (verified Oct 2026 against the 2026-07/2026-10 docs), so the architecture only promises what the platform can deliver.

Goal: stop the recurring pre-order bugs **at the root** by leaning on Shopify's native inventory/selling-plan machinery for everything it does well, and keeping a **small, reconciled custom layer** only for the two things Shopify genuinely can't do.

---

## 1. What we verified (and why it shapes the design)

| Question | Answer | Source |
|---|---|---|
| Can we force **every** checkout (incl. Shop Pay/express/quick-add) through the pre-order plan? | **No.** Pre-order needs "continue selling when out of stock" ON, so any path can buy the variant **without** the plan. | [purchase-options/deferred](https://shopify.dev/docs/apps/build/purchase-options/deferred) |
| Does Shopify natively **cap** the pre-order quantity? | **No.** "Continue selling" oversells without limit. | same |
| Commit inventory at fulfillment, not at sale? | **Yes** — selling plan `inventoryPolicy.reserve: ON_FULFILLMENT`. | [SellingPlanReserve](https://shopify.dev/docs/api/admin-graphql/2026-10/enums/SellingPlanReserve) |
| Create & receive **incoming inventory** (a "PO") via API, with dates? | **Yes** — Inventory **Transfers + Shipments** (`inventoryShipmentCreate/MarkInTransit/Receive`, `dateShipped`/`dateReceived`), plus receive webhooks. | [inventoryShipment](https://shopify.dev/docs/api/admin-graphql/2026-04/queries/inventoryShipment) |
| Native **dispatch date** on the order? | **Partly** — selling plan `fulfillmentTrigger: EXACT_TIME` + `fulfillmentExactTime`; scheduled fulfillment auto-opens on `fulfillAt`. Customer sees it via the plan display, **plan orders only**. | [SellingPlanFulfillmentTrigger](https://shopify.dev/docs/api/admin-graphql/2026-04/enums/SellingPlanFulfillmentTrigger) |

**The two hard truths that drive the whole design:**
1. **Express/no-plan buyers can't be forced through the plan** → we must *capture* them (hold + notify) after the fact. Not removable.
2. **There's no native quantity cap** → we must cap it ourselves.

Everything else, Shopify can own — and we should let it.

---

## 2. Why it's been breaking (unchanged diagnosis)

Pre-order keeps ~7 pieces of state in sync between the **portal DB** and **Shopify**, each written at a *different* event point, with **no single source of truth and no reconciler**. Miss/fail any path → drift. Every big bug (orphaned plans, stuck "available", Maddison not showing, orders not splitting, emails missing, overselling) is an instance of that. See v1 history in git if needed; v2 keeps the reconciler idea but **moves the source of truth into Shopify's own numbers** so there's far less to keep in sync.

---

## 3. Target architecture (hybrid: Shopify-native accounting + a thin reconciled layer)

### 3.1 Shopify owns the accounting (new — the big simplification)

- **Enabling pre-order on a batch creates an Inventory Transfer/Shipment** in Shopify for that batch's quantity, destined to the market's location, with the batch's expected-arrival date. This is the "PO" idea from the brief — **Shopify's `incoming` column becomes the capacity source.**
- **Capacity = Shopify `incoming` − Shopify `committed`** (both native numbers), *not* a custom reservation ledger. This deletes the parallel `PreorderReservation` ledger as the source of truth (we may keep a thin copy for reporting only).
- **Receiving the shipment when stock lands** (via API on shipment load, or the merchant receiving it) moves `incoming → on hand`. That **ends the pre-order automatically** — the orphan class disappears *by construction*, no custom "close the plan" step.
- **Dispatch date = the transfer's expected arrival date** — one source, feeding both the selling plan (`EXACT_TIME`) and the portal.

### 3.2 The selling plan, driven from the above (simplified)

- Selling plan uses `inventoryPolicy.reserve: ON_FULFILLMENT` (commit at fulfillment) and `deliveryPolicy.fixed.fulfillmentTrigger: EXACT_TIME` = the transfer's arrival date.
- **Native order-confirmation email** then carries the "ships by X" info for plan orders — we can drop the custom Klaviyo "pre-order placed" email **for plan orders** (keep it only for the express/no-plan capture path, §3.4).

### 3.3 One declarative reconciler (the backbone — kept from v1, now reading Shopify's numbers)

`reconcilePreorderForProduct(shop, productId)` — the **single writer** of Shopify-side pre-order state. It reads desired state (is there a live batch? what's `incoming − committed`?) and makes Shopify match:
- Selling plan group present iff a live batch exists; deleted otherwise.
- Variant `inventoryPolicy`: **CONTINUE** while `incoming − committed > 0` (capacity left), **DENY** when it hits 0 (**this is the quantity cap** — custom, but now driven by Shopify's own committed count, not our ledger).
- Pre-order tag/metafields set iff live.
- Idempotent; runs on **every relevant event**. **No automatic background sweep** (by decision — nothing runs on its own). The only backstop is a **manual "Reconcile now"** on the health page. Replaces all the scattered one-off syncs (`denyClosedPreorderOrderVariants`, `removeClosedPreorderSellingPlans`, etc.).
  - *Trade-off accepted:* without a sweep, a genuinely dropped webhook won't self-correct — but the health page will **show** the drift so it's caught, and "Reconcile now" fixes it on demand.

### 3.4 Express/no-plan capture + holds, as durable jobs (kept from v1)

Shopify can't force express buyers through the plan, so on `orders/create`:
- Detect lines on a variant that is **currently a live pre-order** (plan exists + `incoming − committed` logic) bought **without** the plan → treat as pre-order.
- Enqueue durable, retried, idempotent **jobs**: **hold** the pre-order lines' fulfillment, **tag** the order, **notify** the customer (the pre-order email that Shopify's native email won't send for no-plan orders).
- Plan-path orders need only the **hold** job (Shopify already tags/emails them).

Holds are the split mechanism (Shopify doesn't auto-split): in-stock lines ship now, pre-order lines' fulfillment is **held** → they **only appear in Pick Pack when released**. Release happens when §3.1 receives the stock.

### 3.5 Health surface (kept from v1)

One admin view + JSON: drift (anything in Shopify not backed by a live batch), failed jobs (orders not held, emails not sent), `write_orders` scope status + re-auth prompt, and a "reconcile now" button. Folds in today's audit/debug endpoints.

---

## 4. Area-by-area target behaviour

| Area | How it works in v2 |
|---|---|
| **Turn pre-order on/off** | Lives in the portal. ON → create the Shopify inventory transfer (incoming) + selling plan (EXACT_TIME date) + reconcile. OFF → delete transfer's incoming + plan + reconcile. **Validated**: can't enable without a destination + matched variants (the Maddison gap). |
| **Capacity / cap** | `incoming − committed` from Shopify. Reconciler flips CONTINUE/DENY. No custom ledger as truth. |
| **Dispatch date** | The transfer's arrival date → selling plan `EXACT_TIME` → native email (plan orders) + portal. |
| **Order split** | Fulfillment **holds** on the pre-order lines (job, retried, verified). In-stock ships now. |
| **Release → Pick Pack** | Receiving the shipment (stock loaded) moves incoming→on-hand, reconciler updates, holds release → **now** visible in Pick Pack. |
| **Emails** | Plan orders → Shopify's native confirmation (with date). No-plan/express orders → the capture job sends the pre-order email. One delivery log for both. |
| **End of pre-order** | Automatic on receive. No orphan cleanup needed. |
| **Express / Shop Pay / any payment** | Captured + held + emailed by the job layer, so behaviour is the same regardless of path (the one thing that *can't* be native). |
| **Storefront vs portal agreement** | One shared `computePreorderAvailability()` used by both; portal can't promise what the storefront won't serve. |

---

## 5. Staged build plan (each phase ships independently; no big-bang)

> Principle: additive and reversible. The 3 currently-live pre-orders keep working throughout — the reconciler is desired-state driven and never turns off a real batch.

### Phase 0 — Safety net now (0.5 day)
Schedule the existing policy sweep + keep the audit/cleanup endpoints, so nothing silently regresses before the reconciler lands. Add counters to measure failure rates before/after.

### Phase 1 — The reconciler + shared eligibility (3–4 days) — *highest value*
- Build `reconcilePreorderForProduct` as the **single writer**; wire into every event + a schedule; delete scattered syncs.
- Cap driven by Shopify `incoming − committed`.
- One `computePreorderAvailability()` for portal + storefront; **block enabling without destination/variants**.
- **Kills: orphans, stuck CONTINUE, overselling drift, and the Maddison class.**

### Phase 2 — Shopify-native accounting (4–5 days)
- On enable: create Inventory Transfer/Shipment (incoming) with the arrival date; on disable: remove it.
- Selling plan → `reserve: ON_FULFILLMENT` + `EXACT_TIME` date from the transfer.
- On shipment load: **receive** the Shopify shipment (incoming→on-hand) instead of our custom close path → pre-order ends natively.
- Drop the custom "pre-order placed" email **for plan orders** (native email covers them).
- **Kills: the whole capacity-ledger-drift + dispatch-date + auto-end class.**

### Phase 3 — Durable jobs for capture / holds / emails (4–5 days)
- `PreorderJob` table + worker + retries + idempotency.
- Express/no-plan capture, fulfillment holds (verified by re-reading the order), and the no-plan pre-order email all become jobs.
- **Kills: orders not splitting, emails not sent (now retried + visible).**

### Phase 4 — Health surface + scope preflight (1–2 days)
Drift, failed jobs, `write_orders` status + re-auth banner, "reconcile now". Fold in existing debug/audit endpoints.

### Phase 5 — Extract + test + docs (ongoing)
Move pre-order logic out of `portal._index.tsx`; unit-test the reconciler + eligibility; integration-test order-create → capture → hold → receive → release; update `docs/PREORDER-SYSTEM.md`.

---

## 6. Guardrails (to stop regressions)

1. **One writer** of Shopify-side pre-order state: only the reconciler.
2. **One rule set**: portal and storefront call the same eligibility function.
3. **Shopify's numbers are the source of truth** for capacity (incoming/committed), not a custom ledger.
4. **No silent side-effects**: every hold/tag/email is a job with a visible outcome — no `.catch(() => undefined)` on customer-impacting actions.
5. **Validate before promising**: can't enable a batch the storefront can't serve.
6. **Every event path calls the reconciler** (no bespoke partial syncs). No automatic sweep by decision; the health page surfaces any drift and "Reconcile now" fixes it manually.
7. **Express is a first-class path**, not an afterthought: the capture job is part of the core flow.

---

## 7. Decisions

**Decided (Oct 2026):**
- ✅ **Scopes + re-auth approved** — we will request `write_orders` (reliable holds/tags) and `read/write_inventory_transfers` + `read/write_inventory_shipments` (native incoming stock) and re-authorise the app.
- ✅ **Payment timing: full charge at checkout** (as today) — no deposit / charge-on-fulfillment for now.
- ✅ **Emails:** plan orders use Shopify's **native** order-confirmation email (carries the dispatch date) — no Klaviyo for those. Klaviyo's role shrinks to the **express/no-plan capture email** and **back-in-stock** (sender for the capture email can move off Klaviyo later; not urgent). Note: a selling plan can't be attached to an order after checkout, so express orders can't be retro-made native — hence the separate capture email.

**Still open:**
- ~~Reconcile cadence~~ — **Decided: no scheduled sweep.** Reconciler runs on events only; manual "Reconcile now" is the backstop.
- **Phase order** — Phase 1 (reconciler) first as written (recommended), or pull Phase 3 (splitting + emails) earlier if those hurt most right now. A specific actively-burning split/email bug can be hotfixed in parallel without reordering.
- **AU/USA** — confirm each market maps to its own Shopify location for the transfer destination (USA fallback-to-AU behaviour stays?).

---

*Next step: confirm phase order, then start Phase 1 — biggest reliability win, safe for the 3 live pre-orders.*
