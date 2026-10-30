# Pre-order System — Redesign Plan

Status: **proposed plan, not yet built.** Goal: stop the recurring pre-order bugs at the root, not one instance at a time.

This document is the agreed list of work before building. It (1) explains *why* the pre-order system keeps breaking, (2) catalogues every failure class we've hit, (3) describes the target architecture that makes those classes impossible, and (4) gives a staged, shippable build plan.

---

## 1. TL;DR — why it keeps breaking

Pre-order needs ~7 pieces of state kept in sync between the **portal DB** (source of production truth) and **Shopify** (what customers see and buy):

| # | State | DB side | Shopify side |
|---|-------|---------|--------------|
| 1 | Selling plan exists? | `preorder-selling-plan-v1:*` registry | Selling plan group on the product |
| 2 | Can oversell? | batch remaining capacity | variant `inventoryPolicy` CONTINUE/DENY |
| 3 | Is a pre-order? | batch enabled | variant pre-order metafields + product tag |
| 4 | Reserved vs capacity | `PreorderReservation` ledger | real paid orders |
| 5 | Ship separately | reservation / combine window | fulfillment **holds** on the pre-order lines |
| 6 | Storefront state | computed live from 1–4 | the theme block + native selling plan |
| 7 | Customer told | — | Klaviyo events → flows (emails) |

Today each of these is written at a **different event point** — activate, deactivate, products/update webhook, orders/create webhook, shipment-load, a scheduler — and **there is no single source of truth and no single reconciler.** The moment *any one* path is missed or fails, Shopify and the DB drift apart.

Every major bug this project has hit is an instance of that one flaw:

- **90+ orphaned selling plans** (products stuck pre-orderable/oversellable): the shipment-load path reverted inventory policy but not the selling plan; a batch closed by any *other* path did neither. → state #1/#2 drifted.
- **Maddison M/L showed "notify me" on the product page while the portal said "available"**: the portal's "available" number and the storefront's eligibility use **different rules** — the portal counted stock from a batch with no destination, which the storefront (correctly) won't serve. → #2/#6 disagree.
- **Orders not splitting in Shopify**: the fulfillment **hold** is best-effort and depends on the `write_orders` scope and a fragile "release-all-then-re-hold" step; when it fails it's only logged. → #5 silently partial.
- **Emails not going out correctly**: notifications are fire-and-forget Klaviyo events whose actual send depends on a **live Klaviyo flow** we can't see, plus subtle plan-vs-no-plan branching. → #7 unverifiable.

**The fix is not a rewrite — it's changing the sync model** from "many bespoke point-syncs" to **one declarative reconciler** plus **durable, retried side-effects**, with **one shared rule set** and a **visible health surface**.

---

## 2. Failure catalogue (every class we've hit) → root cause

| Failure | What the customer/staff saw | Root cause pattern |
|---|---|---|
| Orphaned selling plans | Closed products still pre-orderable / oversellable on collection grid | A. No reconciler; side-effects scattered across event paths |
| Stuck `CONTINUE` inventory policy | "Available/blue" on collection, "sold out" on product page | A + C (best-effort, not reconciled) |
| Pre-order not showing when enabled | Maddison M/L: portal "available", storefront "notify me" | B. Portal and storefront use different eligibility maths; no validation that an enabled batch is actually serveable (e.g. destination set) |
| Orders not splitting | In-stock + pre-order not separating; whole order stuck or pre-order ships early | C + E. Holds are best-effort, depend on `write_orders`, fail soft |
| Emails missing/wrong | No "your order includes a pre-order" / no back-in-stock | D. Fire-and-forget events; depends on unseen Klaviyo flow; plan/no-plan branching |
| Overselling / "ghost" allocations | More sold than the batch holds | #2/#4 drift on express checkout paths |
| Hard to even diagnose | "many more products enabled than I set" | F. No single pre-order health view; failures only in server logs |

**The five cross-cutting root causes:**

- **A. No single source of truth + no reconciler.** The DB batch state *should* be the truth; Shopify should be made to match it. Instead both are edited directly from many places.
- **B. Two different "is this pre-orderable?" computations.** The portal dashboard promises availability the storefront won't serve (and vice-versa).
- **C. Side-effects are best-effort, fire-and-forget.** Holds, tags, emails, policy writes log on failure and move on — leaving silent partial state, no retry, no surfacing.
- **D. Logic is spread across event paths** with no guarantee each path performs the *full* set of required actions.
- **E. Hard dependencies fail soft** (e.g. missing `write_orders` scope) with no clear "action needed" signal.
- **F. No observability.** Nothing shows drift/exceptions; problems are found by customers, not the system.

---

## 3. Target architecture

Four building blocks. Each directly kills one or more root causes.

### 3.1 One declarative reconciler (kills A, D)

A single idempotent function:

```
reconcilePreorderForProduct(shop, productId)
```

It reads the **desired state** from the DB (is there an OPEN, enabled, destined batch with capacity for each variant?) and makes Shopify match it, for the whole product at once:

- Selling plan group: **present** iff a live batch exists for the product; otherwise **deleted**.
- Each variant `inventoryPolicy`: **CONTINUE** iff that variant has remaining capacity in a live batch; else **DENY**.
- Variant pre-order metafields + product tag: set iff live; else cleared.

Rules:
- **Idempotent** — running it twice is a no-op if already correct.
- Called from **every** event that can change desired state: batch enable/disable, destination change, status change, products/update, orders/create (capacity used), orders/cancelled (capacity freed), shipment-load/close, and a **scheduled sweep** (self-healing baseline).
- Replaces the scattered one-off calls (`denyClosedPreorderOrderVariants`, `removeClosedPreorderSellingPlans`, ad-hoc activate/deactivate side-effects, the partial `reconcileAllLivePreorderInventoryPolicies`).

**Result: orphans and drift become impossible** — any plan/policy not backed by a live batch is corrected on the next event or sweep. No more manual cleanup endpoints.

> There is already a seed of this for inventory policy only (`reconcilePreorderInventoryPolicyForVariants` / `reconcileAllLivePreorderInventoryPolicies` in `preorder-inventory-policy.server.ts`). The work is to generalise it to **all** Shopify-side state and make it the *only* writer.

### 3.2 One shared eligibility/capacity module (kills B)

Both the portal dashboard and the storefront resolver must call the **same** function to answer "can this variant be pre-ordered, and how many?":

```
computePreorderAvailability(shop, variantId, market) -> { serveable, availableToPreorder, batchId, reason }
```

- The portal dashboard shows exactly what the storefront will serve — never a number the storefront can't honour.
- Add **validation at enable time**: a batch cannot be enabled/activated unless it has a valid destination (`send_to_au`/`send_to_usa`) and at least one matched variant line — the exact gap that broke Maddison. Enabling without a destination should be refused with a clear message, not silently accepted.
- Today the portal uses `preorder-dashboard.server.ts` capacity and the storefront uses `resolveStorefrontVariantState` (`preorder-storefront-state.ts`). These get unified behind one module; both callers consume it.

### 3.3 Durable, retried side-effect jobs (kills C, E)

Holds, order tags, and customer notifications stop being inline fire-and-forget. They become **idempotent jobs** recorded in a `PreorderJob` table with status (`pending`/`done`/`failed`) + attempt count + last error:

- The orders/create webhook enqueues jobs (hold, tag, notify) and returns fast.
- A worker (scheduler tick + on-demand) runs them with **bounded retries**.
- A job that keeps failing (e.g. `write_orders` missing) surfaces as an **exception** on the health dashboard (§3.4), not a buried log line.
- **Idempotency keys** (order id + job type) make retries safe.

This fixes "orders not splitting" (a failed hold is retried and visible) and "emails not sent" (a failed notify is retried and visible).

### 3.4 A pre-order health surface (kills F)

A single admin view (and a JSON endpoint) that shows, at a glance:

- **Drift**: any Shopify selling plan / CONTINUE policy not backed by a live batch (what the audit endpoint does now, made permanent + automatic).
- **Exceptions**: failed jobs — orders whose pre-order line wasn't held, notifications that didn't send, allocations that need review.
- **Scope status**: is `write_orders` granted? If not, a clear "re-authorise the app" banner (fixes the silent soft-fail class E).
- **Reconcile now** button that runs the sweep.

The existing `/api/preorder-plan-audit`, `/api/preorder-storefront-debug`, `/api/preorder-reconcile-policies`, and `preorder-restock-audit` get folded into this one surface.

---

## 4. Area-by-area target behaviour

### 4.1 Selling-plan / inventory / tag sync
Owned entirely by the reconciler (§3.1). No other code writes these. Closing a batch, loading a shipment, disabling pre-order, deleting a batch — all just change DB state and call the reconciler.

### 4.2 Storefront ↔ portal agreement
One eligibility module (§3.2). Enable-time validation blocks un-serveable configs (no destination, no matched variants). The storefront debug view becomes part of the health surface.

### 4.3 Order splitting (holds)
- Hold = a durable job (§3.3) with retry + verification (re-read the order to confirm the pre-order lines are actually on hold).
- One **unified hold per order** computed from the reservation ledger (keep the current "whole order, once" intent — the bug was re-holding subsets clobbering each other).
- If `write_orders` is missing, the job fails loudly on the health surface with a re-auth prompt, instead of silently not splitting.
- Release on stock-landing is the same job type in reverse, also verified.

### 4.4 Notifications (emails)
- One notification service with a **delivery log** (what fired, for which order, when, success/failure) — so "did the email go out?" is answerable.
- Clear, documented matrix of which event fires on which checkout path (plan/button vs Shop Pay/express/no-plan) so we stop guessing. Pre-order-placed, dispatch, and back-in-stock all go through it.
- A **Klaviyo flow preflight**: the health surface checks the metric exists and warns if the flow likely isn't live (we can't read flow status directly, but we can detect "events fired, zero recipients").

### 4.5 Overselling / ghosts
Falls out of §3.1 + §3.3: capacity drives CONTINUE/DENY on every path, and express/no-plan orders are captured + held as jobs. Covered already in part by `preorder-missed-capture.server.ts`; it becomes a job type.

---

## 5. Staged build plan (each phase ships independently)

> Principle: additive and reversible. No big-bang rewrite of the 33k-line `portal._index.tsx`. Each phase leaves the system working and better than before.

### Phase 0 — Lock in what we just fixed + measure (0.5 day)
- Keep the audit/cleanup endpoints; schedule the existing policy sweep so orphans can't silently return before the reconciler lands.
- Add counters so we can see failure rates before/after.

### Phase 1 — The reconciler (highest value) (2–3 days)
- Build `reconcilePreorderForProduct` as the single writer of selling-plan + inventory-policy + metafields + tag.
- Wire it into every event path; delete the scattered one-off sync calls.
- Add the scheduled sweep (self-healing).
- **Kills: orphans, stuck CONTINUE, most drift.**

### Phase 2 — Shared eligibility + enable-time validation (1–2 days)
- Extract one `computePreorderAvailability`; point the dashboard and storefront at it.
- Refuse enabling a batch with no destination / no matched variants (with a clear reason).
- **Kills: portal/storefront disagreement (Maddison class).**

### Phase 3 — Durable jobs for holds + notifications (3–4 days)
- `PreorderJob` table + worker + retries + idempotency.
- Move holds, tags, and notifications onto it; verify holds by re-reading the order.
- **Kills: orders not splitting, emails not sent (now retried + visible).**

### Phase 4 — Health surface + scope preflight (1–2 days)
- One admin page: drift, exceptions, scope status, "reconcile now".
- `write_orders` detection + re-auth banner.
- Fold existing debug/audit endpoints in.
- **Kills: F (invisible failures) and E (silent scope soft-fail).**

### Phase 5 — Extract + test + document (ongoing)
- Move pre-order business logic out of `portal._index.tsx` into the pre-order modules.
- Unit tests for the reconciler (desired-state → actions) and eligibility; integration test for order-create → reserve → hold → notify.
- Update `docs/PREORDER-SYSTEM.md` to match reality.

---

## 6. Guardrails to prevent regression

1. **One writer per Shopify-side state.** Only the reconciler writes selling plans / inventory policy / pre-order metafields / tags. Grep-enforceable in review.
2. **One rule set.** Portal and storefront must call the same eligibility function. No second capacity calculation.
3. **No silent side-effects.** Every hold/tag/email goes through the job system with a visible outcome. No `.catch(() => undefined)` on customer-impacting actions.
4. **Validate before promising.** A batch can't be enabled in a state the storefront can't serve.
5. **Self-healing by default.** The scheduled sweep means a missed event path self-corrects within the sweep interval instead of becoming a permanent orphan.
6. **Observability first.** If it can fail silently, it must appear on the health surface.

---

## 7. Open decisions (need your call before/early in the build)

- **Reconcile trigger cadence** for the scheduled sweep (e.g. every 15 min vs hourly) — trade-off between Shopify API load and self-heal speed.
- **Email ownership**: keep Klaviyo flows (we fire events) or move transactional pre-order emails somewhere we can verify delivery end-to-end.
- **Scope**: is re-authorising the app for `write_orders` acceptable (needed for reliable holds/tags), and who does it.
- **Phasing order**: this plan does reconciler → eligibility → jobs → health. If "orders not splitting" and "emails" are the most painful *right now*, we can pull Phase 3 earlier at the cost of doing it before the reconciler simplifies things.

---

*Next step: confirm the phase order and the open decisions in §7, then start Phase 1.*
