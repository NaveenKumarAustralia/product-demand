# Pre-order — launch checklist

Living list of what's left before AU pre-orders are fully "on" for real customers.
Priorities: 🔴 blocks going broad · 🟡 do soon · 🟢 later / when USA.
"(live)" = needs the real store / merchant action; "(code)" = Claude can build it.

Updated: 2026-09-03.

---

## ✅ Done so far
- Scopes approved + Railway env + re-auth; readiness panel green.
- App deployed to Shopify (theme block "Pre-order / Notify" + app proxy, absolute URL).
- AU fulfilment location set; USA parked (plugs in later, no code change).
- Per-row Status-cell "Enable pre-order" control (permission-gated), orange→green, one-click enable/date/activate, turn-off removes the Shopify selling plan.
- Status/Destination: not locked — On Order/Cancelled disabled while live; destination editable with an owner alert.
- Safety buffer removed (all incoming units sellable).
- Storefront shows PRE-ORDER on an out-of-stock AU variant (variant-id GID/numeric match fixed).
- Owner-only (superadmin) alert when a live-preorder product's status/destination changes.

---

## 🔴 1. Prove the money path (controlled end-to-end test)  (live, with Claude)
- [ ] Add the pre-order XS to cart → checkout → pay in full → order completes.
- [ ] Confirm the order webhook creates a reservation against **the exact batch** (#1256).
- [ ] Confirm capacity drops (Remaining 5 → 4) and the batch card reflects it.
- [ ] Confirm cart/checkout clearly shows it's a pre-order + the dispatch date.
- [ ] Cancel the test order → confirm the reservation **releases** (Remaining back to 5).
- [ ] Try to order more than available (e.g. 6 when 5 left) → confirm it's blocked, never oversells.
- [ ] Order-confirmation email wording for pre-order customers (clearly says pre-order + dispatch date).
- [ ] Capture unit price on the reservation at this point → unlocks revenue/money-in-advance in the Reports tab (#7).

## 🔴 2. Fulfilment when the stock actually arrives  (needs a defined process)
- [ ] Decide + build the flow: batch lands (status → arrived / received) → how the reserved pre-order orders get fulfilled in Shopify.
- [ ] Confirm inventory received at the AU location correctly flips those variants back to normal in-stock and the block hides.

## 🟡 3. Back-in-stock email (Klaviyo)  (live)
- [ ] Build a Klaviyo **Flow** triggered by the `Karma East Back In Stock Available` event (the app fires the event; no Flow = no email sent).
- [ ] Test with a real (non-@test) email address.
- [ ] Decide who owns "Notify me" store-wide — our block vs the old back-in-stock app — so customers aren't notified twice. (Interim: a global ON/OFF toggle for our notify-me block now exists on Pre-orders → Back in Stock; turn OFF while the other app runs. Pre-order is unaffected.)

## 🟡 4. Destination-change action (the "I don't know what yet")
- [ ] Define what should happen when a live-preorder product's destination changes (e.g. pause the preorder? refund/cancel reservations? move market?). Right now it only alerts you.

## 🟡 5. Overallocation / shortfall handling
- [ ] If incoming production drops below what's reserved, define the admin view + process (handoff rule: never silently unreserve a real customer).

## ✅ 6. Customer "My pre-orders" page  (DONE — one no-code link step left)
- [x] Themed page served via the app proxy at `/apps/karma-east-preorder?view=my-preorders` — renders inside the live theme (header/footer/fonts), shows the logged-in customer's reserved items, sizes, dispatch dates and status. Read-only, no theme editing needed.
- [ ] (Optional, no-code) Add a menu/account link to that URL: Shopify admin → Online Store → Navigation, so customers can find it. Or link it from the account page.

## ✅ 7. Reporting  (DONE except revenue)
- [x] Pre-orders report (Pre-orders → Reports): customer orders, reserved / dispatched / released units, AU vs USA split, most-reserved batches **with a fill-rate bar (reserved vs incoming)**, and allocation exceptions.
- [ ] Revenue / money-taken-in-advance — REQUIRES capturing unit price at reservation time (webhook doesn't carry it yet). Bundled with the checkout/order-processing work in the section below, since that's where price gets captured.

## 🟢 8. USA expansion (when 3PL ready)  (live)
- [ ] Set the USA fulfilment location → USA batches become activatable automatically. Then repeat the end-to-end test for USA and verify AU/USA never cross.

## 🟢 9. Broad rollout
- [ ] After the controlled test passes, decide the rollout: which products/batches to enable, and confirm the theme block is on all relevant product templates.

---

## Known limitations / park for later
- **Store credit can't be used on pre-orders** — Shopify blocks store credit with deferred purchase options ("You can't use store credit with deferred purchase options"). Platform rule, not our code; customers just pay by card/PayPal. Leave as-is.
- **Product-page block design** — 4 mockups built (artifact); waiting on Koku to pick A/B/C/D before building the winner into the live theme block.
- **Order-confirmation email wording** for pre-order customers (deferred with the checkout work).

## Open questions / things to add
- (add here)
