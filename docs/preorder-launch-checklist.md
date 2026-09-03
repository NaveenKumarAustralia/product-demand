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

## 🔴 2. Fulfilment when the stock actually arrives  (needs a defined process)
- [ ] Decide + build the flow: batch lands (status → arrived / received) → how the reserved pre-order orders get fulfilled in Shopify.
- [ ] Confirm inventory received at the AU location correctly flips those variants back to normal in-stock and the block hides.

## 🟡 3. Back-in-stock email (Klaviyo)  (live)
- [ ] Build a Klaviyo **Flow** triggered by the `Karma East Back In Stock Available` event (the app fires the event; no Flow = no email sent).
- [ ] Test with a real (non-@test) email address.
- [ ] Decide who owns "Notify me" store-wide — our block vs the old back-in-stock app — so customers aren't notified twice. (Currently excluded per-product.)

## 🟡 4. Destination-change action (the "I don't know what yet")
- [ ] Define what should happen when a live-preorder product's destination changes (e.g. pause the preorder? refund/cancel reservations? move market?). Right now it only alerts you.

## 🟡 5. Overallocation / shortfall handling
- [ ] If incoming production drops below what's reserved, define the admin view + process (handoff rule: never silently unreserve a real customer).

## 🟡 6. Customer "My pre-orders" page  (code + live)
- [ ] Theme render of the customer-account endpoint (server side exists) so a logged-in customer sees their pre-orders + dispatch dates.

## 🟡 7. Reporting  (code)
- [ ] A pre-orders report: units reserved, revenue taken in advance, by batch/market. (viewReports permission already exists.)

## 🟢 8. USA expansion (when 3PL ready)  (live)
- [ ] Set the USA fulfilment location → USA batches become activatable automatically. Then repeat the end-to-end test for USA and verify AU/USA never cross.

## 🟢 9. Broad rollout
- [ ] After the controlled test passes, decide the rollout: which products/batches to enable, and confirm the theme block is on all relevant product templates.

---

## Open questions / things to add
- (add here)
