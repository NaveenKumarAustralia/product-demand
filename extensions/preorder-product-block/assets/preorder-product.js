(() => {
  const roots = Array.from(document.querySelectorAll('[data-ke-preorder-root]'));
  if (!roots.length) return;

  const shopifyRoot = () => (window.Shopify && window.Shopify.routes && window.Shopify.routes.root) || '/';
  const numericId = (value) => String(value || '').split('/').pop();

  function currentVariantId(root) {
    const fromUrl = new URL(window.location.href).searchParams.get('variant');
    const forms = Array.from(document.querySelectorAll('form[action*="/cart/add"]'));
    const formValue = forms.map((form) => form.querySelector('[name="id"]')?.value).find(Boolean);
    const fallback = root.querySelector('[data-ke-variants]');
    let first = null;
    try { first = JSON.parse(fallback?.textContent || '[]')?.[0]?.id; } catch (_) {}
    return String(fromUrl || formValue || first || '').trim();
  }

  function variantMeta(root, variantId) {
    try {
      const rows = JSON.parse(root.querySelector('[data-ke-variants]')?.textContent || '[]');
      return rows.find((row) => String(row.id) === String(variantId)) || {};
    } catch (_) {
      return {};
    }
  }

  // Date in the visitor's own country format (US -> 09/30/2026, AU -> 30/09/2026, etc.).
  function formatLocalDate(value) {
    if (!value) return 'soon';
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return 'soon';
    const locale = (typeof navigator !== 'undefined' && navigator.language) ? navigator.language : undefined;
    try {
      return new Intl.DateTimeFormat(locale, { year: 'numeric', month: '2-digit', day: '2-digit' }).format(date);
    } catch (_) {
      return new Intl.DateTimeFormat(undefined, { year: 'numeric', month: '2-digit', day: '2-digit' }).format(date);
    }
  }

  // Button label: "Pre-order · ships in 24 days" (counts down as the date nears).
  function shipInLabel(value) {
    if (!value) return 'Pre-order';
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return 'Pre-order';
    const days = Math.ceil((date.getTime() - Date.now()) / 86400000);
    if (days <= 0) return 'Pre-order · ships soon';
    return `Pre-order · ships in ${days} day${days === 1 ? '' : 's'}`;
  }

  function setHidden(element, hidden) {
    if (element) element.hidden = hidden;
  }

  // The theme's own "Sold out" / Add-to-cart + dynamic-checkout buttons. We hide
  // these while a pre-order is showing so the customer can only pre-order (and
  // can't bypass the selling plan via dynamic checkout); restored otherwise. We
  // remember each element's previous inline display so we can put it back.
  var THEME_BUY_SELECTORS = 'form[action*="/cart/add"] [type="submit"], form[action*="/cart/add"] [name="add"], .shopify-payment-button, .product-form__submit';
  function setThemeBuyHidden(hidden) {
    var nodes = Array.from(document.querySelectorAll(THEME_BUY_SELECTORS));
    for (var i = 0; i < nodes.length; i++) {
      var el = nodes[i];
      if (el.closest('[data-ke-preorder-root]')) continue; // never touch our own button
      if (hidden) {
        if (el.dataset.kePrevDisplay === undefined) el.dataset.kePrevDisplay = el.style.display || '';
        el.style.display = 'none';
      } else if (el.dataset.kePrevDisplay !== undefined) {
        el.style.display = el.dataset.kePrevDisplay;
        delete el.dataset.kePrevDisplay;
      }
    }
  }

  async function loadState(root) {
    const variantId = currentVariantId(root);
    if (!variantId) return;
    if (root.dataset.loadingVariant === variantId) return;
    root.dataset.loadingVariant = variantId;

    const proxy = root.dataset.proxyPath || '/apps/karma-east-preorder';
    const market = root.dataset.market || 'AU';
    const url = new URL(proxy, window.location.origin);
    url.searchParams.set('variantId', variantId);
    url.searchParams.set('market', market);

    try {
      const response = await fetch(url.toString(), { credentials: 'same-origin', headers: { Accept: 'application/json' } });
      const result = await response.json();
      if (!response.ok || !result?.state) throw new Error(result?.error || 'Availability could not be loaded.');
      if (currentVariantId(root) !== variantId) return;
      render(root, variantId, result.state);
    } catch (error) {
      console.warn('[Karma East preorder] state load failed', error);
      root.hidden = true;
    } finally {
      if (root.dataset.loadingVariant === variantId) delete root.dataset.loadingVariant;
    }
  }

  function render(root, variantId, state) {
    const eyebrow = root.querySelector('[data-ke-preorder-eyebrow]');
    const title = root.querySelector('[data-ke-preorder-title]');
    const copy = root.querySelector('[data-ke-preorder-copy]');
    const preorderButton = root.querySelector('[data-ke-preorder-button]');
    const notifyForm = root.querySelector('[data-ke-notify-form]');
    const notifyMessage = root.querySelector('[data-ke-notify-message]');

    root.dataset.variantId = variantId;
    root.dataset.state = state.state;
    if (notifyMessage) {
      notifyMessage.textContent = '';
      delete notifyMessage.dataset.kind;
    }

    // Only ever take over the buy area for a real pre-order or a real
    // notify-me. For in_stock, inactive (market not live, e.g. USA before its
    // location is set), or any unknown state, our block does nothing so the
    // theme's normal add-to-cart / out-of-stock UI shows instead. This is what
    // stops in-stock variants (in markets we don't ship) showing the notify box.
    if (state.state === 'preorder') {
      root.hidden = false;
      // No badge — just the out-of-stock copy with the back-in-stock date, and
      // the pre-order button. Hide the theme's Sold out / buy buttons.
      setHidden(eyebrow, true);
      setHidden(title, true);
      copy.textContent = `Currently out of stock. Arriving back in stock ${formatLocalDate(state.expectedShipDate)}.`;
      root.dataset.sellingPlanId = state.sellingPlanId || '';
      root.dataset.batchId = String(state.batchId || '');
      preorderButton.textContent = shipInLabel(state.expectedShipDate);
      setHidden(preorderButton, false);
      setHidden(notifyForm, true);
      setThemeBuyHidden(true);
      return;
    }
    if (state.state === 'notify_me') {
      root.hidden = false;
      setHidden(eyebrow, false);
      setHidden(title, false);
      eyebrow.textContent = 'Currently unavailable';
      title.textContent = 'Notify me when this size is available';
      copy.textContent = 'Join the waitlist and we’ll let you know when this size can be ordered again.';
      delete root.dataset.sellingPlanId;
      delete root.dataset.batchId;
      setHidden(preorderButton, true);
      setHidden(notifyForm, false);
      setThemeBuyHidden(false);
      return;
    }
    root.hidden = true;
    setThemeBuyHidden(false);
  }

  async function addPreorder(root, button) {
    const variantId = root.dataset.variantId;
    const sellingPlanId = numericId(root.dataset.sellingPlanId);
    if (!variantId || !sellingPlanId) return;
    button.disabled = true;
    const original = button.textContent;
    button.textContent = 'Adding…';
    try {
      const form = new FormData();
      form.append('id', variantId);
      form.append('quantity', '1');
      form.append('selling_plan', sellingPlanId);
      const response = await fetch(shopifyRoot() + 'cart/add.js', {
        method: 'POST',
        body: form,
        headers: { Accept: 'application/json' },
      });
      const result = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(result?.description || result?.message || 'Could not add preorder to cart.');
      window.location.assign(shopifyRoot() + 'cart');
    } catch (error) {
      window.alert(error instanceof Error ? error.message : 'Could not add preorder to cart.');
      button.disabled = false;
      button.textContent = original;
    }
  }

  async function joinWaitlist(root, form) {
    const emailInput = form.querySelector('input[name="email"]');
    const message = root.querySelector('[data-ke-notify-message]');
    const submit = form.querySelector('button[type="submit"]');
    const variantId = root.dataset.variantId;
    const email = emailInput?.value?.trim();
    if (!variantId || !email) return;

    const variant = variantMeta(root, variantId);
    submit.disabled = true;
    message.textContent = 'Saving…';
    delete message.dataset.kind;

    try {
      const proxy = root.dataset.proxyPath || '/apps/karma-east-preorder';
      const response = await fetch(proxy, {
        method: 'POST',
        credentials: 'same-origin',
        headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
        body: JSON.stringify({
          email,
          market: root.dataset.market || 'AU',
          productId: root.dataset.productId || '',
          productTitle: root.dataset.productTitle || '',
          variantId,
          variantTitle: variant.title || '',
          sku: variant.sku || '',
        }),
      });
      const result = await response.json().catch(() => ({}));
      if (!response.ok || result.ok !== true) throw new Error(result.error || 'Could not join the waitlist.');
      message.textContent = 'You’re on the list. We’ll email you when it is available.';
      message.dataset.kind = 'success';
      emailInput.value = '';
    } catch (error) {
      message.textContent = error instanceof Error ? error.message : 'Could not join the waitlist.';
      message.dataset.kind = 'error';
    } finally {
      submit.disabled = false;
    }
  }

  for (const root of roots) {
    root.querySelector('[data-ke-preorder-button]')?.addEventListener('click', (event) => addPreorder(root, event.currentTarget));
    root.querySelector('[data-ke-notify-form]')?.addEventListener('submit', (event) => {
      event.preventDefault();
      joinWaitlist(root, event.currentTarget);
    });
    loadState(root);
  }

  let timer;
  const refresh = () => {
    clearTimeout(timer);
    timer = setTimeout(() => roots.forEach(loadState), 80);
  };
  document.addEventListener('change', (event) => {
    if (event.target?.matches?.('[name="id"], input[type="radio"], select')) refresh();
  });
  document.addEventListener('click', (event) => {
    if (event.target?.closest?.('[data-option-value], .option-selector, .variant-input, .variant-picker')) refresh();
  });
  window.addEventListener('popstate', refresh);
})();
