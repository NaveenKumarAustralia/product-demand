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

  function formatExpected(value) {
    if (!value) return 'Expected dispatch date to be confirmed';
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return 'Expected dispatch date to be confirmed';
    return `Expected dispatch ${new Intl.DateTimeFormat(undefined, { day: 'numeric', month: 'short', year: 'numeric' }).format(date)}`;
  }

  function setHidden(element, hidden) {
    if (element) element.hidden = hidden;
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
      eyebrow.textContent = 'Pre-order';
      title.textContent = formatExpected(state.expectedShipDate);
      copy.textContent = 'This item is currently being made. Order now to reserve yours. Payment is taken in full at checkout.';
      root.dataset.sellingPlanId = state.sellingPlanId || '';
      root.dataset.batchId = String(state.batchId || '');
      setHidden(preorderButton, false);
      setHidden(notifyForm, true);
      return;
    }
    if (state.state === 'notify_me') {
      root.hidden = false;
      eyebrow.textContent = 'Currently unavailable';
      title.textContent = 'Notify me when this size is available';
      copy.textContent = 'Join the waitlist and we’ll let you know when this size can be ordered again.';
      delete root.dataset.sellingPlanId;
      delete root.dataset.batchId;
      setHidden(preorderButton, true);
      setHidden(notifyForm, false);
      return;
    }
    root.hidden = true;
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
