/**
 * Fail closed on ambiguous Shopify product identity. No writes or API calls.
 * @template {{id: string, title: string}} T
 * @param {T[]} results
 * @param {string} query
 * @param {string | undefined} explicitId
 * @param {string | undefined} explicitTitle
 * @param {string | undefined} historicalId
 * @returns {T}
 */
export function selectProductForLookup(results, query, explicitId, explicitTitle, historicalId) {
  const normalized = (value) => String(value ?? "").trim().toLowerCase();
  if (!results.length) throw new Error(`Shopify search for "${query}" returned nothing. Retry or search again; no product was selected.`);
  if (explicitId) {
    const exact = results.filter(p => p.id === explicitId);
    if (exact.length !== 1) throw new Error(`Product ID ${explicitId} was not uniquely found. No title fallback is allowed.`);
    if (explicitTitle && normalized(exact[0].title) !== normalized(explicitTitle)) throw new Error(`Product ID ${explicitId} does not match the supplied title. Verify the product before continuing.`);
    return exact[0];
  }
  const byTitle = results.filter(p => normalized(p.title) === normalized(query));
  if (byTitle.length > 1) throw new Error(`Several products are called "${query}". Use product_id: ${byTitle.map(p => p.id).join(", ")}`);
  if (byTitle.length !== 1) throw new Error(`No exact product "${query}". Closest: ${results.slice(0, 8).map(p => `${p.title} (${p.id})`).join("; ")}`);
  if (historicalId && byTitle[0].id !== historicalId) throw new Error(`The current product ID differs from the historical order for "${query}". Verify the correct product_id.`);
  return byTitle[0];
}
