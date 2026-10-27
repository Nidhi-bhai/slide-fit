/* Purchase windows for SlideFit. No file hash, no subscription, no server.
   $5 lasts 24 hours of conversions in this browser.
   A $12 pass lasts 30 days. Neither renews on its own.
   Restore only accepts a PayPal order id: 17 letters or digits. */
(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.SlideFitPurchases = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  const DECK_MS = 24 * 60 * 60 * 1000;
  const PASS_MS = 30 * 24 * 60 * 60 * 1000;

  function duration(product) {
    return product === "pass" ? PASS_MS : DECK_MS;
  }

  const ORDER_ID = /^[A-Z0-9]{17}$/;

  function normalizeOrderId(orderId) {
    return String(orderId == null ? "" : orderId).trim().toUpperCase();
  }

  function assertOrderId(orderId) {
    if (!ORDER_ID.test(orderId)) {
      throw new Error("Enter a PayPal order id: 17 letters or digits.");
    }
  }

  function recordPurchase(list, input) {
    const orderId = normalizeOrderId(input && input.orderId);
    if (!orderId) throw new Error("PayPal did not return an order id.");
    const product = input.product === "pass" ? "pass" : "deck";
    const existing = (list || []).find((item) => normalizeOrderId(item && item.orderId) === orderId);
    if (existing) return { purchases: (list || []).slice(), purchase: existing, alreadyStored: true };
    const purchase = {
      orderId: orderId,
      product: product,
      purchasedAt: input.purchasedAt || Date.now(),
    };
    return { purchases: (list || []).concat(purchase), purchase: purchase, alreadyStored: false };
  }

  function restorePurchase(list, input, now) {
    const orderId = normalizeOrderId(input && input.orderId);
    assertOrderId(orderId);
    const at = now || Date.now();
    const existing = (list || []).find((item) => normalizeOrderId(item && item.orderId) === orderId);
    if (existing) {
      const until = existing.purchasedAt + duration(existing.product);
      return {
        purchases: (list || []).slice(),
        purchase: existing,
        active: until > at,
        until: until,
        fromStoredTime: true,
      };
    }
    // Unknown here. Do not invent a 30-day pass. A 24-hour pass starts now.
    const purchase = { orderId: orderId, product: "day", purchasedAt: at };
    return {
      purchases: (list || []).concat(purchase),
      purchase: purchase,
      active: true,
      until: at + DECK_MS,
      fromStoredTime: false,
    };
  }

  function access(list, now) {
    const at = now || Date.now();
    let best = null;
    (list || []).forEach((item) => {
      if (!item || (item.product !== "deck" && item.product !== "pass" && item.product !== "day")) return;
      const until = item.purchasedAt + duration(item.product);
      if (until <= at) return;
      if (!best || until > best.until) {
        best = { orderId: item.orderId, product: item.product, purchasedAt: item.purchasedAt, until: until };
      }
    });
    if (!best) return { pro: false, reason: "free" };
    return { pro: true, reason: best.product, until: best.until, orderId: best.orderId, product: best.product };
  }

  return {
    DECK_MS: DECK_MS,
    PASS_MS: PASS_MS,
    duration: duration,
    recordPurchase: recordPurchase,
    restorePurchase: restorePurchase,
    access: access,
  };
});
