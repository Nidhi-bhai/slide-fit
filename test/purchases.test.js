const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Purchases = require("../js/purchases.js");

const HOUR = 60 * 60 * 1000;

test("$5 deck lasts 24 hours and is not bound to file bytes", () => {
  const start = Date.parse("2026-10-27T12:00:00Z");
  const saved = Purchases.recordPurchase([], { orderId: "ORDER-DECK", product: "deck", purchasedAt: start });
  assert.equal(saved.purchase.orderId, "ORDER-DECK");
  assert.equal(Object.prototype.hasOwnProperty.call(saved.purchase, "hash"), false);
  const during = Purchases.access(saved.purchases, start + 23 * HOUR);
  assert.equal(during.pro, true);
  assert.equal(during.reason, "deck");
  const after = Purchases.access(saved.purchases, start + Purchases.DECK_MS);
  assert.equal(after.pro, false);
});

test("$12 pass is 30 days, one window, and does not extend itself", () => {
  const start = Date.parse("2026-10-01T00:00:00Z");
  const saved = Purchases.recordPurchase([], { orderId: "ORDER-PASS", product: "pass", purchasedAt: start });
  assert.equal(Purchases.access(saved.purchases, start + 29 * 24 * HOUR).pro, true);
  assert.equal(Purchases.access(saved.purchases, start + Purchases.PASS_MS).reason, "free");
  const again = Purchases.recordPurchase(saved.purchases, { orderId: "ORDER-PASS", product: "pass", purchasedAt: start + Purchases.PASS_MS });
  assert.equal(again.alreadyStored, true);
  assert.equal(again.purchase.purchasedAt, start);
});

test("restore uses the saved purchase time, and otherwise starts from now", () => {
  const start = Date.parse("2026-10-27T08:00:00Z");
  const saved = Purchases.recordPurchase([], { orderId: " ORDER-1 ", product: "deck", purchasedAt: start });
  const later = start + 20 * HOUR;
  const restored = Purchases.restorePurchase(saved.purchases, { orderId: "ORDER-1", product: "pass" }, later);
  assert.equal(restored.fromStoredTime, true);
  assert.equal(restored.purchase.product, "deck");
  assert.equal(restored.purchase.purchasedAt, start);
  assert.equal(restored.active, true);
  assert.equal(restored.until, start + Purchases.DECK_MS);

  const expired = Purchases.restorePurchase(saved.purchases, { orderId: "ORDER-1" }, start + Purchases.DECK_MS + 1);
  assert.equal(expired.fromStoredTime, true);
  assert.equal(expired.active, false);
  assert.equal(expired.purchases.length, 1);

  const unknown = Purchases.restorePurchase(saved.purchases, { orderId: "NEW-ORDER", product: "deck" }, later);
  assert.equal(unknown.fromStoredTime, false);
  assert.equal(unknown.purchase.purchasedAt, later);
  assert.equal(unknown.purchase.product, "pass");
  assert.equal(unknown.active, true);
  assert.equal(unknown.until, later + Purchases.PASS_MS);
});

test("the page no longer sells a subscription or hashes the file", () => {
  const root = path.join(__dirname, "..");
  const app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
  const config = fs.readFileSync(path.join(root, "config.js"), "utf8");
  const index = fs.readFileSync(path.join(root, "index.html"), "utf8");
  const privacy = fs.readFileSync(path.join(root, "privacy.html"), "utf8");
  assert.doesNotMatch(app, /PLAN_ID|subscription|SHA-256|unlockedDecks/);
  assert.doesNotMatch(config, /PLAN_ID/);
  assert.match(config, /PAYPAL_CLIENT_ID/);
  assert.match(index, /first 5 slides/);
  assert.doesNotMatch(index, /five content slides/i);
  assert.match(index, /does not auto-renew/);
  assert.match(index, /24 hours/);
  assert.match(index, /not tied to the file/);
  assert.match(index, /SmartArt/);
  assert.match(index, /bold, colour, and links/i);
  assert.match(index, /id="restore-order"/);
  assert.match(privacy, /stores only the PayPal order id/i);
  assert.doesNotMatch(privacy, /hash of the file/i);
});
