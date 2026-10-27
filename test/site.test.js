const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.join(__dirname, "..");

function read(rel) {
  return fs.readFileSync(path.join(root, rel), "utf8");
}

const pages = [
  "index.html",
  "privacy.html",
  "404.html",
  "how-to/apply-template-to-existing-powerpoint.html",
  "how-to/change-pptx-template.html",
  "how-to/convert-slides-to-company-template.html",
  "how-to/rebrand-pptx-without-uploading.html",
  "how-to/map-slide-layouts.html",
];

test("every page credits Nidhi Karuturi and uses relative assets", () => {
  for (const page of pages) {
    const html = read(page);
    assert.match(html, /Built by <a href="https:\/\/github\.com\/Nidhi-bhai">Nidhi Karuturi<\/a>/);
    assert.equal(html.includes('href="/'), false, page + " should not use domain-root asset paths");
  }
});

test("the landing page targets the template-conversion queries and wires the app", () => {
  const html = read("index.html");
  assert.match(html, /Apply a template to an existing PowerPoint/);
  assert.match(html, /Change a PPTX template/);
  assert.match(html, /convert slides to a company template/i);
  assert.match(html, /js\/slidefit-core\.js/);
  assert.match(html, /js\/vendor\/jszip\.min\.js/);
  assert.match(html, /config\.js/);
  assert.match(html, /id="fit-button"/);
  assert.match(html, /id="paypal-once"/);
  assert.match(html, /id="paypal-month"/);
  const config = read("config.js");
  assert.match(config, /PAYPAL_CLIENT_ID/);
  assert.match(config, /PLAN_ID/);
  assert.doesNotMatch(config, /CLIENT_SECRET/);
});

test("sample decks and the PayPal placeholder config are present", () => {
  assert.ok(fs.existsSync(path.join(root, "samples/northwind-quarterly.pptx")));
  assert.ok(fs.existsSync(path.join(root, "samples/harbor-company-template.pptx")));
  assert.ok(fs.existsSync(path.join(root, ".nojekyll")));
  const sitemap = read("sitemap.xml");
  for (const page of pages.filter((name) => name !== "404.html")) {
    const url = page === "index.html" ? "https://nidhi-bhai.github.io/slide-fit/" : "https://nidhi-bhai.github.io/slide-fit/" + page;
    assert.ok(sitemap.includes(url), url);
  }
});
