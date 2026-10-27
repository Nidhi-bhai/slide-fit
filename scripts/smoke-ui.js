/**
 * Browser check for the static site. Not part of npm test.
 *   python3 -m http.server 8765
 *   node scripts/smoke-ui.js
 */
const fs = require("node:fs");
const path = require("node:path");
const puppeteer = require("puppeteer-core");

const base = process.env.SLIDEFIT_URL || "http://127.0.0.1:8765/";
const outDir = "/tmp/slidefit-ui";
fs.mkdirSync(outDir, { recursive: true });

async function main() {
  const browser = await puppeteer.launch({
    executablePath: process.env.CHROME || "/usr/local/bin/google-chrome",
    headless: "new",
    args: ["--no-sandbox", "--disable-dev-shm-usage"],
  });
  const page = await browser.newPage();
  const errors = [];
  page.on("pageerror", (err) => errors.push(String(err)));
  page.on("console", (msg) => {
    if (msg.type() === "error") errors.push(msg.text());
  });

  await page.setViewport({ width: 1280, height: 900 });
  await page.goto(base, { waitUntil: "load" });
  await page.waitForSelector("#restore-order");
  const h1 = await page.$eval("h1", (node) => node.textContent);
  if (!/Apply a template to an existing PowerPoint/.test(h1)) throw new Error("Unexpected h1: " + h1);
  await page.screenshot({ path: path.join(outDir, "desktop-top.png") });

  await page.click("#load-sample-deck");
  await page.waitForFunction(() => document.querySelector("#drop-deck .drop-title").textContent.includes("northwind"));
  await page.click("#load-sample-template");
  await page.waitForFunction(() => document.querySelector("#drop-template .drop-title").textContent.includes("harbor"));
  await page.click("#fit-button");
  await page.waitForSelector("#preview:not([hidden])");
  await page.waitForFunction(() => document.body.innerText.includes("Northwind Quarterly Review"));
  const freeCount = await page.$$eval("#thumbs .thumb", (nodes) => nodes.length);
  if (freeCount !== 6) throw new Error("Expected 6 free slides, got " + freeCount);
  const freeText = await page.$eval("#stage", (node) => node.innerText);
  if (!/Northwind Quarterly Review/.test(freeText)) throw new Error("Preview missing title");
  await page.screenshot({ path: path.join(outDir, "desktop-preview.png") });

  await page.click('#thumbs .thumb:nth-child(4)');
  await page.waitForFunction(() => document.querySelector("#stage img"));
  await page.click('#thumbs .thumb:nth-child(5)');
  await page.waitForFunction(() => document.querySelector("#stage table"));
  const pitch = await page.evaluate(() => document.body.innerText);
  if (!/does not auto-renew/.test(pitch)) throw new Error("Missing auto-renew copy");
  if (!/first 5 slides/.test(pitch)) throw new Error("Missing first 5 slides copy");
  const appJs = await page.evaluate(async () => (await fetch("js/app.js")).text());
  if (/PLAN_ID|subscription|SHA-256/.test(appJs)) throw new Error("app.js still describes the old billing");

  await page.$eval("#restore-order", (node) => {
    node.value = "SMOKE-ORDER-1";
  });
  await page.click("#restore-form button[type=submit]");
  await page.waitForFunction(() => document.querySelectorAll("#thumbs .thumb").length === 7);
  const proText = await page.$eval("#preview", (node) => node.innerText);
  if (/SlideFit free watermark/.test(proText)) throw new Error("Watermark still present after unlock");
  if (!/Risks to watch/.test(proText) || !/Decisions for Thursday/.test(proText)) {
    throw new Error("Pro preview missing later slides");
  }

  const downloadDir = path.join(outDir, "dl");
  fs.mkdirSync(downloadDir, { recursive: true });
  const client = await page.createCDPSession();
  await client.send("Page.setDownloadBehavior", { behavior: "allow", downloadPath: downloadDir });
  await page.click("#download");
  const started = Date.now();
  let downloaded = "";
  while (Date.now() - started < 8000) {
    const names = fs.readdirSync(downloadDir).filter((name) => name.endsWith(".pptx"));
    if (names.length) {
      downloaded = path.join(downloadDir, names[0]);
      break;
    }
    await new Promise((resolve) => setTimeout(resolve, 200));
  }
  if (!downloaded) throw new Error("Download did not produce a pptx");
  const sig = fs.readFileSync(downloaded).subarray(0, 2).toString();
  if (sig !== "PK") throw new Error("Download is not a zip/pptx");

  await page.setViewport({ width: 390, height: 844 });
  await page.screenshot({ path: path.join(outDir, "mobile-preview.png"), fullPage: false });

  await page.goto(new URL("how-to/apply-template-to-existing-powerpoint.html", base).href, { waitUntil: "load" });
  const guide = await page.$eval("h1", (node) => node.textContent);
  if (!/Apply a template/.test(guide)) throw new Error("Guide h1 missing");
  const credit = await page.$eval("footer", (node) => node.innerText);
  if (!/Nidhi Karuturi/.test(credit)) throw new Error("Footer missing");

  if (errors.length) throw new Error("Page errors:\n" + errors.join("\n"));
  console.log("smoke ok", downloaded, "bytes", fs.statSync(downloaded).size);
  await browser.close();
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
