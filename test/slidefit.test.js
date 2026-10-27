const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");

globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;

const SlideFit = require("../js/slidefit-core.js");
const JSZip = require("jszip");

const root = path.join(__dirname, "..");
const northwind = fs.readFileSync(path.join(root, "samples/northwind-quarterly.pptx"));
const harbor = fs.readFileSync(path.join(root, "samples/harbor-company-template.pptx"));

function layouts() {
  return [
    { name: "Title Slide", placeholders: [{ type: "ctrTitle" }, { type: "subTitle", idx: "1" }] },
    { name: "Title and Content", placeholders: [{ type: "title" }, { type: "obj", idx: "1" }] },
    { name: "Section Header", placeholders: [{ type: "title" }, { type: "body", idx: "1" }] },
    { name: "Two Content", placeholders: [{ type: "title" }, { type: "obj", idx: "1" }, { type: "obj", idx: "2" }] },
    { name: "Title Only", placeholders: [{ type: "title" }] },
    { name: "Blank", placeholders: [] },
    { name: "Picture with Caption", placeholders: [{ type: "title" }, { type: "pic", idx: "1" }, { type: "body", idx: "2" }] },
    { name: "Title and Vertical Text", placeholders: [{ type: "title", orient: "vert" }, { type: "body", idx: "1" }] },
  ];
}

test("classifies title, section, bullets, image, and table slides", () => {
  assert.equal(SlideFit.classifySlide({ title: "Northwind Quarterly Review", subtitle: "Q3 field briefing", bullets: [], images: [], tables: [] }, 0), "title");
  assert.equal(SlideFit.classifySlide({ title: "What changed this quarter", subtitle: "", bullets: [], images: [], tables: [] }, 1), "section");
  assert.equal(
    SlideFit.classifySlide({ title: "Pipeline health", subtitle: "", bullets: [{ text: "One", level: 0 }], images: [], tables: [] }, 2),
    "title-bullets",
  );
  assert.equal(
    SlideFit.classifySlide({ title: "West region coverage", subtitle: "", bullets: [{ text: "Seattle", level: 0 }], images: [{}], tables: [] }, 3),
    "image-text",
  );
  assert.equal(
    SlideFit.classifySlide({ title: "Revenue by line", subtitle: "", bullets: [], images: [], tables: [{ rows: [["Line"]] }] }, 4),
    "table",
  );
});

test("maps each content kind to the closest template layout", () => {
  const all = layouts();
  assert.equal(SlideFit.matchLayout("title", all).name, "Title Slide");
  assert.equal(SlideFit.matchLayout("section", all).name, "Section Header");
  assert.equal(SlideFit.matchLayout("title-bullets", all).name, "Title and Content");
  assert.equal(SlideFit.matchLayout("image-text", all).name, "Picture with Caption");
  assert.equal(SlideFit.matchLayout("table", all).name, "Title and Content");
  assert.equal(SlideFit.matchLayout("watermark", all).name, "Title Slide");
});

test("free plan keeps five slides and adds a watermark; pro keeps the deck", () => {
  const slides = Array.from({ length: 7 }, (_, i) => ({ title: "S" + i, bullets: [], images: [], tables: [] }));
  const free = SlideFit.planSlides(slides, { pro: false });
  assert.equal(free.slides.length, 5);
  assert.equal(free.omitted, 2);
  assert.equal(free.watermark, true);
  const pro = SlideFit.planSlides(slides, { pro: true });
  assert.equal(pro.slides.length, 7);
  assert.equal(pro.omitted, 0);
  assert.equal(pro.watermark, false);
  const short = SlideFit.planSlides(slides.slice(0, 2), { pro: false });
  assert.equal(short.slides.length, 2);
  assert.equal(short.watermark, true);
});

test("reads Northwind content, including the picture and the table", async () => {
  const deck = await SlideFit.parseDeck(northwind);
  assert.equal(deck.slides.length, 7);
  assert.equal(deck.slides[0].title, "Northwind Quarterly Review");
  assert.match(deck.slides[0].subtitle, /Q3 field briefing/);
  assert.equal(deck.slides[1].title, "What changed this quarter");
  assert.equal(deck.slides[2].bullets.length, 4);
  assert.match(deck.slides[2].bullets[2].text, /net > gross & fees/);
  assert.equal(deck.slides[3].images.length, 1);
  assert.equal(deck.slides[3].images[0].bytes[0], 0x89);
  assert.equal(deck.slides[4].tables.length, 1);
  assert.equal(deck.slides[4].tables[0].rows[1][0], "Specialty tea");
  assert.equal(deck.slides[5].title, "Risks to watch");
});

async function slideXmls(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const names = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => Number(a.match(/slide(\d+)/)[1]) - Number(b.match(/slide(\d+)/)[1]));
  const slides = [];
  for (const name of names) {
    const xml = await zip.file(name).async("string");
    const relName = name.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
    const rels = await zip.file(relName).async("string");
    const layoutTarget = (rels.match(/Target="([^"]*slideLayout\d+\.xml)"/) || [])[1];
    const layoutPath = "ppt/slideLayouts/" + layoutTarget.split("/").pop();
    const layoutXml = await zip.file(layoutPath).async("string");
    const layoutName = (layoutXml.match(/name="([^"]+)"/) || [])[1];
    slides.push({ name, xml, rels, layoutName });
  }
  return { zip, slides };
}

test("applies the Harbor template to Northwind on the free plan", async () => {
  const result = await SlideFit.applyTemplate(northwind, harbor, { pro: false });
  assert.equal(result.omitted, 2);
  assert.equal(result.watermark, true);
  assert.equal(result.slides.length, 6);
  assert.equal(result.slides[0].layoutName, "Title Slide");
  assert.equal(result.slides[1].layoutName, "Section Header");
  assert.equal(result.slides[2].layoutName, "Title and Content");
  assert.equal(result.slides[3].layoutName, "Picture with Caption");
  assert.equal(result.slides[4].layoutName, "Title and Content");
  assert.equal(result.slides[5].kind, "watermark");

  const { zip, slides } = await slideXmls(result.bytes);
  assert.equal(slides.length, 6);
  assert.match(slides[0].xml, /Northwind Quarterly Review/);
  assert.match(slides[0].xml, /Q3 field briefing/);
  assert.equal(slides[0].layoutName, "Title Slide");
  assert.match(slides[2].xml, /net &gt; gross &amp; fees/);
  assert.match(slides[2].xml, /\+mn-lt/);
  assert.match(slides[2].xml, /schemeClr val="tx1"/);
  assert.equal(slides[3].layoutName, "Picture with Caption");
  assert.match(slides[3].rels, /image/);
  assert.match(slides[4].xml, /Specialty tea/);
  assert.match(slides[5].xml, /SlideFit free watermark/);
  assert.doesNotMatch(slides.map((s) => s.xml).join("\n"), /Risks to watch/);

  const theme = await zip.file("ppt/theme/theme1.xml").async("string");
  assert.match(theme, /typeface="Georgia"/);
  assert.match(theme, /typeface="Verdana"/);
  assert.match(theme, /0E7C7B/);
  const master = await zip.file("ppt/slideMasters/slideMaster1.xml").async("string");
  assert.match(master, /HARBOR &amp; CO\./);
  const media = Object.keys(zip.files).filter((name) => name.startsWith("ppt/media/slidefit-"));
  assert.ok(media.length >= 1);
});

test("pro conversion keeps every slide and leaves the watermark out", async () => {
  const result = await SlideFit.applyTemplate(northwind, harbor, { pro: true });
  assert.equal(result.omitted, 0);
  assert.equal(result.watermark, false);
  const { slides } = await slideXmls(result.bytes);
  assert.equal(slides.length, 7);
  const all = slides.map((s) => s.xml).join("\n");
  assert.match(all, /Risks to watch/);
  assert.match(all, /Decisions for Thursday/);
  assert.match(all, /Denver route/);
  assert.doesNotMatch(all, /SlideFit free watermark/);
});

test("rejects a file that is not a PowerPoint package", async () => {
  await assert.rejects(() => SlideFit.applyTemplate(Buffer.from("hello"), harbor, { pro: true }), /not a \.pptx/);
});
