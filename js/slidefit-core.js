/* SlideFit core: read a .pptx, map each slide onto a template layout, write a new .pptx.
   Runs in the browser (JSZip + DOMParser) and in Node tests. No network. */
(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.SlideFit = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const SKIP_PH = new Set(["dt", "ftr", "sldNum"]);

  class SlideFitError extends Error {
    constructor(message) {
      super(message);
      this.name = "SlideFitError";
    }
  }

  function zipLib() {
    if (typeof JSZip !== "undefined") return JSZip;
    return require("jszip");
  }

  function parseXml(text) {
    const Parser = globalThis.DOMParser || (typeof DOMParser !== "undefined" ? DOMParser : null);
    if (!Parser) throw new SlideFitError("This browser cannot read Office XML.");
    const doc = new Parser().parseFromString(text, "text/xml");
    if (!doc || !doc.documentElement) throw new SlideFitError("Could not read part of the PowerPoint package.");
    if (doc.documentElement.localName === "parsererror") {
      throw new SlideFitError("Could not read part of the PowerPoint package.");
    }
    return doc;
  }

  function kids(el) {
    const out = [];
    if (!el || !el.childNodes) return out;
    for (let i = 0; i < el.childNodes.length; i++) {
      const node = el.childNodes[i];
      if (node.nodeType === 1) out.push(node);
    }
    return out;
  }

  function child(el, local) {
    return kids(el).find((node) => node.localName === local) || null;
  }

  function allLocal(el, local) {
    const out = [];
    if (!el || !el.getElementsByTagName) return out;
    const list = el.getElementsByTagName("*");
    for (let i = 0; i < list.length; i++) {
      if (list[i].localName === local) out.push(list[i]);
    }
    return out;
  }

  function attr(el, name) {
    if (!el || !el.getAttribute) return null;
    const direct = el.getAttribute(name);
    if (direct != null && direct !== "") return direct;
    if (name.startsWith("r:") && el.getAttributeNS) {
      const value = el.getAttributeNS(NS_REL, name.slice(2));
      if (value) return value;
    }
    return null;
  }

  function decode(data) {
    if (typeof data === "string") return data;
    return new TextDecoder("utf-8").decode(data);
  }

  function toUint8(data) {
    if (!data) return null;
    if (data instanceof Uint8Array) return data;
    if (typeof data === "string") return new TextEncoder().encode(data);
    return new Uint8Array(data);
  }

  function xmlEscape(value) {
    return String(value == null ? "" : value)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function safeDecode(value) {
    try {
      return decodeURIComponent(value);
    } catch (err) {
      return value;
    }
  }

  function resolveTarget(fromPart, target) {
    if (!target) return null;
    if (/^[a-z][a-z0-9+.-]*:/i.test(target)) return null;
    const clean = String(target).split("#")[0].split("?")[0];
    if (clean.startsWith("/")) return safeDecode(clean.replace(/^\/+/, ""));
    const base = fromPart ? fromPart.split("/").slice(0, -1) : [];
    for (const part of clean.split("/")) {
      if (part === "..") base.pop();
      else if (part && part !== ".") base.push(safeDecode(part));
    }
    return base.join("/");
  }

  function relsPath(part) {
    const slash = part.lastIndexOf("/");
    const dir = slash >= 0 ? part.slice(0, slash) : "";
    const name = slash >= 0 ? part.slice(slash + 1) : part;
    return (dir ? dir + "/" : "") + "_rels/" + name + ".rels";
  }

  function ownerOfRels(relsFile) {
    if (relsFile === "_rels/.rels") return "";
    return relsFile.replace(/\/_rels\/([^/]+)\.rels$/, "/$1").replace(/\.rels$/, "");
  }

  function relativeTarget(fromPart, toPart) {
    const fromDir = fromPart.split("/").slice(0, -1);
    const to = toPart.split("/");
    let shared = 0;
    while (shared < fromDir.length && shared < to.length && fromDir[shared] === to[shared]) shared++;
    return "../".repeat(fromDir.length - shared) + to.slice(shared).join("/");
  }

  function parseRels(xml) {
    const doc = parseXml(xml);
    return allLocal(doc, "Relationship")
      .map((el) => ({
        id: el.getAttribute("Id"),
        type: el.getAttribute("Type") || "",
        target: el.getAttribute("Target"),
        targetMode: el.getAttribute("TargetMode") || null,
      }))
      .filter((rel) => rel.id && rel.target);
  }

  function serializeRels(rels) {
    const body = rels
      .map((rel) => {
        const mode = rel.targetMode ? ` TargetMode="${xmlEscape(rel.targetMode)}"` : "";
        return `<Relationship Id="${xmlEscape(rel.id)}" Type="${xmlEscape(rel.type)}" Target="${xmlEscape(rel.target)}"${mode}/>`;
      })
      .join("");
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      body +
      "</Relationships>"
    );
  }

  function loadRels(files, part) {
    const path = relsPath(part);
    if (!files[path]) return [];
    return parseRels(decode(files[path]));
  }

  async function unzip(bytes) {
    let zip;
    try {
      zip = await zipLib().loadAsync(bytes);
    } catch (err) {
      throw new SlideFitError("That file is not a .pptx PowerPoint deck.");
    }
    const files = {};
    const names = Object.keys(zip.files);
    for (const name of names) {
      const entry = zip.files[name];
      if (entry.dir) continue;
      files[name] = await entry.async("uint8array");
    }
    if (!files["ppt/presentation.xml"]) {
      throw new SlideFitError("That file is not a .pptx PowerPoint deck.");
    }
    return files;
  }

  async function zipFiles(files) {
    const zip = new (zipLib())();
    for (const name of Object.keys(files)) zip.file(name, files[name]);
    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
  }

  function shapeBox(el) {
    const spPr = child(el, "spPr");
    const xfrm = (spPr && child(spPr, "xfrm")) || child(el, "xfrm");
    if (!xfrm) return null;
    const off = child(xfrm, "off");
    const ext = child(xfrm, "ext");
    if (!off || !ext) return null;
    return {
      x: Number(off.getAttribute("x") || 0),
      y: Number(off.getAttribute("y") || 0),
      cx: Number(ext.getAttribute("cx") || 0),
      cy: Number(ext.getAttribute("cy") || 0),
    };
  }

  function placeholderInfo(el) {
    const nv = child(el, "nvSpPr") || child(el, "nvPicPr") || child(el, "nvGraphicFramePr");
    const nvPr = nv && child(nv, "nvPr");
    const ph = nvPr && child(nvPr, "ph");
    if (!ph) return null;
    const type = ph.getAttribute("type");
    return {
      type: type || "obj",
      rawType: type,
      idx: ph.getAttribute("idx"),
      orient: ph.getAttribute("orient"),
      sz: ph.getAttribute("sz"),
    };
  }

  function paragraphsOf(txBody) {
    const paras = [];
    for (const node of kids(txBody)) {
      if (node.localName !== "p") continue;
      const pPr = child(node, "pPr");
      const level = pPr && pPr.getAttribute("lvl") ? Number(pPr.getAttribute("lvl")) : 0;
      let text = "";
      let fontSize = 0;
      for (const part of kids(node)) {
        if (part.localName === "r" || part.localName === "fld") {
          for (const t of allLocal(part, "t")) text += t.textContent || "";
          const rPr = child(part, "rPr");
          const sz = rPr && rPr.getAttribute("sz");
          if (sz) fontSize = Math.max(fontSize, Number(sz));
        } else if (part.localName === "br") text += "\n";
      }
      const cleaned = text.replace(/\u00a0/g, " ").trim();
      if (!cleaned) continue;
      paras.push({ text: cleaned, level: Number.isFinite(level) ? level : 0, fontSize });
    }
    return paras;
  }

  function embedOf(el) {
    const blipFill = child(el, "blipFill");
    const blip = blipFill && child(blipFill, "blip");
    if (!blip) return null;
    return attr(blip, "r:embed");
  }

  function cellText(tc) {
    const tx = child(tc, "txBody");
    if (!tx) return "";
    return paragraphsOf(tx)
      .map((p) => p.text)
      .join(" ")
      .trim();
  }

  function walkShapes(container, acc, origin) {
    for (const el of kids(container)) {
      if (el.localName === "sp" || el.localName === "pic" || el.localName === "graphicFrame") {
        acc.push({ el, origin: origin || { x: 0, y: 0 } });
      } else if (el.localName === "grpSp") {
        const box = shapeBox(el);
        const next = {
          x: (origin ? origin.x : 0) + (box ? box.x : 0),
          y: (origin ? origin.y : 0) + (box ? box.y : 0),
        };
        walkShapes(el, acc, next);
      } else if (el.localName === "AlternateContent") {
        const choice = kids(el).find((node) => node.localName === "Choice");
        if (choice) walkShapes(choice, acc, origin);
      }
    }
  }

  function contentTypeFor(ext) {
    const map = {
      png: "image/png",
      jpg: "image/jpeg",
      jpeg: "image/jpeg",
      gif: "image/gif",
      emf: "image/x-emf",
      wmf: "image/x-wmf",
      tif: "image/tiff",
      tiff: "image/tiff",
      bmp: "image/bmp",
    };
    return map[ext] || "image/png";
  }

  function imageSize(bytes) {
    if (!bytes || bytes.length < 24) return { width: 0, height: 0 };
    if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) {
      const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      return { width: view.getUint32(16), height: view.getUint32(20) };
    }
    if (bytes[0] === 0xff && bytes[1] === 0xd8) {
      let offset = 2;
      while (offset + 8 < bytes.length) {
        if (bytes[offset] !== 0xff) break;
        const marker = bytes[offset + 1];
        const size = (bytes[offset + 2] << 8) + bytes[offset + 3];
        if (marker >= 0xc0 && marker <= 0xc3) {
          const height = (bytes[offset + 5] << 8) + bytes[offset + 6];
          const width = (bytes[offset + 7] << 8) + bytes[offset + 8];
          return { width, height };
        }
        if (!size) break;
        offset += 2 + size;
      }
    }
    if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) {
      return { width: bytes[6] + (bytes[7] << 8), height: bytes[8] + (bytes[9] << 8) };
    }
    return { width: 0, height: 0 };
  }

  function imageFromRel(embed, rels, slidePart, files) {
    const rel = rels.find((item) => item.id === embed);
    if (!rel) return null;
    const part = resolveTarget(slidePart, rel.target);
    const bytes = toUint8(files[part]);
    if (!bytes) return null;
    const ext = (part.split(".").pop() || "png").toLowerCase();
    return { bytes, ext, contentType: contentTypeFor(ext), ...imageSize(bytes) };
  }

  function readSlideSize(files) {
    const xml = decode(files["ppt/presentation.xml"]);
    const tag = (xml.match(/<p:sldSz\b[^>]*\/?>/) || [])[0] || "";
    const cx = Number((tag.match(/\bcx="(\d+)"/) || [])[1] || 12192000);
    const cy = Number((tag.match(/\bcy="(\d+)"/) || [])[1] || 6858000);
    return { cx, cy };
  }

  function slidePartOrder(files) {
    const pres = parseXml(decode(files["ppt/presentation.xml"]));
    const rels = loadRels(files, "ppt/presentation.xml");
    const byId = new Map(rels.map((rel) => [rel.id, rel]));
    const parts = [];
    for (const sld of allLocal(pres, "sldId")) {
      const rid = attr(sld, "r:id");
      const rel = byId.get(rid);
      if (!rel || !/\/slide$/.test(rel.type)) continue;
      const part = resolveTarget("ppt/presentation.xml", rel.target);
      if (part && files[part]) parts.push(part);
    }
    return parts;
  }

  function readTheme(xml) {
    const doc = parseXml(xml);
    const colors = {};
    for (const name of ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]) {
      const el = allLocal(doc, name)[0];
      if (!el) continue;
      const srgb = child(el, "srgbClr");
      const sys = child(el, "sysClr");
      const val = srgb ? srgb.getAttribute("val") : sys && (sys.getAttribute("lastClr") || "");
      if (val) colors[name] = val.toUpperCase();
    }
    const major = allLocal(doc, "majorFont")[0];
    const minor = allLocal(doc, "minorFont")[0];
    return {
      colors,
      majorFont: (major && child(major, "latin") && child(major, "latin").getAttribute("typeface")) || "",
      minorFont: (minor && child(minor, "latin") && child(minor, "latin").getAttribute("typeface")) || "",
    };
  }

  function findThemeXml(files) {
    const direct = files["ppt/theme/theme1.xml"];
    if (direct) return decode(direct);
    const rels = Object.keys(files).filter((name) => /ppt\/slideMasters\/_rels\/slideMaster\d+\.xml\.rels$/.test(name));
    for (const relFile of rels) {
      const owner = ownerOfRels(relFile);
      for (const rel of parseRels(decode(files[relFile]))) {
        if (rel.type.endsWith("/theme")) {
          const part = resolveTarget(owner, rel.target);
          if (files[part]) return decode(files[part]);
        }
      }
    }
    return "";
  }

  function masterPlaceholders(files, masterPart) {
    if (!masterPart || !files[masterPart]) return [];
    const doc = parseXml(decode(files[masterPart]));
    const tree = allLocal(doc, "spTree")[0];
    if (!tree) return [];
    const found = [];
    walkShapes(tree, found, { x: 0, y: 0 });
    return found
      .map((item) => {
        const ph = placeholderInfo(item.el);
        const box = shapeBox(item.el);
        if (!ph || !box || SKIP_PH.has(ph.type)) return null;
        return { ...ph, box };
      })
      .filter(Boolean);
  }

  function readLayouts(files) {
    const layouts = [];
    const masterRels = Object.keys(files)
      .filter((name) => /^ppt\/slideMasters\/_rels\/slideMaster\d+\.xml\.rels$/.test(name))
      .sort();
    const seen = new Set();
    for (const relFile of masterRels) {
      const masterPart = ownerOfRels(relFile);
      const masterPh = masterPlaceholders(files, masterPart);
      for (const rel of parseRels(decode(files[relFile]))) {
        if (!rel.type.endsWith("/slideLayout")) continue;
        const part = resolveTarget(masterPart, rel.target);
        if (!part || seen.has(part) || !files[part]) continue;
        seen.add(part);
        const doc = parseXml(decode(files[part]));
        const cSld = allLocal(doc, "cSld")[0];
        const name = (cSld && cSld.getAttribute("name")) || part;
        const tree = allLocal(doc, "spTree")[0];
        const found = [];
        if (tree) walkShapes(tree, found, { x: 0, y: 0 });
        const placeholders = [];
        for (const item of found) {
          const ph = placeholderInfo(item.el);
          if (!ph || SKIP_PH.has(ph.type)) continue;
          const own = shapeBox(item.el);
          const box = own || inheritBox(ph, masterPh);
          placeholders.push({ ...ph, box });
        }
        const vertical = /vertical/i.test(name) || placeholders.some((ph) => ph.orient === "vert");
        layouts.push({ name, part, placeholders, vertical });
      }
    }
    return layouts;
  }

  function inheritBox(ph, masterPh) {
    if (ph.idx != null) {
      const byIdx = masterPh.find((item) => item.idx === ph.idx && item.box);
      if (byIdx) return byIdx.box;
    }
    if (ph.rawType) {
      const byType = masterPh.find((item) => item.rawType === ph.rawType && item.box);
      if (byType) return byType.box;
    }
    if (ph.type === "obj" || ph.type === "body") {
      const body = masterPh.find((item) => (item.type === "body" || item.type === "obj") && item.box);
      if (body) return body.box;
    }
    if (ph.type === "title" || ph.type === "ctrTitle") {
      const title = masterPh.find((item) => (item.type === "title" || item.type === "ctrTitle") && item.box);
      if (title) return title.box;
    }
    return null;
  }

  function readSlideModel(files, part) {
    const doc = parseXml(decode(files[part]));
    const tree = allLocal(doc, "spTree")[0];
    const found = [];
    if (tree) walkShapes(tree, found, { x: 0, y: 0 });
    const rels = loadRels(files, part);
    const texts = [];
    const images = [];
    const tables = [];
    const skipped = [];

    for (const item of found) {
      const el = item.el;
      const ph = placeholderInfo(el);
      if (ph && SKIP_PH.has(ph.type)) continue;
      const box = shapeBox(el);
      const x = item.origin.x + (box ? box.x : 0);
      const y = item.origin.y + (box ? box.y : 0);

      if (el.localName === "graphicFrame") {
        const tbl = allLocal(el, "tbl")[0];
        if (!tbl) {
          skipped.push("chart or diagram");
          continue;
        }
        const rows = [];
        for (const tr of kids(tbl)) {
          if (tr.localName !== "tr") continue;
          const row = [];
          for (const tc of kids(tr)) if (tc.localName === "tc") row.push(cellText(tc));
          if (row.length) rows.push(row);
        }
        if (rows.length) tables.push({ rows });
        continue;
      }

      const embed = embedOf(el);
      if (embed && (el.localName === "pic" || el.localName === "sp")) {
        const image = imageFromRel(embed, rels, part, files);
        if (image) images.push(image);
      }

      if (el.localName === "sp") {
        const tx = child(el, "txBody");
        const paragraphs = tx ? paragraphsOf(tx) : [];
        if (!paragraphs.length) continue;
        const fontSize = paragraphs.reduce((max, p) => Math.max(max, p.fontSize || 0), 0);
        texts.push({ ph, paragraphs, x, y, fontSize });
      }
    }

    return textsToSlide(texts, images, tables, skipped);
  }

  function textsToSlide(texts, images, tables, skipped) {
    const placed = texts.filter((shape) => shape.ph);
    const free = texts.filter((shape) => !shape.ph).sort((a, b) => a.y - b.y || a.x - b.x);
    let title = "";
    let subtitle = "";
    const bullets = [];

    const titleShape = placed.find((shape) => shape.ph.type === "title" || shape.ph.type === "ctrTitle");
    const subShape = placed.find((shape) => shape.ph.type === "subTitle");
    if (titleShape) {
      title = titleShape.paragraphs[0] ? titleShape.paragraphs[0].text : "";
      if (!subShape) {
        titleShape.paragraphs.slice(1).forEach((p) => bullets.push({ text: p.text, level: p.level || 0 }));
      }
    }
    if (subShape) subtitle = subShape.paragraphs.map((p) => p.text).join("\n");
    for (const shape of placed) {
      if (shape === titleShape || shape === subShape) continue;
      if (SKIP_PH.has(shape.ph.type)) continue;
      shape.paragraphs.forEach((p) => bullets.push({ text: p.text, level: p.level || 0 }));
    }
    for (const shape of free) {
      if (!title) {
        title = shape.paragraphs[0] ? shape.paragraphs[0].text : "";
        shape.paragraphs.slice(1).forEach((p) => bullets.push({ text: p.text, level: p.level || 0 }));
        continue;
      }
      shape.paragraphs.forEach((p) => bullets.push({ text: p.text, level: p.level || 0 }));
    }

    return { title: title.trim(), subtitle: subtitle.trim(), bullets, images, tables, skipped };
  }

  function classifySlide(slide, index) {
    const bullets = slide.bullets || [];
    const images = slide.images || [];
    const tables = slide.tables || [];
    const title = (slide.title || "").trim();
    const subtitle = (slide.subtitle || "").trim();
    const words = title.split(/\s+/).filter(Boolean).length;
    if (tables.length) return "table";
    if (images.length) return "image-text";
    if (index === 0 && bullets.length === 0) return "title";
    if (bullets.length === 0 && !subtitle && words > 0 && words <= 8 && title.length <= 72) return "section";
    if (bullets.length) return "title-bullets";
    if (title || subtitle) return "title-bullets";
    return "blank";
  }

  function hasType(layout, type) {
    return layout.placeholders.some((ph) => ph.type === type);
  }

  function scoreLayout(kind, layout) {
    const name = (layout.name || "").toLowerCase();
    const title = hasType(layout, "title") || hasType(layout, "ctrTitle");
    const body = hasType(layout, "body") || hasType(layout, "obj");
    const pic = hasType(layout, "pic");
    let score = 0;
    if (layout.vertical || name.includes("vertical")) score -= 40;
    if (kind === "title" || kind === "watermark") {
      if (name === "title slide") score += 80;
      if (hasType(layout, "ctrTitle")) score += 25;
      if (hasType(layout, "subTitle")) score += 15;
      if (title) score += 6;
      if (pic) score -= 8;
      if (kind === "watermark" && hasType(layout, "ctrTitle") && hasType(layout, "subTitle")) score += 10;
    } else if (kind === "section") {
      if (name === "section header") score += 80;
      if (name.includes("section")) score += 35;
      if (name.includes("title only")) score += 18;
      if (title && !body && !pic) score += 12;
      if (title) score += 6;
      if (pic) score -= 8;
    } else if (kind === "image-text") {
      if (name === "picture with caption") score += 80;
      if (name.includes("picture") || name.includes("photo") || name.includes("image")) score += 20;
      if (pic && title) score += 24;
      if (pic && body) score += 12;
      if (pic) score += 8;
      if (!pic && title && body) score += 6;
    } else {
      if (name === "title and content") score += 70;
      if (name.includes("title") && name.includes("content") && !name.includes("two")) score += 20;
      if (title && body) score += 22;
      if (hasType(layout, "obj")) score += 6;
      if (name.includes("two content") || name.includes("comparison")) score -= 8;
      if (pic) score -= 6;
      if (name === "blank" || name === "title only") score -= 4;
      if (kind === "table" && body) score += 4;
    }
    return score;
  }

  function matchLayout(kind, layouts) {
    let best = null;
    let bestScore = -Infinity;
    (layouts || []).forEach((layout, index) => {
      const score = scoreLayout(kind, layout) - index * 0.001;
      if (score > bestScore) {
        bestScore = score;
        best = layout;
      }
    });
    return best;
  }

  function planSlides(slides, options) {
    const pro = !!(options && options.pro);
    const classified = (slides || []).map((slide, index) => ({
      ...slide,
      sourceIndex: index,
      kind: classifySlide(slide, index),
    }));
    if (pro) return { slides: classified, omitted: 0, watermark: false };
    return {
      slides: classified.slice(0, 5),
      omitted: Math.max(0, classified.length - 5),
      watermark: true,
    };
  }

  function clampBox(box, size) {
    if (!box) return null;
    let x = Math.round(box.x);
    let y = Math.round(box.y);
    let cx = Math.round(box.cx);
    let cy = Math.round(box.cy);
    if (x < 0) x = 0;
    if (y < 0) y = 0;
    if (x + cx > size.cx) cx = Math.max(180000, size.cx - x);
    if (y + cy > size.cy) cy = Math.max(180000, size.cy - y);
    if (cx < 180000) cx = 180000;
    if (cy < 120000) cy = 120000;
    return { x, y, cx, cy };
  }

  function fallbackBox(size, role) {
    const margin = Math.round(size.cx * 0.07);
    if (role === "title") {
      return { x: margin, y: Math.round(size.cy * 0.08), cx: size.cx - margin * 2, cy: Math.round(size.cy * 0.16) };
    }
    if (role === "pic") {
      return {
        x: Math.round(size.cx * 0.52),
        y: Math.round(size.cy * 0.28),
        cx: Math.round(size.cx * 0.4),
        cy: Math.round(size.cy * 0.56),
      };
    }
    return { x: margin, y: Math.round(size.cy * 0.28), cx: size.cx - margin * 2, cy: Math.round(size.cy * 0.58) };
  }

  function linesOf(text) {
    return String(text || "")
      .split(/\n/)
      .map((line) => line.trim())
      .filter(Boolean);
  }

  function phTag(ph) {
    if (!ph) return "";
    const attrs = [];
    if (ph.rawType) attrs.push(`type="${xmlEscape(ph.rawType)}"`);
    else if (ph.type && ph.type !== "obj") attrs.push(`type="${xmlEscape(ph.type)}"`);
    if (ph.idx != null && ph.idx !== "") attrs.push(`idx="${xmlEscape(ph.idx)}"`);
    if (ph.orient) attrs.push(`orient="${xmlEscape(ph.orient)}"`);
    if (ph.sz) attrs.push(`sz="${xmlEscape(ph.sz)}"`);
    return `<p:ph${attrs.length ? " " + attrs.join(" ") : ""}/>`;
  }

  function runProps(role, box) {
    const major = role === "title" || role === "subtitle";
    const face = major ? "+mj-lt" : "+mn-lt";
    const ea = major ? "+mj-ea" : "+mn-ea";
    const cs = major ? "+mj-cs" : "+mn-cs";
    let sz = "";
    if (box && role === "title" && box.cy < 800000) sz = ' sz="2000"';
    else if (box && role === "subtitle" && box.cy < 900000) sz = ' sz="1600"';
    else if (box && role === "body" && box.cy < 1200000) sz = ' sz="1400"';
    return `<a:rPr lang="en-US" dirty="0"${sz}><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="${face}"/><a:ea typeface="${ea}"/><a:cs typeface="${cs}"/></a:rPr>`;
  }

  function paragraphXml(text, role, level, box) {
    const pPr = role === "body" ? `<a:pPr lvl="${level || 0}"/>` : "<a:pPr/>";
    return `<a:p>${pPr}<a:r>${runProps(role, box)}<a:t>${xmlEscape(text)}</a:t></a:r></a:p>`;
  }

  function textShapeXml(id, name, ph, box, role, paragraphs) {
    const nvPr = ph ? `<p:nvPr>${phTag(ph)}</p:nvPr>` : "<p:nvPr/>";
    const anchor = role === "body" ? "t" : "ctr";
    const body = paragraphs
      .map((p) => paragraphXml(p.text, role, p.level || 0, box))
      .join("");
    return (
      "<p:sp>" +
      "<p:nvSpPr>" +
      `<p:cNvPr id="${id}" name="${xmlEscape(name)}"/>` +
      '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
      nvPr +
      "</p:nvSpPr>" +
      "<p:spPr>" +
      `<a:xfrm><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.cx}" cy="${box.cy}"/></a:xfrm>` +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>' +
      "</p:spPr>" +
      "<p:txBody>" +
      `<a:bodyPr wrap="square" anchor="${anchor}"/>` +
      "<a:lstStyle/>" +
      body +
      "</p:txBody></p:sp>"
    );
  }

  function pictureXml(id, relId, box, ph) {
    const nvPr = ph ? `<p:nvPr>${phTag(ph)}</p:nvPr>` : "<p:nvPr/>";
    return (
      "<p:pic><p:nvPicPr>" +
      `<p:cNvPr id="${id}" name="Picture ${id}"/>` +
      '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>' +
      nvPr +
      "</p:nvPicPr>" +
      `<p:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      "<p:spPr>" +
      `<a:xfrm><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.cx}" cy="${box.cy}"/></a:xfrm>` +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      "</p:spPr></p:pic>"
    );
  }

  function contain(box, width, height) {
    if (!box) return null;
    if (!width || !height) return box;
    const scale = Math.min(box.cx / width, box.cy / height);
    const cx = Math.max(1, Math.round(width * scale));
    const cy = Math.max(1, Math.round(height * scale));
    return {
      x: box.x + Math.round((box.cx - cx) / 2),
      y: box.y + Math.round((box.cy - cy) / 2),
      cx,
      cy,
    };
  }

  function gridBox(box, count, index) {
    const cols = Math.ceil(Math.sqrt(count));
    const rows = Math.ceil(count / cols);
    const gap = count > 1 ? 80000 : 0;
    const cw = Math.floor((box.cx - gap * (cols - 1)) / cols);
    const ch = Math.floor((box.cy - gap * (rows - 1)) / rows);
    const col = index % cols;
    const row = Math.floor(index / cols);
    return { x: box.x + col * (cw + gap), y: box.y + row * (ch + gap), cx: cw, cy: ch };
  }

  function border(tag) {
    return `<a:${tag} w="6350" cap="flat" cmpd="sng"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:prstDash val="solid"/></a:${tag}>`;
  }

  function tableXml(id, table, box) {
    const rows = table.rows || [];
    const cols = rows.reduce((max, row) => Math.max(max, row.length), 0);
    if (!cols || !rows.length) return "";
    const colW = Math.floor(box.cx / cols);
    const widths = Array.from({ length: cols }, (_, i) => (i === cols - 1 ? box.cx - colW * (cols - 1) : colW));
    const rowH = Math.floor(box.cy / rows.length);
    const heights = rows.map((_, i) => (i === rows.length - 1 ? box.cy - rowH * (rows.length - 1) : rowH));
    const grid = widths.map((w) => `<a:gridCol w="${w}"/>`).join("");
    const body = rows
      .map((row, r) => {
        const cells = widths
          .map((_, c) => {
            const text = row[c] == null ? "" : String(row[c]);
            const header = r === 0 && rows.length > 1;
            const fill = header ? "accent1" : "bg1";
            const ink = header ? "bg1" : "tx1";
            return (
              "<a:tc>" +
              "<a:txBody><a:bodyPr anchor=\"ctr\"/><a:lstStyle/>" +
              `<a:p><a:pPr/><a:r><a:rPr lang="en-US" sz="1400" dirty="0"><a:solidFill><a:schemeClr val="${ink}"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:rPr><a:t>${xmlEscape(text)}</a:t></a:r></a:p>` +
              "</a:txBody>" +
              `<a:tcPr marL="72000" marR="72000" marT="36000" marB="36000"><a:solidFill><a:schemeClr val="${fill}"/></a:solidFill>${border("lnL")}${border("lnR")}${border("lnT")}${border("lnB")}</a:tcPr>` +
              "</a:tc>"
            );
          })
          .join("");
        return `<a:tr h="${heights[r]}">${cells}</a:tr>`;
      })
      .join("");
    return (
      "<p:graphicFrame>" +
      "<p:nvGraphicFramePr>" +
      `<p:cNvPr id="${id}" name="Table ${id}"/>` +
      '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>' +
      "<p:nvPr/>" +
      "</p:nvGraphicFramePr>" +
      `<p:xfrm><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.cx}" cy="${box.cy}"/></p:xfrm>` +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">' +
      `<a:tbl><a:tblPr firstRow="1"/><a:tblGrid>${grid}</a:tblGrid>${body}</a:tbl>` +
      "</a:graphicData></a:graphic></p:graphicFrame>"
    );
  }

  function assignPlaces(slide, layout) {
    const phs = (layout.placeholders || []).filter((ph) => !SKIP_PH.has(ph.type));
    const used = new Set();
    function take(pred) {
      const index = phs.findIndex((ph, i) => !used.has(i) && pred(ph));
      if (index < 0) return null;
      used.add(index);
      return phs[index];
    }
    const isTitle = (ph) => ph.type === "title" || ph.type === "ctrTitle";
    let titlePh = null;
    let subPh = null;
    let bodyPh = null;
    let picPh = null;
    if (slide.kind === "title" || slide.kind === "watermark") {
      titlePh = take((ph) => ph.type === "ctrTitle") || take(isTitle);
      subPh = take((ph) => ph.type === "subTitle") || take((ph) => ph.type === "body" || ph.type === "obj");
    } else if (slide.kind === "section") {
      titlePh = take(isTitle);
      subPh = take((ph) => ph.type === "body" || ph.type === "subTitle" || ph.type === "obj");
    } else if (slide.kind === "image-text") {
      titlePh = take(isTitle);
      picPh = take((ph) => ph.type === "pic");
      bodyPh = take((ph) => ph.type === "body" || ph.type === "obj");
    } else {
      titlePh = take(isTitle);
      bodyPh = take((ph) => ph.type === "body" || ph.type === "obj");
      picPh = take((ph) => ph.type === "pic");
    }
    return { titlePh, subPh, bodyPh, picPh };
  }

  function buildSlide(slide, template, mediaState) {
    const size = template.size;
    const layout = slide.layout;
    const places = assignPlaces(slide, layout);
    const rels = [
      {
        id: "rId1",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
        target: relativeTarget("ppt/slides/slide1.xml", layout.part),
      },
    ];
    const media = [];
    let shapeId = 2;
    const chunks = [];

    function boxFor(ph, role) {
      const raw = ph && ph.box ? ph.box : fallbackBox(size, role);
      return clampBox(raw, size);
    }

    const titleParas = linesOf(slide.title).map((text) => ({ text, level: 0 }));
    const subtitleParas = linesOf(slide.subtitle).map((text) => ({ text, level: 0 }));
    const bulletParas = (slide.bullets || [])
      .filter((b) => b && String(b.text || "").trim())
      .map((b) => ({ text: String(b.text).trim(), level: b.level || 0 }));

    if (titleParas.length) {
      const ph = places.titlePh;
      chunks.push(textShapeXml(shapeId++, "Title", ph, boxFor(ph, "title"), "title", titleParas));
    }

    let bodyParas = bulletParas;
    let subParas = subtitleParas;
    if ((slide.kind === "title" || slide.kind === "watermark") && bulletParas.length && places.subPh) {
      subParas = subParas.concat(bulletParas);
      bodyParas = [];
    } else if (slide.kind === "section" && places.subPh) {
      subParas = subParas.concat(bulletParas);
      bodyParas = [];
    }

    if (subParas.length) {
      const ph = places.subPh || places.bodyPh;
      const role = places.subPh ? "subtitle" : "body";
      chunks.push(textShapeXml(shapeId++, "Subtitle", ph, boxFor(ph, role === "subtitle" ? "subtitle" : "body"), role, subParas));
      if (!places.subPh) bodyParas = [];
    }

    let bodyBox = places.bodyPh ? boxFor(places.bodyPh, "body") : bodyParas.length ? fallbackBox(size, "body") : null;
    let imageBox = places.picPh ? boxFor(places.picPh, "pic") : null;
    if (!imageBox && slide.images && slide.images.length) {
      if (bodyBox && bodyParas.length) {
        const left = Math.round(bodyBox.cx * 0.56);
        imageBox = clampBox(
          { x: bodyBox.x + left + 100000, y: bodyBox.y, cx: bodyBox.cx - left - 100000, cy: bodyBox.cy },
          size,
        );
        bodyBox = { ...bodyBox, cx: left };
      } else {
        imageBox = bodyBox || fallbackBox(size, "pic");
        if (!bodyParas.length) bodyBox = null;
      }
    }

    if (slide.tables && slide.tables.length) {
      const tableBox = bodyBox || fallbackBox(size, "body");
      if (bodyParas.length) {
        const top = Math.round(tableBox.cy * 0.34);
        chunks.push(textShapeXml(shapeId++, "Body", places.bodyPh, { ...tableBox, cy: top }, "body", bodyParas));
        bodyParas = [];
        slide.tables.forEach((table) => {
          chunks.push(
            tableXml(shapeId++, table, {
              x: tableBox.x,
              y: tableBox.y + top + 60000,
              cx: tableBox.cx,
              cy: Math.max(200000, tableBox.cy - top - 60000),
            }),
          );
        });
      } else {
        slide.tables.forEach((table) => chunks.push(tableXml(shapeId++, table, tableBox)));
      }
    } else if (bodyParas.length && bodyBox) {
      chunks.push(textShapeXml(shapeId++, "Body", places.bodyPh, bodyBox, "body", bodyParas));
    }

    (slide.images || []).forEach((image, index) => {
      if (!imageBox || !image || !image.bytes) return;
      const ext = (image.ext || "png").replace(/[^a-z0-9]/gi, "") || "png";
      const filename = `slidefit-${mediaState.next++}.${ext}`;
      const path = `ppt/media/${filename}`;
      media.push({ path, bytes: toUint8(image.bytes) });
      const relId = "rId" + (rels.length + 1);
      rels.push({
        id: relId,
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        target: "../media/" + filename,
      });
      const cell = gridBox(imageBox, slide.images.length, index);
      const fitted = contain(cell, image.width, image.height);
      const ph = index === 0 ? places.picPh : null;
      chunks.push(pictureXml(shapeId++, relId, fitted, ph));
    });

    const xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      "<p:cSld><p:spTree>" +
      '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
      chunks.join("") +
      "</p:spTree></p:cSld>" +
      "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>";

    return { xml, rels: serializeRels(rels), media };
  }

  function stripExistingSlides(files) {
    const relPath = "ppt/_rels/presentation.xml.rels";
    const rels = parseRels(decode(files[relPath])).filter((rel) => {
      if (!rel.type.endsWith("/slide")) return true;
      const part = resolveTarget("ppt/presentation.xml", rel.target);
      if (part) {
        delete files[part];
        delete files[relsPath(part)];
      }
      return false;
    });
    files[relPath] = serializeRels(rels);
    let pres = decode(files["ppt/presentation.xml"]);
    if (/<p:sldIdLst\/>/.test(pres)) pres = pres.replace(/<p:sldIdLst\/>/, "<p:sldIdLst></p:sldIdLst>");
    if (/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/.test(pres)) {
      pres = pres.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, "<p:sldIdLst></p:sldIdLst>");
    } else {
      pres = pres.replace("</p:sldMasterIdLst>", "</p:sldMasterIdLst><p:sldIdLst></p:sldIdLst>");
    }
    files["ppt/presentation.xml"] = pres;
    return rels;
  }

  function collectKeep(files) {
    const keep = new Set(["[Content_Types].xml"]);
    const queue = ["_rels/.rels"];
    const seen = new Set();
    while (queue.length) {
      const relFile = queue.pop();
      if (!relFile || seen.has(relFile) || !files[relFile]) continue;
      seen.add(relFile);
      keep.add(relFile);
      const owner = ownerOfRels(relFile);
      let rels = [];
      try {
        rels = parseRels(decode(files[relFile]));
      } catch (err) {
        continue;
      }
      for (const rel of rels) {
        if (rel.targetMode === "External") continue;
        const target = resolveTarget(owner, rel.target);
        if (!target) continue;
        keep.add(target);
        const nested = relsPath(target);
        if (files[nested]) queue.push(nested);
      }
    }
    return keep;
  }

  function syncContentTypes(files, slideCount) {
    let xml = decode(files["[Content_Types].xml"]);
    xml = xml.replace(/<Override\b[^>]*\/>/g, (tag) => {
      const partName = (tag.match(/PartName="([^"]+)"/) || [])[1];
      if (!partName) return tag;
      const part = safeDecode(partName.replace(/^\//, ""));
      return files[part] ? tag : "";
    });
    for (let i = 1; i <= slideCount; i++) {
      const partName = `/ppt/slides/slide${i}.xml`;
      if (!xml.includes(`PartName="${partName}"`)) {
        xml = xml.replace(
          "</Types>",
          `<Override PartName="${partName}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`,
        );
      }
    }
    const defaults = {
      png: "image/png",
      jpeg: "image/jpeg",
      jpg: "image/jpeg",
      gif: "image/gif",
      emf: "image/x-emf",
      wmf: "image/x-wmf",
    };
    Object.keys(defaults).forEach((ext) => {
      const used = Object.keys(files).some((name) => name.toLowerCase().endsWith("." + ext));
      if (used && !new RegExp(`Extension="${ext}"`, "i").test(xml)) {
        xml = xml.replace(/(<Types\b[^>]*>)/, `$1<Default Extension="${ext}" ContentType="${defaults[ext]}"/>`);
      }
    });
    files["[Content_Types].xml"] = xml;
    if (files["docProps/app.xml"]) {
      files["docProps/app.xml"] = decode(files["docProps/app.xml"]).replace(/<Slides>\d+<\/Slides>/, `<Slides>${slideCount}</Slides>`);
    }
  }

  function summarize(slide) {
    return {
      kind: slide.kind,
      title: slide.title || "",
      subtitle: slide.subtitle || "",
      bullets: slide.bullets || [],
      tables: slide.tables || [],
      images: slide.images || [],
      skipped: slide.skipped || [],
      layoutName: slide.layoutName || (slide.layout && slide.layout.name) || "",
      sourceIndex: slide.sourceIndex,
    };
  }

  function watermarkSlide(layout) {
    return {
      kind: "watermark",
      title: "SlideFit free watermark",
      subtitle: "Free conversions include up to 5 slides from your deck, plus this slide.",
      bullets: [
        { text: "Pro removes the watermark and converts every slide.", level: 0 },
        { text: "Your files stayed in this browser. Nothing was uploaded.", level: 0 },
      ],
      images: [],
      tables: [],
      skipped: [],
      layout,
      layoutName: layout.name,
    };
  }

  async function parseDeck(bytes) {
    const files = await unzip(bytes);
    const slides = slidePartOrder(files).map((part) => readSlideModel(files, part));
    const themeXml = findThemeXml(files);
    return { slides, size: readSlideSize(files), theme: themeXml ? readTheme(themeXml) : { colors: {}, majorFont: "", minorFont: "" } };
  }

  async function parseTemplate(bytes) {
    const files = await unzip(bytes);
    const themeXml = findThemeXml(files);
    return {
      files,
      layouts: readLayouts(files),
      size: readSlideSize(files),
      theme: themeXml ? readTheme(themeXml) : { colors: {}, majorFont: "", minorFont: "" },
    };
  }

  async function applyTemplate(sourceBytes, templateBytes, options) {
    const template = await parseTemplate(templateBytes);
    const deck = await parseDeck(sourceBytes);
    if (!deck.slides.length) throw new SlideFitError("The deck has no slides.");
    if (!template.layouts.length) throw new SlideFitError("The template has no slide layouts to apply.");
    const plan = planSlides(deck.slides, options || {});
    const produced = plan.slides.map((slide) => {
      const layout = matchLayout(slide.kind, template.layouts);
      return { ...slide, layout, layoutName: layout ? layout.name : "" };
    });
    if (plan.watermark) produced.push(watermarkSlide(matchLayout("watermark", template.layouts)));

    const files = Object.assign({}, template.files);
    let rels = stripExistingSlides(files);
    const relPath = "ppt/_rels/presentation.xml.rels";
    const mediaState = { next: 1 };
    const sldIds = [];
    produced.forEach((slide, index) => {
      const built = buildSlide(slide, template, mediaState);
      const part = `ppt/slides/slide${index + 1}.xml`;
      files[part] = built.xml;
      files[`ppt/slides/_rels/slide${index + 1}.xml.rels`] = built.rels;
      built.media.forEach((item) => {
        files[item.path] = item.bytes;
      });
      const rid = freshId(rels);
      rels.push({
        id: rid,
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
        target: `slides/slide${index + 1}.xml`,
      });
      sldIds.push(`<p:sldId id="${256 + index}" r:id="${rid}"/>`);
    });
    files[relPath] = serializeRels(rels);
    files["ppt/presentation.xml"] = decode(files["ppt/presentation.xml"]).replace(
      "<p:sldIdLst></p:sldIdLst>",
      `<p:sldIdLst>${sldIds.join("")}</p:sldIdLst>`,
    );

    const keep = collectKeep(files);
    Object.keys(files).forEach((name) => {
      if (!keep.has(name)) delete files[name];
    });
    syncContentTypes(files, produced.length);
    const bytes = await zipFiles(files);
    return {
      bytes,
      slides: produced.map(summarize),
      omitted: plan.omitted,
      watermark: plan.watermark,
      theme: template.theme,
      size: template.size,
    };
  }

  function freshId(rels) {
    const used = new Set(rels.map((rel) => rel.id));
    let n = 1;
    while (used.has("rId" + n)) n++;
    return "rId" + n;
  }

  return {
    SlideFitError,
    classifySlide,
    matchLayout,
    planSlides,
    parseDeck,
    parseTemplate,
    applyTemplate,
  };
});
