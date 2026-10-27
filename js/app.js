(function () {
  const cfg = window.SLIDEFIT_CONFIG || {};
  const KEYS = {
    monthly: "slidefit.monthlyUntil",
    decks: "slidefit.unlockedDecks",
    pending: "slidefit.pendingDeck",
  };
  const KIND_LABEL = {
    title: "Title slide",
    section: "Section header",
    "title-bullets": "Title and bullets",
    "image-text": "Image and text",
    table: "Table",
    blank: "Content",
    watermark: "Watermark",
  };

  const state = {
    deckBuf: null,
    deckName: "",
    deckHash: "",
    templateBuf: null,
    templateName: "",
    result: null,
    selected: 0,
  };

  const el = (id) => document.getElementById(id);

  function paypalConfigured() {
    const id = cfg.PAYPAL_CLIENT_ID || "";
    const plan = cfg.PLAN_ID || "";
    return id && !id.includes("YOUR_PAYPAL") && plan && !plan.includes("YOUR_PLAN");
  }

  function readDecks() {
    try {
      const parsed = JSON.parse(localStorage.getItem(KEYS.decks) || "[]");
      return Array.isArray(parsed) ? parsed : [];
    } catch (err) {
      return [];
    }
  }

  function monthlyUntil() {
    return Number(localStorage.getItem(KEYS.monthly) || 0);
  }

  async function digest(buf) {
    if (crypto.subtle) {
      const hash = await crypto.subtle.digest("SHA-256", buf);
      return Array.from(new Uint8Array(hash))
        .map((b) => b.toString(16).padStart(2, "0"))
        .join("");
    }
    return "len-" + buf.byteLength;
  }

  function accessFor(hash) {
    if (monthlyUntil() > Date.now()) return { pro: true, reason: "monthly" };
    const decks = readDecks();
    if (localStorage.getItem(KEYS.pending) === "1" && hash) {
      decks.push(hash);
      localStorage.setItem(KEYS.decks, JSON.stringify(decks));
      localStorage.removeItem(KEYS.pending);
    }
    if (hash && decks.includes(hash)) return { pro: true, reason: "deck" };
    return { pro: false, reason: "free" };
  }

  function setStatus(text) {
    el("status").textContent = text;
  }

  function showError(text) {
    const box = el("alert");
    box.hidden = !text;
    box.textContent = text || "";
  }

  function refreshUnlockNote() {
    const access = accessFor(state.deckHash);
    const note = el("unlock-status");
    if (access.reason === "monthly") {
      const when = new Date(monthlyUntil());
      note.textContent = "Pro monthly is on in this browser until " + when.toLocaleDateString() + ". It lives in local storage, not on a server.";
    } else if (access.reason === "deck") {
      note.textContent = "This deck is unlocked in this browser. Clearing site data removes the unlock.";
    } else if (localStorage.getItem(KEYS.pending) === "1") {
      note.textContent = "A one-deck unlock is waiting. Load the deck, then fit it.";
    } else {
      note.textContent = "Free plan: the first 5 slides, plus a watermark slide.";
    }
  }

  function refreshFitButton() {
    el("fit-button").disabled = !(state.deckBuf && state.templateBuf);
  }

  function markReady(id, name) {
    const zone = el(id);
    zone.classList.add("ready");
    zone.querySelector(".drop-title").textContent = name;
  }

  async function useBuffer(kind, buf, name) {
    if (kind === "deck") {
      state.deckBuf = buf;
      state.deckName = name;
      state.deckHash = await digest(buf);
      markReady("drop-deck", name);
    } else {
      state.templateBuf = buf;
      state.templateName = name;
      markReady("drop-template", name);
    }
    refreshFitButton();
    refreshUnlockNote();
    showError("");
    setStatus("Both files stay in this tab. Nothing is uploaded.");
  }

  async function useFile(kind, file) {
    if (!file) return;
    if (!/\.(pptx|potx|ppsx)$/i.test(file.name)) {
      showError("Use a PowerPoint file (.pptx). A .potx template can be saved as .pptx, and SlideFit also accepts .potx and .ppsx.");
      return;
    }
    const buf = await file.arrayBuffer();
    await useBuffer(kind, buf, file.name);
  }

  async function loadSample(kind) {
    const url = kind === "deck" ? "samples/northwind-quarterly.pptx" : "samples/harbor-company-template.pptx";
    const name = url.split("/").pop();
    setStatus("Loading " + name + "…");
    const res = await fetch(url);
    if (!res.ok) throw new Error("Could not load " + name + ". Open the site over http, not as a local file.");
    await useBuffer(kind, await res.arrayBuffer(), name);
  }

  function bindDrop(zoneId, inputId, kind) {
    const zone = el(zoneId);
    const input = el(inputId);
    ["dragenter", "dragover"].forEach((type) => {
      zone.addEventListener(type, (event) => {
        event.preventDefault();
        zone.classList.add("hot");
      });
    });
    ["dragleave", "drop"].forEach((type) => {
      zone.addEventListener(type, (event) => {
        event.preventDefault();
        zone.classList.remove("hot");
      });
    });
    zone.addEventListener("drop", (event) => {
      const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
      if (file) useFile(kind, file).catch((err) => showError(err.message));
    });
    input.addEventListener("change", () => {
      if (input.files && input.files[0]) useFile(kind, input.files[0]).catch((err) => showError(err.message));
    });
  }

  function fontStack(name, fallback) {
    if (!name || name.charAt(0) === "+") return fallback;
    return '"' + name.replace(/"/g, "") + '", ' + fallback;
  }

  function revokePreviews() {
    (state.result && state.result.slides ? state.result.slides : []).forEach((slide) => {
      (slide.images || []).forEach((image) => {
        if (image.previewUrl) URL.revokeObjectURL(image.previewUrl);
        image.previewUrl = "";
      });
    });
  }

  function preparePreviews(result) {
    result.slides.forEach((slide) => {
      (slide.images || []).forEach((image) => {
        if (!image.bytes) return;
        const blob = new Blob([image.bytes], { type: image.contentType || "image/png" });
        image.previewUrl = URL.createObjectURL(blob);
      });
    });
  }

  function addText(parent, tag, text, className) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    node.textContent = text;
    parent.appendChild(node);
    return node;
  }

  function renderSlide(slide, theme) {
    const view = document.createElement("article");
    view.className = "slide-view" + (slide.kind === "watermark" ? " is-watermark" : "");
    const bg = "#" + ((theme.colors && theme.colors.lt1) || "F7F3EA");
    const ink = "#" + ((theme.colors && theme.colors.dk1) || "1C2B2A");
    const accent = "#" + ((theme.colors && theme.colors.accent1) || "0E7C7B");
    view.style.background = bg;
    view.style.color = ink;
    view.style.setProperty("--accent", accent);
    view.style.fontFamily = fontStack(theme.minorFont, "Calibri, sans-serif");
    addText(view, "p", (KIND_LABEL[slide.kind] || "Slide") + " · " + (slide.layoutName || "Layout"), "slide-kicker");
    if (slide.title) {
      const title = addText(view, "h3", slide.title);
      title.style.fontFamily = fontStack(theme.majorFont, "Georgia, serif");
    }
    if (slide.subtitle) addText(view, "p", slide.subtitle, "subtitle");

    const bullets = slide.bullets || [];
    const images = (slide.images || []).filter((image) => image.previewUrl);
    const tables = slide.tables || [];
    const body = document.createElement("div");
    if (images.length && (bullets.length || slide.subtitle)) body.className = "slide-columns";

    if (bullets.length) {
      const list = document.createElement("ul");
      bullets.forEach((bullet) => {
        const item = document.createElement("li");
        item.textContent = bullet.text;
        if (bullet.level) item.style.marginLeft = bullet.level * 12 + "px";
        list.appendChild(item);
      });
      body.appendChild(list);
    }
    images.forEach((image) => {
      const img = document.createElement("img");
      img.src = image.previewUrl;
      img.alt = slide.title ? "Picture from “" + slide.title + "”" : "Picture from the source slide";
      body.appendChild(img);
    });
    if (body.childNodes.length) view.appendChild(body);

    tables.forEach((table) => {
      const html = document.createElement("table");
      (table.rows || []).forEach((row, r) => {
        const tr = document.createElement("tr");
        row.forEach((cell) => {
          const td = document.createElement(r === 0 ? "th" : "td");
          td.textContent = cell;
          tr.appendChild(td);
        });
        html.appendChild(tr);
      });
      view.appendChild(html);
    });

    if (slide.skipped && slide.skipped.length) {
      addText(view, "p", "Skipped: " + slide.skipped.join(", ") + ".", "fine");
    }
    addText(view, "p", "Preview", "wordmark");
    return view;
  }

  function renderPreview(result) {
    const theme = result.theme || { colors: {}, majorFont: "", minorFont: "" };
    const preview = el("preview");
    preview.hidden = false;
    state.selected = Math.min(state.selected, result.slides.length - 1);
    const stage = el("stage");
    stage.replaceChildren(renderSlide(result.slides[state.selected], theme));

    const thumbs = el("thumbs");
    thumbs.replaceChildren();
    result.slides.forEach((slide, index) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "thumb";
      button.setAttribute("aria-current", index === state.selected ? "true" : "false");
      const label = "Slide " + (index + 1) + ", " + (KIND_LABEL[slide.kind] || "Slide");
      button.setAttribute("aria-label", label + (slide.title ? ", " + slide.title : ""));
      const strong = document.createElement("strong");
      strong.textContent = index + 1 + ". " + (slide.title || KIND_LABEL[slide.kind] || "Slide");
      button.appendChild(strong);
      button.appendChild(document.createTextNode(slide.layoutName || ""));
      button.addEventListener("click", () => {
        state.selected = index;
        renderPreview(result);
      });
      thumbs.appendChild(button);
    });

    const lines = result.slides.map((slide, index) => {
      return index + 1 + ". " + (slide.title || KIND_LABEL[slide.kind]) + " → " + slide.layoutName;
    });
    el("map-note").textContent = lines.join("  ·  ");

    const access = accessFor(state.deckHash);
    if (result.omitted > 0) {
      setStatus("Free plan kept 5 of " + (5 + result.omitted) + " slides and added a watermark. Pro converts the whole deck and leaves that slide out.");
    } else if (result.watermark) {
      setStatus("Free plan added a watermark slide at the end. Pro leaves it out.");
    } else if (access.reason === "monthly") {
      setStatus("Pro monthly: " + result.slides.length + " slides, no watermark.");
    } else {
      setStatus("Pro for this deck: " + result.slides.length + " slides, no watermark.");
    }
  }

  async function fit() {
    if (!window.SlideFit) {
      showError("The converter did not load.");
      return;
    }
    showError("");
    el("fit-button").disabled = true;
    setStatus("Fitting slides in this tab…");
    try {
      const access = accessFor(state.deckHash);
      revokePreviews();
      const result = await window.SlideFit.applyTemplate(state.deckBuf, state.templateBuf, { pro: access.pro });
      preparePreviews(result);
      state.result = result;
      state.selected = 0;
      renderPreview(result);
      el("download").disabled = false;
    } catch (err) {
      showError(err && err.message ? err.message : "Could not fit that deck.");
      setStatus("");
    } finally {
      refreshFitButton();
    }
  }

  function download() {
    if (!state.result) return;
    const blob = new Blob([state.result.bytes], {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const base = (state.deckName || "deck").replace(/\.(pptx|potx|ppsx)$/i, "");
    link.href = url;
    link.download = "SlideFit-" + base + ".pptx";
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2500);
  }

  function unlockDeck() {
    if (state.deckHash) {
      const decks = readDecks();
      if (!decks.includes(state.deckHash)) decks.push(state.deckHash);
      localStorage.setItem(KEYS.decks, JSON.stringify(decks));
    } else {
      localStorage.setItem(KEYS.pending, "1");
    }
    refreshUnlockNote();
    if (state.deckBuf && state.templateBuf) fit();
  }

  function unlockMonthly() {
    localStorage.setItem(KEYS.monthly, String(Date.now() + 30 * 24 * 60 * 60 * 1000));
    refreshUnlockNote();
    if (state.deckBuf && state.templateBuf) fit();
  }

  function forget() {
    localStorage.removeItem(KEYS.monthly);
    localStorage.removeItem(KEYS.decks);
    localStorage.removeItem(KEYS.pending);
    refreshUnlockNote();
    setStatus("Unlocks stored in this browser were removed.");
  }

  function loadScript(src, namespace) {
    return new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = src;
      script.async = true;
      if (namespace) script.setAttribute("data-namespace", namespace);
      script.onload = () => resolve();
      script.onerror = () => reject(new Error("PayPal failed to load."));
      document.head.appendChild(script);
    });
  }

  async function setupPayPal() {
    const note = el("paypal-note");
    if (!paypalConfigured()) {
      el("simulate").hidden = false;
      note.textContent = "PayPal is not configured yet. config.js still has YOUR_PAYPAL_CLIENT_ID and YOUR_PLAN_ID. The simulate buttons only write a local unlock so you can preview Pro. They do not charge a card.";
      return;
    }
    note.textContent = "PayPal opens in a secure window. SlideFit still does not upload the deck.";
    const currency = encodeURIComponent(cfg.CURRENCY || "USD");
    const client = encodeURIComponent(cfg.PAYPAL_CLIENT_ID);
    await loadScript("https://www.paypal.com/sdk/js?client-id=" + client + "&currency=" + currency + "&intent=capture");
    if (window.paypal && window.paypal.Buttons) {
      window.paypal.Buttons({
        style: { layout: "vertical", shape: "rect", label: "pay" },
        createOrder: function (data, actions) {
          return actions.order.create({
            purchase_units: [
              {
                description: "SlideFit Pro — one deck",
                amount: { currency_code: cfg.CURRENCY || "USD", value: cfg.PRICE_DECK || "5.00" },
              },
            ],
          });
        },
        onApprove: function (data, actions) {
          return actions.order.capture().then(unlockDeck);
        },
      }).render("#paypal-once");
    }
    await loadScript(
      "https://www.paypal.com/sdk/js?client-id=" + client + "&vault=true&intent=subscription&currency=" + currency,
      "paypalSubscription",
    );
    const sub = window.paypalSubscription;
    if (sub && sub.Buttons) {
      sub.Buttons({
        style: { layout: "vertical", shape: "rect", label: "subscribe" },
        createSubscription: function (data, actions) {
          return actions.subscription.create({ plan_id: cfg.PLAN_ID });
        },
        onApprove: function () {
          unlockMonthly();
        },
      }).render("#paypal-month");
    }
  }

  function init() {
    bindDrop("drop-deck", "file-deck", "deck");
    bindDrop("drop-template", "file-template", "template");
    el("load-sample-deck").addEventListener("click", () => loadSample("deck").catch((err) => showError(err.message)));
    el("load-sample-template").addEventListener("click", () => loadSample("template").catch((err) => showError(err.message)));
    el("fit-button").addEventListener("click", fit);
    el("download").addEventListener("click", download);
    el("sim-deck").addEventListener("click", unlockDeck);
    el("sim-month").addEventListener("click", unlockMonthly);
    el("forget").addEventListener("click", forget);
    refreshUnlockNote();
    setupPayPal().catch((err) => {
      el("paypal-note").textContent = err.message;
    });
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
