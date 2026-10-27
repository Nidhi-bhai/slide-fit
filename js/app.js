(function () {
  const cfg = window.SLIDEFIT_CONFIG || {};
  const KEYS = { purchases: "slidefit.purchases" };
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
    templateBuf: null,
    templateName: "",
    result: null,
    selected: 0,
  };

  const el = (id) => document.getElementById(id);

  function paypalConfigured() {
    const id = cfg.PAYPAL_CLIENT_ID || "";
    return Boolean(id) && !id.includes("YOUR_PAYPAL");
  }

  function loadPurchases() {
    try {
      const parsed = JSON.parse(localStorage.getItem(KEYS.purchases) || "[]");
      return Array.isArray(parsed) ? parsed : [];
    } catch (err) {
      return [];
    }
  }

  function storePurchases(list) {
    localStorage.setItem(KEYS.purchases, JSON.stringify(list));
  }

  function currentAccess() {
    return window.SlideFitPurchases.access(loadPurchases(), Date.now());
  }

  function productLabel(product) {
    if (product === "pass") return "$12 30-day pass";
    if (product === "day") return "24-hour pass";
    return "$5 for 24 hours";
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
    const access = currentAccess();
    const note = el("unlock-status");
    if (access.pro && access.reason === "pass") {
      note.textContent = "30-day pass is on in this browser until " + new Date(access.until).toLocaleString() + ". Order id " + access.orderId + ". It does not auto-renew. SlideFit stores only that order id, in this browser.";
    } else if (access.pro && access.reason === "day") {
      note.textContent = "24-hour pass is on until " + new Date(access.until).toLocaleString() + ". Order id " + access.orderId + ". This browser had not saved that id, so this is not a 30-day restore. A full restore of a 30-day pass needs the same browser, or email nidhi69n@gmail.com with the order id.";
    } else if (access.pro) {
      note.textContent = "$5 is on until " + new Date(access.until).toLocaleString() + ". Unlimited conversions for 24 hours in this browser. Order id " + access.orderId + ". It does not auto-renew.";
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

    const access = currentAccess();
    if (result.omitted > 0) {
      setStatus("Free plan kept the first 5 of " + (5 + result.omitted) + " slides and added a watermark. $5 covers unlimited conversions for 24 hours, and a $12 pass covers 30 days. Either leaves that slide out.");
    } else if (result.watermark) {
      setStatus("Free plan added a watermark slide after the first 5 slides. A pass leaves it out.");
    } else if (access.reason === "pass") {
      setStatus("30-day pass: " + result.slides.length + " slides, no watermark. This pass does not auto-renew.");
    } else if (access.reason === "day") {
      setStatus("24-hour pass: " + result.slides.length + " slides, no watermark. This is not a 30-day restore. It does not auto-renew.");
    } else {
      setStatus("24 hours: " + result.slides.length + " slides, no watermark. Unlimited conversions in this browser until the window ends. It does not auto-renew.");
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
      const access = currentAccess();
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

  function showReceipt(purchase, mode, until) {
    const box = el("order-receipt");
    box.hidden = false;
    const label = productLabel(purchase.product);
    const end = new Date(until).toLocaleString();
    if (mode === "expired") {
      box.textContent = "PayPal order id " + purchase.orderId + " is saved in this browser, and that " + label + " has ended. It does not auto-renew. Pay again if you want another window. SlideFit stores only the order id, and only here.";
      return;
    }
    if (mode === "paid") {
      box.textContent = "Paid on PayPal. Order id " + purchase.orderId + ". Save this id. " + label + " until " + end + ". It does not auto-renew. SlideFit stores only this order id, in this browser.";
      return;
    }
    if (mode === "stored") {
      box.textContent = "Restored PayPal order id " + purchase.orderId + " from the saved purchase time. " + label + " until " + end + ". It does not auto-renew. SlideFit stores only this order id, in this browser.";
      return;
    }
    box.textContent = "This browser has not saved PayPal order id " + purchase.orderId + ", so Restore granted a 24-hour pass until " + end + ". A full restore of a 30-day pass needs the same browser, or email nidhi69n@gmail.com with the order id. SlideFit has no license server and did not ask PayPal if the id is real.";
  }

  function grant(orderId, product, mode) {
    const saved = window.SlideFitPurchases.recordPurchase(loadPurchases(), {
      orderId: orderId,
      product: product,
      purchasedAt: Date.now(),
    });
    storePurchases(saved.purchases);
    const until = saved.purchase.purchasedAt + window.SlideFitPurchases.duration(saved.purchase.product);
    showReceipt(saved.purchase, mode || "paid", until);
    refreshUnlockNote();
    if (state.deckBuf && state.templateBuf && window.SlideFitPurchases.access(saved.purchases, Date.now()).pro) fit();
  }

  function restore(event) {
    event.preventDefault();
    showError("");
    const orderId = el("restore-order").value;
    let result;
    try {
      result = window.SlideFitPurchases.restorePurchase(loadPurchases(), { orderId: orderId }, Date.now());
    } catch (err) {
      showError(err.message);
      return;
    }
    storePurchases(result.purchases);
    if (!result.active) {
      showReceipt(result.purchase, "expired", result.until);
      refreshUnlockNote();
      return;
    }
    showReceipt(result.purchase, result.fromStoredTime ? "stored" : "fresh", result.until);
    refreshUnlockNote();
    if (state.deckBuf && state.templateBuf) fit();
  }

  function forget() {
    localStorage.removeItem(KEYS.purchases);
    const box = el("order-receipt");
    box.hidden = true;
    box.textContent = "";
    refreshUnlockNote();
    setStatus("Order ids stored in this browser were removed.");
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

  function orderButton(selector, product, value, description) {
    window.paypal.Buttons({
      style: { layout: "vertical", shape: "rect", label: "pay" },
      createOrder: function (data, actions) {
        return actions.order.create({
          purchase_units: [
            {
              description: description,
              amount: { currency_code: cfg.CURRENCY || "USD", value: value },
            },
          ],
        });
      },
      onApprove: function (data, actions) {
        return actions.order.capture().then(function (details) {
          const orderId = (details && details.id) || data.orderID;
          grant(orderId, product, "paid");
        });
      },
    }).render(selector);
  }

  async function setupPayPal() {
    const note = el("paypal-note");
    if (!paypalConfigured()) {
      el("simulate").hidden = false;
      note.textContent = "PayPal is not loaded on this copy of the page. The buttons below only preview a pass. They do not charge a card.";
      return;
    }
    note.textContent = "The payment happens on PayPal. SlideFit does not upload the deck. After you pay, this page shows the PayPal order id and stores only that id in this browser.";
    const currency = encodeURIComponent(cfg.CURRENCY || "USD");
    const client = encodeURIComponent(cfg.PAYPAL_CLIENT_ID);
    await loadScript("https://www.paypal.com/sdk/js?client-id=" + client + "&currency=" + currency + "&intent=capture");
    if (window.paypal && window.paypal.Buttons) {
      orderButton("#paypal-once", "deck", cfg.PRICE_DECK || "5.00", "SlideFit — unlimited conversions for 24 hours");
      orderButton("#paypal-pass", "pass", cfg.PRICE_PASS || "12.00", "SlideFit — 30-day pass, one time");
    }
  }

  function init() {
    bindDrop("drop-deck", "file-deck", "deck");
    bindDrop("drop-template", "file-template", "template");
    el("load-sample-deck").addEventListener("click", () => loadSample("deck").catch((err) => showError(err.message)));
    el("load-sample-template").addEventListener("click", () => loadSample("template").catch((err) => showError(err.message)));
    el("fit-button").addEventListener("click", fit);
    el("download").addEventListener("click", download);
    el("sim-deck").addEventListener("click", function () { grant("LOCAL-DECK-PREVIEW", "deck", "paid"); });
    el("sim-pass").addEventListener("click", function () { grant("LOCAL-PASS-PREVIEW", "pass", "paid"); });
    el("restore-form").addEventListener("submit", restore);
    el("forget").addEventListener("click", forget);
    refreshUnlockNote();
    setupPayPal().catch((err) => {
      el("paypal-note").textContent = err.message;
    });
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
