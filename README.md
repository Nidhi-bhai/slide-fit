# SlideFit

Apply a company PowerPoint template to a deck you already have. SlideFit runs in the browser: drop a `.pptx` and a template `.pptx`, preview the mapped slides, and download a new file. Nothing is uploaded.

Free conversions keep the first 5 slides and add a watermark slide. $5 unlocks unlimited conversions for 24 hours in this browser, not tied to the file bytes. $12 is a one-time 30-day pass. Neither auto-renews. The payment happens on PayPal. SlideFit stores only the order id, in this browser, and there is no license server. Restore accepts a 17-character PayPal order id. A saved purchase keeps its own window. An unknown id gets 24 hours, and a full 30-day restore needs the same browser or an email to nidhi69n@gmail.com.

Built by [Nidhi Karuturi](https://github.com/Nidhi-bhai).

## Try it

Open `index.html` over HTTP (GitHub Pages, or `python3 -m http.server`). Then either drop your own files or use the two samples:

- `samples/northwind-quarterly.pptx` — a 7-slide content deck (title, section, bullets, picture, table, and two further slides)
- `samples/harbor-company-template.pptx` — a widescreen company template (Harbor colours, Georgia / Verdana, master bar)

The free result is six slides (the first 5 plus the watermark). A 24-hour window or a 30-day pass converts all seven.

## What is preserved

Titles, subtitles, bullets, pictures, and tables are written into the closest template layout (Title Slide, Section Header, Title and Content, Picture with Caption). Fonts and colours follow the template theme. Inline formatting (bold, colour, links) is not kept. Charts, SmartArt, notes, and animations are not rebuilt.

## Develop

```bash
npm install
npm test
python3 scripts/generate_samples.py   # rebuild the sample decks
```

Tests cover classification, layout choice, the free/pro cutoff, and a full conversion of the Northwind sample onto the Harbor template. Opening that output in LibreOffice Impress (or PowerPoint) is the check that the package is a real presentation: the theme background, master mark, text, picture, and table should all be there.

Fraunces and Outfit are shipped under the SIL Open Font License. See `css/fonts/`.

## GitHub Pages

The site is a static root (`index.html`, no Jekyll). After this branch is on the remote, enable Pages for the branch you want to publish, folder `/`. The project URL is:

https://nidhi-bhai.github.io/slide-fit/
