// Renders every book/diagrams/*.svg to a crisp PNG next to it (3x scale).
// The book builder embeds the PNGs, so re-run this after editing an SVG.
// Usage: node build/render_diagrams.js [name ...]
const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");

const DIR = path.join(__dirname, "..", "book", "diagrams");

(async () => {
  const only = process.argv.slice(2);
  const files = fs.readdirSync(DIR).filter((f) => f.endsWith(".svg") &&
    (!only.length || only.includes(f.replace(/\.svg$/, ""))));
  const browser = await chromium.launch();
  for (const f of files) {
    const svg = fs.readFileSync(path.join(DIR, f), "utf8");
    const vb = svg.match(/viewBox="0 0 ([\d.]+) ([\d.]+)"/);
    const w = Math.ceil(vb ? +vb[1] : 1400), h = Math.ceil(vb ? +vb[2] : 1000);
    const page = await browser.newPage({ viewport: { width: w, height: h }, deviceScaleFactor: 3 });
    await page.setContent(`<body style="margin:0;background:#fff">${svg}</body>`);
    const out = path.join(DIR, f.replace(/\.svg$/, ".png"));
    await (await page.$("svg")).screenshot({ path: out });
    await page.close();
    console.log(`Rendered ${path.basename(out)} (${w}x${h})`);
  }
  await browser.close();
})();
