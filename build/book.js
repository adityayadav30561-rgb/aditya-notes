// Builds the two-column "book" edition of a phase from book/<phase>/*.md.
// Markup (one block per paragraph, blank-line separated):
//   # Chapter title            starts a chapter (banner, then two columns); "~ subtitle" line optional
//   ## Topic heading           topic heading;  ### Sub-heading  run-in sub-heading
//   plain text                 paragraph; **bold**, *italic* allowed
//   - item                     bullet;  > text  shaded call-out block (first **bold** run is its label)
//   | a | b |                  table (first row = header, |---| row ignored)
//   [[diagram:name|col]]       figure from book/diagrams/name.png, one column wide (or |full)
//   [[full]] ... [[/full]]     everything inside goes on its own single-column page(s)
//   [[wide]] ... [[/wide]]     full-width block in the middle of a page (tables, small figures)
//   [[colbreak]]               column break
// Usage: node build/book.js [phase1]
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, Footer, PageNumber, BorderStyle,
  Table, TableRow, TableCell, WidthType, ShadingType, SectionType, ImageRun, LevelFormat,
  VerticalAlign, PageBreak,
} = require("docx");

const ROOT = path.join(__dirname, "..");
const BOOKS = {
  phase1: {
    dir: "book/phase1", out: "SSB_Service_Knowledge_Phase1.docx",
    title: "Service Knowledge", subtitle: "Phase 1 – The Indian Army: Organisation",
    foot: "SSB Service Knowledge · Phase 1",
  },
};

const BODY = "Cambria", HEAD = "Calibri";
const NAVY = "1F3864", BLUE = "2E5597", PALE = "EAF0FA", GREY = "595959";
const PAGE_W = 11906, PAGE_H = 16838, MARGIN = 800, GAP = 400;
const FULL_W = PAGE_W - 2 * MARGIN;
const COL_W = Math.floor((FULL_W - GAP) / 2);
const px = (twips) => Math.round(twips / 15); // 96 dpi
const BODY_SIZE = 19; // 9.5 pt

// ---------- inline markup ----------
function runs(text, base = {}) {
  const out = [];
  for (const part of text.split(/(\*\*[^*]+\*\*|\*[^*]+\*)/g)) {
    if (!part) continue;
    if (part.startsWith("**")) out.push(new TextRun({ ...base, text: part.slice(2, -2), bold: true }));
    else if (part.startsWith("*")) out.push(new TextRun({ ...base, text: part.slice(1, -1), italics: true }));
    else out.push(new TextRun({ ...base, text: part }));
  }
  return out;
}

// ---------- parsing ----------
function parse(dir) {
  const files = fs.readdirSync(path.join(ROOT, dir)).filter((f) => f.endsWith(".md")).sort();
  const chapters = [];
  let ch = null, full = null, para = [], list = [], quote = [], table = [];
  const target = () => (full ? full : ch.blocks);
  const flush = () => {
    if (para.length) target().push({ t: "p", text: para.join(" ") });
    if (list.length) target().push({ t: "ul", items: list });
    if (quote.length) target().push({ t: "quote", text: quote.join(" ") });
    if (table.length) target().push({ t: "table", rows: table });
    para = []; list = []; quote = []; table = [];
  };
  for (const file of files) {
    for (const raw of fs.readFileSync(path.join(ROOT, dir, file), "utf8").split(/\r?\n/)) {
      const line = raw.trimEnd();
      let m;
      if (!line.trim()) { flush(); continue; }
      if (line.startsWith("# ")) {
        flush();
        ch = { title: line.slice(2).trim(), sub: "", blocks: [] };
        chapters.push(ch);
      } else if (line.startsWith("~ ")) ch.sub = line.slice(2).trim();
      else if (line === "[[full]]" || line === "[[wide]]") {
        flush(); full = []; ch.blocks.push({ t: line === "[[full]]" ? "full" : "wide", blocks: full });
      } else if (line === "[[/full]]" || line === "[[/wide]]") { flush(); full = null; }
      else if (line === "[[colbreak]]") { flush(); target().push({ t: "colbreak" }); }
      else if ((m = line.match(/^\[\[diagram:([\w-]+)(?:\|(col|full))?\]\]$/))) {
        flush(); target().push({ t: "img", name: m[1], size: m[2] || (full ? "full" : "col") });
      } else if (line.startsWith("### ")) { flush(); target().push({ t: "h3", text: line.slice(4) }); }
      else if (line.startsWith("## ")) { flush(); target().push({ t: "h2", text: line.slice(3) }); }
      else if (line.startsWith("- ")) { if (para.length) flush(); list.push(line.slice(2)); }
      else if (line.startsWith("> ")) { if (para.length || list.length) flush(); quote.push(line.slice(2)); }
      else if (line.startsWith("|")) {
        if (para.length || list.length) flush();
        if (!/^\|[\s:|-]+\|?$/.test(line)) table.push(line.replace(/^\||\|$/g, "").split("|").map((c) => c.trim()));
      } else { if (list.length || quote.length || table.length) flush(); para.push(line.trim()); }
    }
    flush();
  }
  return chapters;
}

// ---------- block builders ----------
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

function h2(text) {
  return new Paragraph({
    keepNext: true, keepLines: true, spacing: { before: 200, after: 70 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 1 } },
    children: [new TextRun({ text, font: HEAD, bold: true, size: 25, color: NAVY })],
  });
}
const h3 = (text) => new Paragraph({
  keepNext: true, keepLines: true, spacing: { before: 100, after: 30 },
  children: [new TextRun({ text, font: HEAD, bold: true, size: 21, color: BLUE })],
});
const body = (text) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED, spacing: { after: 90, line: 262 },
  children: runs(text, { font: BODY, size: BODY_SIZE }),
});
const bullet = (text) => new Paragraph({
  numbering: { reference: "bullets", level: 0 }, alignment: AlignmentType.LEFT,
  spacing: { after: 40, line: 255 }, children: runs(text, { font: BODY, size: BODY_SIZE }),
});
const quote = (text) => new Paragraph({
  keepLines: true, spacing: { before: 60, after: 110, line: 255 },
  indent: { left: 120, right: 60 },
  shading: { type: ShadingType.CLEAR, color: "auto", fill: PALE },
  border: { left: { style: BorderStyle.SINGLE, size: 24, color: NAVY, space: 6 } },
  children: runs(text, { font: BODY, size: BODY_SIZE - 1 }),
});

function table(rows, width) {
  const n = Math.max(...rows.map((r) => r.length));
  const longestWord = (c) => Math.max(...rows.map((r) => Math.max(0, ...(r[c] || "").split(/\s+/).map((w) => w.length))));
  const weight = Array.from({ length: n }, (_, c) =>
    Math.max(8, longestWord(c) * 1.7, Math.pow(Math.max(...rows.map((r) => (r[c] || "").length)), 0.72)));
  const sum = weight.reduce((a, b) => a + b, 0);
  const widths = weight.map((w) => Math.floor((w / sum) * width));
  widths[n - 1] += width - widths.reduce((a, b) => a + b, 0);
  const line = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
  const borders = { top: line, bottom: line, left: line, right: line };
  const size = width > COL_W + 100 ? 18 : 16;
  return new Table({
    width: { size: width, type: WidthType.DXA }, columnWidths: widths,
    rows: rows.map((r, i) => new TableRow({
      tableHeader: i === 0, cantSplit: true,
      children: widths.map((w, c) => new TableCell({
        width: { size: w, type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.TOP,
        margins: { top: 30, bottom: 30, left: 60, right: 60 },
        shading: i === 0 ? { type: ShadingType.CLEAR, color: "auto", fill: NAVY }
          : i % 2 === 0 ? { type: ShadingType.CLEAR, color: "auto", fill: "F3F6FB" } : undefined,
        children: [new Paragraph({
          spacing: { after: 0 },
          children: runs(r[c] || "", i === 0
            ? { font: HEAD, size, bold: true, color: "FFFFFF" }
            : { font: c === 0 ? HEAD : BODY, size, bold: c === 0 }),
        })],
      })),
    })),
  });
}

function image(name, size) {
  const png = path.join(ROOT, "book", "diagrams", `${name}.png`);
  if (!fs.existsSync(png)) {
    console.warn(`WARNING: missing diagram ${name}.png (run node build/render_diagrams.js)`);
    return new Paragraph({ children: [new TextRun({ text: `[diagram ${name} missing]`, color: "B03A2E" })] });
  }
  const svg = fs.readFileSync(path.join(ROOT, "book", "diagrams", `${name}.svg`), "utf8");
  const vb = svg.match(/viewBox="0 0 ([\d.]+) ([\d.]+)"/);
  const ratio = +vb[2] / +vb[1];
  let w = px(size === "full" ? FULL_W : COL_W);
  let h = Math.round(w * ratio);
  const maxH = px(PAGE_H - 2 * MARGIN) - 40;
  if (h > maxH) { h = maxH; w = Math.round(h / ratio); }
  return new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 80, after: 120 }, keepLines: true,
    children: [new ImageRun({
      type: "png", data: fs.readFileSync(png), transformation: { width: w, height: h },
      altText: { title: name, description: `Diagram: ${name}`, name },
    })],
  });
}

function render(blocks, width) {
  const out = [];
  for (const b of blocks) {
    if (b.t === "h2") out.push(h2(b.text));
    else if (b.t === "h3") out.push(h3(b.text));
    else if (b.t === "p") out.push(body(b.text));
    else if (b.t === "ul") b.items.forEach((i) => out.push(bullet(i)));
    else if (b.t === "quote") out.push(quote(b.text));
    else if (b.t === "table") { out.push(table(b.rows, width)); out.push(new Paragraph({ spacing: { before: 0, after: 120 }, children: [new TextRun({ text: "", size: 8 })] })); }
    else if (b.t === "img") out.push(image(b.name, b.size));
    else if (b.t === "colbreak") out.push(new Paragraph({ children: [new PageBreak()] }));
  }
  return out;
}

// ---------- document ----------
async function build(cfg) {
  const chapters = parse(cfg.dir);
  const footer = () => ({ default: new Footer({ children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({ text: `${cfg.foot}   ·   `, font: HEAD, size: 16, color: "8C8C8C" }),
      new TextRun({ children: [PageNumber.CURRENT], font: HEAD, size: 16, color: "8C8C8C" }),
    ],
  })] }) });
  const page = { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, bottom: 900, left: MARGIN, right: MARGIN, footer: 400 } };
  const one = (type, children) => ({ properties: { type, page }, footers: footer(), children });
  const two = (type, children) => ({ properties: { type, page, column: { count: 2, space: GAP, separate: true } }, footers: footer(), children });

  const sections = [];
  // cover
  const cover = [
    new Paragraph({ spacing: { before: 2600, after: 120 }, children: [new TextRun({ text: "SSB INTERVIEW", font: HEAD, size: 28, bold: true, color: BLUE, characterSpacing: 120 })] }),
    new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: cfg.title, font: HEAD, size: 84, bold: true, color: NAVY })] }),
    new Paragraph({
      spacing: { after: 500 }, border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: NAVY, space: 8 } },
      children: [new TextRun({ text: cfg.subtitle, font: HEAD, size: 36, color: GREY })],
    }),
    new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: "Contents", font: HEAD, size: 28, bold: true, color: NAVY })] }),
    ...chapters.map((c, i) => new Paragraph({
      spacing: { after: 70 }, indent: { left: 200 },
      children: [
        new TextRun({ text: `${i + 1}.  `, font: HEAD, size: 24, bold: true, color: BLUE }),
        new TextRun({ text: c.title, font: HEAD, size: 24 }),
      ],
    })),
  ];
  sections.push(one(SectionType.NEXT_PAGE, cover));

  chapters.forEach((c, i) => {
    const banner = [
      new Paragraph({
        spacing: { before: 0, after: 0 }, shading: { type: ShadingType.CLEAR, color: "auto", fill: NAVY },
        indent: { left: 140, right: 140 },
        children: [new TextRun({ text: `CHAPTER ${i + 1}`, font: HEAD, size: 18, bold: true, color: "BDD0F0", characterSpacing: 80 })],
      }),
      new Paragraph({
        spacing: { before: 0, after: c.sub ? 0 : 200 }, shading: { type: ShadingType.CLEAR, color: "auto", fill: NAVY },
        indent: { left: 140, right: 140 }, keepNext: true,
        children: [new TextRun({ text: c.title, font: HEAD, size: 44, bold: true, color: "FFFFFF" })],
      }),
    ];
    if (c.sub) banner.push(new Paragraph({
      spacing: { after: 200 }, shading: { type: ShadingType.CLEAR, color: "auto", fill: NAVY },
      indent: { left: 140, right: 140 },
      children: [new TextRun({ text: c.sub, font: HEAD, size: 22, italics: true, color: "DCE3F0" })],
    }));
    sections.push(one(SectionType.NEXT_PAGE, banner));

    // split the chapter at [[full]] pages
    let text = [], cont = true; // cont: next section may start on the same page
    const pushText = () => {
      if (!text.length) return;
      sections.push(two(cont ? SectionType.CONTINUOUS : SectionType.NEXT_PAGE, render(text, COL_W)));
      cont = false; text = [];
    };
    for (const b of c.blocks) {
      if (b.t === "full" || b.t === "wide") {
        pushText();
        const wide = b.t === "wide";
        sections.push(one(wide ? SectionType.CONTINUOUS : SectionType.NEXT_PAGE, render(b.blocks, FULL_W)));
        cont = wide;
      } else text.push(b);
    }
    pushText();
  });

  const doc = new Document({
    title: `${cfg.title} – ${cfg.subtitle}`,
    styles: { default: { document: { run: { font: BODY, size: BODY_SIZE } } } },
    numbering: { config: [{ reference: "bullets", levels: [{
      level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 300, hanging: 200 } }, run: { color: BLUE } },
    }] }] },
    sections,
  });
  fs.writeFileSync(path.join(ROOT, cfg.out), await Packer.toBuffer(doc));
  console.log(`Wrote ${cfg.out} (${chapters.length} chapters)`);
}

(async () => {
  const names = process.argv.slice(2);
  for (const k of names.length ? names : Object.keys(BOOKS)) await build(BOOKS[k]);
})();
