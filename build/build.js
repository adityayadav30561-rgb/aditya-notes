// Builds SSB_Master_Notes.docx from the Markdown files in topics/.
// Usage: node build/build.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  ShadingType, BorderStyle, AlignmentType, LevelFormat, HeadingLevel, Header, Footer,
  PageNumber, Bookmark, InternalHyperlink, TabStopType,
} = require("docx");

const ROOT = path.join(__dirname, "..");
const TOPICS_DIR = path.join(ROOT, "topics");
const OUT = path.join(ROOT, "SSB_Master_Notes.docx");

// Categories appear in this order; any other category is appended alphabetically.
const CATEGORY_ORDER = [
  "International Organisations & Groupings",
  "International Relations & Geopolitics",
  "India: Neighbourhood & Bilateral Ties",
  "Defence & Security",
  "Economy",
  "Polity & Governance",
  "Science & Technology",
  "Environment & Climate",
  "Social Issues",
  "Sports & Awards",
  "SSB Personal & Psychology",
];

const FONT = "Calibri";
const BODY = 19; // half-points (9.5 pt)
const SMALL = 17;
const NAVY = "1F3864";
const ACCENT = "C55A11";
const PAGE_W = 11906, MARGIN = 680; // A4, 0.47"
const CONTENT_W = PAGE_W - 2 * MARGIN;

// ---------- parsing ----------
function parseTopic(file) {
  const raw = fs.readFileSync(file, "utf8").replace(/\r\n/g, "\n");
  const m = raw.match(/^---\n([\s\S]*?)\n---\n([\s\S]*)$/);
  if (!m) throw new Error(`${file}: missing front matter`);
  const meta = {};
  for (const line of m[1].split("\n")) {
    const i = line.indexOf(":");
    if (i > 0) meta[line.slice(0, i).trim()] = line.slice(i + 1).trim();
  }
  for (const k of ["title", "category", "updated"]) {
    if (!meta[k]) throw new Error(`${file}: front matter needs "${k}"`);
  }
  const slug = path.basename(file, ".md").replace(/[^A-Za-z0-9]/g, "_");
  return { ...meta, slug, body: m[2] };
}

function inline(text, base = {}) {
  // **bold** → bold runs
  return text.split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((part) =>
    part.startsWith("**") && part.endsWith("**")
      ? new TextRun({ ...base, text: part.slice(2, -2), bold: true })
      : new TextRun({ ...base, text: part }));
}

const cellsOf = (line) => line.trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map((c) => c.trim());

function table(rows) {
  const cols = rows[0].length;
  const weights = Array.from({ length: cols }, (_, c) =>
    Math.min(60, Math.max(8, ...rows.map((r) => (r[c] || "").replace(/\*\*/g, "").length))));
  const total = weights.reduce((a, b) => a + b, 0);
  const widths = weights.map((w) => Math.floor((w / total) * CONTENT_W));
  widths[cols - 1] += CONTENT_W - widths.reduce((a, b) => a + b, 0);
  const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((r, ri) => new TableRow({
      tableHeader: ri === 0,
      children: widths.map((w, ci) => new TableCell({
        width: { size: w, type: WidthType.DXA },
        borders,
        margins: { top: 20, bottom: 20, left: 70, right: 70 },
        shading: ri === 0 ? { type: ShadingType.CLEAR, color: "auto", fill: "DCE3F0" } : undefined,
        children: [new Paragraph({
          spacing: { before: 0, after: 0 },
          children: inline(r[ci] || "", { size: SMALL, bold: ri === 0 || undefined, font: FONT }),
        })],
      })),
    })),
  });
}

function renderBody(body) {
  const out = [];
  const lines = body.split("\n");
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (!line.trim()) continue;
    if (line.startsWith("## ")) {
      out.push(new Paragraph({ heading: HeadingLevel.HEADING_3, children: inline(line.slice(3)) }));
    } else if (/^\s*- /.test(line)) {
      const level = line.match(/^\s*/)[0].length >= 2 ? 1 : 0;
      out.push(new Paragraph({
        numbering: { reference: "bullets", level },
        spacing: { before: 0, after: 10 },
        children: inline(line.replace(/^\s*- /, "")),
      }));
    } else if (line.trim().startsWith("|")) {
      const rows = [];
      while (i < lines.length && lines[i].trim().startsWith("|")) {
        if (!/^\s*\|[\s:|-]+\|\s*$/.test(lines[i])) rows.push(cellsOf(lines[i]));
        i++;
      }
      i--;
      out.push(table(rows));
    } else {
      out.push(new Paragraph({ spacing: { before: 0, after: 20 }, children: inline(line.trim()) }));
    }
  }
  return out;
}

// ---------- assemble ----------
const topics = fs.readdirSync(TOPICS_DIR).filter((f) => f.endsWith(".md"))
  .map((f) => parseTopic(path.join(TOPICS_DIR, f)));

const catRank = (c) => { const i = CATEGORY_ORDER.indexOf(c); return i < 0 ? 999 : i; };
topics.sort((a, b) => catRank(a.category) - catRank(b.category)
  || a.category.localeCompare(b.category) || a.title.localeCompare(b.title));

const categories = [...new Set(topics.map((t) => t.category))];
const today = new Date().toISOString().slice(0, 10);

const children = [
  new Paragraph({
    spacing: { after: 40 },
    children: [new TextRun({ text: "SSB Interview – Master Notes", bold: true, size: 36, color: NAVY, font: FONT })],
  }),
  new Paragraph({
    spacing: { after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 4 } },
    children: [new TextRun({
      text: `${topics.length} topic${topics.length === 1 ? "" : "s"} · built ${today} · bold = must-know · ★ = high priority · read: Facts → Analysis → "So what for India?"`,
      italics: true, size: SMALL, color: "595959", font: FONT,
    })],
  }),
  new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("Index")] }),
];
// Index table: topic names link to the topic's bookmark.
{
  const header = ["#", "Topic", "Category", "Updated"];
  const widths = [500, 4200, 4146, 1700];
  const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const cell = (w, kids, head) => new TableCell({
    width: { size: w, type: WidthType.DXA }, borders,
    margins: { top: 20, bottom: 20, left: 70, right: 70 },
    shading: head ? { type: ShadingType.CLEAR, color: "auto", fill: "DCE3F0" } : undefined,
    children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: kids })],
  });
  const txt = (t, head) => new TextRun({ text: t, size: SMALL, bold: head || undefined, font: FONT });
  children.push(new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: header.map((h, i) => cell(widths[i], [txt(h, true)], true)) }),
      ...topics.map((t, i) => new TableRow({ children: [
        cell(widths[0], [txt(String(i + 1))]),
        cell(widths[1], [new InternalHyperlink({ anchor: t.slug, children: [
          new TextRun({ text: t.title, size: SMALL, font: FONT, color: "0563C1", underline: {} })] })]),
        cell(widths[2], [txt(t.category)]),
        cell(widths[3], [txt(t.updated)]),
      ] })),
    ],
  }));
}

for (const cat of categories) {
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(cat)] }));
  for (const t of topics.filter((x) => x.category === cat)) {
    children.push(new Paragraph({
      heading: HeadingLevel.HEADING_2,
      tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
      children: [
        new Bookmark({ id: t.slug, children: [new TextRun(t.title)] }),
        new TextRun({ text: `\tupdated ${t.updated}`, size: SMALL, bold: false, italics: true, color: "7F7F7F" }),
      ],
    }));
    children.push(...renderBody(t.body));
  }
}

const doc = new Document({
  creator: "SSB Notes",
  title: "SSB Interview – Master Notes",
  styles: {
    default: { document: { run: { font: FONT, size: BODY }, paragraph: { spacing: { line: 240 } } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, color: "FFFFFF", font: FONT },
        paragraph: { spacing: { before: 240, after: 60 }, outlineLevel: 0, keepNext: true,
          shading: { type: ShadingType.CLEAR, color: "auto", fill: NAVY } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, color: NAVY, font: FONT },
        paragraph: { spacing: { before: 160, after: 40 }, outlineLevel: 1, keepNext: true,
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 2 } } } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 20, bold: true, color: ACCENT, font: FONT },
        paragraph: { spacing: { before: 80, after: 20 }, outlineLevel: 2, keepNext: true } },
    ],
  },
  numbering: { config: [{ reference: "bullets", levels: [
    { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 230, hanging: 170 } } } },
    { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 460, hanging: 170 } } } },
  ] }] },
  sections: [{
    properties: { page: {
      size: { width: PAGE_W, height: 16838 },
      margin: { top: 620, bottom: 620, left: MARGIN, right: MARGIN, header: 300, footer: 300 },
    } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: "SSB Master Notes", size: 14, color: "A6A6A6" })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], size: 14, color: "A6A6A6" })] })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log(`Wrote ${path.relative(ROOT, OUT)} (${topics.length} topics)`);
});
