// Builds the question sheets listed in SHEETS, plus the lecturette notes, from
// their Markdown sources.
// Each source has "## Section" lines and "- question" lines. Each output is one
// compact table (# | Question | Answer): questions numbered continuously, each
// section as a shaded title row, and an empty Answer column to fill in.
// Usage: node build/build.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, Footer,
  PageNumber, BorderStyle, Tab, Table, TableRow, TableCell, WidthType, ShadingType,
} = require("docx");

const ROOT = path.join(__dirname, "..");
const SHEETS = [
  { src: "questions.md", out: "SSB_Question_Sheet.docx", title: "SSB Interview – Question Sheet" },
  { src: "service_knowledge.md", out: "SSB_Service_Knowledge.docx", title: "SSB Interview – Service Knowledge" },
];

const FONT = "Calibri";
const NAVY = "1F3864";
const SIZE = 18; // 9 pt

const PAGE_W = 11906, MARGIN = 600;
const CONTENT_W = PAGE_W - 2 * MARGIN;
const COLS = [520, 5800, CONTENT_W - 520 - 5800]; // # | Question | Answer

const line = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: line, bottom: line, left: line, right: line };
const cell = (w, text, opts = {}) => new TableCell({
  width: { size: w, type: WidthType.DXA },
  borders,
  margins: { top: 15, bottom: 15, left: 60, right: 60 },
  shading: opts.fill ? { type: ShadingType.CLEAR, color: "auto", fill: opts.fill } : undefined,
  children: [new Paragraph({
    alignment: opts.align,
    spacing: { before: 0, after: 0 },
    children: [new TextRun({ text, size: SIZE, bold: opts.bold, color: opts.color })],
  })],
});

function parse(file) {
  const sections = [];
  for (const line of fs.readFileSync(file, "utf8").split(/\r?\n/)) {
    if (line.startsWith("## ")) sections.push({ title: line.slice(3).trim(), questions: [] });
    else if (line.startsWith("- ")) {
      if (!sections.length) sections.push({ title: "Questions", questions: [] });
      sections[sections.length - 1].questions.push(line.slice(2).trim());
    }
  }
  return sections;
}

async function build({ src, out, title }) {
  const sections = parse(path.join(ROOT, src));

  // One continuous table: a header row (repeated on every page), then for each
  // section a shaded title row followed by its questions.
  let n = 0;
  const rows = [new TableRow({ tableHeader: true, children: [
    cell(COLS[0], "#", { bold: true, fill: "DCE3F0", align: AlignmentType.CENTER }),
    cell(COLS[1], "Question", { bold: true, fill: "DCE3F0" }),
    cell(COLS[2], "Answer", { bold: true, fill: "DCE3F0" }),
  ] })];
  for (const s of sections) {
    rows.push(new TableRow({ cantSplit: true, children: [new TableCell({
      columnSpan: 3,
      width: { size: CONTENT_W, type: WidthType.DXA },
      borders,
      margins: { top: 25, bottom: 25, left: 60, right: 60 },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: NAVY },
      children: [new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(s.title)] })],
    })] }));
    for (const q of s.questions) {
      rows.push(new TableRow({ cantSplit: true, children: [
        cell(COLS[0], String(++n), { align: AlignmentType.CENTER, color: "7F7F7F" }),
        cell(COLS[1], q),
        cell(COLS[2], ""),
      ] }));
    }
  }

  const doc = new Document({
    title,
    styles: {
      default: { document: { run: { font: FONT, size: SIZE } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 19, bold: true, color: "FFFFFF" },
          paragraph: { spacing: { before: 0, after: 0 }, outlineLevel: 0, keepNext: true } },
      ],
    },
    sections: [{
      properties: { page: { size: { width: PAGE_W, height: 16838 },
        margin: { top: 600, bottom: 600, left: MARGIN, right: MARGIN, footer: 300 } } },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ children: [PageNumber.CURRENT], size: 14, color: "A6A6A6" })] })] }) },
      children: [
        new Paragraph({
          spacing: { after: 60 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 2 } },
          children: [new TextRun({ text: title, bold: true, size: 28, color: NAVY })],
        }),
        new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: COLS, rows }),
      ],
    }],
  });

  fs.writeFileSync(path.join(ROOT, out), await Packer.toBuffer(doc));
  console.log(`Wrote ${out} (${n} questions)`);
}

// Lecturette notes. Each "## Topic" in lecturettes.md starts a new page and is shown
// in three ways:
//   1. Framework  – "- Stage: a → b → c" lines before any "###" heading, shown as a
//                   table (# | Stage | Points) with one point per line.
//   2. Understand it – plain paragraphs under "### Paragraph", to read, not memorise.
//   3. Flowchart  – "- Stage: key → key" lines under "### Flowchart", shown as boxes
//                   joined by down arrows: only the points to remember, in order.
const LECTURETTES = { src: "lecturettes.md", out: "SSB_Lecturettes.docx", title: "SSB – Lecturette Preparation" };
const L_SIZE = 21; // 10.5 pt
const RED = "C00000";
const L_COLS = [520, 2200, CONTENT_W - 520 - 2200]; // # | Stage | Points

function parseLecturettes(file) {
  const topics = [];
  let t, part;
  for (const raw of fs.readFileSync(file, "utf8").split(/\r?\n/)) {
    const line = raw.trim();
    if (raw.startsWith("## ")) {
      t = { title: raw.slice(3).trim(), framework: [], paragraphs: [], flowchart: [] };
      topics.push(t);
      part = "framework";
    } else if (!t) continue;
    else if (/^### +paragraph/i.test(line)) part = "paragraphs";
    else if (/^### +flowchart/i.test(line)) part = "flowchart";
    else if (part === "paragraphs") { if (line) t.paragraphs.push(line); }
    else if (line.startsWith("- ")) t[part].push(line.slice(2).trim());
  }
  return topics;
}

// "Stage: a → b" → ["Stage", "a → b"]
const splitStage = (line) => {
  const colon = line.indexOf(":");
  return colon > 0 ? [line.slice(0, colon).trim(), line.slice(colon + 1).trim()] : ["", line];
};

function lCell(w, runs, opts = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: opts.borders || borders,
    verticalAlign: "center",
    margins: opts.margins || { top: 70, bottom: 70, left: 100, right: 100 },
    shading: opts.fill ? { type: ShadingType.CLEAR, color: "auto", fill: opts.fill } : undefined,
    children: runs[0] instanceof Paragraph ? runs
      : [new Paragraph({ alignment: opts.align, spacing: { before: 0, after: 0, line: 300 }, children: runs })],
  });
}

// One line per point, each led by a red arrow. A short "Label:" at the start of a
// point (e.g. "Security:") is set in bold.
function flowParas(text) {
  return text.split(/\s*→\s*/).map((part, i) => {
    const colon = part.indexOf(":");
    const label = colon > 0 && colon <= 25 ? part.slice(0, colon + 1) : "";
    return new Paragraph({
      spacing: { before: i ? 40 : 0, after: 0, line: 280 },
      indent: { left: 260, hanging: 260 },
      children: [
        new TextRun({ text: "→", bold: true, color: RED, size: L_SIZE }), new TextRun({ children: [new Tab()] }),
        ...(label ? [new TextRun({ text: label, bold: true, size: L_SIZE })] : []),
        new TextRun({ text: part.slice(label.length), size: L_SIZE }),
      ],
      tabStops: [{ type: "left", position: 260 }],
    });
  });
}

const partHeading = (text, first) => new Paragraph({
  keepNext: true,
  spacing: { before: first ? 0 : 280, after: 100 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 1 } },
  children: [new TextRun({ text, bold: true, size: 24, color: NAVY })],
});

function frameworkTable(lines) {
  const rows = [new TableRow({ tableHeader: true, children: [
    lCell(L_COLS[0], [new TextRun({ text: "#", bold: true, size: L_SIZE })], { fill: "DCE3F0", align: AlignmentType.CENTER }),
    lCell(L_COLS[1], [new TextRun({ text: "Stage", bold: true, size: L_SIZE })], { fill: "DCE3F0" }),
    lCell(L_COLS[2], [new TextRun({ text: "Points", bold: true, size: L_SIZE })], { fill: "DCE3F0" }),
  ] })];
  lines.forEach((line, i) => {
    const [stage, points] = splitStage(line);
    const fill = i % 2 ? "F5F7FB" : undefined;
    rows.push(new TableRow({ cantSplit: true, children: [
      lCell(L_COLS[0], [new TextRun({ text: String(i + 1), color: "7F7F7F", size: L_SIZE })], { align: AlignmentType.CENTER, fill }),
      lCell(L_COLS[1], [new TextRun({ text: stage, bold: true, color: NAVY, size: L_SIZE })], { fill }),
      lCell(L_COLS[2], flowParas(points), { fill }),
    ] }));
  });
  return new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: L_COLS, rows });
}

// Each stage is a box: stage name on the left, its key points joined by red arrows
// on the right. A red ▼ between boxes shows the order.
function flowchart(lines) {
  const BOX_W = Math.round(CONTENT_W * 0.86), LABEL_W = 1900;
  const edge = { style: BorderStyle.SINGLE, size: 8, color: NAVY };
  const boxBorders = { top: edge, bottom: edge, left: edge, right: edge };
  const out = [];
  lines.forEach((line, i) => {
    const [stage, points] = splitStage(line);
    if (i) out.push(new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 0, after: 0 },
      children: [new TextRun({ text: "▼", color: RED, size: 22 })] }));
    const runs = [];
    points.split(/\s*→\s*/).forEach((p, j) => {
      if (j) runs.push(new TextRun({ text: "  →  ", bold: true, color: RED, size: L_SIZE }));
      runs.push(new TextRun({ text: p, size: L_SIZE }));
    });
    out.push(new Table({
      alignment: AlignmentType.CENTER,
      width: { size: BOX_W, type: WidthType.DXA },
      columnWidths: [LABEL_W, BOX_W - LABEL_W],
      rows: [new TableRow({ cantSplit: true, children: [
        lCell(LABEL_W, [new TextRun({ text: stage, bold: true, color: "FFFFFF", size: L_SIZE })],
          { fill: NAVY, borders: boxBorders, align: AlignmentType.CENTER }),
        lCell(BOX_W - LABEL_W, runs, { fill: "EEF2F9", borders: boxBorders,
          margins: { top: 90, bottom: 90, left: 140, right: 140 } }),
      ] })],
    }));
  });
  return out;
}

async function buildLecturettes({ src, out, title }) {
  const topics = parseLecturettes(path.join(ROOT, src));
  const children = [new Paragraph({
    spacing: { after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 2 } },
    children: [new TextRun({ text: title, bold: true, size: 32, color: NAVY })],
  })];

  topics.forEach((t, ti) => {
    children.push(new Paragraph({
      heading: HeadingLevel.HEADING_1,
      pageBreakBefore: ti > 0,
      shading: { type: ShadingType.CLEAR, color: "auto", fill: NAVY },
      spacing: { before: ti > 0 ? 0 : 120, after: 160 },
      indent: { left: 100, right: 100 },
      children: [new TextRun({ text: `Topic ${ti + 1}: ${t.title}` })],
    }));
    children.push(partHeading("1. Framework", true), frameworkTable(t.framework));
    if (t.paragraphs.length) {
      children.push(partHeading("2. Understand it"));
      t.paragraphs.forEach((p) => children.push(new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { before: 0, after: 120, line: 300 },
        children: [new TextRun({ text: p, size: 22 })],
      })));
    }
    if (t.flowchart.length) {
      children.push(partHeading("3. Flowchart: points to remember"), ...flowchart(t.flowchart));
    }
  });

  const doc = new Document({
    title,
    styles: {
      default: { document: { run: { font: FONT, size: L_SIZE } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 26, bold: true, color: "FFFFFF" },
          paragraph: { outlineLevel: 0, keepNext: true } },
      ],
    },
    sections: [{
      properties: { page: { size: { width: PAGE_W, height: 16838 },
        margin: { top: 800, bottom: 800, left: MARGIN, right: MARGIN, footer: 300 } } },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ children: [PageNumber.CURRENT], size: 14, color: "A6A6A6" })] })] }) },
      children,
    }],
  });

  fs.writeFileSync(path.join(ROOT, out), await Packer.toBuffer(doc));
  console.log(`Wrote ${out} (${topics.length} topics)`);
}

(async () => {
  for (const sheet of SHEETS) await build(sheet);
  await buildLecturettes(LECTURETTES);
})();
