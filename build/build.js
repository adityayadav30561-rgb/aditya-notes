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
  PageNumber, BorderStyle, Table, TableRow, TableCell, WidthType, ShadingType,
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

// Lecturette notes: each "## Topic" starts a new page with a navy banner, then a
// table of its flow (# | Stage | Flow). Each "- Stage: a → b → c" line is one row,
// and its points are listed one per line so the flow can be read at a glance.
const LECTURETTES = { src: "lecturettes.md", out: "SSB_Lecturettes.docx", title: "SSB – Lecturette Preparation" };
const L_SIZE = 21; // 10.5 pt
const L_COLS = [520, 2200, CONTENT_W - 520 - 2200]; // # | Stage | Points

function lCell(w, runs, opts = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders,
    verticalAlign: "center",
    margins: { top: 70, bottom: 70, left: 100, right: 100 },
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
        new TextRun({ text: "→\t", bold: true, color: "C00000", size: L_SIZE }),
        ...(label ? [new TextRun({ text: label, bold: true, size: L_SIZE })] : []),
        new TextRun({ text: part.slice(label.length), size: L_SIZE }),
      ],
      tabStops: [{ type: "left", position: 260 }],
    });
  });
}

async function buildLecturettes({ src, out, title }) {
  const topics = parse(path.join(ROOT, src));
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
    const rows = [new TableRow({ tableHeader: true, children: [
      lCell(L_COLS[0], [new TextRun({ text: "#", bold: true, size: L_SIZE })], { fill: "DCE3F0", align: AlignmentType.CENTER }),
      lCell(L_COLS[1], [new TextRun({ text: "Stage", bold: true, size: L_SIZE })], { fill: "DCE3F0" }),
      lCell(L_COLS[2], [new TextRun({ text: "Flow", bold: true, size: L_SIZE })], { fill: "DCE3F0" }),
    ] })];
    t.questions.forEach((line, i) => {
      const colon = line.indexOf(":");
      const stage = colon > 0 ? line.slice(0, colon).trim() : "";
      const points = colon > 0 ? line.slice(colon + 1).trim() : line;
      const fill = i % 2 ? "F5F7FB" : undefined;
      rows.push(new TableRow({ cantSplit: true, children: [
        lCell(L_COLS[0], [new TextRun({ text: String(i + 1), color: "7F7F7F", size: L_SIZE })], { align: AlignmentType.CENTER, fill }),
        lCell(L_COLS[1], [new TextRun({ text: stage, bold: true, color: NAVY, size: L_SIZE })], { fill }),
        lCell(L_COLS[2], flowParas(points), { fill }),
      ] }));
    });
    children.push(new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: L_COLS, rows }));
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
