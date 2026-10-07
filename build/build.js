// Builds the question sheets listed in SHEETS from their Markdown sources.
// Each source has "## Section" lines and "- question" lines; a line may carry an
// answer after " :: ". Each output is one compact, print-friendly table
// (# | Question | Answer): questions numbered continuously and each section as a
// shaded title row. Questions without an answer get an empty Answer cell.
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
const SIZE = 14; // 7 pt

const PAGE_W = 11906, MARGIN = 420;
const CONTENT_W = PAGE_W - 2 * MARGIN;
const COLS = [420, 4300, CONTENT_W - 420 - 4300]; // # | Question | Answer

const line = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: line, bottom: line, left: line, right: line };
const cell = (w, text, opts = {}) => new TableCell({
  width: { size: w, type: WidthType.DXA },
  borders,
  margins: { top: 0, bottom: 0, left: 45, right: 45 },
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
      const [q, ...a] = line.slice(2).split(" :: ");
      sections[sections.length - 1].questions.push({ q: q.trim(), a: a.join(" :: ").trim() });
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
      margins: { top: 10, bottom: 10, left: 45, right: 45 },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: NAVY },
      children: [new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(s.title)] })],
    })] }));
    for (const { q, a } of s.questions) {
      rows.push(new TableRow({ cantSplit: true, children: [
        cell(COLS[0], String(++n), { align: AlignmentType.CENTER, color: "7F7F7F" }),
        cell(COLS[1], q, { bold: true }),
        cell(COLS[2], a),
      ] }));
    }
  }

  const doc = new Document({
    title,
    styles: {
      default: { document: { run: { font: FONT, size: SIZE } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 15, bold: true, color: "FFFFFF" },
          paragraph: { spacing: { before: 0, after: 0 }, outlineLevel: 0, keepNext: true } },
      ],
    },
    sections: [{
      properties: { page: { size: { width: PAGE_W, height: 16838 },
        margin: { top: 400, bottom: 400, left: MARGIN, right: MARGIN, footer: 200 } } },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ children: [PageNumber.CURRENT], size: 14, color: "A6A6A6" })] })] }) },
      children: [
        new Paragraph({
          spacing: { after: 30 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 2 } },
          children: [new TextRun({ text: title, bold: true, size: 22, color: NAVY })],
        }),
        new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: COLS, rows }),
      ],
    }],
  });

  fs.writeFileSync(path.join(ROOT, out), await Packer.toBuffer(doc));
  console.log(`Wrote ${out} (${n} questions)`);
}

(async () => { for (const sheet of SHEETS) await build(sheet); })();
