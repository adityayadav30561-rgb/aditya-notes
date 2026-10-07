// Builds the two-page, two-column master sheets listed in GUIDES (SRT, TAT)
// from their Markdown sources.
// Source format: "# Title", "> tagline", "## Section", "- bullet" (with **bold**),
// "| Group |" for a shaded table group row and "| Situation | Response" for a row.
// Usage: node build/guides.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, Table, TableRow,
  TableCell, WidthType, ShadingType, LevelFormat, Footer, PageNumber,
} = require("docx");

const ROOT = path.join(__dirname, "..");
const GUIDES = [
  { src: "srt_guide.md", out: "SSB_SRT_Guide.docx" },
  { src: "tat_guide.md", out: "SSB_TAT_Guide.docx" },
];

const FONT = "Calibri";
const NAVY = "1F3864";
const SIZE = 15; // 7.5 pt

const PAGE_W = 11906, MARGIN = 380, GAP = 220;
const COL_W = Math.floor((PAGE_W - 2 * MARGIN - GAP) / 2);
const TCOLS = [1500, COL_W - 1500]; // Situation | Response

// "**bold** plain" → runs
const runs = (text, opts = {}) => text.split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((t) =>
  t.startsWith("**") ? new TextRun({ text: t.slice(2, -2), bold: true, size: SIZE, ...opts })
                     : new TextRun({ text: t, size: SIZE, ...opts }));

const line = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: line, bottom: line, left: line, right: line };
const cell = (w, children, fill, span) => new TableCell({
  width: { size: w, type: WidthType.DXA },
  columnSpan: span,
  borders,
  margins: { top: 0, bottom: 0, left: 40, right: 40 },
  shading: fill ? { type: ShadingType.CLEAR, color: "auto", fill } : undefined,
  children: [new Paragraph({ spacing: { before: 0, after: 0 }, children })],
});

function build({ src, out }) {
  let title = "";
  const children = [];
  let rows = null;
  const flush = () => {
    if (!rows) return;
    children.push(new Table({ width: { size: COL_W, type: WidthType.DXA }, columnWidths: TCOLS, rows }));
    rows = null;
  };

  for (const raw of fs.readFileSync(path.join(ROOT, src), "utf8").split(/\r?\n/)) {
    const l = raw.trimEnd();
    if (l.startsWith("| ")) {
      rows = rows || [];
      const parts = l.slice(2).split(" | ").map((s) => s.replace(/ \|$/, "").trim());
      if (l.endsWith(" |") && parts.length === 1) {
        rows.push(new TableRow({ cantSplit: true, children: [
          cell(COL_W, runs(parts[0], { bold: true, color: NAVY }), "DCE3F0", 2)] }));
      } else {
        rows.push(new TableRow({ cantSplit: true, children: [
          cell(TCOLS[0], runs(parts[0], { bold: true })), cell(TCOLS[1], runs(parts[1] || ""))] }));
      }
      continue;
    }
    flush();
    if (l.startsWith("# ")) {
      title = l.slice(2);
      children.push(new Paragraph({
        spacing: { after: 20 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 1 } },
        children: [new TextRun({ text: l.slice(2), bold: true, size: 24, color: NAVY })],
      }));
    } else if (l.startsWith("> ")) {
      children.push(new Paragraph({ spacing: { after: 30 }, children: runs(l.slice(2), { italics: true }) }));
    } else if (l.startsWith("## ")) {
      children.push(new Paragraph({
        spacing: { before: 40, after: 15 },
        keepNext: true,
        shading: { type: ShadingType.CLEAR, color: "auto", fill: NAVY },
        children: [new TextRun({ text: l.slice(3), bold: true, size: 17, color: "FFFFFF" })],
      }));
    } else if (l.startsWith("- ")) {
      children.push(new Paragraph({
        numbering: { reference: "dot", level: 0 },
        spacing: { before: 0, after: 8 },
        children: runs(l.slice(2)),
      }));
    }
  }
  flush();

  const doc = new Document({
    title,
    styles: { default: { document: { run: { font: FONT, size: SIZE } } } },
    numbering: { config: [{ reference: "dot", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 140, hanging: 110 } } } }] }] },
    sections: [{
      properties: {
        page: { size: { width: PAGE_W, height: 16838 },
          margin: { top: 380, bottom: 380, left: MARGIN, right: MARGIN, footer: 180 } },
        column: { count: 2, space: GAP, separate: true },
      },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ children: [PageNumber.CURRENT], size: 12, color: "A6A6A6" })] })] }) },
      children,
    }],
  });

  return Packer.toBuffer(doc).then((buf) => {
    fs.writeFileSync(path.join(ROOT, out), buf);
    console.log(`Wrote ${out}`);
  });
}

(async () => { for (const g of GUIDES) await build(g); })();
