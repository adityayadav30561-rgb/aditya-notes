// Builds SSB_Question_Sheet.docx from questions.md.
// questions.md: "## Section" lines and "- question" lines. Questions are numbered
// continuously, each followed by blank space for the answer.
// Usage: node build/build.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, Footer,
  PageNumber, LevelFormat, BorderStyle,
} = require("docx");

const ROOT = path.join(__dirname, "..");
const SRC = path.join(ROOT, "questions.md");
const OUT = path.join(ROOT, "SSB_Question_Sheet.docx");
const FONT = "Calibri";
const NAVY = "1F3864";

const children = [
  new Paragraph({
    spacing: { after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 4 } },
    children: [new TextRun({ text: "SSB Interview – Question Sheet", bold: true, size: 32, color: NAVY })],
  }),
];

let count = 0;
for (const line of fs.readFileSync(SRC, "utf8").split(/\r?\n/)) {
  if (line.startsWith("## ")) {
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(line.slice(3).trim())] }));
  } else if (line.startsWith("- ")) {
    count++;
    children.push(new Paragraph({
      numbering: { reference: "questions", level: 0 },
      keepNext: true,
      spacing: { before: 120, after: 0 },
      children: [new TextRun({ text: line.slice(2).trim(), bold: true })],
    }));
    children.push(new Paragraph({ spacing: { after: 120 }, indent: { left: 440 }, children: [] }));
  }
}

const doc = new Document({
  title: "SSB Interview – Question Sheet",
  styles: {
    default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, color: NAVY },
        paragraph: { spacing: { before: 280, after: 40 }, outlineLevel: 0, keepNext: true } },
    ],
  },
  numbering: { config: [{ reference: "questions", levels: [
    { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 440, hanging: 440 } } } },
  ] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 },
      margin: { top: 1000, bottom: 1000, left: 1000, right: 1000 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "A6A6A6" })] })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log(`Wrote ${path.relative(ROOT, OUT)} (${count} questions)`);
});
