/* Generates Install_Guide_Kiser_Invoice_Tools.docx */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, LevelFormat, BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageBreak, HeadingLevel, ExternalHyperlink, HeightRule,
} = require("docx");

const GREEN = "1F5130";
const GREEN_DK = "186E3E";
const GREEN_SOFT = "E2EFE0";
const INK = "1C2733";
const INK_SOFT = "6B7682";
const GRAY = "F5F6F8";
const LINE = "D8DCE0";

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120, ...(opts.spacing || {}) },
    alignment: opts.alignment,
    children: [new TextRun({
      text, bold: opts.bold, color: opts.color || INK,
      size: opts.size || 22, italics: opts.italics, font: "Arial",
    })],
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: INK })],
  });
}

function h1(text) {
  return new Paragraph({
    spacing: { before: 200, after: 160 },
    children: [new TextRun({ text, bold: true, size: 32, color: GREEN, font: "Arial" })],
  });
}

function screenshotPlaceholder(caption) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [
      new TableRow({
        height: { value: 2600, rule: HeightRule.ATLEAST },
        children: [
          new TableCell({
            width: { size: 9360, type: WidthType.DXA },
            verticalAlign: VerticalAlign.CENTER,
            shading: { fill: GRAY, type: ShadingType.CLEAR },
            borders: {
              top: { style: BorderStyle.DASHED, size: 6, color: "B0B5BA" },
              bottom: { style: BorderStyle.DASHED, size: 6, color: "B0B5BA" },
              left: { style: BorderStyle.DASHED, size: 6, color: "B0B5BA" },
              right: { style: BorderStyle.DASHED, size: 6, color: "B0B5BA" },
            },
            margins: { top: 320, bottom: 320, left: 240, right: 240 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER, spacing: { after: 80 },
                children: [new TextRun({ text: caption, italics: true, color: INK_SOFT, size: 22, font: "Arial" })],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: "(paste your screenshot here)", color: "8C95A0", size: 18, font: "Arial" })],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function step(num, title, body, caption) {
  return [
    new Paragraph({
      spacing: { before: 280, after: 80 },
      children: [
        new TextRun({ text: `Step ${num}`, bold: true, color: GREEN, size: 22, font: "Arial" }),
        new TextRun({ text: "   " + title, bold: true, color: INK, size: 26, font: "Arial" }),
      ],
    }),
    p(body, { spacing: { after: 180 } }),
    screenshotPlaceholder(caption),
    p(" ", { spacing: { after: 80 } }),
  ];
}

const brandIcon = fs.readFileSync(path.join(__dirname, "..", "assets", "icon-128.png"));

const titleChildren = [
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 1800, after: 280 },
    children: [new ImageRun({
      type: "png", data: brandIcon,
      transformation: { width: 110, height: 110 },
      altText: { title: "K", description: "Kiser Invoice Tools", name: "logo" },
    })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 80 },
    children: [new TextRun({ text: "Kiser Invoice Tools", bold: true, size: 56, color: GREEN, font: "Arial" })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 240 },
    children: [new TextRun({ text: "Excel Add-in  —  Install Guide", size: 30, color: INK, font: "Arial" })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 240, after: 80 },
    children: [new TextRun({ text: "Ed Ball Landscape Architecture", color: INK_SOFT, font: "Arial", size: 22 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 1200 },
    children: [new TextRun({ text: "Version 1.0", color: INK_SOFT, font: "Arial", size: 20 })],
  }),
  new Table({
    width: { size: 8640, type: WidthType.DXA },
    columnWidths: [8640],
    alignment: AlignmentType.CENTER,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 8640, type: WidthType.DXA },
            shading: { fill: GREEN_SOFT, type: ShadingType.CLEAR },
            borders: {
              top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            },
            margins: { top: 260, bottom: 260, left: 320, right: 320 },
            children: [
              new Paragraph({
                spacing: { after: 140 },
                children: [new TextRun({ text: "What you need", bold: true, color: GREEN, size: 26, font: "Arial" })],
              }),
              bullet("Excel for Microsoft 365 (Excel on the web works best on Office 2019/older desktop)"),
              bullet("The file manifest.prod.xml (sent along with this guide)"),
              bullet("Internet connection — about 2 minutes per laptop"),
            ],
          }),
        ],
      }),
    ],
  }),
  new Paragraph({ children: [new PageBreak()] }),
];

const stepsChildren = [
  h1("Step-by-step installation"),
  ...step(
    1, "Open Excel for the web",
    "Go to office.com and sign in with your Microsoft 365 account. Open any Excel workbook (your invoice workbook is fine, or create a blank one).",
    "Screenshot 1 — Excel on the web open with a workbook"
  ),
  ...step(
    2, "Open the Add-ins menu",
    "On the Home tab, click the Add-ins dropdown on the ribbon. At the bottom of the menu, click Advanced…",
    "Screenshot 2 — Add-ins dropdown with 'Advanced…' highlighted"
  ),
  ...step(
    3, "Click Upload My Add-in",
    "In the Office Add-ins dialog, look at the top right corner and click Upload My Add-in.",
    "Screenshot 3 — Office Add-ins dialog with 'Upload My Add-in' highlighted"
  ),
  new Paragraph({ children: [new PageBreak()] }),
  ...step(
    4, "Choose the manifest file",
    "Click Browse, select the manifest.prod.xml file your team sent you, then click Upload.",
    "Screenshot 4 — Upload Add-in dialog with Browse and Upload highlighted"
  ),
  ...step(
    5, "Open the panel",
    "The Kiser Invoice Tools button appears on the Home tab (green K icon). Click it to open the side panel.",
    "Screenshot 5 — Home tab with the 'Kiser Invoice Tools' button visible"
  ),
  ...step(
    6, "Start working",
    "Use the Invoice Form tab to add LDP / LCP sections and Payment Requests, the PO Form tab to manage POs, and the Help tab to see what each button does and which parts of the workbook must not be changed.",
    "Screenshot 6 — Kiser Invoice Tools panel open with the Invoice Form tab"
  ),
  new Paragraph({ children: [new PageBreak()] }),
];

const troubleshootingChildren = [
  h1("Troubleshooting"),
  p("If something doesn't look right, try these in order:", { spacing: { after: 140 } }),
  bullet("Hard refresh the browser: Ctrl + Shift + R."),
  bullet("Remove the add-in from MY ADD-INS and upload manifest.prod.xml again."),
  bullet("Close the browser tab completely, reopen Excel for the web, open the workbook again."),
  bullet("On Office 2019/older desktop the Upload option doesn't exist — use Excel for the web instead."),
  bullet("If your IT team deployed the add-in centrally (Microsoft 365 Admin Center), it appears under ADMIN MANAGED automatically and no upload is needed."),

  h1("Automatic updates"),
  p("You don't need to reinstall when the add-in is updated. The code lives on a hosted page; every time you open the panel, you get the latest version automatically. The manifest only changes if the hosting URL itself moves."),

  h1("Support & source"),
  p("If you hit any issue, send a screenshot to your IT or to the add-in maintainer."),
  new Paragraph({
    alignment: AlignmentType.LEFT, spacing: { before: 100 },
    children: [new ExternalHyperlink({
      link: "https://github.com/dilranjankr/forms",
      children: [new TextRun({ text: "https://github.com/dilranjankr/forms", style: "Hyperlink", font: "Arial", size: 22 })],
    })],
  }),
];

const doc = new Document({
  styles: { default: { document: { run: { font: "Arial", size: 22, color: INK } } } },
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } },
      }],
    }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
      },
    },
    children: [...titleChildren, ...stepsChildren, ...troubleshootingChildren],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  const out = path.join(__dirname, "..", "Install_Guide_Kiser_Invoice_Tools.docx");
  fs.writeFileSync(out, buffer);
  console.log("OK:", out);
});
