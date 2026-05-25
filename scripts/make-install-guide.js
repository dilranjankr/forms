/* Generates Install_Guide_Kiser_Invoice_Tools.docx
 * Designed for non-IT users. Auto-embeds any PNG screenshots placed in
 *   assets/screenshots/  named  01.png, 02.png, ... 06.png
 * If a screenshot is missing, a clean placeholder box appears instead.
 */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, LevelFormat, BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageBreak, HeadingLevel, ExternalHyperlink, HeightRule, Footer, Header, PageNumber,
} = require("docx");

/* -------- palette -------- */
const GREEN     = "1F5130"; // brand dark green
const GREEN_DK  = "13402A";
const GREEN_SOFT= "E2EFE0";
const YELLOW    = "FFF4C2";
const YELLOW_DK = "9A7A00";
const RED_SOFT  = "FDECEC";
const RED_DK    = "B23A3A";
const INK       = "1C2733";
const INK_SOFT  = "6B7682";
const GRAY      = "F5F6F8";
const LINE      = "D8DCE0";

/* -------- helpers -------- */
function run(text, opts = {}) {
  return new TextRun({
    text,
    bold: opts.bold,
    italics: opts.italics,
    color: opts.color || INK,
    size: opts.size || 22,
    font: "Arial",
  });
}
function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after != null ? opts.after : 120, ...(opts.spacing || {}) },
    alignment: opts.alignment,
    children: [run(text, opts)],
  });
}
function bullet(text, opts = {}) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
    children: [run(text, { size: opts.size || 22, color: opts.color || INK, bold: opts.bold })],
  });
}
function h1(text) {
  return new Paragraph({
    spacing: { before: 240, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: GREEN, space: 4 } },
    children: [run(text, { bold: true, size: 32, color: GREEN })],
  });
}
function h2(text, color) {
  return new Paragraph({
    spacing: { before: 200, after: 120 },
    children: [run(text, { bold: true, size: 26, color: color || GREEN })],
  });
}

/* -------- coloured callout box -------- */
function callout({ fill, border, title, titleColor, bodyChildren }) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            width: { size: 9360, type: WidthType.DXA },
            shading: { fill, type: ShadingType.CLEAR },
            borders: {
              top:    { style: BorderStyle.SINGLE, size: 8, color: border },
              bottom: { style: BorderStyle.SINGLE, size: 8, color: border },
              left:   { style: BorderStyle.SINGLE, size: 24, color: border },
              right:  { style: BorderStyle.SINGLE, size: 8, color: border },
            },
            margins: { top: 240, bottom: 240, left: 320, right: 320 },
            children: [
              new Paragraph({
                spacing: { after: 100 },
                children: [run(title, { bold: true, size: 24, color: titleColor || border })],
              }),
              ...bodyChildren,
            ],
          }),
        ],
      }),
    ],
  });
}

/* -------- screenshot block: image if present, else placeholder -------- */
function screenshotBlock(index, caption) {
  const file = path.join(__dirname, "..", "Screenshots", `${String(index).padStart(2, "0")}.png`);
  const exists = fs.existsSync(file);
  if (exists) {
    // size: width 600px keeps it tidy on US Letter with ~9360 DXA content area
    const data = fs.readFileSync(file);
    // try to read dimensions; if undecidable, force 600x375 (16:10)
    let w = 600, h = 375;
    try {
      const buf = data;
      // PNG header: width @16, height @20 (big-endian uint32)
      if (buf[0] === 0x89 && buf[1] === 0x50) {
        const W = buf.readUInt32BE(16);
        const H = buf.readUInt32BE(20);
        if (W > 0 && H > 0) {
          // shrink so 2-3 steps fit per page
          w = 460;
          h = Math.round((H / W) * 460);
          // cap tall portraits
          if (h > 320) { h = 320; w = Math.round((W / H) * 320); }
        }
      }
    } catch (_) {}
    return [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 40, after: 40 },
        keepNext: true,
        children: [new ImageRun({
          type: "png",
          data,
          transformation: { width: w, height: h },
          altText: { title: caption, description: caption, name: `step-${index}` },
        })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 160 },
        children: [run(caption, { italics: true, size: 18, color: INK_SOFT })],
      }),
    ];
  }
  // placeholder
  return [
    new Table({
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
                top:    { style: BorderStyle.DASHED, size: 6, color: "B0B5BA" },
                bottom: { style: BorderStyle.DASHED, size: 6, color: "B0B5BA" },
                left:   { style: BorderStyle.DASHED, size: 6, color: "B0B5BA" },
                right:  { style: BorderStyle.DASHED, size: 6, color: "B0B5BA" },
              },
              margins: { top: 320, bottom: 320, left: 240, right: 240 },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER, spacing: { after: 80 },
                  children: [run(caption, { italics: true, color: INK_SOFT, size: 22 })],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [run("(screenshot goes here)", { color: "8C95A0", size: 18 })],
                }),
              ],
            }),
          ],
        }),
      ],
    }),
    p(" ", { after: 80 }),
  ];
}

/* -------- step card: green numbered badge + title + body + screenshot -------- */
function stepCard(num, title, body, tip, caption) {
  const badge = new TableCell({
    width: { size: 900, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    shading: { fill: GREEN, type: ShadingType.CLEAR },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: GREEN },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: GREEN },
      left: { style: BorderStyle.SINGLE, size: 4, color: GREEN },
      right: { style: BorderStyle.SINGLE, size: 4, color: GREEN },
    },
    margins: { top: 120, bottom: 120, left: 120, right: 120 },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [run(String(num), { bold: true, size: 52, color: "FFFFFF" })],
      }),
    ],
  });
  const contentChildren = [
    new Paragraph({
      spacing: { after: 100 },
      children: [run(title, { bold: true, size: 28, color: INK })],
    }),
    new Paragraph({
      spacing: { after: 0 },
      children: [run(body, { size: 22, color: INK })],
    }),
  ];
  if (tip) {
    contentChildren.push(
      new Paragraph({
        spacing: { before: 100, after: 0 },
        children: [
          run("Tip: ", { bold: true, color: GREEN_DK, size: 20 }),
          run(tip, { italics: true, size: 20, color: INK_SOFT }),
        ],
      })
    );
  }
  const content = new TableCell({
    width: { size: 8460, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    shading: { fill: GREEN_SOFT, type: ShadingType.CLEAR },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: GREEN_SOFT },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: GREEN_SOFT },
      left: { style: BorderStyle.SINGLE, size: 4, color: GREEN_SOFT },
      right: { style: BorderStyle.SINGLE, size: 4, color: GREEN_SOFT },
    },
    margins: { top: 200, bottom: 200, left: 280, right: 240 },
    children: contentChildren,
  });

  return [
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [900, 8460],
      rows: [new TableRow({ cantSplit: true, children: [badge, content] })],
    }),
    new Paragraph({ spacing: { after: 40 }, children: [run(" ", { size: 8 })] }),
    ...screenshotBlock(num, caption),
  ];
}

/* -------- pieces -------- */
const brandIcon = fs.readFileSync(path.join(__dirname, "..", "assets", "icon-128.png"));

const MANIFEST_URL = "https://dilranjankr.github.io/forms/install.html"; // auto-downloads manifest.prod.xml
const REPO_URL = "https://github.com/dilranjankr/forms";

/* TITLE PAGE — compact */
const titleChildren = [
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 600, after: 180 },
    children: [new ImageRun({
      type: "png", data: brandIcon,
      transformation: { width: 96, height: 96 },
      altText: { title: "K", description: "LDP and LCP form", name: "logo" },
    })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 60 },
    children: [run("LDP and LCP form", { bold: true, size: 52, color: GREEN })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 400 },
    children: [run("Excel Add-in  •  Install Guide  —  about 2 minutes", { italics: true, size: 22, color: INK_SOFT })],
  }),

  /* DOWNLOAD MANIFEST CALLOUT (yellow, highlight) */
  callout({
    fill: YELLOW,
    border: YELLOW_DK,
    title: "Download the manifest file first",
    titleColor: YELLOW_DK,
    bodyChildren: [
      new Paragraph({
        spacing: { after: 100 },
        children: [run("Click the link below and save the file to your computer. You will need it during Step 4.", { size: 22 })],
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [new ExternalHyperlink({
          link: MANIFEST_URL,
          children: [new TextRun({
            text: MANIFEST_URL,
            style: "Hyperlink", font: "Arial", size: 22, bold: true,
          })],
        })],
      }),
      new Paragraph({
        spacing: { after: 0 },
        children: [run("The file (manifest.prod.xml) will download automatically. Save it anywhere — Desktop or Downloads is fine.", { italics: true, size: 20, color: INK_SOFT })],
      }),
    ],
  }),
  p(" ", { after: 160 }),

  /* WHAT YOU NEED CALLOUT (green) */
  callout({
    fill: GREEN_SOFT,
    border: GREEN,
    title: "What you need",
    bodyChildren: [
      bullet("Excel for the web  (any laptop, any browser)"),
      bullet("Your Microsoft 365 work account"),
      bullet("The manifest file from above  +  internet connection"),
    ],
  }),

  new Paragraph({ children: [new PageBreak()] }),
];

/* STEPS */
const stepsChildren = [
  h1("Step-by-step  —  install in 5 quick steps"),
  p("Open your invoice workbook in Excel for the web first, then follow these in order.", { after: 200, italics: true, color: INK_SOFT }),

  ...stepCard(
    1, "Click Add-ins on the Home tab",
    "Open your workbook in Excel for the web. On the Home tab of the ribbon, click the Add-ins button (small puzzle-piece icon).",
    null,
    "Add-ins button"
  ),
  ...stepCard(
    2, "Click Advanced… at the bottom",
    "The Add-ins panel opens. Scroll to the bottom and click the Advanced… link in the bottom-left corner.",
    null,
    "Advanced link"
  ),
  ...stepCard(
    3, "Click Upload My Add-in",
    "The Office Add-ins dialog opens. In the top-right corner, click Upload My Add-in.",
    null,
    "Upload My Add-in"
  ),
  ...stepCard(
    4, "Choose the manifest file → Upload",
    "Click Browse, pick the manifest.prod.xml file you saved earlier, then click Upload.",
    null,
    "Upload Add-in dialog"
  ),
  ...stepCard(
    5, "Click Kiser Invoice Tools",
    "A new Kiser Invoice Tools button (green K) appears on the Home tab. Click it — the panel opens on the right side of Excel.",
    null,
    "Kiser Invoice Tools button"
  ),
  new Paragraph({ children: [new PageBreak()] }),
];

/* DONE + TROUBLESHOOTING + LINKS  (one compact page) */
const closingChildren = [
  h1("You’re done"),
  callout({
    fill: GREEN_SOFT, border: GREEN,
    title: "Inside the panel",
    bodyChildren: [
      bullet("LDP and LCP form  —  add sections, payment requests, generate invoices"),
      bullet("PO form  —  add and move POs"),
      bullet("Help  —  rules and tips"),
    ],
  }),
  p(" ", { after: 100 }),

  h1("Troubleshooting"),
  callout({
    fill: RED_SOFT, border: RED_DK,
    title: "Panel stuck on ‘Loading…’",
    titleColor: RED_DK,
    bodyChildren: [
      bullet("Press Ctrl + Shift + R to hard-refresh."),
      bullet("Close the Excel browser tab and reopen the workbook."),
    ],
  }),
  p(" ", { after: 100 }),
  callout({
    fill: RED_SOFT, border: RED_DK,
    title: "Icons look black / text not styled",
    titleColor: RED_DK,
    bodyChildren: [
      bullet("Office cache is stale. Press Windows + R, paste %LOCALAPPDATA%\\Microsoft\\Office\\16.0\\Wef, delete everything inside, reopen Excel."),
    ],
  }),
  p(" ", { after: 100 }),
  callout({
    fill: RED_SOFT, border: RED_DK,
    title: "‘Upload My Add-in’ button is missing",
    titleColor: RED_DK,
    bodyChildren: [
      bullet("Old desktop Excel (2016 / 2019) — use Excel for the web (office.com) instead."),
    ],
  }),
  p(" ", { after: 100 }),

  callout({
    fill: YELLOW, border: YELLOW_DK,
    title: "Auto-updates  —  nothing to do",
    titleColor: YELLOW_DK,
    bodyChildren: [
      new Paragraph({
        children: [run("Latest version loads automatically every time you open the panel.", { size: 22 })],
      }),
    ],
  }),
];

/* -------- DOCUMENT -------- */
const doc = new Document({
  creator: "Ed Ball Landscape Architecture",
  title: "Kiser Invoice Tools — Install Guide",
  description: "Install guide for the Kiser Invoice Tools Excel add-in.",
  styles: {
    default: { document: { run: { font: "Arial", size: 22, color: INK } } },
    paragraphStyles: [
      { id: "Hyperlink", name: "Hyperlink", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { color: "0563C1", underline: { type: "single", color: "0563C1" } } },
    ],
  },
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 540, hanging: 280 } } },
      }],
    }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 }, // US Letter
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
      },
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: GREEN, space: 4 } },
            children: [run("LDP and LCP form", { bold: true, color: GREEN, size: 18 })],
          }),
        ],
      }),
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              run("Page ", { color: INK_SOFT, size: 16 }),
              new TextRun({ children: [PageNumber.CURRENT], color: INK_SOFT, size: 16, font: "Arial" }),
              run(" / ", { color: INK_SOFT, size: 16 }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], color: INK_SOFT, size: 16, font: "Arial" }),
            ],
          }),
        ],
      }),
    },
    children: [
      ...titleChildren,
      ...stepsChildren,
      ...closingChildren,
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  const out = path.join(__dirname, "..", "Install_Guide_Kiser_Invoice_Tools.docx");
  fs.writeFileSync(out, buffer);
  console.log("OK:", out);
});
