const fs = require("fs");
const path = require("path");
const { JSDOM } = require("jsdom");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  LevelFormat,
  BorderStyle,
  WidthType,
  Table,
  TableRow,
  TableCell,
  ShadingType,
} = require("docx");

const root = __dirname;
const htmlPath = path.join(root, "creative-proposal.html");
const outputPath = path.join(root, "creative-proposal.docx");

const html = fs.readFileSync(htmlPath, "utf8");
const dom = new JSDOM(html);
const documentNode = dom.window.document;

function textOf(node) {
  return node ? node.textContent.replace(/\s+/g, " ").trim() : "";
}

function addParagraph(children, text, options = {}) {
  const clean = String(text || "").replace(/\s+/g, " ").trim();
  if (!clean) return;

  const run = new TextRun({
    text: clean,
    bold: options.bold,
    italics: options.italics,
    color: options.color,
    size: options.size,
    font: "Microsoft YaHei",
  });

  children.push(
    new Paragraph({
      heading: options.heading,
      alignment: options.alignment,
      spacing: options.spacing || { after: 160, line: 320 },
      border: options.border,
      shading: options.shading,
      children: [run],
    })
  );
}

function addBullet(children, text) {
  const clean = String(text || "").replace(/\s+/g, " ").trim();
  if (!clean) return;
  children.push(
    new Paragraph({
      numbering: { reference: "bullets", level: 0 },
      spacing: { after: 100, line: 300 },
      children: [
        new TextRun({
          text: clean,
          font: "Microsoft YaHei",
        }),
      ],
    })
  );
}

function addSectionTitle(children, title, intro) {
  addParagraph(children, title, {
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 260, after: 180, line: 360 },
  });
  if (intro) {
    addParagraph(children, intro, {
      color: "475569",
      spacing: { after: 180, line: 320 },
    });
  }
}

function addCard(children, title, paragraphs = [], bullets = []) {
  if (title) {
    addParagraph(children, title, {
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 180, after: 120, line: 320 },
    });
  }
  paragraphs.forEach((item) => addParagraph(children, item));
  bullets.forEach((item) => addBullet(children, item));
}

function getDirectChildrenText(container, selector) {
  return Array.from(container.querySelectorAll(selector)).map((item) => textOf(item)).filter(Boolean);
}

const children = [];

const heroTitle = textOf(documentNode.querySelector(".hero-copy h1"));
const heroLead = textOf(documentNode.querySelector(".hero-copy .lead"));
const heroQuote = textOf(documentNode.querySelector(".hero-quote"));
const heroTags = getDirectChildrenText(documentNode.querySelector(".hero-tags") || documentNode, ".tag");
const miniItems = Array.from(documentNode.querySelectorAll(".stack .mini")).map((item) => ({
  label: textOf(item.querySelector("small")),
  value: textOf(item.querySelector("strong")),
}));

addParagraph(children, heroTitle, {
  heading: HeadingLevel.TITLE,
  alignment: AlignmentType.CENTER,
  spacing: { after: 180, line: 360 },
});

const subtitle = miniItems.find((item) => item.label.includes("副标题"))?.value || "";
if (subtitle) {
  addParagraph(children, subtitle, {
    alignment: AlignmentType.CENTER,
    color: "7C3AED",
    bold: true,
    size: 28,
    spacing: { after: 180, line: 320 },
  });
}

addParagraph(children, heroLead, {
  alignment: AlignmentType.CENTER,
  spacing: { after: 180, line: 360 },
});

addParagraph(children, heroQuote, {
  alignment: AlignmentType.CENTER,
  italics: true,
  color: "0F766E",
  spacing: { after: 180, line: 320 },
  border: {
    bottom: { style: BorderStyle.SINGLE, size: 6, color: "D8B4FE", space: 1 },
  },
});

if (heroTags.length) {
  addParagraph(children, `关键词：${heroTags.join(" / ")}`, {
    alignment: AlignmentType.CENTER,
    color: "64748B",
  });
}

if (miniItems.length) {
  const tableWidth = 9360;
  const colWidth = 4680;
  const border = { style: BorderStyle.SINGLE, size: 1, color: "E2E8F0" };
  const rows = [];
  for (let i = 0; i < miniItems.length; i += 2) {
    const pair = miniItems.slice(i, i + 2);
    if (pair.length === 1) {
      pair.push({ label: "", value: "" });
    }
    rows.push(
      new TableRow({
        children: pair.map((item) =>
          new TableCell({
            width: { size: colWidth, type: WidthType.DXA },
            borders: { top: border, bottom: border, left: border, right: border },
            shading: { fill: "FAF5FF", type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 140, right: 140 },
            children: [
              new Paragraph({
                spacing: { after: 60 },
                children: [new TextRun({ text: item.label, bold: true, color: "6D28D9", font: "Microsoft YaHei" })],
              }),
              new Paragraph({
                children: [new TextRun({ text: item.value, font: "Microsoft YaHei" })],
              }),
            ],
          })
        ),
      })
    );
  }

  children.push(
    new Table({
      width: { size: tableWidth, type: WidthType.DXA },
      columnWidths: [colWidth, colWidth],
      rows,
    })
  );
}

const sectionCards = Array.from(documentNode.querySelectorAll("main .section-card"));

sectionCards.forEach((card, index) => {
  const title = textOf(card.querySelector(".section-head h2"));
  const intro = textOf(card.querySelector(".section-head p"));
  addSectionTitle(children, title, intro);

  if (index === 0) {
    const conceptTitle = textOf(card.querySelector(".concept-panel strong"));
    const conceptDesc = textOf(card.querySelector(".concept-panel p"));
    const pills = getDirectChildrenText(card.querySelector(".pill-row") || card, ".pill");
    addCard(children, conceptTitle, [conceptDesc], pills);
    return;
  }

  if (index === 1) {
    Array.from(card.querySelectorAll(".grid-3 .card")).forEach((item) => {
      addCard(children, textOf(item.querySelector("h3")), [textOf(item.querySelector("p"))]);
    });
    return;
  }

  if (index === 2) {
    Array.from(card.querySelectorAll(".grid-2 .card")).forEach((item) => {
      addCard(
        children,
        textOf(item.querySelector("h3")),
        [],
        getDirectChildrenText(item, "li")
      );
    });
    return;
  }

  if (index === 3) {
    const mvpBox = card.querySelector(".mvp-box");
    if (mvpBox) {
      addCard(
        children,
        textOf(mvpBox.querySelector("strong")),
        [textOf(mvpBox.querySelector("p"))],
        getDirectChildrenText(mvpBox, "li")
      );
    }
    const noDoCard = card.querySelector(".grid-2 .card");
    if (noDoCard) {
      addCard(children, textOf(noDoCard.querySelector("h3")), [], getDirectChildrenText(noDoCard, "li"));
    }
    Array.from(card.querySelectorAll(".grid-3 .card")).forEach((item) => {
      addCard(children, textOf(item.querySelector("h3")), [textOf(item.querySelector("p"))]);
    });
    return;
  }

  if (index === 4) {
    Array.from(card.querySelectorAll(".grid-3 .card")).forEach((item) => {
      addCard(children, textOf(item.querySelector("h3")), [textOf(item.querySelector("p"))]);
    });
    return;
  }

  if (index === 5) {
    Array.from(card.querySelectorAll(".flow-item")).forEach((item) => {
      const step = textOf(item.querySelector(".step"));
      const name = textOf(item.querySelector("h3"));
      const desc = textOf(item.querySelector("p"));
      addCard(children, `${step} ${name}`.trim(), [desc]);
    });
    return;
  }

  if (index === 6) {
    Array.from(card.querySelectorAll(".compare-card")).forEach((item) => {
      addCard(children, textOf(item.querySelector("h3")), [], getDirectChildrenText(item, "li"));
    });
    return;
  }

  if (index === 7) {
    Array.from(card.querySelectorAll(".grid-3 .card")).forEach((item) => {
      addCard(children, textOf(item.querySelector("h3")), [textOf(item.querySelector("p"))]);
    });
    return;
  }

  if (index === 8) {
    Array.from(card.querySelectorAll(".stat")).forEach((item) => {
      const label = textOf(item.querySelector(".label"));
      const name = textOf(item.querySelector("h3"));
      const desc = textOf(item.querySelector("p"));
      addCard(children, `${name}（${label}）`, [desc]);
    });
  }
});

const finale = documentNode.querySelector(".finale");
if (finale) {
  addSectionTitle(children, "结语", "");
  addParagraph(children, textOf(finale.querySelector("strong")), {
    bold: true,
    color: "0F172A",
    spacing: { after: 120, line: 320 },
  });
  addParagraph(children, textOf(finale.querySelector("p")));
}

const doc = new Document({
  creator: "TRAE",
  title: "栖序 创意文档",
  description: "由 HTML 创意文档导出的 Word 版本",
  styles: {
    default: {
      document: {
        run: {
          font: "Microsoft YaHei",
          size: 22,
          color: "111827",
        },
      },
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { font: "Microsoft YaHei", size: 30, bold: true, color: "0F172A" },
        paragraph: { spacing: { before: 260, after: 160 }, outlineLevel: 0 },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { font: "Microsoft YaHei", size: 26, bold: true, color: "1D4ED8" },
        paragraph: { spacing: { before: 180, after: 120 }, outlineLevel: 1 },
      },
      {
        id: "Title",
        name: "Title",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { font: "Microsoft YaHei", size: 42, bold: true, color: "4C1D95" },
        paragraph: { spacing: { before: 120, after: 180 }, outlineLevel: 0 },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
            alignment: AlignmentType.LEFT,
            style: {
              paragraph: {
                indent: { left: 720, hanging: 360 },
                spacing: { after: 100, line: 300 },
              },
            },
          },
        ],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: {
            width: 11906,
            height: 16838,
          },
          margin: { top: 1440, right: 1260, bottom: 1440, left: 1260 },
        },
      },
      children,
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(outputPath, buffer);
  console.log(`Created: ${outputPath}`);
});
